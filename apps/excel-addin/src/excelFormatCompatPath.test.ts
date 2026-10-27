import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { captureWorkbookSourceFingerprint, currentExcelFormatCapability, executePlan, undoExecution } from "./excel";
import type { AnalysisPlan, ExcelAction } from "./contracts";
import { createOfficeTestHost, installOfficeTestHost, type TestOfficeHost } from "./officeTestHost";

/**
 * **直接格式兼容路径**在 ExcelApi ≤ 1.17 宿主上的端到端验收。
 *
 * 背景（Excel 实机故障，2026-10-13）：`ProPlus2024Volume` / `16.0.17932` 这类
 * 批量许可的永久版 Office，按微软要求集表**最高只到 ExcelApi 1.17**，
 * 1.18/1.19/1.20 一律不支持。旧实现强制依赖 1.19 的 `getDisplayedCellProperties`，
 * 于是 setFont 在写前被拒 —— 格式动作在这台宿主机上完全不可用。
 *
 * 修复方向不是"请升级 Office"，而是补一条 1.9–1.17 的兼容路径：
 *  - 读：`getCellProperties` 只返回**直接设置在单元格上**的属性，缺席 = 继承/默认；
 *  - 恢复：先 `clear(Formats)` 清掉区域**全部直接格式**，再 `setCellProperties`
 *    写回捕获到的**直接格式集合** ⇒ 直接设置过的属性逐格还原，未直接设置的属性回到继承态。
 *
 * 本文件用**真实产品代码**（`executePlan` / `undoExecution`）跑在有真实状态的替身宿主上，
 * 证明用户点名的四类场景（**加粗、水平居中、混合字号、行高**）在兼容路径上闭环，
 * 并证明不可靠恢复的场景（条件格式 / 合并区域）被**明确拒绝且原表零写入**。
 *
 * 证据边界：替身宿主自动化证据，**非**真实 Excel 实机证据。
 */

const SHEET = "数据";

function plan(actions: ExcelAction[], overrides: Partial<AnalysisPlan> = {}): AnalysisPlan {
  return {
    id: "plan-compat-format",
    title: "格式",
    summary: "格式",
    assumptions: [],
    warnings: [],
    actions,
    ...overrides
  };
}


/**
 * 计划带上**生产形态**的来源锚点。
 *
 * 写闸门要求：写了固定值的计划必须记录"它是按哪一版源算出来的"（缺失即拒绝）。
 * 这里用产品自己的读取器在"计划生成那一刻"读一次 —— 与 App 给桥绑定记的那枚是
 * 同一个口径。审阅/撤销的多数用例不需要它，但凡动手 **写值** 的用例都必须有。
 */
async function anchoredPlan(
  actions: ExcelAction[],
  sheets: string[] = [SHEET]
): Promise<AnalysisPlan> {
  return {
    ...plan(actions),
    sourceFingerprintSheets: [...sheets],
    sourceFingerprint: await captureWorkbookSourceFingerprint(sheets)
  };
}

/** 只声明到 ExcelApi 1.17：批量许可永久版的实际上限 ⇒ 不支持 1.19，支持 1.9。 */
const LEGACY_API_VERSIONS = Array.from({ length: 17 }, (_, index) => `1.${index + 1}`);

function legacyHost(overrides: Partial<Parameters<typeof createOfficeTestHost>[0]> = {}): TestOfficeHost {
  return createOfficeTestHost({
    url: "https://contoso.sharepoint.com/format.xlsx",
    activeSheet: SHEET,
    selectedRange: "A1:B2",
    excelApiVersions: LEGACY_API_VERSIONS,
    sheets: [
      {
        name: SHEET,
        values: [
          ["甲", "乙"],
          ["丙", "丁"]
        ]
      }
    ],
    ...overrides
  });
}

let host: TestOfficeHost;

beforeEach(() => {
  host = legacyHost();
  installOfficeTestHost(host);
});

afterEach(() => {
  // 宿主替身只装到全局，无其它副作用需清理。
});

/** 用真实 `Excel.run` 读回某区域的数字格式（替身没有直接暴露 numberFormat 读口）。 */
async function numberFormatAt(address: string): Promise<string> {
  return (await Excel.run(async (context) => {
    const range = context.workbook.worksheets.getItem(SHEET).getRange(address);
    range.load("numberFormat");
    await context.sync();
    return range.numberFormat[0][0];
  })) as string;
}

describe("宿主能力判定：1.17 走兼容路径，不是一刀切拒绝", () => {
  it("ExcelApi 1.19=false、1.9=true ⇒ path=direct，并在 note 里说清缺什么、怎么恢复", () => {
    const capability = currentExcelFormatCapability();
    expect(capability.path).toBe("direct");
    expect(capability.missingApis.join("")).toContain("ExcelApi 1.19");
    expect(capability.blockedActionTypes).toEqual([]);
    expect(capability.note).toContain("直接格式");
    expect(capability.note).toContain("条件格式");
    expect(capability.note).toContain("合并单元格");
  });
});

describe("兼容路径闭环 · 用户点名的四类格式动作", () => {
  it("加粗：setFont{bold} 执行后逐格直接加粗；撤销后回到「未直接设置加粗」而不是显式 false", async () => {
    // A1 执行前**未直接设置**加粗（继承默认）。B1 直接设置过倾斜，用来钉"只改被点名属性"。
    host.setCellStyle(SHEET, "B1", { fontItalic: true });

    const result = await executePlan(
      plan([{ type: "setFont", sheet: SHEET, range: "A1:B1", bold: true }]),
      { authorizedSheetNames: [SHEET] }
    );

    // 执行后逐格直接加粗（动作确实生效）。
    expect(host.cellStyle(SHEET, "A1")?.fontBold).toBe(true);
    expect(host.cellStyle(SHEET, "B1")?.fontBold).toBe(true);

    const saved = result.undoSnapshot!.ranges[0].format!;
    // 走的是兼容路径：有直接格式矩阵，没有显示值矩阵。
    expect(saved.compatPath).toBe(true);
    expect(saved.cells).toBeNull();
    expect(saved.directCells).not.toBeNull();
    // A1 执行前未直接设置加粗 ⇒ 记录为继承属性（这是"未采集 ≠ 默认值"的来源记录）。
    expect(saved.inheritedFields).toContain("fontBold");

    await undoExecution(result.undoSnapshot!);

    // A1 回到"未直接设置加粗"：直接格式里**没有** fontBold，而不是被写成 bold:false。
    const a1 = host.cellStyle(SHEET, "A1");
    expect(a1?.fontBold).toBeUndefined();
    expect(a1?.fontBold).not.toBe(false);
    // B1 的直接倾斜没有被这次加粗动作牵动（动作只改 bold），恢复后仍在。
    expect(host.cellStyle(SHEET, "B1")?.fontItalic).toBe(true);
    expect(host.cellStyle(SHEET, "B1")?.fontBold).toBeUndefined();
  });

  it("水平居中：setAlignment{horizontal:center} 执行后逐格居中；撤销后各自回到原对齐（含未直接设置）", async () => {
    // A1 未直接设置对齐；B1 直接设置为 Left。
    host.setCellStyle(SHEET, "B1", { horizontalAlignment: "Left" });

    const result = await executePlan(
      plan([{ type: "setAlignment", sheet: SHEET, range: "A1:B1", horizontal: "center" }]),
      { authorizedSheetNames: [SHEET] }
    );

    // 执行后逐格居中（枚举忽略大小写：动作给的是 "center"，宿主存 "Center"）。
    expect(host.cellStyle(SHEET, "A1")?.horizontalAlignment).toBe("Center");
    expect(host.cellStyle(SHEET, "B1")?.horizontalAlignment).toBe("Center");
    expect(result.undoSnapshot!.ranges[0].format!.inheritedFields).toContain(
      "horizontalAlignment"
    );

    await undoExecution(result.undoSnapshot!);

    // A1 回到"未直接设置"（继承 General），不是被写成 "General" 这个显式值。
    expect(host.cellStyle(SHEET, "A1")?.horizontalAlignment).toBeUndefined();
    // B1 精确回到执行前的 Left。
    expect(host.cellStyle(SHEET, "B1")?.horizontalAlignment).toBe("Left");
  });

  it("混合字号：各格字号原本不同 → setFont 统一 → 撤销后逐格还原原始混合字号", async () => {
    // A1=20、B1=9 直接设置；A2 未直接设置（继承）。
    host.setCellStyle(SHEET, "A1", { fontSize: 20 });
    host.setCellStyle(SHEET, "B1", { fontSize: 9 });

    const result = await executePlan(
      plan([{ type: "setFont", sheet: SHEET, range: "A1:B2", size: 12 }]),
      { authorizedSheetNames: [SHEET] }
    );

    // 执行后四格统一 12（旧的"区域级单值"实现会在这里就把混合态拍平）。
    expect(host.cellStyle(SHEET, "A1")?.fontSize).toBe(12);
    expect(host.cellStyle(SHEET, "B1")?.fontSize).toBe(12);
    expect(host.cellStyle(SHEET, "A2")?.fontSize).toBe(12);

    const saved = result.undoSnapshot!.ranges[0].format!;
    expect(saved.directCells).not.toBeNull();

    await undoExecution(result.undoSnapshot!);

    // 逐格还原：A1 回 20、B1 回 9、A2 回到"未直接设置"（继承）。
    expect(host.cellStyle(SHEET, "A1")?.fontSize).toBe(20);
    expect(host.cellStyle(SHEET, "B1")?.fontSize).toBe(9);
    expect(host.cellStyle(SHEET, "A2")?.fontSize).toBeUndefined();
  });

  it("行高：逐行直接设置 → resizeRange 统一 → 撤销后逐行还原；纯 resizeRange 不因区域含条件格式被误拒", async () => {
    // 目标区域**含条件格式**：但 resizeRange 只改行级属性，不该被"逐格不可恢复"拦住。
    const withConditional = legacyHost({ conditionalFormatCount: 3 });
    installOfficeTestHost(withConditional);
    withConditional.setRowHeight(SHEET, 0, 20);
    withConditional.setRowHeight(SHEET, 1, 30);

    const result = await executePlan(
      plan([{ type: "resizeRange", sheet: SHEET, range: "A1:B2", rowHeight: 40 }]),
      { authorizedSheetNames: [SHEET] }
    );

    // 没被拒（这正是 needsCellCapture 只在真要改逐格属性时才查条件格式的意义）。
    expect(withConditional.rowHeight(SHEET, 0)).toBe(40);
    expect(withConditional.rowHeight(SHEET, 1)).toBe(40);
    // 只给了 rowHeight ⇒ 列宽一个都没动。
    expect(withConditional.columnWidth(SHEET, 0)).toBe(8.43);
    expect(withConditional.columnWidth(SHEET, 1)).toBe(8.43);

    await undoExecution(result.undoSnapshot!);

    // 逐行还原，而不是被拍平成同一条行高。
    expect(withConditional.rowHeight(SHEET, 0)).toBe(20);
    expect(withConditional.rowHeight(SHEET, 1)).toBe(30);
    // 撤销也不该顺手改列宽。
    expect(withConditional.columnWidth(SHEET, 0)).toBe(8.43);
  });

  it("行高是行级属性：区域内任一格所在整行共享同一条行高", async () => {
    // 第 1 行（A1/B1 同属第 1 行）执行前 22。resizeRange 只覆盖 A1:B2。
    host.setRowHeight(SHEET, 0, 22);
    host.setRowHeight(SHEET, 1, 22);

    await executePlan(
      plan([{ type: "resizeRange", sheet: SHEET, range: "A1:B2", rowHeight: 50 }]),
      { authorizedSheetNames: [SHEET] }
    );

    // 行高落在行上：读第 1 行只有一个值，A1 与 B1 共享。
    expect(host.rowHeight(SHEET, 0)).toBe(50);
    expect(host.rowHeight(SHEET, 1)).toBe(50);
  });
});

describe("兼容路径 · 无法可靠恢复的区域在执行前明确拒绝（原表零写入）", () => {
  it("目标区域含条件格式：setFill 被拒，点名条件格式与缺失的 ExcelApi 1.19，且零写入", async () => {
    const withConditional = legacyHost({ conditionalFormatCount: 2 });
    installOfficeTestHost(withConditional);

    const failure = await executePlan(
      plan([{ type: "setFill", sheet: SHEET, range: "A1:B2", color: "#FF0000" }]),
      { authorizedSheetNames: [SHEET] }
    ).then(
      () => null,
      (error: unknown) => error as Error
    );

    expect(failure).not.toBeNull();
    const message = failure!.message;
    // 指名哪一步、哪个动作、哪类区域、缺哪个 API —— 不是一句笼统错误。
    expect(message).toContain("第 1 步");
    expect(message).toContain("setFill");
    expect(message).toContain("条件格式");
    expect(message).toContain("ExcelApi 1.19");
    expect(message).toContain("已拒绝执行");

    // 原表零写入：拒绝发生在写入之前。
    expect(withConditional.writes()).toEqual([]);
  });

  it("目标区域含合并单元格：setFont 被拒，点名合并区域，且零写入", async () => {
    const withMerge = legacyHost({ mergedAreas: true });
    installOfficeTestHost(withMerge);

    const failure = await executePlan(
      plan([{ type: "setFont", sheet: SHEET, range: "A1:B2", bold: true }]),
      { authorizedSheetNames: [SHEET] }
    ).then(
      () => null,
      (error: unknown) => error as Error
    );

    expect(failure).not.toBeNull();
    expect(failure!.message).toContain("合并单元格");
    expect(failure!.message).toContain("setFont");
    expect(withMerge.writes()).toEqual([]);
  });

  it("目标区域含数据验证：setFont 被拒，点名数据验证（clear(Formats) 会清除且无法还原），零写入", async () => {
    // 数据验证不是格式动作的管辖范围，本机也读不回逐格验证来源；clear(Formats) 会清除它。
    // ⇒ 与条件格式 / 合并单元格同口径：写入前 fail-closed，而不是静默丢验证。
    const withValidation = legacyHost({ dataValidationCells: ["B2"] });
    installOfficeTestHost(withValidation);

    const failure = await executePlan(
      plan([{ type: "setFont", sheet: SHEET, range: "A1:B2", bold: true }]),
      { authorizedSheetNames: [SHEET] }
    ).then(
      () => null,
      (error: unknown) => error as Error
    );

    expect(failure).not.toBeNull();
    expect(failure!.message).toContain("数据验证");
    expect(failure!.message).toContain("setFont");
    expect(withValidation.writes()).toEqual([]);
    // 拒绝发生在写入前 ⇒ 数据验证原样保留（1 格），没被 clear 掉。
    expect(withValidation.dataValidationCount(SHEET)).toBe(1);
  });

  it("区域无数据验证（type=「None」）：正常放行，不被误拒", async () => {
    // 默认替身无任何数据验证 ⇒ dataValidation.type === "None"，不得被"含数据验证"误伤。
    const plain = legacyHost();
    installOfficeTestHost(plain);
    expect(plain.dataValidationCount(SHEET)).toBe(0);

    const result = await executePlan(
      plan([{ type: "setFont", sheet: SHEET, range: "A1:B2", bold: true }]),
      { authorizedSheetNames: [SHEET] }
    );

    expect(result.undoSnapshot).not.toBeNull();
    expect(plain.writes().length).toBeGreaterThan(0);
  });

  it("连 ExcelApi 1.9 都没有的宿主：格式动作在**执行前检查**就被拒，点名缺哪个 API", async () => {
    const ancient = legacyHost({
      excelApiVersions: Array.from({ length: 8 }, (_, index) => `1.${index + 1}`)
    });
    installOfficeTestHost(ancient);

    const failure = await executePlan(
      plan([{ type: "setFont", sheet: SHEET, range: "A1:B2", bold: true }]),
      { authorizedSheetNames: [SHEET] }
    ).then(
      () => null,
      (error: unknown) => error as Error
    );

    expect(failure).not.toBeNull();
    // 能力检测前移：这句话来自预检，不是执行期捕获。
    expect(failure!.message).toContain("执行前检查未通过");
    expect(failure!.message).toContain("ExcelApi 1.19");
    expect(failure!.message).toContain("ExcelApi 1.9");
    expect(ancient.writes()).toEqual([]);
  });
});

describe("兼容路径不碰数据：值 / 公式 / 数字格式", () => {  it("setFont + setFill 执行并撤销后，值、公式与数字格式都原样保留（clear(Formats) 不动内容）", async () => {
    // 换一份含**空单元格**的初始数据：空单元格是最容易被"格式操作"误伤的。
    const dataHost = legacyHost({
      sheets: [{ name: SHEET, values: [["", "乙"], ["丙", "丁"]] }]
    });
    installOfficeTestHost(dataHost);
    dataHost.setCell(SHEET, "B1", 123);
    dataHost.setCell(SHEET, "A2", "文本");
    dataHost.setCell(SHEET, "B2", "=1+1", { asFormula: true });
    // B1 直接设置一个数字格式：兼容路径的 clear(Formats) 会清掉它，撤销必须补回来。
    dataHost.setCellDisplay(SHEET, "B1", { format: "0.00" });
    expect(await numberFormatAt("B1")).toBe("0.00");
    host = dataHost;

    const result = await executePlan(
      plan([
        { type: "setFont", sheet: SHEET, range: "A1:B2", bold: true },
        { type: "setFill", sheet: SHEET, range: "A1:B2", color: "#FFFF00" }
      ]),
      { authorizedSheetNames: [SHEET] }
    );

    // 执行后只有格式变了，数据内容一字未动。
    expect(host.cell(SHEET, "A1")).toBe("");
    expect(host.cell(SHEET, "B1")).toBe(123);
    expect(host.cell(SHEET, "A2")).toBe("文本");
    expect(host.cell(SHEET, "B2")).toBe("=1+1");

    await undoExecution(result.undoSnapshot!);

    // 撤销：数据内容仍不变（兼容路径的 clear(Formats) 只清直接格式、不碰值）。
    expect(host.cell(SHEET, "A1")).toBe("");
    expect(host.cell(SHEET, "B1")).toBe(123);
    expect(host.cell(SHEET, "A2")).toBe("文本");
    expect(host.cell(SHEET, "B2")).toBe("=1+1");
    // 数字格式被恢复（撤销顺序：先恢复格式、后写 numberFormat）。
    expect(await numberFormatAt("B1")).toBe("0.00");
  });
});

describe("值/公式动作的撤销不碰任何格式（按实际影响拆分恢复）", () => {
  // 这一组是**回归守卫**，守的是一条曾被两度搞错的性质：值/公式动作的撤销**只写值/公式**。
  //
  // 旧实现无差别地恢复值 + 填充 + 字体 + 全部格式 + 行高列宽，于是"只改值"的撤销也会写格式；
  // 曾经的补丁是"给值动作也逐格采集格式"—— 治标：写回的是采集到的值，但**写**这个动作本身
  // 就会把用户在执行之后对格式做的编辑覆盖掉，而且多了一次逐格读取。
  // 现在按 `undoRestoreScope` 收窄：值动作压根不进格式范围（`format` 为 null），
  // 于是混合格式**不被写回 = 原样保留**，不依赖"采得全不全"。
  it("≤1.17 宿主：writeValues 的撤销只写值，格式范围为空且混合对齐原样保留", async () => {
    host.setCellStyle(SHEET, "A1", { horizontalAlignment: "Left" });
    host.setCellStyle(SHEET, "B1", { horizontalAlignment: "Center" });
    host.setCellStyle(SHEET, "A1", { fontSize: 12 });
    host.setCellStyle(SHEET, "B1", { fontSize: 16 });

    const result = await executePlan(
      await anchoredPlan([{ type: "writeValues", sheet: SHEET, range: "A1:B1", values: [["新甲", "新乙"]] }]),
      { authorizedSheetNames: [SHEET] }
    );

    const saved = result.undoSnapshot!.ranges[0];
    // 值动作的恢复范围就只有值/公式：不采逐格格式（format 为 null），也不恢复格式。
    expect(saved.scope).toEqual({
      formulas: true,
      numberFormat: false,
      cellFormat: false,
      rowHeight: false,
      columnWidth: false
    });
    expect(saved.format).toBeNull();
    expect(host.cell(SHEET, "A1")).toBe("新甲");

    const writesBeforeUndo = host.writes().length;
    await undoExecution(result.undoSnapshot!);

    // 值回到执行前，且两种不同的对齐与字号各自保留（不是被拍平成同一个值）。
    expect(host.cell(SHEET, "A1")).toBe("甲");
    expect(host.cell(SHEET, "B1")).toBe("乙");
    expect(host.cellStyle(SHEET, "A1")?.horizontalAlignment).toBe("Left");
    expect(host.cellStyle(SHEET, "B1")?.horizontalAlignment).toBe("Center");
    expect(host.cellStyle(SHEET, "A1")?.fontSize).toBe(12);
    expect(host.cellStyle(SHEET, "B1")?.fontSize).toBe(16);
    // **只写了值/公式**：撤销期间没有任何格式类写入（既没有 clear(Formats)，也没有
    // setCellProperties）。这一条是"按实际影响拆分恢复"的直接取证点。
    const undoWrites = host.writes().slice(writesBeforeUndo);
    expect(undoWrites.length).toBeGreaterThan(0);
    expect(undoWrites.some((write) => write.kind === "values" || write.kind === "formulas")).toBe(
      true
    );
    expect(undoWrites.filter((write) => write.kind === "format" || write.kind === "clear")).toEqual(
      []
    );
  });
});
