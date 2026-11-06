import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { captureWorkbookSourceFingerprint, executePlan, undoExecution } from "./excel";
import type { AnalysisPlan, ExcelAction } from "./contracts";
import { createOfficeTestHost, installOfficeTestHost, type TestOfficeHost } from "./officeTestHost";

/**
 * **撤销按"动作实际影响"拆分恢复**的四组点名场景（ExcelApi ≤ 1.17 兼容路径）。
 *
 * 背景：旧撤销是"整块重写"——值动作也逐格采格式、格式动作无条件 `clear(Formats)` 后写回
 * 全量属性。后果有两个，且都**静默**：
 *  ① 把用户在本动作之后对别的属性做的编辑一起抹掉；
 *  ② 把"样式/继承来源"的格式固化成直接格式（显示看起来没变，来源语义已经变了）。
 *
 * 本文件的判据统一是"**只写本次动作真正会写回的属性**"：
 *  - `setFont{bold}` 的撤销只动 `fontBold`，`formatFields` 也只有它；
 *  - 只改行高/列宽时，撤销**不调用任何单元格清格式操作**；
 *  - 需要 `clear(Formats)` 时（本动作把某个执行前"未直接设置"的属性设上过），先把
 *    非本动作属性的**当前**取值并回去，清完再写回。
 *
 * 证据边界：**替身宿主自动化证据，非真实 Excel 实机证据**。替身按真实 `CellPropertiesFormat`
 * 可读全集建模（含删除线/上下标/图案色/八条边/收缩适应/文字方向/阅读顺序/保护/命名样式），
 * 且 `clear(Formats)` 清的是"整格全部格式 + 数字格式 + 命名样式"，不是"删掉某个局部对象"。
 */

const SHEET = "数据";

function plan(actions: ExcelAction[], overrides: Partial<AnalysisPlan> = {}): AnalysisPlan {
  return {
    id: "plan-undo-scope",
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

/** 只声明到 ExcelApi 1.17：批量许可永久版（如 ProPlus2024Volume）的实际上限。 */
const LEGACY_API_VERSIONS = Array.from({ length: 17 }, (_, index) => `1.${index + 1}`);

function legacyHost(
  overrides: Partial<Parameters<typeof createOfficeTestHost>[0]> = {}
): TestOfficeHost {
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

/** 撤销期间**格式类**写入（`clear(Formats)` 与 `setCellProperties` 都记在这里）。 */
function formatWritesSince(index: number) {
  return host
    .writes()
    .slice(index)
    .filter((write) => write.kind === "format" || write.kind === "clear");
}

describe("① 原有删除线/上下标/图案色/斜线边框：执行加粗后撤销仍逐格保留", () => {
  it("A1 已有删除线+上标+图案色+斜线边框，加粗再撤销：只掉加粗，其余四项逐项保留", async () => {
    host.setCellStyle(SHEET, "A1", {
      fontStrikethrough: true,
      fontSuperscript: true,
      fillPattern: "LightUp",
      fillColor: "#00FF00",
      fillPatternColor: "#0000FF",
      borders: { diagonalDown: { style: "Continuous", color: "#FF00FF", weight: "Thick" } }
    });

    const result = await executePlan(
      plan([{ type: "setFont", sheet: SHEET, range: "A1", bold: true }]),
      { authorizedSheetNames: [SHEET] }
    );

    // 动作确实生效，且快照只声明改了 `fontBold`（这就是"按动作实际影响"的第一层）。
    expect(host.cellStyle(SHEET, "A1")?.fontBold).toBe(true);
    expect(result.undoSnapshot!.ranges[0].formatFields).toEqual(["fontBold"]);

    await undoExecution(result.undoSnapshot!);

    const after = host.cellStyle(SHEET, "A1");
    // 加粗回到"未直接设置"（执行前它本就不是直接格式）。
    expect(after?.fontBold).toBeUndefined();
    // 四项原有直接格式逐项保留 —— 少采一个字段、或少合并一次"当前值"，这里就会掉。
    expect(after?.fontStrikethrough).toBe(true);
    expect(after?.fontSuperscript).toBe(true);
    expect(after?.fillPattern).toBe("LightUp");
    expect(after?.fillColor).toBe("#00FF00");
    expect(after?.fillPatternColor).toBe("#0000FF");
    expect(after?.borders?.diagonalDown?.style).toBe("Continuous");
    expect(after?.borders?.diagonalDown?.color).toBe("#FF00FF");
    expect(after?.borders?.diagonalDown?.weight).toBe("Thick");
  });

  it("另一格（B1）不受牵连：本轮撤销只作用于 A1，B1 的直接格式一个字节不动", async () => {
    host.setCellStyle(SHEET, "B1", { fontStrikethrough: true, fillPatternColor: "#112233" });

    const result = await executePlan(
      plan([{ type: "setFont", sheet: SHEET, range: "A1", bold: true }]),
      { authorizedSheetNames: [SHEET] }
    );
    await undoExecution(result.undoSnapshot!);

    expect(host.cellStyle(SHEET, "B1")?.fontStrikethrough).toBe(true);
    expect(host.cellStyle(SHEET, "B1")?.fillPatternColor).toBe("#112233");
  });
});

describe("② 加粗后用户又改文字方向/收缩适应/保护：撤销不覆盖这些后续修改", () => {
  it("撤销前用户改了 textOrientation/shrinkToFit/locked/formulaHidden，撤销后四项原样保留", async () => {
    const result = await executePlan(
      plan([{ type: "setFont", sheet: SHEET, range: "A1", bold: true }]),
      { authorizedSheetNames: [SHEET] }
    );
    expect(host.cellStyle(SHEET, "A1")?.fontBold).toBe(true);

    // —— 用户在本动作之后，自己改了这四个属性。
    host.setCellStyle(SHEET, "A1", {
      textOrientation: 45,
      shrinkToFit: true,
      locked: false,
      formulaHidden: true
    });

    // 这四个属性都不在本次恢复范围内 ⇒ 既不报冲突，也不该被写回。
    await undoExecution(result.undoSnapshot!);

    const after = host.cellStyle(SHEET, "A1");
    expect(after?.fontBold).toBeUndefined(); // 加粗按要求撤销
    expect(after?.textOrientation).toBe(45); // 后续修改保留
    expect(after?.shrinkToFit).toBe(true);
    expect(after?.locked).toBe(false);
    expect(after?.formulaHidden).toBe(true);
  });

  it("加粗后用户改了**本动作没碰过的另一个受管属性**（字色）：撤销保留用户的新字色", async () => {
    // 字色是"受管属性"（属于撤销会处理的那一类），但**不是**本次动作改的属性。
    // 这一条正是"属性子集收窄"的判别点：写回时若按**全集**取执行前的值，
    // 就会把用户在加粗之后改的新字色一起覆盖掉。
    host.setCellStyle(SHEET, "A1", { fontColor: "#FF0000" });

    const result = await executePlan(
      plan([{ type: "setFont", sheet: SHEET, range: "A1", bold: true }]),
      { authorizedSheetNames: [SHEET] }
    );
    // 快照只声明改了 `fontBold`（字色不在内）。
    expect(result.undoSnapshot!.ranges[0].formatFields).toEqual(["fontBold"]);

    // 用户在本动作之后把字色改成绿色。
    host.setCellStyle(SHEET, "A1", { fontColor: "#00FF00" });

    await undoExecution(result.undoSnapshot!);

    const after = host.cellStyle(SHEET, "A1");
    expect(after?.fontBold).toBeUndefined(); // 加粗撤销
    expect(after?.fontColor).toBe("#00FF00"); // ★ 用户的新字色保留，而不是被写回 #FF0000
  });

  it("反例锚点：若用户改的**正是**本次要写回的属性（fontBold），撤销必须在任何写入前报冲突", async () => {
    const result = await executePlan(
      plan([{ type: "setFont", sheet: SHEET, range: "A1", bold: true }]),
      { authorizedSheetNames: [SHEET] }
    );

    // 用户自己把加粗关掉 —— 这是本动作会写回的属性。
    host.setCellStyle(SHEET, "A1", { fontBold: false });

    const before = host.writes().length;
    const failure = await undoExecution(result.undoSnapshot!).then(
      () => null,
      (error: unknown) => error as Error
    );

    expect(failure).not.toBeNull();
    expect(failure!.message).toContain("在本次撤销将要写回的字段");
    expect(failure!.message).toContain("加粗");
    // 冲突在**任何**写入之前抛出：整次撤销保持只读失败。
    expect(host.writes().slice(before)).toEqual([]);
    // 用户的修改原样保留（没被静默覆盖）。
    expect(host.cellStyle(SHEET, "A1")?.fontBold).toBe(false);
  });
});

describe("③ 含条件格式的区域只改行高：撤销不调用单元格清格式，规则保持不变", () => {
  it("区域带 1 条条件格式 + 直接字体色，resizeRange 后撤销：零格式写入、规则与字体色都不动", async () => {
    // 区域含 1 条条件格式规则；A1 另有直接字体色（证明"没被清掉"）。
    host.setConditionalFormatCount(SHEET, 1);
    host.setCellStyle(SHEET, "A1", { fontColor: "#123456" });

    const result = await executePlan(
      plan([{ type: "resizeRange", sheet: SHEET, range: "A1:B2", rowHeight: 30 }]),
      { authorizedSheetNames: [SHEET] }
    );

    // 只改行高不该被"区域含条件格式"误拒 —— 这是本轮修复的关键点之一。
    expect(host.rowHeight(SHEET, 0)).toBe(30);

    const saved = result.undoSnapshot!.ranges[0];
    expect(saved.scope).toEqual({
      formulas: false,
      numberFormat: false,
      cellFormat: false,
      rowHeight: true,
      columnWidth: false
    });
    // 有为行高采的快照，但**没有**逐格格式矩阵（`cellFormat` 为假 ⇒ 不采）。
    expect(saved.format).not.toBeNull();
    expect(saved.format?.directCells ?? null).toBeNull();

    const before = host.writes().length;
    await undoExecution(result.undoSnapshot!);

    // 行高回到"未直接设置"（标准行高）。
    expect(host.rowHeight(SHEET, 0)).toBe(15);
    // ★ 判据：撤销期间**没有**任何格式类写入 ⇒ 没调用 `clear(Formats)`、也没 `setCellProperties`。
    expect(formatWritesSince(before)).toEqual([]);
    // 直接格式与条件格式规则都原样。
    expect(host.cellStyle(SHEET, "A1")?.fontColor).toBe("#123456");
    expect(host.conditionalFormatCount(SHEET)).toBe(1);
  });

  it("只改列宽同理：撤销零格式写入，且不碰区域内的条件格式与直接格式", async () => {
    host.setConditionalFormatCount(SHEET, 2);
    host.setCellStyle(SHEET, "A1", { fontStrikethrough: true });

    const result = await executePlan(
      plan([{ type: "resizeRange", sheet: SHEET, range: "A1:A2", columnWidth: 24 }]),
      { authorizedSheetNames: [SHEET] }
    );
    expect(host.columnWidth(SHEET, 0)).toBe(24);

    const before = host.writes().length;
    await undoExecution(result.undoSnapshot!);

    expect(host.columnWidth(SHEET, 0)).toBe(8.43);
    expect(formatWritesSince(before)).toEqual([]);
    expect(host.cellStyle(SHEET, "A1")?.fontStrikethrough).toBe(true);
    expect(host.conditionalFormatCount(SHEET)).toBe(2);
  });
});

describe("⑤ 同区域多个格式动作：反序撤销不得用旧现场把较晚动作刚恢复的属性写回", () => {
  it("先加粗后居中：撤销后加粗与居中**都**回到执行前，混合字号原样保留", async () => {
    // 两格字号不同（A1=12、A2=16），且两格都**没有**直接的对齐格式。
    host.setCellStyle(SHEET, "A1", { fontSize: 12 });
    host.setCellStyle(SHEET, "A2", { fontSize: 16 });

    const result = await executePlan(
      plan([
        { type: "setFont", sheet: SHEET, range: "A1:B2", bold: true },
        { type: "setAlignment", sheet: SHEET, range: "A1:B2", horizontal: "center" }
      ]),
      { authorizedSheetNames: [SHEET] }
    );
    expect(host.cellStyle(SHEET, "A1")?.fontBold).toBe(true);
    expect(host.cellStyle(SHEET, "A1")?.horizontalAlignment).toBe("Center");

    await undoExecution(result.undoSnapshot!);

    const a1 = host.cellStyle(SHEET, "A1");
    expect(a1?.fontBold).toBeUndefined();
    // ★ 判别点：加粗是**较早**的动作，恢复时排在居中之后。它在 `clear(Formats)` 后
    //   合并"非本动作属性的**当前**值"时，必须读**恢复那一刻**的现场；若沿用复核阶段
    //   一次性读回的快照（那份还带着居中），居中就会被重新写回来 —— 撤销看起来"做过了"，
    //   实际留下残留。这条断言就是钉住这一点的唯一判别点。
    expect(a1?.horizontalAlignment).toBeUndefined();
    expect(a1?.fontSize).toBe(12);
    const a2 = host.cellStyle(SHEET, "A2");
    expect(a2?.fontBold).toBeUndefined();
    expect(a2?.horizontalAlignment).toBeUndefined();
    expect(a2?.fontSize).toBe(16);
  });

  it("部分重叠区域改不同属性：重叠格两种属性都恢复，非重叠格只恢复各自被改的属性", async () => {
    // 动作 1：A1:B1 加粗；动作 2：B1:B2 水平居中。重叠格 = B1（同时被两个动作改过）。
    const result = await executePlan(
      plan([
        { type: "setFont", sheet: SHEET, range: "A1:B1", bold: true },
        { type: "setAlignment", sheet: SHEET, range: "B1:B2", horizontal: "center" }
      ]),
      { authorizedSheetNames: [SHEET] }
    );
    expect(host.cellStyle(SHEET, "B1")?.fontBold).toBe(true);
    expect(host.cellStyle(SHEET, "B1")?.horizontalAlignment).toBe("Center");

    await undoExecution(result.undoSnapshot!);

    // 重叠格 B1：加粗与居中**都**撤销（较早的加粗恢复时不得把居中写回）。
    const b1 = host.cellStyle(SHEET, "B1");
    expect(b1?.fontBold).toBeUndefined();
    expect(b1?.horizontalAlignment).toBeUndefined();
    // 非重叠格 A1：只撤销加粗，对齐从未被改、保持未直接设置。
    const a1 = host.cellStyle(SHEET, "A1");
    expect(a1?.fontBold).toBeUndefined();
    expect(a1?.horizontalAlignment).toBeUndefined();
    // 非重叠格 B2：只撤销居中，加粗从未被改。
    const b2 = host.cellStyle(SHEET, "B2");
    expect(b2?.horizontalAlignment).toBeUndefined();
    expect(b2?.fontBold).toBeUndefined();
  });

  it("同一属性连续覆盖（字号 12→14→16）：撤销后回到执行前各自的混合字号，不是中途值", async () => {
    // 执行前 A1=12、A2=16（混合字号）。
    host.setCellStyle(SHEET, "A1", { fontSize: 12 });
    host.setCellStyle(SHEET, "A2", { fontSize: 16 });

    const result = await executePlan(
      plan([
        { type: "setFont", sheet: SHEET, range: "A1:A2", size: 14 },
        { type: "setFont", sheet: SHEET, range: "A1:A2", size: 16 }
      ]),
      { authorizedSheetNames: [SHEET] }
    );
    expect(host.cellStyle(SHEET, "A1")?.fontSize).toBe(16);
    expect(host.cellStyle(SHEET, "A2")?.fontSize).toBe(16);

    await undoExecution(result.undoSnapshot!);

    // 逆序撤销：先撤"16"（回到 14 的中间态），再撤"14"（回到 12/16 的初始混合态）。
    // 判别点：较早动作（14）恢复时读的是**恢复那一刻**的现场，最终每格回到自己的初始字号，
    // 而不是被统一写回某个中途值。
    expect(host.cellStyle(SHEET, "A1")?.fontSize).toBe(12);
    expect(host.cellStyle(SHEET, "A2")?.fontSize).toBe(16);
  });

  it("恢复中途失败：如实抛错，不把部分恢复当成成功", async () => {
    // 用 writeValues：撤销时经 `range.formulas` 写回（宿主里走 commitWrite，可被 failWrites 拦截）。
    const result = await executePlan(
      await anchoredPlan([{ type: "writeValues", sheet: SHEET, range: "A1:A2", values: [["X"], ["Y"]] }]),
      { authorizedSheetNames: [SHEET] }
    );
    expect(host.cell(SHEET, "A1")).toBe("X");

    // 模拟撤销期间的值写回被宿主拒绝：撤销必须如实抛错，而不是静默吞掉、让调用方以为成功。
    host.failWrites((event) => event.kind === "formulas", "宿主拒绝了撤销写回");

    const failure = await undoExecution(result.undoSnapshot!).then(
      () => null,
      (error: unknown) => error as Error
    );

    expect(failure).not.toBeNull();
  });
});

describe("④ 非 Normal 命名样式 + 直接格式覆盖的恢复", () => {
  it("A1 是 Good 样式且被红色直接填充覆盖：加粗撤销后样式仍为 Good、红色覆盖仍在", async () => {
    host.setCellStyleName(SHEET, "A1", "Good");
    host.setCellStyle(SHEET, "A1", { fillColor: "#FF0000" });

    const result = await executePlan(
      plan([{ type: "setFont", sheet: SHEET, range: "A1", bold: true }]),
      { authorizedSheetNames: [SHEET] }
    );
    expect(host.cellStyle(SHEET, "A1")?.fontBold).toBe(true);

    await undoExecution(result.undoSnapshot!);

    // 命名样式还原成 Good —— 不是被 `clear(Formats)` 打回 Normal。
    expect(host.cellStyleName(SHEET, "A1")).toBe("Good");
    // 直接格式覆盖（红色填充）也还原。
    expect(host.cellStyle(SHEET, "A1")?.fillColor).toBe("#FF0000");
    // 加粗撤掉。
    expect(host.cellStyle(SHEET, "A1")?.fontBold).toBeUndefined();
  });

  it("命名样式 + 数字格式：clear(Formats) 会清掉两者，撤销必须都写回", async () => {
    host.setCellStyleName(SHEET, "A1", "Neutral");
    // 通过真实 API 写数字格式，让替身的单元格式数字格式与 range.numberFormat 同源。
    await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getItem(SHEET).getRange("A1");
      range.numberFormat = [["0.00"]];
      await context.sync();
    });

    const result = await executePlan(
      plan([{ type: "setFont", sheet: SHEET, range: "A1", bold: true }]),
      { authorizedSheetNames: [SHEET] }
    );
    await undoExecution(result.undoSnapshot!);

    expect(host.cellStyleName(SHEET, "A1")).toBe("Neutral");
    expect(host.cellNumberFormat(SHEET, "A1")).toBe("0.00");
  });
});

it('宿主静默未保留恢复值时，回读拒绝宣告全部恢复成功', async () => {
  const result = await executePlan(await anchoredPlan([{type:'writeValues',sheet:SHEET,range:'A1',values:[['新值']]}]), {authorizedSheetNames:[SHEET]});
  const runtime = globalThis as unknown as {Excel:{run:(callback:(context:Excel.RequestContext)=>Promise<void>)=>Promise<void>}};
  const originalRun = runtime.Excel.run;
  runtime.Excel.run = callback => originalRun(async context => {
    const sync = context.sync.bind(context);
    context.sync = async <T>(passThroughValue?: T): Promise<T> => {
      const value = await sync(passThroughValue);
      // Simulate a host that accepts the call but fails to retain the restored value.
      if (host.cell(SHEET,'A1') === '甲') host.setCell(SHEET,'A1','新值');
      return value;
    };
    await callback(context);
  });
  try { await expect(undoExecution(result.undoSnapshot!)).rejects.toThrow('恢复后回读不一致'); }
  finally { runtime.Excel.run = originalRun; }
});
