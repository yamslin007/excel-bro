import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { executePlan, undoExecution } from "./excel";
import type { AnalysisPlan, ExcelAction } from "./contracts";
import { createOfficeTestHost, installOfficeTestHost, type TestOfficeHost } from "./officeTestHost";

/**
 * Excel 混合格式**按粒度**保存与恢复（Task #37 的核心核对）。
 *
 * 旧撤销快照把格式存成**区域级单值**（一个 `horizontalAlignment`、一个 `rowHeight`、
 * 一个 `fillColor`…），混合格式被折叠成单一值，撤销时统一写回单值 —— 逐格差异全丢。
 *
 * 本文件证明：新捕获走 `getCellProperties`（逐格）＋ `getRowProperties`/`getColumnProperties`
 * （逐行/逐列），撤销走 `setCellProperties`＋`setRowProperties`/`setColumnProperties`，
 * 因此"执行前确实各不相同的格式"在撤销后能**逐格/逐行/逐列**精确还原，而不是被拍平。
 *
 * 全部跑在 `officeTestHost`（有真实状态、逐格 style / 逐行行高 / 逐列列宽的替身）上，
 * 走**真实产品代码** `executePlan` 与 `undoExecution`，不 mock 执行器或撤销流程。
 * 证据边界：替身宿主自动化证据，非真实 Excel 实机证据。
 */

const SHEET = "数据";

function plan(actions: ExcelAction[], overrides: Partial<AnalysisPlan> = {}): AnalysisPlan {
  return {
    id: "plan-mixed-format",
    title: "格式",
    summary: "格式",
    assumptions: [],
    warnings: [],
    actions,
    ...overrides
  };
}

let host: TestOfficeHost;

function createHost(): TestOfficeHost {
  return createOfficeTestHost({
    url: "https://contoso.sharepoint.com/format.xlsx",
    activeSheet: SHEET,
    selectedRange: "A1:B2",
    // 声明支持 ExcelApi 1.19：`getDisplayedCellProperties`（显示值捕获）需要它。
    // 这是本测试文件"正常路径"的前提；"宿主不支持 1.19 ⇒ 执行前拒绝"由专门用例覆盖。
    excelApiVersions: Array.from({ length: 19 }, (_, index) => `1.${index + 1}`),
    sheets: [
      {
        name: SHEET,
        values: [
          ["甲", "乙"],
          ["丙", "丁"]
        ]
      }
    ]
  });
}

beforeEach(() => {
  host = createHost();
  installOfficeTestHost(host);
});

afterEach(() => {
  // 宿主替身只装到全局，无其它副作用需清理。
});

describe("混合格式按粒度保存与恢复", () => {
  it("逐格填充色：执行前各不相同 → setFill 统一 → 撤销后逐格回到各自色", async () => {
    // 执行前：A1 红 / B1 蓝 / A2 绿 / B2 黄（确实各不相同，不掩盖缺口）。
    host.setCellStyle(SHEET, "A1", { fillColor: "#FF0000" });
    host.setCellStyle(SHEET, "B1", { fillColor: "#0000FF" });
    host.setCellStyle(SHEET, "A2", { fillColor: "#00FF00" });
    host.setCellStyle(SHEET, "B2", { fillColor: "#FFFF00" });

    const result = await executePlan(
      plan([{ type: "setFill", sheet: SHEET, range: "A1:B2", color: "#888888" }]),
      { authorizedSheetNames: [SHEET] }
    );

    expect(result.undoSnapshot).not.toBeNull();
    // 回读：执行后四格统一为灰色。
    expect(host.cellStyle(SHEET, "A1")?.fillColor).toBe("#888888");
    expect(host.cellStyle(SHEET, "B1")?.fillColor).toBe("#888888");
    expect(host.cellStyle(SHEET, "A2")?.fillColor).toBe("#888888");
    expect(host.cellStyle(SHEET, "B2")?.fillColor).toBe("#888888");

    // 撤销：逐格回到各自初始色，而不是被拍平成同一个值。
    await undoExecution(result.undoSnapshot!);
    expect(host.cellStyle(SHEET, "A1")?.fillColor).toBe("#FF0000");
    expect(host.cellStyle(SHEET, "B1")?.fillColor).toBe("#0000FF");
    expect(host.cellStyle(SHEET, "A2")?.fillColor).toBe("#00FF00");
    expect(host.cellStyle(SHEET, "B2")?.fillColor).toBe("#FFFF00");
  });

  it("逐格字体与对齐：setFont/setAlignment 后撤销逐格还原", async () => {
    // 执行前：A1 加粗左对齐、B1 斜体居中、A2 大字右对齐、B2 微软雅黑。
    host.setCellStyle(SHEET, "A1", { fontBold: true, horizontalAlignment: "Left" });
    host.setCellStyle(SHEET, "B1", { fontItalic: true, horizontalAlignment: "Center" });
    host.setCellStyle(SHEET, "A2", { fontSize: 20, horizontalAlignment: "Right" });
    host.setCellStyle(SHEET, "B2", { fontName: "微软雅黑" });

    const result = await executePlan(
      plan([
        { type: "setFont", sheet: SHEET, range: "A1:B2", bold: true, name: "Arial", size: 12 },
        { type: "setAlignment", sheet: SHEET, range: "A1:B2", horizontal: "center" }
      ]),
      { authorizedSheetNames: [SHEET] }
    );

    // 回读：执行后统一字体名与加粗。
    expect(host.cellStyle(SHEET, "A1")?.fontBold).toBe(true);
    expect(host.cellStyle(SHEET, "B2")?.fontName).toBe("Arial");
    expect(host.cellStyle(SHEET, "A2")?.fontSize).toBe(12);

    await undoExecution(result.undoSnapshot!);

    // 逐格还原：B2 字体名回到微软雅黑，A2 字号回到 20，B1 斜体保持（未被 setFont 覆盖）。
    expect(host.cellStyle(SHEET, "B2")?.fontName).toBe("微软雅黑");
    expect(host.cellStyle(SHEET, "A2")?.fontSize).toBe(20);
    expect(host.cellStyle(SHEET, "B1")?.fontItalic).toBe(true);
    expect(host.cellStyle(SHEET, "A1")?.horizontalAlignment).toBe("Left");
    expect(host.cellStyle(SHEET, "B1")?.horizontalAlignment).toBe("Center");
  });

  it("逐行行高 / 逐列列宽：resizeRange 后撤销逐行逐列还原", async () => {
    // 执行前：第 1 行 20、第 2 行 30；A 列 10、B 列 20（各不相同）。
    host.setRowHeight(SHEET, 0, 20);
    host.setRowHeight(SHEET, 1, 30);
    host.setColumnWidth(SHEET, 0, 10);
    host.setColumnWidth(SHEET, 1, 20);

    const result = await executePlan(
      plan([{ type: "resizeRange", sheet: SHEET, range: "A1:B2", rowHeight: 40, columnWidth: 15 }]),
      { authorizedSheetNames: [SHEET] }
    );

    // 回读：执行后统一行高列宽。
    expect(host.rowHeight(SHEET, 0)).toBe(40);
    expect(host.rowHeight(SHEET, 1)).toBe(40);
    expect(host.columnWidth(SHEET, 0)).toBe(15);
    expect(host.columnWidth(SHEET, 1)).toBe(15);

    await undoExecution(result.undoSnapshot!);

    // 逐行/逐列还原，而不是被拍平成单一尺寸。
    expect(host.rowHeight(SHEET, 0)).toBe(20);
    expect(host.rowHeight(SHEET, 1)).toBe(30);
    expect(host.columnWidth(SHEET, 0)).toBe(10);
    expect(host.columnWidth(SHEET, 1)).toBe(20);
  });

  it("逐格边框：setBorders 后撤销逐格还原（不同格不同边框）", async () => {
    // 执行前：A1 只有上边、B1 只有下边，样式/颜色/粗细各不相同。
    host.setCellStyle(SHEET, "A1", {
      borders: { top: { style: "Continuous", color: "#FF0000", weight: "Thick" } }
    });
    host.setCellStyle(SHEET, "B1", {
      borders: { bottom: { style: "Dashed", color: "#0000FF", weight: "Medium" } }
    });

    const result = await executePlan(
      plan([
        {
          type: "setBorders",
          sheet: SHEET,
          range: "A1:B1",
          sides: ["top", "bottom", "left", "right"],
          style: "continuous",
          color: "#000000",
          weight: "thin"
        }
      ]),
      { authorizedSheetNames: [SHEET] }
    );

    // 回读：执行后 A1 上边变成黑色细实线。
    expect(host.cellStyle(SHEET, "A1")?.borders?.top?.color).toBe("#000000");
    expect(host.cellStyle(SHEET, "A1")?.borders?.top?.weight).toBe("Thin");

    await undoExecution(result.undoSnapshot!);

    // 逐格还原：A1 上边回到红色粗线，B1 下边回到蓝色虚线。
    expect(host.cellStyle(SHEET, "A1")?.borders?.top?.color).toBe("#FF0000");
    expect(host.cellStyle(SHEET, "A1")?.borders?.top?.weight).toBe("Thick");
    expect(host.cellStyle(SHEET, "B1")?.borders?.bottom?.style).toBe("Dashed");
    expect(host.cellStyle(SHEET, "B1")?.borders?.bottom?.color).toBe("#0000FF");
  });
});

describe("纯格式动作的执行前拒绝", () => {
  it("目标区域过大、格式无法逐格捕获时，格式动作在执行前被拒，不产生不可撤销修改", async () => {
    // 构造一个超过 formatCaptureMaxCells(20000) 的区域：150 行 × 150 列 = 22500 格。
    // setFill 是纯格式动作，区域过大 ⇒ captureRangeFormat 降级为 null ⇒ 必须执行前拒绝。
    const bigRange = "A1:EY150";
    await expect(
      executePlan(
        plan([{ type: "setFill", sheet: SHEET, range: bigRange, color: "#FF0000" }]),
        { authorizedSheetNames: [SHEET] }
      )
    ).rejects.toThrow("无法在执行前可靠保存恢复依据");

    // 负向断言的正向锚点：确实一个格式写都没有发生（被拒在写入之前）。
    const formatWrites = host.writes().filter((write) => write.kind === "format");
    expect(formatWrites).toEqual([]);
  });
});

describe("格式采集缺失在修改前拒绝（原表零修改）", () => {
  // 这些用例证明：生产代码不只查"区域格数上限"，还把**逐格矩阵尺寸、逐行/逐列数量、
  // 缺项**一并当作"格式不可完整捕获"的判据，任一不满足都在写入前拒绝 —— 而不是执行完
  // 才发现撤销不了。

  function corruptingHost(
    corrupt: (value: unknown, kind: "cell" | "row" | "column") => unknown
  ): TestOfficeHost {
    return createOfficeTestHost({
      url: "https://contoso.sharepoint.com/format.xlsx",
      activeSheet: SHEET,
      selectedRange: "A1:B2",
      corruptFormatRead: corrupt,
      sheets: [{ name: SHEET, values: [["甲", "乙"], ["丙", "丁"]] }]
    });
  }

  it("逐格矩阵被截断（行数不足）⇒ 执行前拒绝，原表零写入", async () => {
    // getDisplayedCellProperties 返回的矩阵比目标区域少一行：模拟宿主/网络把逐格矩阵截断。
    const shortHost = corruptingHost((value, kind) =>
      kind === "cell" ? (value as unknown[]).slice(0, (value as unknown[]).length - 1) : value
    );
    installOfficeTestHost(shortHost);

    await expect(
      executePlan(
        plan([{ type: "setFill", sheet: SHEET, range: "A1:B2", color: "#FF0000" }]),
        { authorizedSheetNames: [SHEET] }
      )
    ).rejects.toThrow("无法在执行前可靠保存恢复依据");

    // 原表零写入：被拒在写入之前，没有任何 format 写、也没有任何 values 写。
    expect(shortHost.writes()).toEqual([]);
  });

  it("逐列矩阵数量不足 ⇒ 执行前拒绝，原表零写入", async () => {
    // getColumnProperties 返回的列数比目标区域少一列。
    const shortHost = corruptingHost((value, kind) =>
      kind === "column" ? (value as unknown[]).slice(0, (value as unknown[]).length - 1) : value
    );
    installOfficeTestHost(shortHost);

    await expect(
      executePlan(
        plan([{ type: "resizeRange", sheet: SHEET, range: "A1:B2", columnWidth: 20 }]),
        { authorizedSheetNames: [SHEET] }
      )
    ).rejects.toThrow("行高/列宽无法在执行前可靠确定");

    expect(shortHost.writes()).toEqual([]);
  });

  it("逐行矩阵数量不足 ⇒ 执行前拒绝，原表零写入", async () => {
    // getRowProperties 返回的行数比目标区域少一行。
    const shortHost = corruptingHost((value, kind) =>
      kind === "row" ? (value as unknown[]).slice(0, (value as unknown[]).length - 1) : value
    );
    installOfficeTestHost(shortHost);

    await expect(
      executePlan(
        plan([{ type: "resizeRange", sheet: SHEET, range: "A1:B2", rowHeight: 30 }]),
        { authorizedSheetNames: [SHEET] }
      )
    ).rejects.toThrow("行高/列宽无法在执行前可靠确定");

    expect(shortHost.writes()).toEqual([]);
  });

  it("逐格矩阵出现短行（某行内列数不足）⇒ 执行前拒绝，原表零写入", async () => {
    // 矩阵某一行被截断（列数比目标区域少）：模拟"短行"。
    const shortHost = corruptingHost((value, kind) => {
      if (kind !== "cell") return value;
      const matrix = value as unknown[][];
      const cloned = matrix.map((row) => [...row]);
      cloned[0] = cloned[0].slice(0, cloned[0].length - 1); // 第一行少一列
      return cloned;
    });
    installOfficeTestHost(shortHost);

    await expect(
      executePlan(
        plan([{ type: "setFill", sheet: SHEET, range: "A1:B2", color: "#FF0000" }]),
        { authorizedSheetNames: [SHEET] }
      )
    ).rejects.toThrow("无法在执行前可靠保存恢复依据");

    expect(shortHost.writes()).toEqual([]);
  });

  it("逐格矩阵出现缺项（元素为 null）⇒ 执行前拒绝，原表零写入", async () => {
    // 矩阵里某个元素变成 null：模拟采集失败（宿主异常、返回被截断成 undefined）。
    const shortHost = corruptingHost((value, kind) => {
      if (kind !== "cell") return value;
      const matrix = value as unknown[][];
      const cloned = matrix.map((row) => [...row]);
      cloned[0][0] = null;
      return cloned;
    });
    installOfficeTestHost(shortHost);

    await expect(
      executePlan(
        plan([{ type: "setFill", sheet: SHEET, range: "A1:B2", color: "#FF0000" }]),
        { authorizedSheetNames: [SHEET] }
      )
    ).rejects.toThrow("无法在执行前可靠保存恢复依据");

    expect(shortHost.writes()).toEqual([]);
  });
});

describe("合法无填充的保存与恢复（未采集 ≠ 默认值）", () => {
  it("无填充格：setFill 后撤销回到「无填充」而非白色", async () => {
    // A1 执行前没有任何直接设置的填充（合法空态）；B1 有蓝色填充。
    // setFill 把两者都刷成灰色。撤销后 A1 必须回到「无填充」（pattern "None"、无颜色），
    // 而不是被写回一个白色 `#FFFFFF` 或别的"看起来像默认"的值。
    host.setCellStyle(SHEET, "B1", { fillColor: "#0000FF" });

    const result = await executePlan(
      plan([{ type: "setFill", sheet: SHEET, range: "A1:B1", color: "#888888" }]),
      { authorizedSheetNames: [SHEET] }
    );

    // 回读：执行后两格统一灰色。
    expect(host.cellStyle(SHEET, "A1")?.fillColor).toBe("#888888");
    expect(host.cellStyle(SHEET, "B1")?.fillColor).toBe("#888888");

    await undoExecution(result.undoSnapshot!);

    // A1 回到「无填充」：fillPattern 为 None、无填充色（不是白色）。
    expect(host.cellStyle(SHEET, "A1")?.fillPattern).toBe("None");
    expect(host.cellStyle(SHEET, "A1")?.fillColor).toBe("");
    // B1 回到蓝色。
    expect(host.cellStyle(SHEET, "B1")?.fillColor).toBe("#0000FF");
    expect(host.cellStyle(SHEET, "B1")?.fillPattern).toBe("Solid");
  });

  it("行高/列宽全部未直接设置：resizeRange 后撤销回到标准行高/列宽", async () => {
    // 区域内所有行/列都未直接设置行高/列宽（使用标准行高/列宽）。
    const result = await executePlan(
      plan([{ type: "resizeRange", sheet: SHEET, range: "A1:B2", rowHeight: 40, columnWidth: 18 }]),
      { authorizedSheetNames: [SHEET] }
    );

    // 回读：执行后统一 40/18。
    expect(host.rowHeight(SHEET, 0)).toBe(40);
    expect(host.rowHeight(SHEET, 1)).toBe(40);
    expect(host.columnWidth(SHEET, 0)).toBe(18);
    expect(host.columnWidth(SHEET, 1)).toBe(18);

    await undoExecution(result.undoSnapshot!);

    // 撤销回到标准行高/列宽（未直接设置 = 读默认 15/8.43），而不是被写死成 40/18。
    expect(host.rowHeight(SHEET, 0)).toBe(15);
    expect(host.rowHeight(SHEET, 1)).toBe(15);
    expect(host.columnWidth(SHEET, 0)).toBe(8.43);
    expect(host.columnWidth(SHEET, 1)).toBe(8.43);
  });

  it("行高/列宽混合了直接设置与标准：resizeRange 执行前拒绝，原表零写入", async () => {
    // 第 1 行 / A 列直接设置行高/列宽，第 2 行 / B 列保持标准（未直接设置）。
    // 混合状态无法逐行恢复（setRowProperties 只能写具体值、useStandardHeight 只能整区域
    // 统一重置），因此 resizeRange 必须在执行前拒绝。
    host.setRowHeight(SHEET, 0, 25);
    host.setColumnWidth(SHEET, 0, 12);

    await expect(
      executePlan(
        plan([{ type: "resizeRange", sheet: SHEET, range: "A1:B2", rowHeight: 40, columnWidth: 18 }]),
        { authorizedSheetNames: [SHEET] }
      )
    ).rejects.toThrow("行高/列宽无法在执行前可靠确定");

    // 原表零写入。
    expect(host.writes()).toEqual([]);
    // 且行高/列宽未被改动。
    expect(host.rowHeight(SHEET, 0)).toBe(25);
    expect(host.columnWidth(SHEET, 0)).toBe(12);
  });
});

describe("格式操作不改变数据内容", () => {
  it("setFill/setFont 不改变空单元格、数值、文本与公式", async () => {
    // 构造混合数据：空单元格、数值、文本、公式。
    host.setCell(SHEET, "A1", "");           // 空
    host.setCell(SHEET, "B1", 123);           // 数值
    host.setCell(SHEET, "A2", "文本");        // 文本
    host.setCell(SHEET, "B2", "=1+1", { asFormula: true }); // 公式

    const result = await executePlan(
      plan([
        { type: "setFill", sheet: SHEET, range: "A1:B2", color: "#FFFF00" },
        { type: "setFont", sheet: SHEET, range: "A1:B2", bold: true }
      ]),
      { authorizedSheetNames: [SHEET] }
    );

    // 执行后数据内容不变（只有格式变了）。
    expect(host.cell(SHEET, "A1")).toBe("");
    expect(host.cell(SHEET, "B1")).toBe(123);
    expect(host.cell(SHEET, "A2")).toBe("文本");
    expect(host.cell(SHEET, "B2")).toBe("=1+1");

    await undoExecution(result.undoSnapshot!);

    // 撤销后数据内容仍然不变（格式撤销不该碰值/公式）。
    expect(host.cell(SHEET, "A1")).toBe("");
    expect(host.cell(SHEET, "B1")).toBe(123);
    expect(host.cell(SHEET, "A2")).toBe("文本");
    expect(host.cell(SHEET, "B2")).toBe("=1+1");
  });
});

describe("值动作的撤销不碰任何格式（按实际影响拆分恢复）", () => {
  // 曾经的做法是"给值动作也逐格采集格式"，治标：写回的虽是采集值，但**写**这个动作本身
  // 就会覆盖用户在执行之后对格式做的编辑。现在值动作不进格式范围（`format` 为 null），
  // 混合格式因此"不被写回 = 原样保留"。
  it("writeValues 在 1.19 宿主上：撤销只写值，格式范围为空且混合对齐与字号原样保留", async () => {
    host.setCellStyle(SHEET, "A1", { horizontalAlignment: "Left" });
    host.setCellStyle(SHEET, "B1", { horizontalAlignment: "Center" });
    host.setCellStyle(SHEET, "A1", { fontSize: 12 });
    host.setCellStyle(SHEET, "B1", { fontSize: 16 });

    const result = await executePlan(
      plan([{ type: "writeValues", sheet: SHEET, range: "A1:B1", values: [["新甲", "新乙"]] }]),
      { authorizedSheetNames: [SHEET] }
    );

    const saved = result.undoSnapshot!.ranges[0];
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

    expect(host.cell(SHEET, "A1")).toBe("甲");
    expect(host.cell(SHEET, "B1")).toBe("乙");
    expect(host.cellStyle(SHEET, "A1")?.horizontalAlignment).toBe("Left");
    expect(host.cellStyle(SHEET, "B1")?.horizontalAlignment).toBe("Center");
    expect(host.cellStyle(SHEET, "A1")?.fontSize).toBe(12);
    expect(host.cellStyle(SHEET, "B1")?.fontSize).toBe(16);
    // 撤销期间没有任何格式类写入 —— 这是本组的取证点。
    const undoWrites = host.writes().slice(writesBeforeUndo);
    expect(undoWrites.length).toBeGreaterThan(0);
    expect(undoWrites.filter((write) => write.kind === "format" || write.kind === "clear")).toEqual(
      []
    );
  });
});

describe("端到端反例：显示值语义（未直接设置 / 样式来源 / 明确无填充）", () => {
  // 这组用例直接对应验收原则：撤销恢复的是"执行前那一刻的实际显示状态"，
  // 不是"直接设置状态"。`getDisplayedCellProperties` 提供确定显示值，恢复据此写回。

  it("原本未直接设置加粗（显示不加粗）：setFont 加粗后撤销回到「显示不加粗」", async () => {
    // A1 未直接设置加粗，显示上也不加粗（无样式背景）。
    const result = await executePlan(
      plan([{ type: "setFont", sheet: SHEET, range: "A1", bold: true }]),
      { authorizedSheetNames: [SHEET] }
    );

    // 执行后显示加粗。
    expect(host.cellStyle(SHEET, "A1")?.fontBold).toBe(true);

    await undoExecution(result.undoSnapshot!);

    // 撤销后回到"显示不加粗"（bold: false 被显式写回，恢复显示效果）。
    expect(host.cellStyle(SHEET, "A1")?.fontBold).toBe(false);
  });

  it("原本未直接设置加粗、但样式让显示加粗：setFont 只改字号，撤销只动字号、不把加粗写成直接格式", async () => {
    // 表格样式给 A1 加粗（显示加粗），但 A1 未直接设置加粗。
    host.setDisplayedDefault(SHEET, { fontBold: true });

    const result = await executePlan(
      plan([{ type: "setFont", sheet: SHEET, range: "A1", size: 20 }]),
      { authorizedSheetNames: [SHEET] }
    );

    // 快照按**实际影响**收窄：本动作只改了字号。
    expect(result.undoSnapshot!.ranges[0].formatFields).toEqual(["fontSize"]);

    await undoExecution(result.undoSnapshot!);

    // ① 用户看到的仍是加粗（显示未被破坏）—— 但这是样式来源给的，不是撤销写上的。
    expect(host.cellDisplayedStyle(SHEET, "A1").fontBold).toBe(true);
    // ② 单元格**没有**被直接写上加粗：一旦写回，就等于把"样式来源的加粗"固化成直接格式，
    //    从此用户或表格样式再改都盖不住它。撤销恢复的是**本动作改过的属性**，不是显示效果。
    expect(host.cellStyle(SHEET, "A1")?.fontBold).toBeUndefined();
    // ③ 受影响的那一个属性（字号）回到执行前的显示值。
    expect(host.cellStyle(SHEET, "A1")?.fontSize).toBe(11);
  });

  it("来源于样式的填充：setFill 改变后撤销回到样式来源的显示色", async () => {
    // 表格样式给 A1 蓝色填充（显示蓝色），但 A1 未直接设置填充。
    host.setDisplayedDefault(SHEET, { fillColor: "#0000FF" });

    const result = await executePlan(
      plan([{ type: "setFill", sheet: SHEET, range: "A1", color: "#888888" }]),
      { authorizedSheetNames: [SHEET] }
    );

    // 执行后显示灰色。
    expect(host.cellStyle(SHEET, "A1")?.fillColor).toBe("#888888");

    await undoExecution(result.undoSnapshot!);

    // 撤销后回到样式来源的蓝色（显示值被正确捕获并写回）。
    expect(host.cellStyle(SHEET, "A1")?.fillColor).toBe("#0000FF");
    expect(host.cellStyle(SHEET, "A1")?.fillPattern).toBe("Solid");
  });

  it("明确无填充：setFill 设置颜色后撤销回到「无填充」而非白色", async () => {
    // A1 显示上无填充（无直接设置、无样式背景）。
    const result = await executePlan(
      plan([{ type: "setFill", sheet: SHEET, range: "A1", color: "#FF0000" }]),
      { authorizedSheetNames: [SHEET] }
    );

    // 执行后显示红色。
    expect(host.cellStyle(SHEET, "A1")?.fillColor).toBe("#FF0000");

    await undoExecution(result.undoSnapshot!);

    // 撤销后回到"明确无填充"（pattern None，无颜色），不是白色 #FFFFFF。
    expect(host.cellStyle(SHEET, "A1")?.fillPattern).toBe("None");
    expect(host.cellStyle(SHEET, "A1")?.fillColor).toBe("");
  });

  it("宿主只有 ExcelApi 1.18（无 getDisplayedCellProperties）：走直接格式兼容路径，不再一刀切拒绝", async () => {
    // 批量许可永久版 Office（如 ProPlus2024Volume）最高只到 ExcelApi 1.17 —— 这是真实用户
    // 会遇到的宿主。旧实现在这里直接拒绝，等于格式动作在这个宿主上完全不可用。
    // 现在改走**直接格式兼容路径**：逐格读 `getCellProperties`，撤销时先清空区域直接格式、
    // 再写回捕获到的直接格式集合。本用例证明它确实跑通且逐格还原。
    const oldHost = createOfficeTestHost({
      url: "https://contoso.sharepoint.com/format.xlsx",
      activeSheet: SHEET,
      selectedRange: "A1:B2",
      excelApiVersions: Array.from({ length: 18 }, (_, index) => `1.${index + 1}`),
      sheets: [{ name: SHEET, values: [["甲", "乙"], ["丙", "丁"]] }]
    });
    installOfficeTestHost(oldHost);
    // A1 执行前直接设置过红色填充；其余格**未直接设置**（继承/默认）。
    oldHost.setCellStyle(SHEET, "A1", { fillColor: "#FF0000" });

    const result = await executePlan(
      plan([{ type: "setFill", sheet: SHEET, range: "A1:B2", color: "#888888" }]),
      { authorizedSheetNames: [SHEET] }
    );

    // 回读：四格统一灰色（动作确实生效）。
    expect(oldHost.cellStyle(SHEET, "A1")?.fillColor).toBe("#888888");
    expect(oldHost.cellStyle(SHEET, "B1")?.fillColor).toBe("#888888");

    // 快照走的是兼容路径（有直接格式矩阵、没有显示值矩阵）。
    const saved = result.undoSnapshot!.ranges[0].format!;
    expect(saved.compatPath).toBe(true);
    expect(saved.cells).toBeNull();
    expect(saved.directCells).not.toBeNull();

    await undoExecution(result.undoSnapshot!);

    // A1 回到红色；B1 回到"**未直接设置**"（不是被写成某个默认色）——
    // 这就是兼容路径与"写回显示值"的区别：恢复的是格式的来源，不是显示效果。
    expect(oldHost.cellStyle(SHEET, "A1")?.fillColor).toBe("#FF0000");
    expect(oldHost.cellStyle(SHEET, "B1")?.fillColor).toBeUndefined();
    expect(oldHost.cellStyle(SHEET, "B1")?.fillPattern).toBeUndefined();
  });

  it("宿主连 ExcelApi 1.9 都没有：格式动作在**执行前检查**就被拒，且点名缺哪个 API、影响哪一步", async () => {
    // 1.9 是 `getCellProperties`/`setCellProperties` 的最低要求。没有它，逐格格式既读不出
    // 也写不回 —— 必须 fail-closed，而不是交一份"半份快照"。
    // 而且这条拒绝已经**前移到预检**（`preflightPlanActions`）：用户在确认前就知道原因，
    // 不必等执行期报一句笼统错误。
    const ancientHost = createOfficeTestHost({
      url: "https://contoso.sharepoint.com/format.xlsx",
      activeSheet: SHEET,
      selectedRange: "A1:B2",
      excelApiVersions: Array.from({ length: 8 }, (_, index) => `1.${index + 1}`),
      sheets: [{ name: SHEET, values: [["甲", "乙"], ["丙", "丁"]] }]
    });
    installOfficeTestHost(ancientHost);

    const failure = await executePlan(
      plan([{ type: "setFill", sheet: SHEET, range: "A1:B2", color: "#FF0000" }]),
      { authorizedSheetNames: [SHEET] }
    ).then(
      () => null,
      (error: unknown) => error as Error
    );
    expect(failure).not.toBeNull();
    const message = failure!.message;
    // 拒绝发生（不是静默写入），并且**指名**缺 1.19 / 1.9 与受影响的那一步。
    expect(message).toContain("执行前检查未通过");
    expect(message).toContain("ExcelApi 1.19");
    expect(message).toContain("ExcelApi 1.9");
    expect(message).toContain("第 1 步");
    expect(message).toContain("setFill");

    expect(ancientHost.writes()).toEqual([]);
  });
});
