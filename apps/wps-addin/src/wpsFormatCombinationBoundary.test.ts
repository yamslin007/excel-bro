import { describe, expect, it } from "vitest";
import type { AnalysisPlan, CellValue } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange, WpsWorkbook, WpsWorksheet } from "./wpsHost";
import {
  composedFormatSnapshot,
  prepareWpsPlan,
  executeWpsPlan,
  undoWpsExecution,
  latestWpsUndo
} from "./wpsPlanExecutor";
import { wpsRecoveryDescriptor } from "./wpsExecutionLedger";
import {
  stageWpsNativeComparison,
  verifyWpsNativeComparison,
  compareCommittedWpsResultToCopy,
  discardWpsNativeComparison
} from "./wpsNativeComparison";
import { describeWpsFormatAction } from "./wpsFormatCapability";
import { buildPlanChangePreview } from "./wpsPlanChangePreview";

/**
 * Task #53 组合边界测试（第 18 轮）。
 *
 * 保留**真实**执行器 / 审阅 / 回读 / 撤销逻辑与**真实** WPS 宿主替身（有状态），
 * 只用替身宿主驱动。与 `wpsFormatFlow.test.ts` 的差别：
 * 1. **行高按「整行共享状态」实现** —— 行 3 的行高是全行唯一的，用 A 列格写、
 *    用 C 列格读，读到的是同一份行高。旧替身给每个单元格独立存一个 `rowHeight`，
 *    掩盖了"行高是行级属性"的真实影响范围。
 * 2. 覆盖**部分重叠区域**（A1:B3 ∩ B2:C4）、**行高的真实影响范围**（A2:A3 → 20、
 *    C3:C4 → 25）、**外部修改**（改副本 / 改原表 / 改声明属性 / 改无关属性）、
 *    **值与公式保护**（目标区域**内部**含真实公式、文本、数字）。
 */

type Rect = { startRow: number; startColumn: number; endRow: number; endColumn: number };

type CellFormat = {
  bold: boolean;
  italic: boolean;
  name: string;
  size: number;
  color: number;
  horizontalAlignment: number;
  wrapText: boolean;
};

type CellState = CellFormat & { value: CellValue; formula: CellValue };

/** 行高是**行级**属性：整行共享一份；另外单列 cell 格式（字体/对齐/换行）。 */
type SheetState = {
  rows: CellState[][];
  /** 1-based 行号 → 数组下标 row-1。整行唯一。 */
  rowHeights: number[];
};

const DEFAULT_ROW_HEIGHT = 13.5;

function defaultCellFormat(): CellFormat {
  return {
    bold: false,
    italic: false,
    name: "宋体",
    size: 11,
    color: 0,
    horizontalAlignment: 1, // xlGeneral
    wrapText: false
  };
}

function makeCell(value: CellValue, formula: CellValue = value): CellState {
  return { ...defaultCellFormat(), value, formula };
}

function columnName(index: number): string {
  let value = index;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function parseA1(address: string): Rect {
  const [start, end = start] = address.split(":");
  const cell = (text: string) => {
    const match = /^([A-Z]+)(\d+)$/i.exec(text.trim());
    if (!match) throw new Error(`无法解析区域：${address}`);
    let column = 0;
    for (const char of match[1].toUpperCase()) column = column * 26 + (char.charCodeAt(0) - 64);
    return { column, row: Number(match[2]) };
  };
  const from = cell(start);
  const to = cell(end);
  return { startRow: from.row, startColumn: from.column, endRow: to.row, endColumn: to.column };
}

/**
 * 矩形区域替身。`Item(r, c)`（1-based）返回单格；`RowHeight` **落在行上**（整行共享），
 * `Font`/`HorizontalAlignment`/`WrapText` 落在格上。`Value2`/`Formula` 是矩阵访问面。
 */
function makeRange(state: SheetState, rect: Rect, withItem = true): WpsRange {
  const rows = rect.endRow - rect.startRow + 1;
  const columns = rect.endColumn - rect.startColumn + 1;

  const ensure = (absRow: number, absCol: number): CellState => {
    while (state.rows.length < absRow) state.rows.push([]);
    const row = state.rows[absRow - 1];
    while (row.length < absCol) row.push(makeCell(null, null));
    return row[absCol - 1];
  };
  const rowHeightOf = (absRow: number): number => state.rowHeights[absRow - 1] ?? DEFAULT_ROW_HEIGHT;
  const setRowHeight = (absRow: number, height: number): void => {
    state.rowHeights[absRow - 1] = height;
  };

  const cellProxy = (r: number, c: number): WpsRange => {
    const absRow = rect.startRow + r;
    const absCol = rect.startColumn + c;
    return {
      get Font() {
        const target = ensure(absRow, absCol);
        return {
          get Bold() { return target.bold; }, set Bold(v: unknown) { target.bold = v as boolean; },
          get Italic() { return target.italic; }, set Italic(v: unknown) { target.italic = v as boolean; },
          get Name() { return target.name; }, set Name(v: unknown) { target.name = v as string; },
          get Size() { return target.size; }, set Size(v: unknown) { target.size = v as number; },
          get Color() { return target.color; }, set Color(v: unknown) { target.color = v as number; }
        };
      },
      get HorizontalAlignment() { return ensure(absRow, absCol).horizontalAlignment; },
      set HorizontalAlignment(v: unknown) { ensure(absRow, absCol).horizontalAlignment = v as number; },
      get WrapText() { return ensure(absRow, absCol).wrapText; },
      set WrapText(v: unknown) { ensure(absRow, absCol).wrapText = v as boolean; },
      // 行高：整行共享。写某一列即写整行；任意列读到的是同一份。
      get RowHeight() { return rowHeightOf(absRow); },
      set RowHeight(v: unknown) { setRowHeight(absRow, v as number); },
      get Value2() { return ensure(absRow, absCol).value; },
      set Value2(v: unknown) { ensure(absRow, absCol).value = v as CellValue; },
      get Formula() { return ensure(absRow, absCol).formula; },
      set Formula(v: unknown) { ensure(absRow, absCol).formula = v as CellValue; }
    };
  };

  const valueMatrix = (): CellValue[][] =>
    Array.from({ length: rows }, (_, r) =>
      Array.from({ length: columns }, (_, c) => ensure(rect.startRow + r, rect.startColumn + c).value));
  const formulaMatrix = (): CellValue[][] =>
    Array.from({ length: rows }, (_, r) =>
      Array.from({ length: columns }, (_, c) => ensure(rect.startRow + r, rect.startColumn + c).formula));

  const range: WpsRange = {
    MergeCells: false,
    Rows: { Count: rows },
    Columns: { Count: columns },
    Interior: { Color: 0 },
    get Font() {
      const first = ensure(rect.startRow, rect.startColumn);
      return {
        get Bold() { return first.bold; },
        get Italic() { return first.italic; },
        get Name() { return first.name; },
        get Size() { return first.size; },
        get Color() { return first.color; }
      };
    },
    get HorizontalAlignment() { return ensure(rect.startRow, rect.startColumn).horizontalAlignment; },
    get WrapText() { return ensure(rect.startRow, rect.startColumn).wrapText; },
    get RowHeight() { return rowHeightOf(rect.startRow); }
  };
  if (withItem) {
    range.Item = (r: number, c: number) => cellProxy(r - 1, c - 1);
  }
  Object.defineProperties(range, {
    Value2: {
      get: () => {
        const matrix = valueMatrix();
        return rows === 1 && columns === 1 ? matrix[0][0] : matrix;
      },
      set: (incoming: unknown) => {
        const normalized = Array.isArray(incoming) ? (incoming as unknown[][]) : [[incoming]];
        for (let r = 0; r < rows; r += 1) {
          for (let c = 0; c < columns; c += 1) {
            const value = normalized[r]?.[c];
            if (value !== undefined) ensure(rect.startRow + r, rect.startColumn + c).value = value as CellValue;
          }
        }
      },
      enumerable: true
    },
    Formula: {
      get: () => {
        const matrix = formulaMatrix();
        return rows === 1 && columns === 1 ? matrix[0][0] : matrix;
      },
      set: (incoming: unknown) => {
        const normalized = Array.isArray(incoming) ? (incoming as unknown[][]) : [[incoming]];
        for (let r = 0; r < rows; r += 1) {
          for (let c = 0; c < columns; c += 1) {
            const value = normalized[r]?.[c];
            if (value !== undefined) ensure(rect.startRow + r, rect.startColumn + c).formula = value as CellValue;
          }
        }
      },
      enumerable: true
    }
  });
  return range;
}

type CombinationHost = {
  application: WpsApplication;
  originalSheet: WpsWorksheet;
  /** 直接读替身状态里的格（字体/对齐/换行/值/公式）。 */
  cell: (row: number, column: number) => CellState;
  /** 直接读替身状态里的行高（整行唯一）。 */
  rowHeight: (row: number) => number;
  /** 通过宿主 Range 面读某一行某一列的行高（验证整行语义：任意列读到同一份）。 */
  readRowHeightVia: (row: number, column: number) => number;
  /** 通过宿主 Range 面写某一行某一列的行高（模拟用户手工改行高）。 */
  setRowHeightVia: (row: number, column: number, height: number) => void;
  copyClosed: () => boolean;
};

/**
 * 初始数据：4 行 × 3 列。
 * - A 列文本、B 列数字、C 列**真公式**（值/公式分离）。
 * - 每格字号不同（10..21），用于"逐格恢复原始字号"。
 * - 每行行高不同（15/18/22/30），用于"整行恢复原始行高"。
 */
const INITIAL_SIZES = [
  [10, 11, 12],
  [13, 14, 15],
  [16, 17, 18],
  [19, 20, 21]
];
const INITIAL_ROW_HEIGHTS = [15, 18, 22, 30];

function combinationHost(): CombinationHost {
  const rows: CellState[][] = [
    [makeCell("编号"), makeCell("名称"), makeCell("备注")],
    [makeCell("甲"), makeCell(100), makeCell("100元", "=B2&\"元\"")],
    [makeCell("乙"), makeCell(200), makeCell("200元", "=B3&\"元\"")],
    [makeCell("丙"), makeCell(300), makeCell("300元", "=B4&\"元\"")]
  ];
  rows.forEach((row, r) => row.forEach((cell, c) => { cell.size = INITIAL_SIZES[r][c]; }));
  const state: SheetState = { rows, rowHeights: [...INITIAL_ROW_HEIGHTS] };

  let active: WpsWorkbook;
  let copyClosed = false;

  const sheetCache = new Map<string, WpsWorksheet>();
  function buildSheet(name: string, sheetState: SheetState): WpsWorksheet {
    return {
      Name: name,
      ProtectContents: false,
      Range: (address: unknown) => makeRange(sheetState, parseA1(String(address)))
    };
  }
  function sheet(name: string, sheetState: SheetState): WpsWorksheet {
    const cached = sheetCache.get(name);
    if (cached) return cached;
    const built = buildSheet(name, sheetState);
    sheetCache.set(name, built);
    return built;
  }

  const originalSheet = sheet("数据", state);
  const original: WpsWorkbook = {
    Name: "格式数据.xlsx",
    FullName: "C:\\Data\\格式数据.xlsx",
    Path: "C:\\Data",
    ReadOnly: false,
    Worksheets: {
      get Count() { return 1; },
      Item: (key: number | string) => (key === 1 || key === "数据" ? originalSheet : undefined)
    } as unknown as WpsWorkbook["Worksheets"],
    ActiveSheet: originalSheet,
    Activate: () => { active = original; }
  };
  (originalSheet as unknown as { Copy: () => void }).Copy = () => {
    // 副本深拷贝：值与公式逐格拷，**行高整行拷**（行级属性）。
    const copyState: SheetState = {
      rows: state.rows.map((row) => row.map((cell) => ({ ...cell }))),
      rowHeights: [...state.rowHeights]
    };
    const copySheet = buildSheet("数据", copyState);
    const copy: WpsWorkbook = {
      Name: "Book1",
      FullName: "",
      Path: "",
      ReadOnly: false,
      Worksheets: {
        get Count() { return 1; },
        Item: () => copySheet
      } as unknown as WpsWorkbook["Worksheets"],
      ActiveSheet: copySheet,
      Windows: { Count: 1, Item: () => ({ Caption: "Book1" }) },
      Activate: () => { active = copy; },
      Close: (saveChanges?: boolean) => {
        if (saveChanges !== false) throw new Error("副本不得保存");
        copyClosed = true;
        if (active === copy) active = original;
        return true;
      }
    };
    active = copy;
  };
  active = original;

  const storage = new Map<string, unknown>();
  const application: WpsApplication = {
    Version: "12.0",
    get ActiveWorkbook() { return active; },
    get ActiveSheet() { return active.ActiveSheet!; },
    Windows: {
      Count: 2,
      Item: () => ({ Caption: "Book1" }),
      CompareSideBySideWith: () => true,
      SyncScrollingSideBySide: false
    },
    PluginStorage: {
      getItem: (key) => storage.get(key),
      setItem: (key, value) => { storage.set(key, value); }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  };

  return {
    application,
    originalSheet,
    cell: (row, column) => state.rows[row - 1][column - 1],
    rowHeight: (row) => state.rowHeights[row - 1],
    readRowHeightVia: (row, column) =>
      originalSheet.Range!(`${columnName(column)}${row}`).Item!(1, 1).RowHeight as number,
    setRowHeightVia: (row, column, height) => {
      originalSheet.Range!(`${columnName(column)}${row}`).Item!(1, 1).RowHeight = height;
    },
    copyClosed: () => copyClosed
  };
}

function planOf(actions: AnalysisPlan["actions"]): AnalysisPlan {
  return {
    id: "plan-format-combination",
    title: "格式组合",
    summary: "格式组合动作",
    assumptions: [],
    warnings: [],
    actions,
    acceptanceCriteria: []
  };
}

/* ------------------------------------------------------------------------- *
 * 1. 部分重叠区域
 * ------------------------------------------------------------------------- */

describe("格式动作组合边界 · 部分重叠区域", () => {
  it("A1:B3 字号14 → B2:C4 字号16：重叠区 B2:B3 最终为 16，其余按动作顺序，撤销逐格恢复原始字号", async () => {
    const host = combinationHost();
    const prepared = prepareWpsPlan(host.application, [1], planOf([
      { type: "setFont", sheet: "数据", range: "A1:B3", size: 14 },
      { type: "setFont", sheet: "数据", range: "B2:C4", size: 16 }
    ]));

    const staged = stageWpsNativeComparison(host.application, prepared);
    // 确认前原表零写入：字号仍是初始值。
    expect(host.cell(1, 1).size).toBe(10);
    expect(host.cell(3, 2).size).toBe(17);
    expect(() => verifyWpsNativeComparison(host.application, staged)).not.toThrow();

    const record = await executeWpsPlan(host.application, prepared);
    expect(record.state).toBe("verified");

    // A 列：只在动作1范围内 → 14。
    expect(host.cell(1, 1).size).toBe(14);
    expect(host.cell(2, 1).size).toBe(14);
    expect(host.cell(3, 1).size).toBe(14);
    // B1：只在动作1范围内 → 14。
    expect(host.cell(1, 2).size).toBe(14);
    // 重叠区 B2:B3：动作2 覆盖 → 16。
    expect(host.cell(2, 2).size).toBe(16);
    expect(host.cell(3, 2).size).toBe(16);
    // B4、C2:C4：只在动作2范围内 → 16。
    expect(host.cell(4, 2).size).toBe(16);
    expect(host.cell(2, 3).size).toBe(16);
    expect(host.cell(3, 3).size).toBe(16);
    expect(host.cell(4, 3).size).toBe(16);
    // C1：两个动作都不覆盖 → 保持原始 12。
    expect(host.cell(1, 3).size).toBe(12);

    // 提交后对照按整份计划最终状态。
    expect(compareCommittedWpsResultToCopy(staged)).toBe("matched");

    // 撤销：逐格恢复原始字号（14 → 10..17 各自的原值）。
    await undoWpsExecution(host.application, record.executionId);
    for (let r = 1; r <= 4; r += 1) {
      for (let c = 1; c <= 3; c += 1) {
        expect(host.cell(r, c).size).toBe(INITIAL_SIZES[r - 1][c - 1]);
      }
    }
  });
});

/* ------------------------------------------------------------------------- *
 * 2. 行高的真实影响范围（整行共享）
 * ------------------------------------------------------------------------- */

describe("格式动作组合边界 · 行高的真实影响范围", () => {
  it("预览明确提示影响整行", () => {
    expect(describeWpsFormatAction({ type: "resizeRange", rowHeight: 20 })).toContain("影响整行");
    const host = combinationHost();
    const prepared = prepareWpsPlan(host.application, [1], planOf([
      { type: "resizeRange", sheet: "数据", range: "A2:A3", rowHeight: 20 }
    ]));
    const preview = buildPlanChangePreview(prepared);
    expect(preview.actions[0].formatDescription ?? "").toContain("影响整行");
    expect(preview.actions[0].rows).toEqual([]); // 格式动作不走值矩阵对照
  });

  it("A2:A3 行高20 → C3:C4 行高25：最终第2行20、第3/4行25；A 列与 C 列读到同一份行高；撤销恢复原始行高", async () => {
    const host = combinationHost();
    const prepared = prepareWpsPlan(host.application, [1], planOf([
      { type: "resizeRange", sheet: "数据", range: "A2:A3", rowHeight: 20 },
      { type: "resizeRange", sheet: "数据", range: "C3:C4", rowHeight: 25 }
    ]));

    const staged = stageWpsNativeComparison(host.application, prepared);
    // 审阅指纹遵循整行语义：副本里用 A 列 / C 列读同一行，读到同一份行高。
    const copySheet = staged.copy.Worksheets.Item(1);
    expect(copySheet.Range!("A2").Item!(1, 1).RowHeight).toBe(20);
    expect(copySheet.Range!("C2").Item!(1, 1).RowHeight).toBe(20);
    expect(copySheet.Range!("A3").Item!(1, 1).RowHeight).toBe(25);
    expect(copySheet.Range!("C3").Item!(1, 1).RowHeight).toBe(25);
    expect(() => verifyWpsNativeComparison(host.application, staged)).not.toThrow();

    const record = await executeWpsPlan(host.application, prepared);
    expect(record.state).toBe("verified");

    // 整行语义：第 2 行 = 20（仅动作1）、第 3 行 = 25（动作2 覆盖动作1）、第 4 行 = 25。
    expect(host.rowHeight(1)).toBe(15); // 未声明行不变
    expect(host.rowHeight(2)).toBe(20);
    expect(host.rowHeight(3)).toBe(25);
    expect(host.rowHeight(4)).toBe(25);
    // 通过不同列读同一行，确认"整行共享"。
    expect(host.readRowHeightVia(3, 1)).toBe(25); // A3
    expect(host.readRowHeightVia(3, 2)).toBe(25); // B3（不在任何动作范围，但同行）
    expect(host.readRowHeightVia(3, 3)).toBe(25); // C3
    expect(compareCommittedWpsResultToCopy(staged)).toBe("matched");

    // 撤销：逆序恢复，整行回到原始行高。
    await undoWpsExecution(host.application, record.executionId);
    for (let r = 1; r <= 4; r += 1) {
      expect(host.rowHeight(r)).toBe(INITIAL_ROW_HEIGHTS[r - 1]);
    }
  });

  it("执行后外部改同一行（即便改的是无关列）= 撤销冲突（整行语义）", async () => {
    const host = combinationHost();
    const prepared = prepareWpsPlan(host.application, [1], planOf([
      { type: "resizeRange", sheet: "数据", range: "A2:A3", rowHeight: 20 },
      { type: "resizeRange", sheet: "数据", range: "C3:C4", rowHeight: 25 }
    ]));
    const record = await executeWpsPlan(host.application, prepared);
    // 用户把第 3 行行高改回 40 —— 通过 B3（既不在 A2:A3、也不在 C3:C4）。
    // 行高是整行属性，B3 属于第 3 行，所以这仍是对声明行的外部改动。
    host.setRowHeightVia(3, 2, 40);
    await expect(undoWpsExecution(host.application, record.executionId)).rejects.toThrow();
    expect(latestWpsUndo(host.application)?.state).toBe("undo_conflict");
  });

  it("执行后外部改无关行（非声明行）→ 撤销成功且保留该行改动，只恢复声明行", async () => {
    const host = combinationHost();
    const prepared = prepareWpsPlan(host.application, [1], planOf([
      { type: "resizeRange", sheet: "数据", range: "A2:A3", rowHeight: 20 },
      { type: "resizeRange", sheet: "数据", range: "C3:C4", rowHeight: 25 }
    ]));
    const record = await executeWpsPlan(host.application, prepared);
    // 第 1 行不在任何动作范围内 —— 用户改动是"无关行"，不该触发冲突、也不该被撤销恢复。
    host.setRowHeightVia(1, 1, 42);
    await undoWpsExecution(host.application, record.executionId);
    expect(host.rowHeight(1)).toBe(42); // 用户对无关行的改动保留
    expect(host.rowHeight(2)).toBe(INITIAL_ROW_HEIGHTS[1]); // 声明行恢复原值
    expect(host.rowHeight(3)).toBe(INITIAL_ROW_HEIGHTS[2]);
    expect(host.rowHeight(4)).toBe(INITIAL_ROW_HEIGHTS[3]);
  });

  it("合成写前现场按**整行**范围叠加：A 列与 C 列不相交，但第 3 行确实被前序动作改过", async () => {
    const host = combinationHost();
    const prepared = prepareWpsPlan(host.application, [1], planOf([
      { type: "resizeRange", sheet: "数据", range: "A2:A3", rowHeight: 20 },
      { type: "resizeRange", sheet: "数据", range: "C3:C4", rowHeight: 25 }
    ]));
    // 第 2 行只属于动作 1、第 4 行只属于动作 2、第 3 行两个动作都有份。
    expect(composedFormatSnapshot(prepared.actions, 1)?.rowHeights).toEqual([
      20,
      INITIAL_ROW_HEIGHTS[3]
    ]);
    // 第一个动作前面没有别的动作 —— 合成结果就是预览快照本身。
    expect(composedFormatSnapshot(prepared.actions, 0)).toEqual(prepared.actions[0].formatSnapshot);

    // 登记进账本的写前/写后指纹，必须逐字等于执行器真正发出去的检查点 ——
    // 否则服务端会以"证据与登记的写前状态不一致"拒掉写前检查点。
    const descriptor = wpsRecoveryDescriptor(prepared);
    const seen: Array<{ phase: string; actionIndex: number; fingerprint: string | null }> = [];
    await executeWpsPlan(host.application, prepared, undefined, undefined, async (phase, checkpoint) => {
      seen.push({
        phase,
        actionIndex: checkpoint.actionIndex,
        fingerprint: phase === "before" ? checkpoint.beforeFingerprint : checkpoint.afterFingerprint
      });
    });
    for (const [index, action] of descriptor.actions.entries()) {
      expect(action.formatFields).toEqual(["rowHeight"]);
      expect(
        seen.find((item) => item.phase === "before" && item.actionIndex === index)?.fingerprint
      ).toBe(action.previewFingerprint);
      expect(
        seen.find((item) => item.phase === "after" && item.actionIndex === index)?.fingerprint
      ).toBe(action.expectedFingerprint);
    }
    // 两个行高动作的写前现场不同（第 3 行被前面的动作改过）。
    expect(descriptor.actions[0].previewFingerprint).not.toBe(
      descriptor.actions[1].previewFingerprint
    );
  });
});

/* ------------------------------------------------------------------------- *
 * 3. 外部修改
 * ------------------------------------------------------------------------- */

describe("格式动作组合边界 · 外部修改", () => {
  const overlapPlan = (): AnalysisPlan["actions"] => ([
    { type: "setFont", sheet: "数据", range: "A1:B3", size: 14 },
    { type: "setFont", sheet: "数据", range: "B2:C4", size: 16 }
  ]);

  it("审阅后真正改动副本的声明属性 → 拒绝写回", () => {
    const host = combinationHost();
    const prepared = prepareWpsPlan(host.application, [1], planOf(overlapPlan()));
    const staged = stageWpsNativeComparison(host.application, prepared);
    // 用户在副本重叠区把 B2 的字号又改了。
    staged.copy.Worksheets.Item(1).Range!("B2").Item!(1, 1).Font!.Size = 99;
    expect(() => verifyWpsNativeComparison(host.application, staged)).toThrow();
  });

  it("审阅后真正改动原表的声明属性 → 拒绝写回", () => {
    const host = combinationHost();
    const prepared = prepareWpsPlan(host.application, [1], planOf(overlapPlan()));
    const staged = stageWpsNativeComparison(host.application, prepared);
    host.cell(1, 1).size = 99; // 原表 A1（在动作1范围内）被外部改动
    expect(() => verifyWpsNativeComparison(host.application, staged)).toThrow();
  });

  it("执行后修改本次涉及的属性 → 识别撤销冲突", async () => {
    const host = combinationHost();
    const prepared = prepareWpsPlan(host.application, [1], planOf(overlapPlan()));
    const record = await executeWpsPlan(host.application, prepared);
    host.cell(1, 1).size = 99; // A1 属于动作1声明范围（size）
    await expect(undoWpsExecution(host.application, record.executionId)).rejects.toThrow();
    expect(latestWpsUndo(host.application)?.state).toBe("undo_conflict");
  });

  it("执行后修改无关属性 → 撤销保留这些修改", async () => {
    const host = combinationHost();
    const prepared = prepareWpsPlan(host.application, [1], planOf(overlapPlan()));
    const record = await executeWpsPlan(host.application, prepared);
    // 改的是 color / wrapText（本次动作都没声明），不是 size。
    host.cell(1, 1).color = 0xff0000;
    host.cell(1, 1).wrapText = true;
    await undoWpsExecution(host.application, record.executionId);
    // 声明的 size 被恢复为原始值。
    expect(host.cell(1, 1).size).toBe(10);
    // 无关属性保留。
    expect(host.cell(1, 1).color).toBe(0xff0000);
    expect(host.cell(1, 1).wrapText).toBe(true);
  });
});

/* ------------------------------------------------------------------------- *
 * 4. 值与公式保护（目标区域内部含公式/文本/数字）
 * ------------------------------------------------------------------------- */

describe("格式动作组合边界 · 值与公式保护", () => {
  it("格式目标区域 A2:C2 内部含文本/数字/真公式：确认前原表零写入，执行与撤销后值与公式保持不变", async () => {
    const host = combinationHost();
    const prepared = prepareWpsPlan(host.application, [1], planOf([
      { type: "setFont", sheet: "数据", range: "A2:C2", bold: true, size: 18 }
    ]));
    // 目标区域**内部**就有真公式（C2）。
    expect(host.cell(2, 3).formula).toBe("=B2&\"元\"");

    const staged = stageWpsNativeComparison(host.application, prepared);
    // 确认前原表零写入：值/公式/字号全部保持初始。
    expect(host.cell(2, 1).value).toBe("甲");
    expect(host.cell(2, 2).value).toBe(100);
    expect(host.cell(2, 3).value).toBe("100元");
    expect(host.cell(2, 3).formula).toBe("=B2&\"元\"");
    expect(host.cell(2, 1).bold).toBe(false);
    expect(host.cell(2, 1).size).toBe(13);

    expect(() => verifyWpsNativeComparison(host.application, staged)).not.toThrow();
    const record = await executeWpsPlan(host.application, prepared);
    expect(record.state).toBe("verified");

    // 执行后：声明的格式生效，值与公式一字不变。
    expect(host.cell(2, 1).bold).toBe(true);
    expect(host.cell(2, 1).size).toBe(18);
    expect(host.cell(2, 3).size).toBe(18);
    expect(host.cell(2, 1).value).toBe("甲");
    expect(host.cell(2, 2).value).toBe(100);
    expect(host.cell(2, 3).value).toBe("100元");
    expect(host.cell(2, 3).formula).toBe("=B2&\"元\"");
    expect(compareCommittedWpsResultToCopy(staged)).toBe("matched");

    // 撤销后：格式回到初始，值与公式仍一字不变。
    await undoWpsExecution(host.application, record.executionId);
    expect(host.cell(2, 1).bold).toBe(false);
    expect(host.cell(2, 1).size).toBe(13);
    expect(host.cell(2, 1).value).toBe("甲");
    expect(host.cell(2, 2).value).toBe(100);
    expect(host.cell(2, 3).value).toBe("100元");
    expect(host.cell(2, 3).formula).toBe("=B2&\"元\"");
  });

  it("放弃副本时不碰原表（值/公式/格式全不变）", () => {
    const host = combinationHost();
    const prepared = prepareWpsPlan(host.application, [1], planOf([
      { type: "setFont", sheet: "数据", range: "A2:C2", size: 18 }
    ]));
    const staged = stageWpsNativeComparison(host.application, prepared);
    discardWpsNativeComparison(staged);
    expect(host.copyClosed()).toBe(true);
    expect(host.cell(2, 1).size).toBe(13);
    expect(host.cell(2, 3).formula).toBe("=B2&\"元\"");
  });
});
