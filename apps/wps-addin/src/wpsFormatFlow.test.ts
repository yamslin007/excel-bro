import { describe, expect, it } from "vitest";
import type { AnalysisPlan, CellValue } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange, WpsWorkbook, WpsWorksheet } from "./wpsHost";
import { prepareWpsPlan, executeWpsPlan, undoWpsExecution } from "./wpsPlanExecutor";
import {
  stageWpsNativeComparison,
  verifyWpsNativeComparison,
  compareCommittedWpsResultToCopy,
  discardWpsNativeComparison
} from "./wpsNativeComparison";

/**
 * 带**值 + 公式 + 逐格格式**的完整替身宿主（Task #53 格式动作面板闭环）。
 *
 * 与 `wpsFormatExecutor.test.ts` 的 `formatHost` 不同：这里必须同时支持
 * - `Value2` / `Formula` 矩阵（`snapshot()` / 值动作 / 副本对照用）
 * - `Item(r, c)` 返回带 Font/对齐/换行/行高的单格（格式动作逐格采集/写回用）
 *
 * 这样 `stageWpsNativeComparison`（真实审阅逻辑）才能对一个格式动作计划走完
 * 副本审阅 → 确认 → 写回 → 回读 → 撤销的完整链路，而不是被"格式动作没有值矩阵"
 * 挡住。
 */

type FormatState = {
  bold: boolean;
  italic: boolean;
  name: string;
  size: number;
  color: number;
  horizontalAlignment: number;
  wrapText: boolean;
  rowHeight: number;
};

type CellState = FormatState & { value: CellValue; formula: CellValue };

function defaultFormat(): FormatState {
  return {
    bold: false,
    italic: false,
    name: "宋体",
    size: 11,
    color: 0,
    horizontalAlignment: 1, // xlGeneral
    wrapText: false,
    rowHeight: 13.5
  };
}

function parseA1(address: string): { startRow: number; startColumn: number; endRow: number; endColumn: number } {
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
 * 一个"工作表状态"：二维格状态矩阵。值/公式/格式都按 [row-1][column-1] 存。
 */
function makeSheetState(initial: (CellValue | undefined)[][]) {
  const rows = initial.map((row) =>
    row.map((value) => ({ ...defaultFormat(), value: value ?? null, formula: value ?? null } as CellState))
  );
  const maxCols = Math.max(0, ...rows.map((row) => row.length));
  // 保证矩阵是矩形，便于 `slice()`。
  for (const row of rows) {
    while (row.length < maxCols) row.push({ ...defaultFormat(), value: null, formula: null });
  }
  return rows;
}

type SheetState = CellState[][];

/**
 * 构造一个矩形区域替身。`range.Rows.Count`/`Columns.Count` 给尺寸，
 * `Item(r, c)` 返回单格（带格式），`Value2`/`Formula` 是矩阵 getter/setter。
 *
 * `withItem`：false 时**不提供** `Item`（模拟宿主缺逐格定位能力），
 * 用于复现"多格缺 Item 却回退到整块"的缺陷。
 */
function makeRange(
  state: SheetState,
  rect: { startRow: number; startColumn: number; endRow: number; endColumn: number },
  withItem = true
): WpsRange {
  const rows = rect.endRow - rect.startRow + 1;
  const columns = rect.endColumn - rect.startColumn + 1;
  const slice = (): CellValue[][] =>
    Array.from({ length: rows }, (_, r) =>
      Array.from({ length: columns }, (_, c) => {
        const cell = state[rect.startRow - 1 + r]?.[rect.startColumn - 1 + c];
        return cell?.value ?? null;
      })
    );
  const formulaSlice = (): CellValue[][] =>
    Array.from({ length: rows }, (_, r) =>
      Array.from({ length: columns }, (_, c) => {
        const cell = state[rect.startRow - 1 + r]?.[rect.startColumn - 1 + c];
        return cell?.formula ?? null;
      })
    );

  // 单格格式访问器：共享 state 里的某一行某一列。
  const cellProxy = (r: number, c: number): WpsRange => {
    const key = () => {
      const cell = state[rect.startRow - 1 + r]?.[rect.startColumn - 1 + c];
      if (!cell) {
        // 越界：扩容。
        while (state.length <= rect.startRow - 1 + r) {
          const newRow = Array.from({ length: Math.max(columns, state[0]?.length ?? 1) }, () => ({
            ...defaultFormat(), value: null, formula: null
          }));
          state.push(newRow);
        }
        const width = state[0]?.length ?? columns;
        while (state[rect.startRow - 1 + r].length <= rect.startColumn - 1 + c) {
          state[rect.startRow - 1 + r].push({ ...defaultFormat(), value: null, formula: null });
        }
        void width;
      }
      return state[rect.startRow - 1 + r][rect.startColumn - 1 + c];
    };
    return {
      get Font() {
        const cell = key();
        return {
          get Bold() { return cell.bold; }, set Bold(v: unknown) { cell.bold = v as boolean; },
          get Italic() { return cell.italic; }, set Italic(v: unknown) { cell.italic = v as boolean; },
          get Name() { return cell.name; }, set Name(v: unknown) { cell.name = v as string; },
          get Size() { return cell.size; }, set Size(v: unknown) { cell.size = v as number; },
          get Color() { return cell.color; }, set Color(v: unknown) { cell.color = v as number; }
        };
      },
      get HorizontalAlignment() { return key().horizontalAlignment; },
      set HorizontalAlignment(v: unknown) { key().horizontalAlignment = v as number; },
      get WrapText() { return key().wrapText; },
      set WrapText(v: unknown) { key().wrapText = v as boolean; },
      get RowHeight() { return key().rowHeight; },
      set RowHeight(v: unknown) { key().rowHeight = v as number; },
      get Value2() { return key().value; },
      set Value2(v: unknown) { key().value = v as CellValue; },
      get Formula() { return key().formula; },
      set Formula(v: unknown) { key().formula = v as CellValue; }
    };
  };

  // 区域级格式：真 WPS 里 `Range.Font.Bold` 对混合格式返回不确定值，但 `Range`
  // **确实有** `Font` 属性。这里用"首格"模拟区域级读取 —— 若执行器在缺 `Item`
  // 时回退到整块，就会读到首格格式并**伪造**出"逐格已采集"的快照。
  const range: WpsRange = {
    MergeCells: false,
    Rows: { Count: rows },
    Columns: { Count: columns },
    Interior: { Color: 0 },
    get Font() {
      const first = state[rect.startRow - 1]?.[rect.startColumn - 1];
      return {
        get Bold() { return first?.bold ?? false; },
        get Italic() { return first?.italic ?? false; },
        get Name() { return first?.name ?? "宋体"; },
        get Size() { return first?.size ?? 11; },
        get Color() { return first?.color ?? 0; }
      };
    },
    get HorizontalAlignment() { return state[rect.startRow - 1]?.[rect.startColumn - 1]?.horizontalAlignment ?? 1; },
    get WrapText() { return state[rect.startRow - 1]?.[rect.startColumn - 1]?.wrapText ?? false; },
    get RowHeight() { return state[rect.startRow - 1]?.[rect.startColumn - 1]?.rowHeight ?? 13.5; }
  };
  if (withItem) {
    range.Item = (r: number, c: number) => cellProxy(r - 1, c - 1);
  }
  Object.defineProperties(range, {
    Value2: {
      get: () => {
        const current = slice();
        return rows === 1 && columns === 1 ? current[0][0] : current;
      },
      set: (incoming: unknown) => {
        const normalized = Array.isArray(incoming) ? (incoming as unknown[][]) : [[incoming]];
        for (let r = 0; r < rows; r += 1) {
          for (let c = 0; c < columns; c += 1) {
            const value = normalized[r]?.[c];
            if (value !== undefined) {
              const cell = cellProxy(r, c);
              cell.Value2 = value;
            }
          }
        }
      },
      enumerable: true
    },
    Formula: {
      get: () => {
        const current = formulaSlice();
        return rows === 1 && columns === 1 ? current[0][0] : current;
      },
      set: (incoming: unknown) => {
        const normalized = Array.isArray(incoming) ? (incoming as unknown[][]) : [[incoming]];
        for (let r = 0; r < rows; r += 1) {
          for (let c = 0; c < columns; c += 1) {
            const value = normalized[r]?.[c];
            if (value !== undefined) {
              const cell = cellProxy(r, c);
              cell.Formula = value;
            }
          }
        }
      },
      enumerable: true
    }
  });
  return range;
}

/**
 * 完整宿主：一张写入目标表，值 + 公式 + 逐格格式 + 原生副本审阅能力。
 */
function formatFlowHost(options: { withItem?: boolean } = {}) {
  const withItem = options.withItem ?? true;
  const initialState = makeSheetState([
    ["姓名", "金额", "备注"],
    ["甲", 100, "=A2&\"元\""],
    ["乙", 200, "=A3&\"元\""]
  ]);
  let active: WpsWorkbook;
  let copyClosed = false;
  let copyState: SheetState | null = null;

  const sheetCache = new Map<string, WpsWorksheet>();
  function buildSheet(name: string, state: SheetState): WpsWorksheet {
    const sheet: WpsWorksheet = {
      Name: name,
      ProtectContents: false,
      Range: (address: unknown) => makeRange(state, parseA1(String(address)), withItem)
    };
    return sheet;
  }
  function sheet(name: string, state: SheetState): WpsWorksheet {
    const cached = sheetCache.get(name);
    if (cached) return cached;
    const built = buildSheet(name, state);
    sheetCache.set(name, built);
    return built;
  }

  const originalSheet = sheet("数据", initialState);
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
  // Copy：把原表状态深拷贝到副本（值 + 公式 + 格式全带走）。
  (originalSheet as unknown as { Copy: () => void }).Copy = () => {
    copyState = initialState.map((row) => row.map((cell) => ({ ...cell })));
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
    state: initialState,
    cell: (row: number, column: number) => initialState[row - 1][column - 1],
    copyClosed: () => copyClosed
  };
}

function formatPlan(actions: AnalysisPlan["actions"]): AnalysisPlan {
  return {
    id: "plan-format-flow",
    title: "设置格式",
    summary: "加粗并调整对齐",
    assumptions: [],
    warnings: [],
    actions,
    acceptanceCriteria: []
  };
}

describe("WPS 格式动作面板闭环（副本审阅 → 确认 → 写回 → 回读 → 撤销）", () => {
  it("阻断点 1：格式动作走副本审阅，值/公式不变，副本只改格式，写回后逐格回读、撤销恢复", async () => {
    const host = formatFlowHost();
    const prepared = prepareWpsPlan(host.application, [1], formatPlan([
      { type: "setFont", sheet: "数据", range: "A2:B3", bold: true }
    ]));
    // 审阅副本：原表零写入、副本只改格式不改值/公式。
    const staged = stageWpsNativeComparison(host.application, prepared);
    expect(host.cell(2, 1).bold).toBe(false); // 原表未加粗
    expect(host.cell(2, 1).value).toBe("甲");  // 原表值未变
    expect(host.cell(2, 3).formula).toBe("=A2&\"元\""); // 原表公式未变

    verifyWpsNativeComparison(host.application, staged);
    // 写回原表。
    const record = await executeWpsPlan(host.application, prepared);
    expect(record.state).toBe("verified");
    expect(host.cell(2, 1).bold).toBe(true);   // 加粗生效
    expect(host.cell(2, 1).value).toBe("甲");  // 值不变
    expect(host.cell(2, 3).formula).toBe("=A2&\"元\""); // 公式不变
    // 与副本对照一致。
    expect(compareCommittedWpsResultToCopy(staged)).toBe("matched");

    // 撤销：逐格回到不加粗，值/公式不变。
    await undoWpsExecution(host.application, record.executionId);
    expect(host.cell(2, 1).bold).toBe(false);
    expect(host.cell(2, 1).value).toBe("甲");
    expect(host.cell(2, 3).formula).toBe("=A2&\"元\"");
  });

  it("阻断点 2：加粗后用户改字色/行高，再撤销只恢复加粗、不覆盖用户后续修改", async () => {
    const host = formatFlowHost();
    const prepared = prepareWpsPlan(host.application, [1], formatPlan([
      { type: "setFont", sheet: "数据", range: "A2:B3", bold: true }
    ]));
    const record = await executeWpsPlan(host.application, prepared);
    expect(host.cell(2, 1).bold).toBe(true);
    // 执行后用户手动改了字色和行高（加粗仍保留，属于"对其他属性的后续修改"）。
    host.cell(2, 1).color = 0x0000ff;
    host.cell(2, 1).rowHeight = 30;
    // 撤销：属性级恢复只恢复 bold，不覆盖用户改的字色/行高。
    await undoWpsExecution(host.application, record.executionId);
    expect(host.cell(2, 1).bold).toBe(false);  // 加粗已恢复
    expect(host.cell(2, 1).color).toBe(0x0000ff);  // 用户的字色保留
    expect(host.cell(2, 1).rowHeight).toBe(30);    // 用户的行高保留
  });

  it("阻断点 3：多格区域缺少 Item 能力时拒绝且零写入", async () => {
    const host = formatFlowHost({ withItem: false });
    expect(() => prepareWpsPlan(host.application, [1], formatPlan([
      { type: "setFont", sheet: "数据", range: "A2:B3", bold: true }
    ]))).toThrow();
    // 零写入：所有格都保持不加粗。
    expect(host.cell(2, 1).bold).toBe(false);
    expect(host.cell(3, 2).bold).toBe(false);
  });

  it("同一区域多个格式动作：逐个执行、组合回读、逐个撤销（不误报冲突）", async () => {
    const host = formatFlowHost();
    // 同区域 setFont(bold) + setAlignment(center/wrapText)：属性级恢复各恢复各的字段。
    const prepared = prepareWpsPlan(host.application, [1], formatPlan([
      { type: "setFont", sheet: "数据", range: "A2:B3", bold: true },
      { type: "setAlignment", sheet: "数据", range: "A2:B3", horizontal: "center", wrapText: true }
    ]));
    const record = await executeWpsPlan(host.application, prepared);
    expect(record.state).toBe("verified");
    // 两个动作的效果都生效。
    expect(host.cell(2, 1).bold).toBe(true);
    expect(host.cell(2, 1).horizontalAlignment).toBe(-4108); // xlCenter
    expect(host.cell(2, 1).wrapText).toBe(true);
    // 撤销：两个字段都回到写前。
    await undoWpsExecution(host.application, record.executionId);
    expect(host.cell(2, 1).bold).toBe(false);
    expect(host.cell(2, 1).horizontalAlignment).toBe(1); // xlGeneral
    expect(host.cell(2, 1).wrapText).toBe(false);
  });

  it("格式动作副本审阅后原表零写入，放弃副本不碰原表", () => {
    const host = formatFlowHost();
    const prepared = prepareWpsPlan(host.application, [1], formatPlan([
      { type: "setFont", sheet: "数据", range: "A2:B3", bold: true }
    ]));
    const staged = stageWpsNativeComparison(host.application, prepared);
    expect(host.cell(2, 1).bold).toBe(false);
    discardWpsNativeComparison(staged);
    expect(host.copyClosed()).toBe(true);
    expect(host.cell(2, 1).bold).toBe(false);
  });
});

/**
 * Task #53 收口：组合状态 —— 确认/提交后对照/撤销都要按**整份计划的最终状态**
 * 判断，不能拿"逐动作中间快照"当计划产出。
 *
 * 两条独立的组合错误：
 * 1. 副本审阅逐动作存中间格式快照，确认与提交后对照却拿最终状态比中间快照：
 *    同区域"加粗→居中"会在确认阶段被误报成"副本在生成对照后发生变化"。
 * 2. 撤销前逐动作要求"本动作声明的目标值仍成立"：同属性后续覆盖（字号先 12/13→14
 *    再 14→16）时，第一步声明的 14 早已被第二步覆盖成 16，按目标值判会误报冲突。
 *    应按**最终恢复影响范围**（本次动作声明过的属性，现场是否仍等于计划产出的最终
 *    格式）检查冲突，再**逆序**恢复。
 */
describe("WPS 格式动作组合状态（Task #53 收口：中间态 vs 最终状态）", () => {
  it("同区域「加粗→居中」完整走 stage→verify→execute→compare→undo，不被逐动作中间快照误报", async () => {
    const host = formatFlowHost();
    const prepared = prepareWpsPlan(host.application, [1], formatPlan([
      { type: "setFont", sheet: "数据", range: "A2:B3", bold: true },
      { type: "setAlignment", sheet: "数据", range: "A2:B3", horizontal: "center" }
    ]));
    const staged = stageWpsNativeComparison(host.application, prepared);
    // 确认阶段比的是"整份计划最终状态"，不是逐动作中间快照 —— 不得误报。
    expect(() => verifyWpsNativeComparison(host.application, staged)).not.toThrow();
    const record = await executeWpsPlan(host.application, prepared);
    expect(record.state).toBe("verified");
    expect(host.cell(2, 1).bold).toBe(true);
    expect(host.cell(2, 1).horizontalAlignment).toBe(-4108); // xlCenter
    // 提交后对照同样按最终状态。
    expect(compareCommittedWpsResultToCopy(staged)).toBe("matched");
    await undoWpsExecution(host.application, record.executionId);
    expect(host.cell(2, 1).bold).toBe(false);
    expect(host.cell(2, 1).horizontalAlignment).toBe(1); // xlGeneral
  });

  it("同区域字号被后续动作覆盖（原始混合 12/13 → 14 → 16），完整链路 + 撤销恢复原始混合字号", async () => {
    const host = formatFlowHost();
    // 原始混合字号：A2=12、A3=13。
    host.cell(2, 1).size = 12;
    host.cell(3, 1).size = 13;
    const prepared = prepareWpsPlan(host.application, [1], formatPlan([
      { type: "setFont", sheet: "数据", range: "A2:A3", size: 14 },
      { type: "setFont", sheet: "数据", range: "A2:A3", size: 16 }
    ]));
    const staged = stageWpsNativeComparison(host.application, prepared);
    expect(() => verifyWpsNativeComparison(host.application, staged)).not.toThrow();
    const record = await executeWpsPlan(host.application, prepared);
    expect(record.state).toBe("verified");
    expect(host.cell(2, 1).size).toBe(16);
    expect(host.cell(3, 1).size).toBe(16);
    expect(compareCommittedWpsResultToCopy(staged)).toBe("matched");
    // 撤销按最终恢复影响范围检查冲突（第一步声明的 14 已被覆盖成 16，不算冲突），逆序恢复。
    await undoWpsExecution(host.application, record.executionId);
    expect(host.cell(2, 1).size).toBe(12);
    expect(host.cell(3, 1).size).toBe(13);
  });

  it("审阅后真正改动副本 → 仍拒绝写回", () => {
    const host = formatFlowHost();
    const prepared = prepareWpsPlan(host.application, [1], formatPlan([
      { type: "setFont", sheet: "数据", range: "A2:B3", bold: true }
    ]));
    const staged = stageWpsNativeComparison(host.application, prepared);
    // 用户在副本里手动把 A2 的加粗去掉。
    staged.copy.Worksheets.Item(1).Range!("A2").Item!(1, 1).Font!.Bold = false;
    expect(() => verifyWpsNativeComparison(host.application, staged)).toThrow();
  });

  it("审阅后真正改动原表 → 仍拒绝写回", () => {
    const host = formatFlowHost();
    const prepared = prepareWpsPlan(host.application, [1], formatPlan([
      { type: "setFont", sheet: "数据", range: "A2:B3", bold: true }
    ]));
    const staged = stageWpsNativeComparison(host.application, prepared);
    host.cell(2, 1).bold = true; // 原表被外部改动（本次声明的属性）
    expect(() => verifyWpsNativeComparison(host.application, staged)).toThrow();
  });

  it("执行后修改其他属性（非本次声明），撤销只恢复声明属性、保留其余修改", async () => {
    const host = formatFlowHost();
    const prepared = prepareWpsPlan(host.application, [1], formatPlan([
      { type: "setFont", sheet: "数据", range: "A2:B3", bold: true }
    ]));
    const record = await executeWpsPlan(host.application, prepared);
    host.cell(2, 1).color = 0x0000ff;
    host.cell(2, 1).rowHeight = 30;
    await undoWpsExecution(host.application, record.executionId);
    expect(host.cell(2, 1).bold).toBe(false);
    expect(host.cell(2, 1).color).toBe(0x0000ff);
    expect(host.cell(2, 1).rowHeight).toBe(30);
  });
});
