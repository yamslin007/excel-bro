// @vitest-environment jsdom
/**
 * 格式动作**恢复账本链路**的真实跨进程证据（Task #61 / #62）。
 *
 * 复现的实机失败：真机 WPS 上点「确认写回原表」时，登记请求在**出站契约校验**就被拒，
 * 面板如实中止并报「WPS 恢复描述的动作定位或形状无效」；而只替身账本接口的面板测试
 * 看不到这一层，所以一直是绿的。
 *
 * 本文件把这件事的两半接在**同一根链**上：
 *
 *   真实 `<WpsAgentPanel/>` 的入口按钮 → 真实 `./wpsNativeComparison`（副本审阅）
 *   → 真实 `./wpsPlanExecutor`（写前检查点 / 写后回执 / 回读 / 撤销）
 *   → **真实 `../../excel-addin/src/api` 出站校验**（`assertWpsRecoveryDescriptor` /
 *   `assertWpsEvidence` / `assertWpsExecutionState`）→ **真实 HTTP**
 *   → **真实 ASGI 服务端**（`server.app.main:app`，真实 SQLite 账本 `office_executions`）
 *   → 真实 `register / actions/start / actions/complete / complete` 检查点推进。
 *
 * 允许的替身（仅此三类，其余全真；如实标注，不升级口径）：
 *
 * 1. **WPS 原生对象** —— `./wpsHost` 的 `getWpsApplication()`。本文件底部是有状态宿主
 *    替身：`Range().Value2/Formula` 读写同一份矩阵，**行高按「整行共享」实现**（行级
 *    属性），`Worksheet.Copy()` 真的生成一张**单表**未保存副本。这是唯一与真实产品
 *    不同的"设备"。
 * 2. **模型规划** —— `./wpsReadonlyAgent` 的 `runWpsReadonlyTurn` 与
 *    `api.understandMessage` 固定返回计划。**规划是固定的**，本文件不构成
 *    "真实模型自主规划"的证据。
 * 3. **模型目录** —— `api.listModels` 固定返回一个可用连接。
 *
 * 明确**不做**的事：不为账本接口（register / start / complete / finish / status）摆
 * fetch 桩；不预制服务端返回值；不从旧工件回放；测试代码不直接给宿主写格式或值来
 * 代替产品写回。
 *
 * 环境阻断即不报成功：没有 `EB_WPS_LEDGER=1` + `VITE_API_BASE_URL` 时整个套件跳过，
 * 由 `scripts/run_wps_format_ledger_link.mjs` 判定「未完成」。本轮**不**声称真实
 * WPS / 真实 Excel 客户端已经通过 —— 那两项仍需实机验收。
 */
import { act, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { getWpsExecutionStatus, listModels, understandMessage } from "../../excel-addin/src/api";
import type { AnalysisPlan, CellValue } from "../../excel-addin/src/contracts";
import WpsAgentPanel from "./WpsAgentPanel";
import { runWpsReadonlyTurn } from "./wpsReadonlyAgent";
import {
  getWpsApplication,
  type WpsApplication,
  type WpsRange,
  type WpsWorkbook,
  type WpsWorksheet
} from "./wpsHost";

const LEDGER = process.env.EB_WPS_LEDGER === "1";
const API_BASE = process.env.VITE_API_BASE_URL ?? "";
if (LEDGER && !API_BASE) {
  throw new Error("EB_WPS_LEDGER=1 但没有 VITE_API_BASE_URL：跨进程用例不能指向真实服务端。");
}

/** 面板自己的本地恢复索引键（`wpsExecutionLedger` 的唯一来源）。只读它拿 operationId。 */
const PENDING_KEY = "excel-bro.pending-wps-execution.v1";

/* ------------------------------------------------------------------------- *
 * 替身声明
 * ------------------------------------------------------------------------- */

/**
 * 只替身「规划怎么来」。**账本接口与出站校验一律保留真实实现** —— 本文件存在的
 * 理由就是让真实的 `assertWpsRecoveryDescriptor` 参与，而不是被 mock 绕过去。
 */
vi.mock("../../excel-addin/src/api", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../../excel-addin/src/api")>()),
  understandMessage: vi.fn(async () => ({
    kind: "needs_workbook_context",
    summary: "需要字段结构",
    provider: "model"
  })),
  listModels: vi.fn()
}));

/** 只替身「那一轮规划怎么跑」，不替掉整个只读代理模块。 */
vi.mock("./wpsReadonlyAgent", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./wpsReadonlyAgent")>()),
  runWpsReadonlyTurn: vi.fn()
}));

/** 唯一被替身的宿主入口：拿到原生 WPS Application 对象那一步。 */
vi.mock("./wpsHost", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./wpsHost")>()),
  getWpsApplication: () => host.application
}));

let host: LedgerHost;

(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true;

if (!Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => undefined;
}

const listModelsMock = vi.mocked(listModels);
const understandMessageMock = vi.mocked(understandMessage);
const runTurnMock = vi.mocked(runWpsReadonlyTurn);

const connection = {
  id: "connection:first",
  label: "第一个连接",
  provider: "model" as const,
  available: true,
  supportsVision: false
};

/* ------------------------------------------------------------------------- *
 * 有状态 WPS 宿主替身：值 + 公式 + 逐格格式 + **整行**行高 + 原生单表副本
 * ------------------------------------------------------------------------- */

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

/** 行高是**行级**属性：整行共享一份；其余格式属性落在格上。 */
type SheetState = {
  rows: CellState[][];
  /** 1-based 行号 → 下标 row-1。整行唯一。 */
  rowHeights: number[];
};

const DEFAULT_ROW_HEIGHT = 13.5;
const XL_GENERAL = 1;
const XL_CENTER = -4108;

function defaultCellFormat(): CellFormat {
  return {
    bold: false,
    italic: false,
    name: "宋体",
    size: 11,
    color: 0,
    horizontalAlignment: XL_GENERAL,
    wrapText: false
  };
}

function makeCell(value: CellValue, formula: CellValue = value): CellState {
  return { ...defaultCellFormat(), value, formula };
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
 * 矩形区域替身。
 *
 * `countWrite` 只在**原表**上打开：每次产品调用写 API（字体/对齐/换行/行高/值/公式）
 * 记一次，用于"登记失败 ⇒ 原表零写入"的取证。副本侧不计数。
 */
function makeRange(state: SheetState, rect: Rect, options: { withItem?: boolean; countWrite?: () => void } = {}): WpsRange {
  const withItem = options.withItem ?? true;
  const count = options.countWrite ?? (() => undefined);
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
          get Bold() { return target.bold; },
          set Bold(v: unknown) { count(); target.bold = v as boolean; },
          get Italic() { return target.italic; },
          set Italic(v: unknown) { count(); target.italic = v as boolean; },
          get Name() { return target.name; },
          set Name(v: unknown) { count(); target.name = v as string; },
          get Size() { return target.size; },
          set Size(v: unknown) { count(); target.size = v as number; },
          get Color() { return target.color; },
          set Color(v: unknown) { count(); target.color = v as number; }
        };
      },
      get HorizontalAlignment() { return ensure(absRow, absCol).horizontalAlignment; },
      set HorizontalAlignment(v: unknown) { count(); ensure(absRow, absCol).horizontalAlignment = v as number; },
      get WrapText() { return ensure(absRow, absCol).wrapText; },
      set WrapText(v: unknown) { count(); ensure(absRow, absCol).wrapText = v as boolean; },
      // 行高整行共享：写任意一列即写整行，任意一列读到同一份。
      get RowHeight() { return rowHeightOf(absRow); },
      set RowHeight(v: unknown) { count(); setRowHeight(absRow, v as number); },
      get Value2() { return ensure(absRow, absCol).value; },
      set Value2(v: unknown) { count(); ensure(absRow, absCol).value = v as CellValue; },
      get Formula() { return ensure(absRow, absCol).formula; },
      set Formula(v: unknown) { count(); ensure(absRow, absCol).formula = v as CellValue; }
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
  if (withItem) range.Item = (r: number, c: number) => cellProxy(r - 1, c - 1);
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
            if (value !== undefined) {
              count();
              ensure(rect.startRow + r, rect.startColumn + c).value = value as CellValue;
            }
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
            if (value !== undefined) {
              count();
              ensure(rect.startRow + r, rect.startColumn + c).formula = value as CellValue;
            }
          }
        }
      },
      enumerable: true
    }
  });
  return range;
}

type LedgerHost = {
  application: WpsApplication;
  originalSheet: WpsWorksheet;
  cell: (row: number, column: number) => CellState;
  rowHeight: (row: number) => number;
  /** 记录原表被"产品写入"碰过的次数（宿主 setter 计数）。 */
  writes: () => number;
};

/** 初始：4 行 × 3 列；A 列文本 / B 列数字 / C 列真公式；整行行高 15/18/22/30。 */
const INITIAL_ROWS: CellValue[][] = [
  ["姓名", "金额", "备注"],
  ["甲", 100, null],
  ["乙", 200, null],
  ["丙", 300, null]
];
const INITIAL_FORMULAS: CellValue[][] = [
  ["姓名", "金额", "备注"],
  ["甲", 100, "=B2&\"元\""],
  ["乙", 200, "=B3&\"元\""],
  ["丙", 300, "=B4&\"元\""]
];
const INITIAL_ROW_HEIGHTS = [15, 18, 22, 30];

function ledgerHost(): LedgerHost {
  const rows: CellState[][] = INITIAL_ROWS.map((row, r) =>
    row.map((value, c) => makeCell(value, INITIAL_FORMULAS[r][c])));
  const state: SheetState = { rows, rowHeights: [...INITIAL_ROW_HEIGHTS] };

  /** 只统计**原表**上的产品写入（副本不计）：用于"登记失败 ⇒ 原表零写入"。 */
  let originalWrites = 0;
  let active: WpsWorkbook;

  function buildSheet(sheetName: string, sheetState: SheetState, countWrite: boolean): WpsWorksheet {
    return {
      Name: sheetName,
      ProtectContents: false,
      Range: (address: unknown) =>
        makeRange(sheetState, parseA1(String(address)), {
          countWrite: countWrite ? () => { originalWrites += 1; } : undefined
        })
    };
  }

  const sheetCache = new Map<string, WpsWorksheet>();
  function sheet(cacheKey: string, sheetName: string, sheetState: SheetState, countWrite: boolean): WpsWorksheet {
    const cached = sheetCache.get(cacheKey);
    if (cached) return cached;
    const built = buildSheet(sheetName, sheetState, countWrite);
    sheetCache.set(cacheKey, built);
    return built;
  }

  const originalSheet = sheet("数据", "数据", state, true);
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
    const copySheet = sheet("数据#copy", "数据", copyState, false);
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
    writes: () => originalWrites
  };
}

/* ------------------------------------------------------------------------- *
 * 计划
 * ------------------------------------------------------------------------- */

/**
 * setFont ＋ setAlignment 落在**同一区域**（格式动作之间允许重叠），再叠一个
 * resizeRange（行高，**影响整行**）。合计 12 格 —— 必须不超过测试服务端的单元格
 * 限额（`scripts/run_wps_format_ledger_link.mjs` 把它压到 12，用于另一条用例真实地
 * 触发服务端拒绝）。
 */
const formatPlan: AnalysisPlan = {
  id: "plan-wps-format-ledger",
  title: "加粗居中并调行高",
  summary: "A2:B3 加粗＋居中，第 3–4 行行高 25",
  assumptions: [],
  warnings: [],
  actions: [
    { type: "setFont", sheet: "数据", range: "A2:B3", bold: true, size: 14 },
    { type: "setAlignment", sheet: "数据", range: "A2:B3", horizontal: "center" },
    { type: "resizeRange", sheet: "数据", range: "A3:B4", rowHeight: 25 }
  ],
  acceptanceCriteria: []
};

/** 60 格：客户端预览放行（内置上限 1000），测试服务端（上限 12）真实拒绝登记。 */
const oversizedPlan: AnalysisPlan = {
  id: "plan-wps-format-over-limit",
  title: "整块加粗",
  summary: "A1:F10 加粗",
  assumptions: [],
  warnings: [],
  actions: [{ type: "setFont", sheet: "数据", range: "A1:F10", bold: true }],
  acceptanceCriteria: []
};

/* ------------------------------------------------------------------------- *
 * 面板驱动
 * ------------------------------------------------------------------------- */

let mounted: Array<{ container: HTMLElement; unmount: () => void }> = [];

async function mount(element: ReactElement) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  await act(async () => {
    root.render(element);
  });
  const entry = {
    container,
    unmount: () => {
      act(() => {
        root.unmount();
      });
      container.remove();
    }
  };
  mounted.push(entry);
  return entry;
}

async function settle(): Promise<void> {
  await act(async () => {
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
  });
}

async function until(check: () => boolean, message: string): Promise<void> {
  for (let index = 0; index < 200; index += 1) {
    if (check()) return;
    await settle();
  }
  throw new Error(message);
}

function textOf(container: HTMLElement): string {
  return container.textContent ?? "";
}

function findButton(container: HTMLElement, name: string): HTMLButtonElement | null {
  return (
    Array.from(container.querySelectorAll("button")).find(
      (item) => (item.textContent ?? "").trim() === name
    ) ?? null
  );
}

async function click(element: Element | null): Promise<void> {
  if (!element) throw new Error("要点击的元素不存在");
  await act(async () => {
    element.dispatchEvent(new MouseEvent("click", { bubbles: true }));
  });
  await settle();
}

function cardUndoButton(container: HTMLElement): HTMLButtonElement | null {
  return container.querySelector<HTMLButtonElement>(
    '.wps-verification-card button[aria-label="撤销本次执行"]'
  );
}

/**
 * 审阅卡里的「我已核对副本」确认项。
 *
 * 写回入口在勾选前是**禁用**的（确认框是真闸）：不勾就点，写入会被拦下、
 * 原表一格不写 —— 所以走完写回链路前必须先勾这里。
 */
async function acknowledgeCopy(container: HTMLElement): Promise<void> {
  const review = container.querySelector<HTMLElement>(".wps-native-review-status");
  const checkbox = review?.querySelector<HTMLInputElement>('input[type="checkbox"]') ?? null;
  if (!checkbox) throw new Error("审阅卡里应有「我已核对副本」确认项");
  if (!checkbox.checked) await click(checkbox);
}

async function type(element: HTMLTextAreaElement | null, value: string): Promise<void> {
  if (!element) throw new Error("没有找到输入框");
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, "value")?.set;
  await act(async () => {
    setter?.call(element, value);
    element.dispatchEvent(new Event("input", { bubbles: true }));
  });
}

async function submit(element: HTMLTextAreaElement | null): Promise<void> {
  const form = element?.closest("form");
  if (!form) throw new Error("没有找到输入区表单");
  await act(async () => {
    form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
  });
  await settle();
}

function panel() {
  return (
    <WpsAgentPanel
      selectedSheetIndexes={[1]}
      scopeDock={<div className="eb-scope-dock">当前表 数据</div>}
    />
  );
}

async function mountPreview(plan: AnalysisPlan, task: string) {
  listModelsMock.mockResolvedValue({ defaultModelId: connection.id, models: [connection] });
  runTurnMock.mockResolvedValue({ kind: "plan", plan, turnId: "turn-wps-format-ledger" });
  understandMessageMock.mockResolvedValue({
    kind: "needs_workbook_context",
    summary: "需要字段结构",
    provider: "model"
  } as never);
  const { container } = await mount(panel());
  await until(
    () => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false,
    "应加载模型目录"
  );
  const textarea = container.querySelector("textarea");
  await type(textarea, task);
  await submit(textarea);
  await until(() => textOf(container).includes("写入前确认"), `应展示写入计划预览；实际界面：${textOf(container).slice(-600)}`);
  return container;
}

/**
 * 从面板自己的本地恢复索引里读出本次 operationId（只读观测，不替身任何产品代码）。
 *
 * 索引在登记成功后写入、完成回执提交后清掉，所以只能在**确认到结果卡出现之间**
 * 轮询捕获 —— 这里把它并进同一个 `until` 循环，避免与别的 `act` 交叠。
 */
function pendingOperationId(): string {
  const raw = localStorage.getItem(PENDING_KEY);
  if (!raw) return "";
  try {
    const parsed = JSON.parse(raw) as { operationId?: unknown };
    return typeof parsed.operationId === "string" ? parsed.operationId : "";
  } catch {
    return "";
  }
}

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
});

afterEach(() => {
  mounted.forEach((entry) => entry.unmount());
  mounted = [];
  document.body.innerHTML = "";
});

/* ------------------------------------------------------------------------- *
 * 用例
 * ------------------------------------------------------------------------- */

describe.skipIf(!LEDGER)(
  "格式动作恢复账本链路（真实面板 + 真实出站校验 + 真实 HTTP + 真实服务端账本；仅替身 WPS 宿主、规划固定）",
  () => {
    it("setFont＋setAlignment＋resizeRange：副本审阅 → 登记 → 确认写回 → 回读 → 撤销，账本留下格式证据", async () => {
      host = ledgerHost();
      const container = await mountPreview(formatPlan, "把 A2:B3 加粗居中，第 3–4 行行高改成 25");

      // 首屏：真实 `buildPlanChangePreview` + 真实 `describeWpsFormatAction`。
      expect(textOf(container)).toContain("设置字体（加粗、字号改为 14）");
      expect(textOf(container)).toContain("调整尺寸（行高改为 25（影响整行））");

      // 副本审阅（真实 `stageWpsNativeComparison`）。
      await click(findButton(container, "生成右侧副本并排审阅"));
      await until(() => findButton(container, "确认写回原表") !== null, "应出现「确认写回原表」");
      // 确认前原表零写入。
      expect(host.cell(2, 1).bold).toBe(false);
      expect(host.cell(2, 1).size).toBe(11);
      expect(host.cell(2, 1).horizontalAlignment).toBe(XL_GENERAL);
      expect(host.rowHeight(3)).toBe(22);
      expect(host.rowHeight(4)).toBe(30);

      // 确认写回：真实出站校验 → 真实 HTTP → 真实服务端登记 → 真实写前/写后检查点。
      let operationId = "";
      await acknowledgeCopy(container);
      await click(findButton(container, "确认写回原表"));
      // 一边等结果卡，一边把面板自己写下的 operationId 捕获下来（完成回执后它会被清掉）。
      await until(() => {
        operationId = pendingOperationId() || operationId;
        return container.querySelector(".wps-verification-card") !== null;
      }, `确认后应给出结果卡；实际界面：${textOf(container).slice(-400)}`);
      await until(
        () => container.querySelector(".wps-verification-card.verified") !== null,
        `写入成功应给出回读核对结果；实际界面：${textOf(container).slice(-400)}`
      );

      // 格式真的落到了原表（且行高按整行生效）。
      expect(host.cell(2, 1).bold).toBe(true);
      expect(host.cell(2, 1).size).toBe(14);
      expect(host.cell(2, 1).horizontalAlignment).toBe(XL_CENTER);
      expect(host.rowHeight(3)).toBe(25);
      expect(host.rowHeight(4)).toBe(25);
      // 区域外的第 2 行行高没被牵连。
      expect(host.rowHeight(2)).toBe(18);
      // 值与公式保护：目标区域**内部**的真公式/文本/数字一字未动。
      expect(host.cell(2, 1).value).toBe("甲");
      expect(host.cell(3, 2).value).toBe(200);
      expect(host.cell(3, 3).formula).toBe("=B3&\"元\"");
      expect(host.cell(3, 3).value).toBeNull();

      // 真实服务端账本：描述里带**格式证据**（属性声明 + 格式指纹），且不含任何矩阵。
      const state = await getWpsExecutionStatus(operationId);
      expect(state.status).toBe("completed");
      expect(state.totalActions).toBe(3);
      expect(state.descriptor.actions.map((action) => action.type)).toEqual([
        "setFont",
        "setAlignment",
        "resizeRange"
      ]);
      expect(state.descriptor.actions[0].formatFields).toEqual(["bold", "size"]);
      expect(state.descriptor.actions[1].formatFields).toEqual(["horizontalAlignment"]);
      expect(state.descriptor.actions[2].formatFields).toEqual(["rowHeight"]);
      for (const action of state.descriptor.actions) {
        // 格式证据是**格式指纹**（f1:），不是值矩阵指纹（v2:）—— 后者对格式改动没有区分力。
        expect(action.previewFingerprint).toMatch(/^f1:/);
        expect(action.expectedFingerprint).toMatch(/^f1:/);
        expect(action.previewFingerprint).not.toBe(action.expectedFingerprint);
      }
      // 描述里只有定位与指纹：没有单元格内容、没有格式快照、没有完整计划。
      const encoded = JSON.stringify(state.descriptor);
      expect(encoded).not.toContain("甲");
      expect(encoded).not.toContain("=B3");
      expect(encoded).not.toContain("cells");
      expect(encoded).not.toContain("rowHeights");
      expect(encoded).not.toContain("values");

      // 撤销（真实 `undoWpsExecution`，属性级恢复）。
      await click(cardUndoButton(container));
      await until(() => textOf(container).includes("已撤销并核对恢复"), "撤销后应给出核对过的恢复结果");
      expect(host.cell(2, 1).bold).toBe(false);
      expect(host.cell(2, 1).size).toBe(11);
      expect(host.cell(2, 1).horizontalAlignment).toBe(XL_GENERAL);
      expect(host.rowHeight(3)).toBe(22);
      expect(host.rowHeight(4)).toBe(30);
      // 值与公式自始至终未被动过。
      expect(host.cell(2, 1).value).toBe("甲");
      expect(host.cell(3, 2).value).toBe(200);
      expect(host.cell(3, 3).formula).toBe("=B3&\"元\"");
    });

    it("服务端拒绝登记（超出服务端单元格限额）⇒ 原表零写入，不出现写回结果卡", async () => {
      host = ledgerHost();
      const container = await mountPreview(oversizedPlan, "把 A1:F10 加粗");

      await click(findButton(container, "生成右侧副本并排审阅"));
      await until(() => findButton(container, "确认写回原表") !== null, "应出现「确认写回原表」");

      await acknowledgeCopy(container);
      await click(findButton(container, "确认写回原表"));
      // 面板必须如实报"登记失败、已取消写入"，而不是吞掉错误继续写。
      await until(
        () => textOf(container).includes("无法持久登记本次写入现场，已取消写入；原表未被修改。"),
        `登记失败应给出中止说明；实际界面：${textOf(container).slice(-400)}`
      );
      // 服务端的真实拒绝理由出现在界面上 —— 证明这次失败真的来自服务端账本，而非本地桩。
      expect(textOf(container)).toContain("超过 WPS 写入单元格限额");
      // 原表零写入：格式一格未改，值/公式一格未动，行高整行未改。
      for (let row = 1; row <= 4; row += 1) {
        for (let column = 1; column <= 3; column += 1) {
          expect(host.cell(row, column).bold).toBe(false);
          expect(host.cell(row, column).size).toBe(11);
        }
      }
      expect(host.cell(2, 1).value).toBe("甲");
      expect(host.cell(3, 3).formula).toBe("=B3&\"元\"");
      expect([1, 2, 3, 4].map((row) => host.rowHeight(row))).toEqual([15, 18, 22, 30]);
      // 原表上一次产品写入都没有发生（副本上的写入不计）。
      expect(host.writes()).toBe(0);
      // 也不该出现任何"已写回"的结果卡。
      expect(container.querySelector(".wps-verification-card")).toBeNull();
      expect(textOf(container)).not.toContain("已撤销并核对恢复");
    });
  }
);
