import { act, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { listModels, validatePlan } from "../../excel-addin/src/api";
import WpsAgentPanel from "./WpsAgentPanel";
import { runWpsReadonlyTurn } from "./wpsReadonlyAgent";
import { getWpsApplication } from "./wpsHost";
import type { AnalysisPlan, CellValue } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange, WpsWorkbook, WpsWorksheet } from "./wpsHost";

/**
 * WPS 面板**点击**证据（Task #53 收口）。
 *
 * 与 `wpsFormatFlow.test.ts` 的区别：那是**模块集成证据**（直接调
 * `stageWpsNativeComparison` / `executeWpsPlan` / `undoWpsExecution`）。这里是
 * **真实面板组件**点击证据：挂真实 `<WpsAgentPanel/>`，走真实入口按钮
 * （"生成右侧副本并排审阅" → "确认写回原表" → "撤销本次执行"）。
 *
 * **模拟范围（不得笼统写成"仅替身宿主"）** —— 本文件替身了三层：
 * 1. **宿主**：`./wpsHost` 的 `getWpsApplication`（有状态替身宿主，含值/公式/逐格格式/原生副本）。
 * 2. **模型规划**：`./wpsReadonlyAgent` 的 `runWpsReadonlyTurn`（固定返回计划，不跑真实模型）
 *    与 `api.understandMessage`。
 * 3. **计划校验 + 服务端执行账本接口**：`api.validatePlan`（直接回传）与
 *    `api.registerWpsExecution` / `getWpsExecutionStatus` / `startWpsAction` /
 *    `completeWpsAction` / `suspendWpsExecution` / `claimWpsExecutionResume` /
 *    `reconcileWpsActionReceipt` / `finishWpsExecution`（内存替身，不落真实账本）。
 *
 * **保留的真实部分（本文件的证据价值所在）**：真实 React 面板组件 `<WpsAgentPanel/>`
 * ＋真实审阅模块 `./wpsNativeComparison` ＋真实执行模块 `./wpsPlanExecutor`（含真实的
 * `stageWpsNativeComparison` / `verifyWpsNativeComparison` / `executeWpsPlan` /
 * `undoWpsExecution`）。与 `WpsAgentPanel.test.tsx` 刻意不同：那边为了测"两次确认"
 * 状态机把审阅模块整体替身，会把审阅接线一起藏掉。
 *
 * **不升级口径**：本文件**不能**当作真实模型、真实服务端或真实客户端证据。
 */

vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../excel-addin/src/api")>();
  const discovery = await import("./wpsPendingDiscovery.testHost");
  const recovery = await import("./wpsRecoveryStore.testHost");
  const audit = await import("@excel-bro/core/hostWriteAudit");
  const wpsRunningState = (
    operationId: string,
    revision = 1,
    overrides: Record<string, unknown> = {}
  ) => ({
    operationId,
    status: "running" as const,
    revision,
    workbookIdentity: "wps-path:c:\\data\\格式数据.xlsx",
    taskId: "task-test",
    descriptor: { actions: [] },
    authorizedSheetNames: ["数据"],
    nextActionIndex: 0,
    totalActions: 1,
    inFlightActionIndex: null,
    beforeSnapshot: null,
    checkpoint: null,
    result: null,
    error: null,
    canResume: true,
    ownerActive: false,
    requiresReconciliation: false,
    ownerToken: "owner-token-0123456789abcdef0123456789abcdef",
    ...overrides
  });
  return {
    ...actual, ...discovery.pendingDiscoveryApiStubs(),
    understandMessage: vi.fn(async () => ({ kind: "needs_workbook_context", summary: "需要字段结构", provider: "model" })),
    listModels: vi.fn(),
    getModelSettings: vi.fn(),
    validatePlan: vi.fn(),
    registerWpsExecution: vi.fn(async (input: { operationId: string }) => wpsRunningState(input.operationId)),
    getWpsExecutionStatus: vi.fn(async (operationId: string) => wpsRunningState(operationId)),
    startWpsAction: vi.fn(async (operationId: string, input: { expectedRevision: number }) =>
      wpsRunningState(operationId, input.expectedRevision + 1)),
    completeWpsAction: vi.fn(async (operationId: string, input: { expectedRevision: number }) =>
      wpsRunningState(operationId, input.expectedRevision + 1)),
    suspendWpsExecution: vi.fn(async (operationId: string) => wpsRunningState(operationId)),
    claimWpsExecutionResume: vi.fn(async (operationId: string) => wpsRunningState(operationId)),
    reconcileWpsActionReceipt: vi.fn(async (operationId: string) => wpsRunningState(operationId)),
    /* 用户处置留痕（放弃/撤销）：真实网络出口，必须替身。 */
    recordExecutionDisposition: vi.fn(
      async (_namespace: string, _operationId: string, _disposition: string) => undefined
    ),
    finishWpsExecution: vi.fn(async (operationId: string) => ({
      ...wpsRunningState(operationId),
      status: "completed" as const
    })),
    /*
     * 三个**真实网络出口**，以前漏了替身：宿主写入次数补报、恢复材料的保存与回收。
     * 面板会在账本提交之后 best-effort 补报、在终态回收材料 —— 调用方 `void` + `catch`
     * 掉，所以"测试在向用户那台 8765 发请求"能安静很久。现在按本机后端替身的语义给：
     * 补报只回一份自洽的汇总，材料读写走同一个内存后端。
     */
    reportWpsHostWrites: vi.fn(async (operationId: string, input: { entries: import("@excel-bro/core/hostWriteAudit").HostWriteAudit[] }) => ({
      operationId,
      entries: input.entries,
      totals: audit.summarizeHostWrites(input.entries),
      observation: {
        observation: "complete" as const,
        callEvidence: input.entries.some((entry) => entry.hostWriteCalls > 0) ? "nonzero" as const : "zero" as const,
        attemptedActions: input.entries.length,
        reportedActions: input.entries.length,
        missingActions: 0,
        unexpectedReports: 0,
        unit: "wps_api_calls" as const,
        lastReportedAt: Date.now()
      }
    })),
    saveExecutionRecoveryPackage: vi.fn(async (id: string, text: string) => recovery.backendSave(id, text)),
    loadExecutionRecoveryPackage: vi.fn(async (id: string) => recovery.backendLoad(id)),
    deleteExecutionRecoveryPackage: vi.fn(async (id: string) => recovery.backendRemove(id))
  };
});

/** 只替身"那一轮规划怎么跑"，不替掉整个只读代理模块（桥适配器还要用它）。 */
vi.mock("./wpsReadonlyAgent", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./wpsReadonlyAgent")>()),
  runWpsReadonlyTurn: vi.fn()
}));

/** 唯一被替身的是**宿主**：`getWpsApplication` 返回一个带格式能力的替身宿主。 */
vi.mock("./wpsHost", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsHost")>();
  return { ...actual, getWpsApplication: vi.fn(() => ({})) };
});

(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true;

if (!Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => undefined;
}

const listModelsMock = vi.mocked(listModels);
const validatePlanMock = vi.mocked(validatePlan);
const runTurnMock = vi.mocked(runWpsReadonlyTurn);
const getWpsApplicationMock = vi.mocked(getWpsApplication);

const connection = {
  id: "connection:first",
  label: "第一个连接",
  provider: "model" as const,
  available: true,
  supportsVision: false
};

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
  for (let index = 0; index < 120; index += 1) {
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
 * 写回入口在勾选前是**禁用**的（确认框是真闸），所以走完写回链路前必须先勾这里；
 * 直接点按钮会被拦下、原表一格不写。
 */
async function acknowledgeCopy(container: HTMLElement): Promise<void> {
  const review = container.querySelector<HTMLElement>(".wps-native-review-status");
  const checkbox = review?.querySelector<HTMLInputElement>('input[type="checkbox"]') ?? null;
  if (!checkbox) throw new Error("审阅卡里应有「我已核对副本」确认项");
  if (!checkbox.checked) await click(checkbox);
}

async function type(element: HTMLTextAreaElement | null, value: string): Promise<void> {
  if (!element) throw new Error("没有找到输入框");
  const setter = Object.getOwnPropertyDescriptor(
    HTMLTextAreaElement.prototype,
    "value"
  )?.set;
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

/* ------------------------------------------------------------------------- *
 * 替身宿主：一张写入目标表，值 + 公式 + 逐格格式 + 原生副本审阅能力。
 * ------------------------------------------------------------------------- */

type FormatState = {
  bold: boolean; italic: boolean; name: string; size: number; color: number;
  horizontalAlignment: number; wrapText: boolean; rowHeight: number;
};
type CellState = FormatState & { value: CellValue; formula: CellValue };

function defaultFormat(): FormatState {
  return {
    bold: false, italic: false, name: "宋体", size: 11, color: 0,
    horizontalAlignment: 1, wrapText: false, rowHeight: 13.5
  };
}

function parseA1(address: string) {
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

function makeSheetState(values: (CellValue | undefined)[][]): CellState[][] {
  const rows = values.map((row) =>
    row.map((value) => ({ ...defaultFormat(), value: value ?? null, formula: value ?? null } as CellState))
  );
  const maxCols = Math.max(0, ...rows.map((row) => row.length));
  for (const row of rows) {
    while (row.length < maxCols) row.push({ ...defaultFormat(), value: null, formula: null });
  }
  return rows;
}

function makeRange(
  state: CellState[][],
  rect: { startRow: number; startColumn: number; endRow: number; endColumn: number }
): WpsRange {
  const rows = rect.endRow - rect.startRow + 1;
  const columns = rect.endColumn - rect.startColumn + 1;
  const cell = (r: number, c: number): CellState => {
    while (state.length <= rect.startRow - 1 + r) {
      state.push(Array.from({ length: Math.max(columns, state[0]?.length ?? 1) }, () =>
        ({ ...defaultFormat(), value: null, formula: null })));
    }
    while (state[rect.startRow - 1 + r].length <= rect.startColumn - 1 + c) {
      state[rect.startRow - 1 + r].push({ ...defaultFormat(), value: null, formula: null });
    }
    return state[rect.startRow - 1 + r][rect.startColumn - 1 + c];
  };

  const cellProxy = (r: number, c: number): WpsRange => ({
    get Font() {
      const target = cell(r, c);
      return {
        get Bold() { return target.bold; }, set Bold(v: unknown) { target.bold = v as boolean; },
        get Italic() { return target.italic; }, set Italic(v: unknown) { target.italic = v as boolean; },
        get Name() { return target.name; }, set Name(v: unknown) { target.name = v as string; },
        get Size() { return target.size; }, set Size(v: unknown) { target.size = v as number; },
        get Color() { return target.color; }, set Color(v: unknown) { target.color = v as number; }
      };
    },
    get HorizontalAlignment() { return cell(r, c).horizontalAlignment; },
    set HorizontalAlignment(v: unknown) { cell(r, c).horizontalAlignment = v as number; },
    get WrapText() { return cell(r, c).wrapText; },
    set WrapText(v: unknown) { cell(r, c).wrapText = v as boolean; },
    get RowHeight() { return cell(r, c).rowHeight; },
    set RowHeight(v: unknown) { cell(r, c).rowHeight = v as number; },
    get Value2() { return cell(r, c).value; },
    set Value2(v: unknown) { cell(r, c).value = v as CellValue; },
    get Formula() { return cell(r, c).formula; },
    set Formula(v: unknown) { cell(r, c).formula = v as CellValue; }
  });

  const range: WpsRange = {
    MergeCells: false,
    Rows: { Count: rows },
    Columns: { Count: columns },
    Interior: { Color: 0 },
    Item: (r: number, c: number) => cellProxy(r - 1, c - 1),
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
  Object.defineProperties(range, {
    Value2: {
      get: () => {
        const matrix = Array.from({ length: rows }, (_, r) =>
          Array.from({ length: columns }, (_, c) => cell(r, c).value));
        return rows === 1 && columns === 1 ? matrix[0][0] : matrix;
      },
      set: (incoming: unknown) => {
        const normalized = Array.isArray(incoming) ? (incoming as unknown[][]) : [[incoming]];
        for (let r = 0; r < rows; r += 1) {
          for (let c = 0; c < columns; c += 1) {
            const value = normalized[r]?.[c];
            if (value !== undefined) cell(r, c).value = value as CellValue;
          }
        }
      },
      enumerable: true
    },
    Formula: {
      get: () => {
        const matrix = Array.from({ length: rows }, (_, r) =>
          Array.from({ length: columns }, (_, c) => cell(r, c).formula));
        return rows === 1 && columns === 1 ? matrix[0][0] : matrix;
      },
      set: (incoming: unknown) => {
        const normalized = Array.isArray(incoming) ? (incoming as unknown[][]) : [[incoming]];
        for (let r = 0; r < rows; r += 1) {
          for (let c = 0; c < columns; c += 1) {
            const value = normalized[r]?.[c];
            if (value !== undefined) cell(r, c).formula = value as CellValue;
          }
        }
      },
      enumerable: true
    }
  });
  return range;
}

/** 完整宿主：一张写入目标表（名「数据」），值 + 公式 + 逐格格式 + 原生副本审阅能力。 */
function installFormatHost() {
  const state = makeSheetState([
    ["姓名", "金额", "备注"],
    ["甲", 100, "=B2&\"元\""],
    ["乙", 200, "=B3&\"元\""]
  ]);
  let active: WpsWorkbook;
  let copyClosed = false;
  /** 副本是否真的被创建过 —— 用来区分"没建"与"建了没关"。 */
  let copyCreated = false;

  const sheetCache = new Map<string, WpsWorksheet>();
  function buildSheet(name: string, sheetState: CellState[][]): WpsWorksheet {
    return {
      Name: name,
      ProtectContents: false,
      Range: (address: unknown) => makeRange(sheetState, parseA1(String(address)))
    };
  }
  function sheet(name: string, sheetState: CellState[][]): WpsWorksheet {
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
    copyCreated = true;
    const copyState = state.map((row) => row.map((item) => ({ ...item })));
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
  getWpsApplicationMock.mockImplementation(() => application as never);
  return {
    cell: (row: number, column: number) => state[row - 1][column - 1],
    copyClosed: () => copyClosed,
    copyCreated: () => copyCreated
  };
}

const boldPlan: AnalysisPlan = {
  id: "plan-format-review",
  title: "加粗 A2:B3",
  summary: "把 A2:B3 加粗",
  assumptions: [],
  warnings: [],
  actions: [{ type: "setFont", sheet: "数据", range: "A2:B3", bold: true }],
  acceptanceCriteria: []
};

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
  runTurnMock.mockResolvedValue({ kind: "plan", plan, turnId: "turn-format-1" });
  validatePlanMock.mockImplementation(async (submitted) => submitted);
  const { container } = await mount(panel());
  await until(
    () => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false,
    "应加载模型目录"
  );
  const textarea = container.querySelector("textarea");
  await type(textarea, task);
  await submit(textarea);
  await until(() => findButton(container, "确认变更") !== null, "应显示主表执行结果：" + textOf(container));
  return container;
}

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
  getWpsApplicationMock.mockReturnValue({ Version: "12.0" } as never);
});

afterEach(() => {
  mounted.forEach((entry) => entry.unmount());
  mounted = [];
  document.body.innerHTML = "";
});

describe("WPS 面板 · 格式动作点击链路（真实面板 + 真实审阅/执行模块；模型规划·计划校验·服务端账本·宿主均为替身）", () => {
  it("新流程在主表显示结果，确认不重复执行，放弃真正恢复", async () => {
    const host = installFormatHost();
    const container = await mountPreview(boldPlan, "把 A2:B3 加粗");
    await until(() => findButton(container, "确认变更") !== null,
      `应自动执行并显示固定确认入口：${textOf(container)}`);
    expect(host.cell(2, 1).bold).toBe(true);
    expect(host.cell(2, 1).value).toBe("甲");
    await click(findButton(container, "放弃变更"));
    await until(() => textOf(container).includes("已放弃并恢复"), "应真正恢复主表");
    expect(host.cell(2, 1).bold).toBe(false);
  });
  it("主表结果 → 确认变更 → 撤销按钮接线，全程真实模块", async () => {
    const host = installFormatHost();
    const container = await mountPreview(boldPlan, "把 A2:B3 加粗");

    // 首屏格式描述：真实 `buildPlanChangePreview` + 真实 `describeWpsFormatAction`。
    expect(host.cell(2, 1).bold).toBe(true);
    expect(host.cell(2, 1).value).toBe("甲");
    await click(findButton(container, "确认变更"));
    await until(
      () => container.querySelector(".wps-verification-card.verified") !== null,
      "写入成功应给出回读核对结果"
    );
    expect(host.cell(2, 1).bold).toBe(true);
    expect(host.cell(2, 1).value).toBe("甲");
    expect(host.cell(2, 3).formula).toBe("=B2&\"元\"");

    // 撤销按钮接线（真实 `undoWpsExecution`）。
    const undo = cardUndoButton(container);
    expect(undo).not.toBeNull();
    const resultCard = undo!.closest(".execution-result-card");
    expect(resultCard?.textContent).toContain("设置字体 ·");
    expect(resultCard?.textContent).not.toContain("写入值 ·");
    expect(resultCard?.querySelector("details")?.open).toBe(false);
    await click(undo);
    await until(() => textOf(container).includes("已撤销并恢复"), "撤销后应给出核对过的恢复结果");
    expect(host.cell(2, 1).bold).toBe(false);
    expect(host.cell(2, 1).value).toBe("甲");
  });

  it("同区域「加粗→居中」在真实面板链路里不被组合状态误报", async () => {
    const host = installFormatHost();
    const comboPlan: AnalysisPlan = {
      ...boldPlan,
      id: "plan-format-combo",
      title: "加粗并居中",
      actions: [
        { type: "setFont", sheet: "数据", range: "A2:B3", bold: true },
        { type: "setAlignment", sheet: "数据", range: "A2:B3", horizontal: "center" }
      ]
    };
    const container = await mountPreview(comboPlan, "把 A2:B3 加粗并居中");

    await click(findButton(container, "确认变更"));
    await until(
      () => container.querySelector(".wps-verification-card.verified") !== null,
      "组合格式动作应能通过真实审阅并写回"
    );
    expect(host.cell(2, 1).bold).toBe(true);
    expect(host.cell(2, 1).horizontalAlignment).toBe(-4108);

    await click(cardUndoButton(container));
    await until(() => textOf(container).includes("已撤销并恢复"), "撤销后应给出核对过的恢复结果");
    expect(host.cell(2, 1).bold).toBe(false);
    expect(host.cell(2, 1).horizontalAlignment).toBe(1);
  });

  it("可选双视图展示执行前状态，确认仍位于固定入口", async () => {
    const host = installFormatHost();
    const container = await mountPreview(boldPlan, "把 A2:B3 加粗");
    expect(findButton(container, "重新排列")).toBeNull();
    expect(host.cell(2, 1).bold).toBe(true);
    await click(findButton(container, "双视图对比"));
    // 双视图是执行前状态的副本：真的建了出来（不是拿执行后的原表顶替）。
    expect(host.copyCreated()).toBe(true);
    // 确认入口始终在固定操作条上，不随对话滚走。
    expect(findButton(container, "确认变更")).not.toBeNull();
    await click(findButton(container, "确认变更"));
    await until(() => textOf(container).includes("已确认变更"), "应确认结果");
    expect(host.cell(2, 1).bold).toBe(true);
    expect(host.cell(2, 1).value).toBe("甲");
    // 只关本次副本。
    expect(host.copyClosed()).toBe(true);
  });
});
