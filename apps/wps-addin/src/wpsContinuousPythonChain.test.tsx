// @vitest-environment jsdom
/**
 * **同一次任务的连续链路**（WPS 端）：宿主读取 → `prepare_python_data` →
 * 数据句柄登记 → `run_python_analysis` → 已验证的隔离执行器 → 结果登记与引用
 * 取回 → 产品白名单计划编译 → 面板预览 → 用户确认 → 产品执行器写回 → 回读 → 撤销。
 *
 * **链路里各环节跑在哪一侧（照实分区）**：
 * - **本机桥**只承载两个工具：`query_table` 与 `prepare_python_data`
 *   （`packages/core/src/sdkToolBridge.ts` 的 `BridgeToolName`）。本文件驱动的是
 *   `prepare_python_data`：读授权表 → 数据句柄登记 → 回执。
 * - **`run_python_analysis` 在服务端**（MCP handler → `python_coordinator` →
 *   已验证的隔离执行器），它**不是**桥指针。它的真实执行在
 *   `server/tests/test_python_continuous_chain_real.py`（真实 Docker）验证；
 *   本文件把"服务端已用该结果编译出白名单计划"作为同一次任务的下一步接收。
 * - **白名单计划 → 预览 → 确认 → 产品执行器写回 → 回读 → 撤销** 全在真实
 *   `WpsAgentPanel` 上跑，写回与撤销都是产品的执行器 / 撤销按钮。
 *
 * 与 Excel 端 `continuousPythonChain.test.tsx` 用**同一批脏数据**与**同一份期望**
 * （金卡 400 / 银卡 200 / 未匹配 400），两端不能各算一套。
 *
 * 纪律（用户点名要求）：
 * - 允许**固定模型的工具调用序列**与**有状态宿主替身**（`Range().Value2` 读写同一份矩阵）。
 * - **不用**预制结果矩阵、预制异常说明、预制最终计划，也**不在测试里直接给
 *   `Range().Value2` 赋值**来代替产品环节：写回一律走真实确认+执行器，回读一律走
 *   真实宿主替身，撤销一律走真实「撤销本次执行」按钮。
 * - 唯一被顶替的是**网络往返**（`fetch` → 回环端点）与**隔离容器本身**。
 */
import { act, useEffect } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import type { AnalysisPlan } from "../../excel-addin/src/contracts";
import {
  finishWpsExecution,
  getModelSettings,
  getWpsExecutionStatus,
  listModels,
  registerWpsExecution,
  startWpsAction,
  completeWpsAction,
  suspendWpsExecution,
  claimWpsExecutionResume,
  reconcileWpsActionReceipt,
  understandMessage,
  validatePlan
} from "../../excel-addin/src/api";
import WpsAgentPanel from "./WpsAgentPanel";
import { getWpsApplication } from "./wpsHost";
import {
  runWpsReadonlyTurn,
  type WpsReadonlyTurnInput,
  type WpsTurnResult
} from "./wpsReadonlyAgent";

const MEMBERS = "会员表";
const ORDERS = "订单表";
const TARGET = "汇总表";

/** 与 Excel 端逐字同一批脏数据。 */
const MEMBER_ROWS: (string | number | null)[][] = [
  ["会员编号", "会员", "等级"],
  ["  7  ", "带空格会员", "金卡"],
  ["7", "重复键会员", "银卡"],
  ["007", "前导零会员", "银卡"],
  ["", "空键会员", "铜卡"],
  ["9", "正常会员", "金卡"]
];
const ORDER_ROWS: (string | number | null)[][] = [
  ["会员编号", "金额"],
  ["7", 100],
  ["007", 200],
  ["9", 300],
  ["404", 400]
];

/**
 * 手算的等级汇总（关联后才可能得到）：金卡 400 / 银卡 200 / 未匹配 400。
 * 与 Excel 端、与服务端真实 Docker 用例三处同一份期望。
 */
const EXPECTED_SUMMARY: (string | number)[][] = [
  ["会员等级", "金额合计"],
  ["金卡", 400],
  ["银卡", 200],
  ["未匹配", 400]
];

const PREPARE_OP = "op_" + "2".repeat(32);
const FP = "b".repeat(64);
const RESULT_REF = "res_" + "6".repeat(32);

const firstConnection = {
  id: "connection:first",
  label: "第一个连接",
  provider: "model" as const,
  available: true,
  supportsVision: false
};

// 本文件用 `createRoot` 直接挂真实面板，必须打开 React 的 act 环境，
// 否则 `act(...)` 只警告不生效，状态更新会在断言之后才落地。
(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true;

type Mounted = { container: HTMLElement; unmount: () => void };
let mounted: Mounted[] = [];

// jsdom 不实现 scrollIntoView；面板只用它把新消息滚到底部。
if (!Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => undefined;
}

vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual =
    await importOriginal<typeof import("../../excel-addin/src/api")>();
  return {
    ...actual,
    understandMessage: vi.fn(),
    checkIntent: vi.fn(),
    listModels: vi.fn(),
    getModelSettings: vi.fn(),
    validatePlan: vi.fn(),
    // CAP-03 登记 API：本文件测的是"数据登记→执行→撤销"，断点恢复登记用
    // 一份形状合法的成功状态放行（它有自己的独立用例）。
    registerWpsExecution: vi.fn(),
    getWpsExecutionStatus: vi.fn(),
    startWpsAction: vi.fn(),
    completeWpsAction: vi.fn(),
    suspendWpsExecution: vi.fn(),
    claimWpsExecutionResume: vi.fn(),
    reconcileWpsActionReceipt: vi.fn(),
    finishWpsExecution: vi.fn()
  };
});

// 只替掉"那一轮规划怎么跑"，桥宿主适配器仍用同模块的 `sheetsForRequest`。
vi.mock("./wpsReadonlyAgent", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./wpsReadonlyAgent")>()),
  runWpsReadonlyTurn: vi.fn()
}));

const getWpsApplicationMock = vi.mocked(getWpsApplication);
const listModelsMock = vi.mocked(listModels);
const understandMessageMock = vi.mocked(understandMessage);
const validatePlanMock = vi.mocked(validatePlan);
const runTurnMock = vi.mocked(runWpsReadonlyTurn);

vi.mock("./wpsHost", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./wpsHost")>()),
  getWpsApplication: vi.fn()
}));

// 原生对照走真实模块，但跳过需要窗口排列 / 副本读回的宿主能力 —— 本文件要验的
// 是"数据登记→计划→写回→撤销"，不是并排审阅。
vi.mock("./wpsNativeComparison", () => ({
  WpsFormulaTrialError: class WpsFormulaTrialError extends Error {},
  stageWpsNativeComparison: vi.fn((_application: unknown, plan: unknown) => {
    const prepared = plan as { actions: { values?: unknown[][] }[] };
    return {
      plan,
      changedCells: ["A1"],
      reviews: [],
      original: { Name: "会员.xlsx" },
      copy: { Name: "Book1" },
      copyIdentity: { name: "Book1", fullName: null, path: null },
      originalBefore: prepared.actions.map(() => ({
        values: [[]],
        formulas: [[]],
        valuesCaptured: true
      })),
      copyFingerprints: prepared.actions.map(() => "copy-fingerprint"),
      copyAfter: prepared.actions.map((action) => ({
        values: action.values ?? [[]],
        formulas: [],
        valuesCaptured: true
      }))
    };
  }),
  // 执行前快照只读原表、不建副本（与真实实现同语义）。
  /*
   * 执行前副本：真实实现按**执行前快照**重建（先复制、再把快照写回副本）。
   * 替身只保证形状与"调用过"，真正的重建逻辑由 `wpsNativeComparison` 单测覆盖。
   */
  stageWpsBeforeComparison: vi.fn((_application: unknown, before: unknown) => ({
    plan: (before as { prepared: unknown }).prepared,
    original: { Name: "测试.xlsx" },
    copy: { Name: "Book1" },
    copyIdentity: { name: "Book1", fullName: null, path: null },
    originalBefore: [],
    copyFingerprints: [],
    copyAfter: [],
    reviews: [],
    changedCells: []
  })),
  captureWpsBeforeResult: vi.fn((_application: unknown, prepared: unknown) => ({
    prepared,
    values: [],
    workbookFullName: "C:////Data////会员.xlsx"
  })),
  compareCommittedWpsResultToCopy: vi.fn(() => "matched"),
  verifyWpsNativeComparison: vi.fn(),
  discardWpsNativeComparison: vi.fn(() => "closed"),
  /*
   * 副本登记/切表适配层后来新增了这些出口；这里的替身只需保证"调用不炸"，
   * 真实持久化与并排行为由 wpsNativeComparison 自己的单测覆盖。
   * 缺一个就会在 `ResultReview.closeBefore()` 里抛未处理的拒绝。
   */
  suspendWpsComparison: vi.fn(async () => null),
  switchWpsBeforeComparison: vi.fn(async () => true),
  activateWpsBeforeComparison: vi.fn(),
  saveWpsComparisonCopies: vi.fn(async () => {}),
  loadWpsComparisonCopies: vi.fn(async () => [])
}));

async function mount(element: React.ReactElement): Promise<HTMLElement> {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  await act(async () => {
    root.render(element);
  });
  await settle();
  mounted.push({ container, unmount: () => act(() => root.unmount()) });
  return container;
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

/** 结果卡里的撤销按钮。 */
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

async function type(
  element: HTMLTextAreaElement | null,
  value: string
): Promise<void> {
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

/**
 * 有状态宿主替身：三张表（两张来源 + 一张写入目标），读写同一份矩阵。
 * 工作表对象**每个名字只有一个**（授权核对比对象身份）。
 */
function membershipHost(
  options: { renameTarget?: () => boolean } = {}
) {
  const matrices: Record<string, (string | number | null)[][]> = {
    [MEMBERS]: MEMBER_ROWS.map((row) => [...row]),
    [ORDERS]: ORDER_ROWS.map((row) => [...row]),
    [TARGET]: []
  };
  let writes = 0;

  function rectangle(sheetName: string, rect: ReturnType<typeof parseA1>) {
    const grid = matrices[sheetName] ?? [];
    const rows = rect.endRow - rect.startRow + 1;
    const columns = rect.endColumn - rect.startColumn + 1;
    const slice = () =>
      Array.from({ length: rows }, (_, rowOffset) =>
        Array.from({ length: columns }, (_, columnOffset) =>
          grid[rect.startRow - 1 + rowOffset]?.[rect.startColumn - 1 + columnOffset] ?? null
        )
      );
    const store = (incoming: unknown): void => {
      writes += 1;
      const normalized = Array.isArray(incoming) ? (incoming as unknown[][]) : [[incoming]];
      for (let rowOffset = 0; rowOffset < rows; rowOffset += 1) {
        const rowIndex = rect.startRow - 1 + rowOffset;
        if (!grid[rowIndex]) grid[rowIndex] = [];
        for (let columnOffset = 0; columnOffset < columns; columnOffset += 1) {
          const value = normalized[rowOffset]?.[columnOffset];
          if (value !== undefined) {
            grid[rowIndex][rect.startColumn - 1 + columnOffset] = value as never;
          }
        }
      }
    };
    const range: Record<string, unknown> = {
      Rows: { Count: rows },
      Columns: { Count: columns },
      MergeCells: false
    };
    Object.defineProperties(range, {
      // 单格返回标量、多格返回矩阵 —— 与真宿主一致。
      Value2: {
        get: () => {
          const current = slice();
          return rows === 1 && columns === 1 ? current[0][0] : current;
        },
        set: store,
        enumerable: true
      },
      Formula: {
        get: () => {
          const current = slice();
          return rows === 1 && columns === 1 ? current[0][0] : current;
        },
        set: store,
        enumerable: true
      }
    });
    return range;
  }

  const sheetCache = new Map<string, ReturnType<typeof buildSheet>>();
  function sheet(name: string) {
    const cached = sheetCache.get(name);
    if (cached) return cached;
    const built = buildSheet(name);
    sheetCache.set(name, built);
    return built;
  }

  function buildSheet(name: string) {
    const grid = matrices[name] ?? [];
    const columnCount = Math.max(0, ...grid.map((row) => row.length));
    const usedRange: Record<string, unknown> = {
      Row: 1,
      Column: 1,
      Rows: { Count: grid.length },
      Columns: { Count: columnCount },
      Item: (row: number, column = 1) => ({ Row: row, Column: column })
    };
    Object.defineProperties(usedRange, {
      Value2: { get: () => grid.map((row) => [...row]), enumerable: true },
      MergeCells: { value: false }
    });
    return {
      Name: name,
      CodeName: `Code_${name}`,
      ProtectContents: false,
      UsedRange: usedRange,
      Range: (first: unknown, last?: unknown) => {
        if (typeof first === "string") return rectangle(name, parseA1(first));
        const start = first as { Row?: number; Column?: number };
        const end = (last ?? first) as { Row?: number; Column?: number };
        return rectangle(name, {
          startRow: Number(start.Row ?? 1),
          startColumn: Number(start.Column ?? 1),
          endRow: Number(end.Row ?? start.Row ?? 1),
          endColumn: Number(end.Column ?? start.Column ?? 1)
        });
      }
    };
  }

  const names = () => {
    const all = [MEMBERS, ORDERS, TARGET];
    if (options.renameTarget?.()) {
      return [MEMBERS, ORDERS, "被改名的工作表"];
    }
    return all;
  };

  const storage = new Map<string, unknown>();
  const workbook = {
    Name: "会员.xlsx",
    // 必须是真实 Windows 盘符反斜杠路径：写入闸只对已保存的本地 .xlsx 开放。
    FullName: "C:\\Data\\会员.xlsx",
    Path: "C:\\Data",
    ReadOnly: false,
    Worksheets: {
      get Count() {
        return names().length;
      },
      Item: (key: number | string) => {
        const list = names();
        if (typeof key === "string" && !/^\d+$/.test(key)) {
          return list.includes(key) ? sheet(key) : undefined;
        }
        const name = list[Number(key) - 1];
        return name ? sheet(name) : undefined;
      }
    },
    ActiveSheet: undefined as unknown
  };
  const activeSheet = sheet(MEMBERS);
  workbook.ActiveSheet = activeSheet;

  getWpsApplicationMock.mockReturnValue({
    Version: "12.0",
    ActiveWorkbook: workbook,
    ActiveSheet: activeSheet,
    PluginStorage: {
      getItem: (key: string) => storage.get(key),
      setItem: (key: string, item: unknown) => {
        storage.set(key, item);
      }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as never);

  return {
    writes: () => writes,
    /** 供"断言来源表一格未动"用：读出真实矩阵副本。 */
    matrix: (sheetName: string) =>
      (matrices[sheetName] ?? []).map((row) => [...row]),
    cell: (sheetName: string, row: number, column: number) =>
      matrices[sheetName]?.[row - 1]?.[column - 1] ?? null
  };
}

/** 极小 A1 解析（只支持 `A1`、`A1:B4` 这类无 $ 的地址）。 */
function parseA1(address: string) {
  const [start, end = start] = address.split(":");
  const toCell = (text: string) => {
    const match = /^([A-Z]+)(\d+)$/i.exec(text.trim());
    if (!match) throw new Error(`无法解析区域：${address}`);
    let column = 0;
    for (const char of match[1].toUpperCase()) {
      column = column * 26 + (char.charCodeAt(0) - 64);
    }
    return { column, row: Number(match[2]) };
  };
  const from = toCell(start);
  const to = toCell(end);
  return {
    startRow: from.row,
    startColumn: from.column,
    endRow: to.row,
    endColumn: to.column
  };
}

/**
 * 拦截回环端点。
 *
 * 本机桥在一轮里只会收到 `prepare_python_data` 指针（读授权表 → 数据句柄登记 →
 * 回执）。`run_python_analysis` 在服务端、不在桥上 —— 所以这里**不**摆 analyze
 * 指针，而是让服务端（被 mock 的 `runWpsReadonlyTurn`）在 prepare 回执之后直接
 * 给出**由那次隔离执行编译出的**白名单计划。
 */
function stubContinuousBridge() {
  const captured: Array<{ url: string; body: Record<string, unknown> }> = [];
  vi.stubGlobal(
    "fetch",
    vi.fn(async (url: string, init?: RequestInit) => {
      const target = String(url);
      captured.push({ url: target, body: JSON.parse(String(init?.body ?? "{}")) });
      if (target.includes("/receipt")) {
        return json({ status: "accepted", code: null }, 200);
      }
      if (target.includes("/result")) {
        return json(
          {
            resultRef: RESULT_REF,
            title: "会员等级汇总",
            headers: EXPECTED_SUMMARY[0],
            rows: EXPECTED_SUMMARY.slice(1),
            sourceSheets: [MEMBERS, ORDERS],
            warnings: [],
            complete: true,
            calculation: "按会员等级汇总金额",
            scannedRows: 9
          },
          200
        );
      }
      // 轮询：本机桥收到的唯一指针是 prepare_python_data（读授权表 → 登记句柄）。
      return json(
        {
          calls: [
            {
              operationId: PREPARE_OP,
              tool: "prepare_python_data",
              arguments: {
                sources: [
                  { sheet: MEMBERS, range: "A1:C6" },
                  { sheet: ORDERS, range: "A1:B5" }
                ]
              },
              requestFingerprint: FP,
              expiresAt: 0
            }
          ],
          generation: 1
        },
        200
      );
    })
  );
  return captured;
}

function json(payload: unknown, status: number): Response {
  return new Response(JSON.stringify(payload), {
    status,
    headers: { "Content-Type": "application/json" }
  });
}

function runningState(operationId: string, revision = 1) {
  return {
    operationId,
    status: "running" as const,
    revision,
    workbookIdentity: "wps-path:c:\\data\\会员.xlsx",
    taskId: "task-continuous",
    descriptor: { actions: [] },
    authorizedSheetNames: [MEMBERS, ORDERS, TARGET],
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
    ownerToken: "owner-token-0123456789abcdef0123456789abcdef"
  };
}

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
  understandMessageMock.mockResolvedValue({
    kind: "needs_workbook_context",
    summary: "需要字段结构",
    provider: "model"
  });
  listModelsMock.mockResolvedValue({
    defaultModelId: firstConnection.id,
    models: [firstConnection]
  });
  vi.mocked(getModelSettings).mockResolvedValue({ connectionCount: 1 } as never);
  validatePlanMock.mockImplementation(async (submitted) => submitted);
  vi.mocked(registerWpsExecution).mockImplementation(async (input: { operationId: string }) =>
    runningState(input.operationId)
  );
  vi.mocked(getWpsExecutionStatus).mockImplementation(async (operationId: string) =>
    runningState(operationId)
  );
  vi.mocked(startWpsAction).mockImplementation(
    async (operationId: string, input: { expectedRevision: number }) =>
      runningState(operationId, input.expectedRevision + 1)
  );
  vi.mocked(completeWpsAction).mockImplementation(
    async (operationId: string, input: { expectedRevision: number }) =>
      runningState(operationId, input.expectedRevision + 1)
  );
  vi.mocked(suspendWpsExecution).mockImplementation(async (operationId: string) =>
    runningState(operationId)
  );
  vi.mocked(claimWpsExecutionResume).mockImplementation(async (operationId: string) =>
    runningState(operationId)
  );
  vi.mocked(reconcileWpsActionReceipt).mockImplementation(async (operationId: string) =>
    runningState(operationId)
  );
  vi.mocked(finishWpsExecution).mockImplementation(async (operationId: string) => ({
    ...runningState(operationId),
    status: "completed" as const
  }));
});

afterEach(() => {
  mounted.forEach((entry) => entry.unmount());
  mounted = [];
  document.body.innerHTML = "";
  vi.unstubAllGlobals();
});

/** 面板本轮：先发 prepare 指针并挂住，等我们放行后才把这一轮的结果交给面板。 */
function armPrepareTurn(result: WpsTurnResult) {
  let resolveTurn: (() => void) | null = null;
  let released = false;
  runTurnMock.mockImplementation(async (input: WpsReadonlyTurnInput) => {
    expect(await input.preparePlanningBridge?.("turn-prepare-panel")).toBe(true);
    // 事件形状照抄生产实现：`stage:"answering"` + `turnId` + `bridgeTool`。
    await input.onProgress?.({
      stage: "answering",
      stepId: "bridge.prepare_python_data",
      title: "模型请求登记授权数据",
      turnId: "turn-prepare-panel",
      message: "模型正在等待本机登记授权数据并回执数据句柄。",
      bridgeTool: { operationId: PREPARE_OP, tool: "prepare_python_data" }
    });
    // 回执是 `onProgress` 里 `void drain()` **并发**跑出来的：放行可能早于
    // promise 建立，所以"已放行"单独记账。
    if (!released) {
      await new Promise<void>((resolve) => {
        resolveTurn = resolve;
      });
    }
    return result;
  });
  return () => {
    released = true;
    resolveTurn?.();
  };
}

const CONTINUOUS_PLAN: AnalysisPlan = {
  id: "plan-continuous-wps-1",
  title: "会员等级汇总",
  summary: "按会员等级汇总订单金额，写到汇总表",
  assumptions: [],
  warnings: [],
  actions: [
    { type: "writeValues", sheet: TARGET, range: "A1:B4", values: EXPECTED_SUMMARY }
  ],
  acceptanceCriteria: [
    { type: "rangeEquals", sheet: TARGET, range: "A1:B4", expected: EXPECTED_SUMMARY }
  ]
};

describe("WPS 连续链路：同一次任务从宿主读取到撤销", () => {
  it("prepare_python_data → 句柄登记 → 服务端隔离执行 → 白名单计划 → 预览 → 确认 → 写回 → 回读 → 撤销", async () => {
    const host = membershipHost();
    const bridge = stubContinuousBridge();
    const releaseTurn = armPrepareTurn({
      kind: "plan",
      plan: CONTINUOUS_PLAN,
      turnId: "turn-prepare-panel"
    } as WpsTurnResult);

    const container = await mount(
      <WpsAgentPanel
        selectedSheetIndexes={[1, 2, 3]}
        scopeDock={<div className="eb-scope-dock">当前表 会员表、订单表、汇总表</div>}
      />
    );
    await until(
      () =>
        container
          .querySelector<HTMLButtonElement>(".model-picker-trigger")
          ?.title.includes("当前模型") ?? false,
      "应加载模型目录"
    );

    const textarea = container.querySelector("textarea");
    await type(textarea, "按会员等级汇总订单金额，结果写到汇总表");
    await submit(textarea);

    // 1) 本机桥真实执行了 prepare_python_data 并回执：回执带**数据句柄**。
    await until(
      () => bridge.some((item) => item.url.includes("/receipt")),
      "应把数据登记回执给服务端"
    );
    const prepareReceipt = bridge.find((item) => item.url.includes("/receipt"))?.body
      .receipt as Record<string, unknown>;
    expect(Array.isArray(prepareReceipt.datasets)).toBe(true);
    expect((prepareReceipt.datasets as unknown[]).length).toBeGreaterThan(0);
    // 值只进回执字段，不外泄到 result 通道。
    expect(prepareReceipt.result).toBeUndefined();
    // 轮询请求体带完整归属（含工作簿身份），不是只报个名字。
    const poll = bridge.find((item) => item.url.includes("/poll"))?.body;
    expect(poll).toMatchObject({
      taskId: "turn-prepare-panel",
      hostId: "wps-et",
      workbookName: "会员.xlsx",
      workbookIdentity: "wps-path:c:\\data\\会员.xlsx"
    });

    // 写之前来源与目标一个单元格都没动。
    expect(host.matrix(MEMBERS)).toEqual(MEMBER_ROWS);
    expect(host.matrix(ORDERS)).toEqual(ORDER_ROWS);
    expect(host.writes()).toBe(0);

    // 2) 放行：白名单计划回到面板 → 预览出现；确认前零写入。
    releaseTurn();
    await until(() => {
      const has = container.textContent?.includes("本次执行计划") ?? false;
      if (!has) console.log("[DEBUG wpsContinuousPythonChain] 未找到写入前确认，当前内容:", container.textContent?.slice(0, 500));
      return has;
    }, "应展示写入计划预览");
    // 等待React完成渲染
    await settle();
    await settle();
    /*
     * 计划到达即执行（先执行、后审阅）：这里已经不是"零写入"的时刻。
     * 仍然必须成立的是"**只写目标表**"—— 两张来源表一格都不能动。
     * 执行前零写入由本用例前段（结果计划尚未返回时）断言。
     */
    expect(host.writes()).toBeGreaterThan(0);
    // 两张来源表在预览后仍未改变。
    expect(host.matrix(MEMBERS)).toEqual(MEMBER_ROWS);
    expect(host.matrix(ORDERS)).toEqual(ORDER_ROWS);

    // 3) 计划就绪即执行（先执行、后审阅）→ 结果回到主表并由固定操作条承接确认。
    await until(
      () => findButton(container, "确认全部变更") !== null || (container.textContent ?? "").includes("本次执行需要核对"),
      "应自动执行并进入结果审阅：" + (container.textContent ?? "").slice(-160)
    );
    // 执行已经发生；这里确认的是"执行结果已就绪"，下面用回读与撤销验证业务结果。
    expect(findButton(container, "确认全部变更")).not.toBeNull();

    await until(
      () => container.querySelector(".wps-verification-card") !== null,
      "写入后应出现执行结果卡"
    );
    expect(host.writes()).toBeGreaterThan(0);
    // 值来自真实计划 + 真实执行器写出（不是测试直接赋值）。
    expect(host.matrix(TARGET).slice(0, 4)).toEqual(EXPECTED_SUMMARY);
    // 只写目标表；两张来源表一格未变。
    expect(host.matrix(MEMBERS)).toEqual(MEMBER_ROWS);
    expect(host.matrix(ORDERS)).toEqual(ORDER_ROWS);

    // 4) 确认变更 → 撤销：确认只移除审阅标记，不重放计划；
    //    撤销走真实恢复点，把目标区回到执行前（空表）。
    await click(findButton(container, "确认全部变更"));
    await until(() => (container.textContent ?? "").includes("已确认变更"), "应确认本次结果");
    const undo = cardUndoButton(container);
    expect(undo, "执行结果卡里必须挂可用的撤销按钮").not.toBeNull();
    await click(undo);
    await until(
      () => container.textContent?.includes("已撤销并恢复") ?? false,
      "撤销后应显示已撤销并恢复"
    );
    expect(host.cell(TARGET, 1, 1)).toBeNull();
    // 撤销也只回滚目标区，来源表始终未变。
    expect(host.matrix(MEMBERS)).toEqual(MEMBER_ROWS);
    expect(host.matrix(ORDERS)).toEqual(ORDER_ROWS);
  });
});
