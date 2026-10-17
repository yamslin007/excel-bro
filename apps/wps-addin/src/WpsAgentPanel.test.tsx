import { readFileSync } from "node:fs";
import { act, useEffect, useState, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { renderToStaticMarkup } from "react-dom/server";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { MODEL_STORAGE_KEY } from "../../excel-addin/src/conversation";
import { useModelManagement } from "../../excel-addin/src/hooks/useModelManagement";
import { getModelSettings, listModels, understandMessage, validatePlan, registerWpsExecution, getWpsExecutionStatus, startWpsAction, completeWpsAction, finishWpsExecution, suspendWpsExecution, claimWpsExecutionResume, reconcileWpsActionReceipt } from "../../excel-addin/src/api";
import WpsAgentPanel, { wpsWriteBackBlockReason, wpsWriteBackStillCurrent } from "./WpsAgentPanel";
import type { ReviewWriteBackTarget } from "@excel-bro/core/reviewWriteBackGate";
import { matrixFingerprintV2 } from "./wpsPlanExecutor";
import WpsModelSettingsPanel from "./WpsModelSettingsPanel";
import {
  runWpsReadonlyTurn,
  type WpsReadonlyTurnInput,
  type WpsTurnResult
} from "./wpsReadonlyAgent";
import { getWpsApplication } from "./wpsHost";
import { inspectWpsAuthorizedTarget } from "./wpsAuthorizedTarget";
import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import {
  COMPOSER_AUTO_GROW_MAX_HEIGHT,
  COMPOSER_RESIZE_LABEL
} from "../../excel-addin/src/hooks/useComposerResize";
import { compareCommittedWpsResultToCopy, discardWpsNativeComparison, stageWpsNativeComparison, WpsFormulaTrialError } from "./wpsNativeComparison";
import type { AnalysisPlan } from "../../excel-addin/src/contracts";
import { loadTools } from "../../excel-addin/src/storage";

vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual =
    await importOriginal<typeof import("../../excel-addin/src/api")>();
  // CAP-03 服务端登记 API：面板测试不测真实网络，返回一个"已授予写入权"的
  // 成功状态，让写路径能走完动作级检查点。真实登记失败/中止由独立测试覆盖。
  const wpsRunningState = (
    operationId: string,
    revision = 1,
    overrides: Record<string, unknown> = {}
  ) => ({
    operationId,
    status: "running" as const,
    revision,
    workbookIdentity: "wps-path:c:\\data\\测试.xlsx",
    taskId: "task-test",
    descriptor: { actions: [] },
    authorizedSheetNames: ["Sheet1"],
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
    ...actual,
    understandMessage: vi.fn(async () => ({ kind: "needs_workbook_context", summary: "需要字段结构", provider: "model" })),
    listModels: vi.fn(),
    getModelSettings: vi.fn(),
    validatePlan: vi.fn(),
    registerWpsExecution: vi.fn(async (input: { operationId: string }) =>
      wpsRunningState(input.operationId)),
    getWpsExecutionStatus: vi.fn(async (operationId: string) => wpsRunningState(operationId)),
    startWpsAction: vi.fn(async (operationId: string, input: { expectedRevision: number }) =>
      wpsRunningState(operationId, input.expectedRevision + 1)),
    completeWpsAction: vi.fn(async (operationId: string, input: { expectedRevision: number }) =>
      wpsRunningState(operationId, input.expectedRevision + 1)),
    suspendWpsExecution: vi.fn(async (operationId: string) => wpsRunningState(operationId)),
    claimWpsExecutionResume: vi.fn(async (operationId: string) => wpsRunningState(operationId)),
    reconcileWpsActionReceipt: vi.fn(async (operationId: string) => wpsRunningState(operationId)),
    finishWpsExecution: vi.fn(async (operationId: string) => ({
      ...wpsRunningState(operationId),
      status: "completed" as const
    }))
  };
});

/*
 * 只替掉"那一轮规划怎么跑"，**不**替掉整个模块。
 *
 * 桥宿主适配器还要用同模块里的 `sheetsForRequest`（把读到的矩阵交给确定性计算）。
 * 早先这里是整体替换，于是桥的执行路径在本文件里永远跑不起来 —— 这正是
 * "面板接线测试"要覆盖的东西，被 mock 顺手藏掉了。
 */
vi.mock("./wpsReadonlyAgent", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./wpsReadonlyAgent")>()),
  runWpsReadonlyTurn: vi.fn()
}));

// 面板测试只验证“两次确认”状态机；原生复制、着色及失效检测由
// wpsNativeComparison.test.ts 的严格宿主替身独立覆盖。
// 注意：替身必须给全 `WpsNativeComparison` 的**每个**字段。少给字段时
// `vi.fn` 的类型很宽松、tsc 也放过它，但侧栏一旦真的用到就会在运行期崩
// —— 这跟宿主替身少给参数是同一类"替身比真实现宽容"的缺口。
vi.mock("./wpsNativeComparison", () => ({
  WpsFormulaTrialError: class WpsFormulaTrialError extends Error {},
  stageWpsNativeComparison: vi.fn((_application: unknown, plan: unknown) => {
    const prepared = plan as { actions: { values: unknown[][] }[] };
    return {
      plan,
      changedCells: ["A1"],
      reviews: [],
      original: { Name: "测试.xlsx" },
      copy: { Name: "Book1" },
      copyIdentity: { name: "Book1", fullName: null, path: null },
      originalBefore: prepared.actions.map(() => ({ values: [[]], formulas: [[]], valuesCaptured: true })),
      copyFingerprints: prepared.actions.map(() => "copy-fingerprint"),
      // 真实路径里回读值已与计划校对一致，替身照同一语义给。
      copyAfter: prepared.actions.map((action) => ({
        values: action.values,
        formulas: [],
        valuesCaptured: true
      })),
      copyFormatFinal: prepared.actions.map(() => undefined),
      originalFormatBefore: prepared.actions.map(() => undefined)
    };
  }),
  compareCommittedWpsResultToCopy: vi.fn(() => "matched"),
  verifyWpsNativeComparison: vi.fn(),
  // 宿主真实返回的是**关闭结果**（不再是 void/抛异常）。
  discardWpsNativeComparison: vi.fn(() => "closed")
}));

vi.mock("./wpsHost", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsHost")>();
  return { ...actual, getWpsApplication: vi.fn(() => ({})) };
});

/*
 * 授权身份检查是写回入口里**唯一**的异步等待。替身默认跑真实实现（本文件其它
 * 用例要靠它读真实身份），只有"在途窗口"那条用例会临时把它挂起 —— 这样才能把
 * "用户在这个窗口里点了放弃"这一个瞬间固定下来，而不是靠计时碰运气。
 */
vi.mock("./wpsAuthorizedTarget", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsAuthorizedTarget")>();
  return { ...actual, inspectWpsAuthorizedTarget: vi.fn(actual.inspectWpsAuthorizedTarget) };
});

// React 19 从 react 包导出 act；这里刻意不用 @testing-library/react ——
// 它只装在 excel-addin 工作区，会带进根目录那份 React（19.2.x），
// 与本工作区组件使用的 React（19.1.x）形成双实例，触发 Invalid hook call。
(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true;

// jsdom 不实现 scrollIntoView；面板只用它把新消息滚到底部，
// 在测试里补一个空实现即可，不需要断言滚动行为。
if (!Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => undefined;
}

const styles = readFileSync("src/styles.css", "utf8");
const excelStyles = readFileSync("../excel-addin/src/styles.css", "utf8");
const listModelsMock = vi.mocked(listModels);
const getModelSettingsMock = vi.mocked(getModelSettings);
const validatePlanMock = vi.mocked(validatePlan);
const registerWpsExecutionMock = vi.mocked(registerWpsExecution);
const getWpsExecutionStatusMock = vi.mocked(getWpsExecutionStatus);
const startWpsActionMock = vi.mocked(startWpsAction);
const completeWpsActionMock = vi.mocked(completeWpsAction);
const finishWpsExecutionMock = vi.mocked(finishWpsExecution);
const suspendWpsExecutionMock = vi.mocked(suspendWpsExecution);
const claimWpsExecutionResumeMock = vi.mocked(claimWpsExecutionResume);
const reconcileWpsActionReceiptMock = vi.mocked(reconcileWpsActionReceipt);
const runTurnMock = vi.mocked(runWpsReadonlyTurn);
const getWpsApplicationMock = vi.mocked(getWpsApplication);
const stageComparisonMock = vi.mocked(stageWpsNativeComparison);
const compareCommittedMock = vi.mocked(compareCommittedWpsResultToCopy);

const firstConnection = {
  id: "connection:first",
  label: "第一个连接",
  provider: "model" as const,
  available: true,
  supportsVision: false
};

const secondConnection = {
  id: "connection:second",
  label: "第二个连接",
  provider: "model" as const,
  available: true,
  supportsVision: true
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

/** 让已排队的 promise（模型目录读取等）落定，并让 React 处理由此产生的更新。 */
async function settle(): Promise<void> {
  await act(async () => {
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
  });
}

async function until(check: () => boolean, message: string): Promise<void> {
  for (let index = 0; index < 80; index += 1) {
    if (check()) return;
    await settle();
  }
  throw new Error(message);
}

function textOf(container: HTMLElement): string {
  return container.textContent ?? "";
}

function findButton(
  container: HTMLElement,
  name: string
): HTMLButtonElement | null {
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

/**
 * 「备份与恢复」抽屉里的撤销入口。抽屉仍保留这个快捷入口，所以先打开它，
 * 再按容器限定取按钮 —— 消息流里的执行结果卡现在也有同名按钮，
 * 只用文本查找会指错对象。
 */
async function undoButton(container: HTMLElement): Promise<HTMLButtonElement | null> {
  if (!container.querySelector(".wps-backup-card")) {
    await click(
      container.querySelector<HTMLButtonElement>('button[aria-label="备份与恢复"]')
    );
  }
  return container.querySelector<HTMLButtonElement>(".wps-backup-card button");
}

/** 执行结果卡里的撤销按钮（抽屉里那个不算）。 */
function cardUndoButton(container: HTMLElement): HTMLButtonElement | null {
  return container.querySelector<HTMLButtonElement>(
    '.wps-verification-card button[aria-label="撤销本次执行"]'
  );
}

/**
 * 消息流里所有名为「撤销本次执行」的按钮（抽屉里的不算）。每一个都必须落在
 * **某张卡**里 —— 早先那张无条件挂在消息流末尾的全局撤销卡会在这里暴露成
 * "没有卡片容器的裸按钮"。
 */
function flowUndoButtons(container: HTMLElement): HTMLButtonElement[] {
  return Array.from(container.querySelectorAll<HTMLButtonElement>("button")).filter(
    (item) =>
      (item.textContent ?? "").trim() === "↶ 撤销本次执行" &&
      !item.closest(".tool-drawer")
  );
}

/*
 * 把授权身份检查**挂起 / 放行**：让"旧请求在途"这个瞬间可以被固定下来。
 *
 * 不这么做就只能靠计时碰运气 —— 而碰运气的用例要么假红要么假绿。挂起的是
 * `inspectWpsAuthorizedTarget`，它是写回入口里唯一那个 `await`。
 */
const realInspect = vi.mocked(inspectWpsAuthorizedTarget).getMockImplementation();
let inspectHold: Array<() => void> = [];

function holdInspect(): void {
  vi.mocked(inspectWpsAuthorizedTarget).mockImplementation(
    (selected) =>
      new Promise<AuthorizedTarget>((resolve, reject) => {
        inspectHold.push(() => {
          void realInspect?.(selected).then(resolve, reject);
        });
      })
  );
  inspectHold = [];
}
const inspectPending = (): boolean => inspectHold.length > 0;
async function releaseInspect(): Promise<void> {
  if (realInspect) vi.mocked(inspectWpsAuthorizedTarget).mockImplementation(realInspect);
  const queued = inspectHold;
  inspectHold = [];
  await act(async () => {
    queued.forEach((run) => run());
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
  });
}

/**
 * 审阅卡里的「我已核对副本」确认项。
 *
 * 写回入口在勾选前是**禁用**的 —— 确认框是真闸，加载项不替用户回答"我看过了"。
 * 所以凡是要走完写回链路的用例都必须先勾这里；直接点按钮只会被拦下，原表一格都不写。
 * 返回该复选框，供需要断言"勾选前后按钮可用性"的用例复用。
 */
async function acknowledgeCopy(container: HTMLElement): Promise<HTMLInputElement> {
  const review = container.querySelector<HTMLElement>(".wps-native-review-status");
  const checkbox = review?.querySelector<HTMLInputElement>('input[type="checkbox"]') ?? null;
  if (!checkbox) throw new Error("审阅卡里应有「我已核对副本」确认项");
  if (!checkbox.checked) await click(checkbox);
  return checkbox;
}

async function stageAndSave(container: HTMLElement): Promise<void> {
  await until(() => findButton(container, "确认变更") !== null || textOf(container).includes("本次执行需要核对"), "应完成主表执行或明确失败");
  const confirm = findButton(container, "确认变更");
  if (confirm) {
    expect(confirm.disabled).toBe(false);
    await click(confirm);
    await until(() => textOf(container).includes("已确认变更"), "确认应完成");
  }
}

/** 受控输入必须走原生 value setter，否则 React 收不到这次变更。 */
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

function panel() {
  return (
    <WpsAgentPanel
      selectedSheetIndexes={[1]}
      scopeDock={<div className="eb-scope-dock">当前表 Sheet1</div>}
    />
  );
}

/**
 * 最小可用宿主：只提供**澄清复核**需要的东西 —— 一个稳定的工作簿对象、
 * 一张稳定的工作表对象、以及一个结构不变的已用区域。
 *
 * 澄清复核现在额外要求"能证明还是建卡时那个对象"（见 `wpsObjectIdentity`），
 * 所以凡是要走完澄清作答的用例都得先有一个真实宿主：给 `{}` 会 fail-closed。
 * 这里刻意把同一份 `sheet`/`workbook` 对象复用于每次读取，模拟真宿主
 * "同一张表始终是同一个包装对象"的行为。
 */
function clarificationHost(options: { sheetNames?: string[]; workbookName?: string } = {}) {
  const sheetNames = options.sheetNames ?? ["Sheet1"];
  const usedRange = { Rows: { Count: 3 }, Columns: { Count: 2 }, Row: 1, Column: 1 } as Record<
    string,
    unknown
  >;
  Object.defineProperties(usedRange, {
    Value2: { get: () => [["类别", "归属"], ["甲", "华东"], ["乙", "华南"]] },
    Item: { value: () => usedRange },
    MergeCells: { value: false }
  });
  const range = { Rows: { Count: 1 }, Columns: { Count: 2 }, MergeCells: false } as Record<
    string,
    unknown
  >;
  Object.defineProperties(range, {
    Value2: { get: () => [["类别", "归属"]] }
  });
  const sheets = sheetNames.map((Name, position) => ({
    Name,
    Index: position + 1,
    CodeName: `SheetCode${position + 1}`,
    ProtectContents: false,
    UsedRange: usedRange,
    Range: () => range
  }));
  const workbook = {
    Name: options.workbookName ?? "测试.xlsx",
    FullName: `C:\\Data\\${options.workbookName ?? "测试.xlsx"}`,
    Path: "C:\\Data",
    ReadOnly: false,
    Worksheets: { Count: sheets.length, Item: (index: number) => sheets[Number(index) - 1] },
    ActiveSheet: sheets[0]
  };
  const storage = new Map<string, unknown>();
  return {
    Version: "12.0",
    ActiveWorkbook: workbook,
    ActiveSheet: sheets[0],
    PluginStorage: {
      getItem: (key: string) => storage.get(key),
      setItem: (key: string, item: unknown) => { storage.set(key, item); }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as never;
}

/**
 * 一块替身单元格。默认写入生效（执行器据此判定回读一致）；
 * `staleWrite` 时写入不落盘，于是回读核对**真的**失败 ——
 * 这样才能断言界面在失败时不谎报"验证通过"，而不是靠 mock 抛错绕开真实链路。
 */function writeRangeHost(options: {
  staleWrite?: boolean;
  /** 每次宿主写入后回调（参数是累计写入次数），用来在"第一次写入发生的那一刻"取证。 */
  onWrite?: (writes: number) => void;
  /** 写入一旦发生就让 `PluginStorage.getItem` 抛错：模拟恢复点读取失败。 */
  breakStorageReadsAfterWrite?: boolean;
  /** 带逐格格式状态（Font/对齐/换行/行高），供格式动作（setFont 等）走通。 */
  withFormat?: boolean;
} = {}) {
  let value = "旧值";
  let formula = "旧值";
  let writes = 0;
  const format = { bold: false, italic: false, name: "宋体", size: 11, color: 0, horizontalAlignment: 1, wrapText: false, rowHeight: 13.5 };
  const range = { Rows: { Count: 1 }, Columns: { Count: 1 }, MergeCells: false } as Record<string, unknown>;
  if (options.withFormat) {
    range.Font = {
      get Bold() { return format.bold; }, set Bold(v: unknown) { format.bold = v as boolean; },
      get Italic() { return format.italic; }, set Italic(v: unknown) { format.italic = v as boolean; },
      get Name() { return format.name; }, set Name(v: unknown) { format.name = v as string; },
      get Size() { return format.size; }, set Size(v: unknown) { format.size = v as number; },
      get Color() { return format.color; }, set Color(v: unknown) { format.color = v as number; }
    };
    Object.defineProperties(range, {
      HorizontalAlignment: {
        get: () => format.horizontalAlignment,
        set: (v: unknown) => { format.horizontalAlignment = v as number; },
        enumerable: true
      },
      WrapText: {
        get: () => format.wrapText,
        set: (v: unknown) => { format.wrapText = v as boolean; },
        enumerable: true
      },
      RowHeight: {
        get: () => format.rowHeight,
        set: (v: unknown) => { format.rowHeight = v as number; },
        enumerable: true
      }
    });
  }
  Object.defineProperties(range, {
    Value2: {
      get: () => value,
      set: (next: string) => {
        writes += 1;
        if (!options.staleWrite) {
          value = next;
          formula = next;
        }
        options.onWrite?.(writes);
      },
      enumerable: true
    },
    Formula: {
      get: () => formula,
      set: (next: string) => {
        writes += 1;
        if (!options.staleWrite) {
          formula = next;
          value = next;
        }
        options.onWrite?.(writes);
      },
      enumerable: true
    }
  });
  const sheet = { Name: "Sheet1", ProtectContents: false, Range: () => range };
  const storage = new Map<string, unknown>();
  getWpsApplicationMock.mockReturnValue({
    Version: "12.0",
    ActiveWorkbook: {
      Name: "测试.xlsx",
      FullName: "C:\\Data\\测试.xlsx",
      Path: "C:\\Data",
      ReadOnly: false,
      Worksheets: { Count: 1, Item: () => sheet },
      ActiveSheet: sheet
    },
    ActiveSheet: sheet,
    PluginStorage: {
      getItem: (key: string) => {
        if (options.breakStorageReadsAfterWrite && writes > 0) {
          throw new Error("PluginStorage 读取失败");
        }
        return storage.get(key);
      },
      setItem: (key: string, item: unknown) => { storage.set(key, item); }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as never);
  return {
    writes: () => writes,
    cellValue: () => value,
    format: () => format,
    /** 模拟"执行之后用户又手工改了这块区域"：撤销前指纹复核必须能识别出来。 */
    setCell: (next: string) => {
      value = next;
      formula = next;
    }
  };
}

/** 走到"写入前确认"这一步：两端对照用的同一任务状态。 */
async function mountPlanPreview(plan: AnalysisPlan, task: string) {
  listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
  runTurnMock.mockResolvedValue({ kind: "plan", plan, turnId: "turn-preview-1" });
  validatePlanMock.mockImplementation(async (submitted) => submitted);
  const { container } = await mount(panel());
  await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型目录");
  const textarea = container.querySelector("textarea");
  await type(textarea, task);
  await submit(textarea);
  await until(() => findButton(container, "确认变更") !== null || textOf(container).includes("本次执行需要核对"), "应展示执行结果或明确失败");
  return { container, unmount: () => unmountMounted(container) };
}

/**
 * 卸载历史里挂载过的某个面板，并从 `mounted` 里摘掉它 —— `mount` 的清理是
 * 按数组走的，重复卸载同一个 root 会报错。供"面板重开"这类用例使用。
 */
function unmountMounted(container: HTMLElement): void {
  const entry = mounted.find((item) => item.container === container);
  if (!entry) return;
  entry.unmount();
  mounted = mounted.filter((item) => item !== entry);
}

const singleWritePlan: AnalysisPlan = {
  id: "plan-single-write",
  title: "更新 A1",
  summary: "把 A1 更新为新值",
  assumptions: [],
  warnings: [],
  actions: [{ type: "writeValues", sheet: "Sheet1", range: "A1", values: [["90"]] }],
  acceptanceCriteria: []
};

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
  // 上一例若把授权检查挂起过，这里必须还原：否则后续用例的身份检查永远不返回。
  if (realInspect) vi.mocked(inspectWpsAuthorizedTarget).mockImplementation(realInspect);
  inspectHold = [];
  vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "需要字段结构", provider: "model" });
  getWpsApplicationMock.mockReturnValue({ ...(clarificationHost() as unknown as object), Version: "99" } as never);
});

afterEach(() => {
  mounted.forEach((entry) => entry.unmount());
  mounted = [];
  document.body.innerHTML = "";
});

describe("WPS Excel Bro chat UI", () => {
  it("when the real scope entry is available, target changes use it and never open a second dialog", async () => {
    let app = clarificationHost();
    getWpsApplicationMock.mockImplementation(() => app);
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({ kind: "answer", provider: "system", message: "已重新读取所选表" });
    const open = vi.fn();
    let authorize: ((target: AuthorizedTarget) => Promise<boolean>) | null = null;
    function ExistingScopeEntry() {
      const [pickerOpen, setPickerOpen] = useState(false);
      return <WpsAgentPanel selectedSheetIndexes={[1]} scopeDock={<span>既有数据范围选择器</span>}
        scopePickerOpen={pickerOpen} onOpenScope={() => { open(); setPickerOpen(true); }}
        onScopeConfirmationReady={(handler) => { authorize = handler; }} />;
    }
    const { container } = await mount(<ExistingScopeEntry />);
    await until(() => authorize !== null && Boolean(container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型")), "连接就绪");
    app = clarificationHost({ workbookName: "另一份.xlsx" });
    const textarea = container.querySelector("textarea");
    await type(textarea, "按我刚才补充的规则处理"); await submit(textarea);
    await until(() => Boolean(container.querySelector(".target-scope-notice")), "范围条提示");
    expect(container.querySelector('[role="dialog"]')).toBeNull();
    expect(container.querySelector('input[type="radio"]')).toBeNull();
    expect(container.querySelectorAll(".target-scope-notice button")).toHaveLength(1);
    expect(textOf(container.querySelector(".target-scope-notice")!)).not.toContain("暂不操作");
    await act(async () => { window.dispatchEvent(new Event("focus")); });
    expect(container.querySelector(".target-scope-notice")).not.toBeNull();
    expect(runTurnMock).not.toHaveBeenCalled();
    expect(textarea?.value).toBe("按我刚才补充的规则处理");
    await click(findButton(container, "检查数据范围")); expect(open).toHaveBeenCalledOnce();
    expect(container.querySelector(".target-scope-notice")).toBeNull();
    expect(runTurnMock).not.toHaveBeenCalled();
    await act(async () => { expect(await authorize!(await inspectWpsAuthorizedTarget([1]))).toBe(true); });
    await until(() => runTurnMock.mock.calls.length === 1, "现有范围确认后继续分析");
    expect(runTurnMock.mock.calls[0][0].prompt).toBe("按我刚才补充的规则处理");
    expect(container.querySelector(".target-scope-notice")).toBeNull();
  });
  it.each(["谢谢帮助", "你是谁"])("ordinary message %s does not invoke the workbook agent or show a field-read step", async (question) => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    vi.mocked(understandMessage).mockResolvedValue({ kind: "conversation", message: "不客气，随时告诉我。", provider: "model" });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型") ?? false, "应加载模型");
    const textarea = container.querySelector("textarea");
    await type(textarea, question);
    await submit(textarea);
    await until(() => textOf(container).includes("不客气，随时告诉我。"), "应直接回复");
    expect(runTurnMock).not.toHaveBeenCalled();
    const toggle = container.querySelector<HTMLButtonElement>(".task-progress-toggle");
    expect(toggle?.getAttribute("aria-expanded")).toBe("false");
    expect(textOf(container)).not.toContain("理解消息");
    await click(toggle);
    expect(toggle?.getAttribute("aria-expanded")).toBe("true");
    expect(textOf(container)).toContain("理解消息");
    expect(textOf(container)).not.toContain("读取已授权表结构");
    expect(vi.mocked(understandMessage).mock.calls[0][0]).not.toHaveProperty("scope");
  });

  it("a failed message-only route keeps the request and never invokes the workbook agent", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    vi.mocked(understandMessage).mockRejectedValue(new Error("消息理解失败"));
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型") ?? false, "应加载模型");
    const textarea = container.querySelector("textarea");
    await type(textarea, "检查当前表");
    await submit(textarea);
    await until(() => textOf(container).includes("消息理解失败"), "应提示失败");
    expect(runTurnMock).not.toHaveBeenCalled();
    expect(textarea?.value).toBe("检查当前表");
    expect(textOf(container)).not.toContain("读取已授权表结构");
  });
  it("confirmed range with fresh native wrappers is reused for consecutive follow-ups, but another file still pauses", async () => {
    let app: ReturnType<typeof getWpsApplication> = clarificationHost();
    getWpsApplicationMock.mockImplementation(() => ({ ...app,
      ActiveWorkbook: { ...app.ActiveWorkbook }
    } as never));
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({ kind: "answer", provider: "system", message: "追问已分析" });
    let authorize: ((target: AuthorizedTarget) => Promise<boolean>) | null = null;
    const { container } = await mount(<WpsAgentPanel selectedSheetIndexes={[1]} scopeDock={<span>当前表 Sheet1</span>}
      onScopeConfirmationReady={(handler) => { authorize = handler; }} />);
    await until(() => authorize !== null && Boolean(container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：")), "确认入口与模型就绪");
    await act(async () => { expect(await authorize!(await inspectWpsAuthorizedTarget([1]))).toBe(true); });
    expect(runTurnMock).not.toHaveBeenCalled();
    const textarea = container.querySelector("textarea");
    for (let turn = 1; turn <= 2; turn++) {
      await type(textarea, `补充第${turn}条规则`); await submit(textarea);
      await until(() => runTurnMock.mock.calls.length === turn, "连续追问不重复选表");
      await until(() => !container.querySelector(".composer textarea:disabled"), "等待本轮结束");
      expect(container.querySelector('[role="dialog"]')).toBeNull();
    }
    app = clarificationHost({ workbookName: "另一份.xlsx" });
    await type(textarea, "接着处理新问题"); await submit(textarea);
    await until(() => Boolean(container.querySelector('[role="dialog"]')), "换文件仍然暂停");
    expect(runTurnMock).toHaveBeenCalledTimes(2);
    expect(JSON.stringify(localStorage)).not.toContain("C:\\\\Data");
  });
  it("发送前检测文件变化：零模型请求，用户选文件并勾表确认后才继续原输入", async () => {
    let app = clarificationHost() as unknown as { ActiveWorkbook: { Name: string; FullName: string } };
    getWpsApplicationMock.mockImplementation(() => app as never);
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({ kind: "answer", provider: "system", message: "已按新目标分析" });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "模型加载");
    app = clarificationHost({ workbookName: "另一份.xlsx" }) as never;
    const textarea = container.querySelector("textarea");
    await type(textarea, "按已经确认的规则处理");
    await submit(textarea);
    await until(() => Boolean(container.querySelector('[role="dialog"]')), "发送前弹出目标确认");
    expect(runTurnMock).not.toHaveBeenCalled();
    expect(textarea?.value).toBe("按已经确认的规则处理");
    const dialog = container.querySelector('[role="dialog"]')!;
    expect(Array.from(dialog.querySelectorAll<HTMLInputElement>('input[type="radio"]')).every((input) => !input.checked)).toBe(true);
    await click(dialog.querySelectorAll('input[type="radio"]')[1]);
    expect(findButton(container, "确认目标并继续")?.disabled).toBe(true);
    await click(dialog.querySelector('input[type="checkbox"]'));
    await click(findButton(container, "确认目标并继续"));
    await until(() => runTurnMock.mock.calls.length === 1, "确认后继续需求");
    expect(runTurnMock.mock.calls[0][0].prompt).toBe("按已经确认的规则处理");
    expect(runTurnMock.mock.calls[0][0].selectedSheetIndexes).toEqual([1]);
    expect(textOf(container)).not.toContain("执行失败");
    expect(JSON.stringify(localStorage)).not.toContain("C:\\\\Data");
  });

  it("重新获得焦点就提醒变化，不必先点击发送或让任务失败", async () => {
    let app = clarificationHost();
    getWpsApplicationMock.mockImplementation(() => app);
    const { container } = await mount(panel());
    await settle();
    app = clarificationHost({ workbookName: "新文件.xlsx" });
    await act(async () => { window.dispatchEvent(new Event("focus")); });
    await until(() => Boolean(container.querySelector('[role="dialog"]')), "焦点回来主动提示");
    expect(runTurnMock).not.toHaveBeenCalled();
    expect(textOf(container)).toContain("操作已暂停");
  });
  it("shows a structured planning stop as failure, not completed or a write preview", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({ kind: "answer", provider: "system", outcome: "failed", message: "规划预算已用完" });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "模型加载");
    const textarea = container.querySelector("textarea");
    await type(textarea, "生成预览");
    await submit(textarea);
    await until(() => textOf(container).includes("规划预算已用完"), "失败提示");
    expect(textOf(container)).toContain("执行失败");
    expect(textOf(container)).not.toContain("已完成");
    expect(textOf(container)).not.toContain("生成右侧副本并排审阅");
    expect(validatePlanMock).not.toHaveBeenCalled();
  });
  it("launches the read-only wide-window probe from the real chat pane without replacing the conversation", async () => {
    const showDialog = vi.fn((..._args: unknown[]) => true);
    getWpsApplicationMock.mockReturnValue({
      Version: "12.0",
      ShowDialog: showDialog
    } as never);
    listModelsMock.mockResolvedValue({ defaultModelId: null, models: [] });

    const { container } = await mount(panel());
    const textarea = container.querySelector("textarea");
    await type(textarea, "当前对话仍在这里");
    await click(container.querySelector('[aria-label="更多"]'));
    const probeEntry = Array.from(container.querySelectorAll('[role="menuitem"]'))
      .find((item) => item.textContent?.includes("探测独立宽窗口"));
    await click(probeEntry ?? null);

    expect(showDialog).toHaveBeenCalledTimes(1);
    expect(showDialog.mock.calls[0][4]).toBe(false);
    expect(container.querySelector('[aria-label="宽窗口探测结果"]')).not.toBeNull();
    expect(container.querySelector("textarea")?.value).toBe("当前对话仍在这里");
    expect(container.querySelector(".wps-chat-shell")).not.toBeNull();
  });

  it("keeps header actions icon-only, named and connected to their existing surfaces", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "模型加载完成");
    expect(container.querySelector(".brand-copy")?.textContent?.trim()).toBe("Excel Bro");
    for (const [selector, glyph] of [
      [".model-picker-trigger", ""],
      [".new-chat-entry", "＋"],
      [".tools-entry", ""],
      [".folder-backup-entry", "↶"]
    ]) {
      const button = container.querySelector<HTMLButtonElement>(selector);
      expect(button?.textContent).toBe(glyph);
      expect(button?.getAttribute("aria-label")).toBeTruthy();
      expect(button?.title).toBeTruthy();
      expect(button?.querySelector('svg, [aria-hidden="true"]')).not.toBeNull();
    }
    await click(container.querySelector(".model-picker-trigger"));
    expect(container.querySelector('.model-menu[role="menu"]')).not.toBeNull();
    expect(container.querySelector('[role="menuitemradio"][aria-checked="true"]')?.textContent).toContain("第一个连接");
    await click(container.querySelector(".tools-entry"));
    expect(container.querySelector('aside[aria-label="我的工具"]')).not.toBeNull();
    await click(container.querySelector(".folder-backup-entry"));
    expect(container.querySelector('aside[aria-label="备份与恢复"]')).not.toBeNull();
    expect(runTurnMock).not.toHaveBeenCalled();
  });

  it("shows the Excel-style chat shell and keeps guarded write boundaries visible", () => {
    const markup = renderToStaticMarkup(panel());

    expect(markup).toContain("wps-chat-shell chat-shell");
    expect(markup).toContain("model-picker-trigger");
    const modelTrigger =
      markup.match(/<button class="model-picker-trigger"([^>]*)>/)?.[1] ?? "";
    expect(modelTrigger).not.toContain("disabled");
    // 目录尚未加载时不得出现任何“基础模式”一类占位模型。
    expect(markup).not.toContain("基础模式");
    expect(markup).toContain("正在读取模型配置");
    expect(markup).toContain('class="composer"');
    expect(markup).toContain('class="app-logo"');
    expect(markup).toContain('aria-label="隐藏格仔"');
    expect(markup).toContain("pet-float mood-idle");
    expect(markup).toContain("当前表 Sheet1");
    expect(markup).toContain("新对话");
    expect(markup).toContain("工具");
    expect(markup).toContain("备份");
    expect(markup).toContain("当前仍可查询");
    expect(markup).toContain("theme-panel");
    expect(markup).toContain("预设主题");
    expect(markup).toContain("自定义背景");
    expect(markup).not.toContain("删除工作表");
  });

  it("offers the add-model entry when the catalog loads empty", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: null, models: [] });

    const { container } = await mount(panel());
    await until(
      () => textOf(container).includes("尚未配置模型"),
      "目录加载成功后应显示未配置状态"
    );

    expect(container.querySelector('.model-picker-trigger')?.getAttribute("aria-label")).toBe("添加模型");
    expect(textOf(container)).not.toContain("基础模式");
  });

  it("keeps the typed task and guides to add a model instead of answering", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: null, models: [] });

    const { container } = await mount(panel());
    await until(
      () => textOf(container).includes("尚未配置模型"),
      "目录加载成功后应显示未配置状态"
    );

    const textarea = container.querySelector("textarea");
    await type(textarea, "按部门统计平均分");
    await submit(textarea);

    expect(textOf(container)).toContain("尚未配置模型。添加一个模型连接后");
    // 任务文本必须留在输入框里，不能因为缺少模型就被清掉。
    expect(textarea?.value).toBe("按部门统计平均分");
    expect(runTurnMock).not.toHaveBeenCalled();
    expect(container.querySelector(".message-row.system")).toBeNull();
  });

  it("explains a disconnected backend without exposing Failed to fetch and keeps the task", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockRejectedValue(new TypeError("Failed to fetch"));
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");

    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "按部门统计平均分");
    await submit(textarea);
    await until(() => textOf(container).includes("无法连接 Excel Bro 后端"), "应解释连接失败");

    expect(textOf(container)).not.toContain("Failed to fetch");
    expect(textarea?.value).toBe("按部门统计平均分");
    expect(textOf(container)).toContain("本次未修改工作表");
  });

  it("does not restore an executable plan from saved chat history", async () => {
    const now = new Date().toISOString();
    localStorage.setItem("excel-bro.wps.chat.v1", JSON.stringify({
      activeConversationId: "old-chat",
      conversations: [{
        id: "old-chat", title: "修改 A1", createdAt: now, updatedAt: now,
        messages: [
          { id: "question", role: "user", text: "修改 A1", createdAt: now },
          { id: "old-plan", role: "assistant", text: "", createdAt: now,
            planPreview: { title: "更新 A1", status: "pending" } }
        ]
      }]
    }));
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    const { container } = await mount(panel());
    expect(textOf(container)).toContain("计划预览已失效");
    expect(container.querySelector(".wps-plan-card")).toBeNull();
    expect(findButton(container, "生成右侧副本并排审阅")).toBeNull();
  });

  it("does not present a refreshed in-flight WPS activity as still running", async () => {
    const now = new Date().toISOString();
    localStorage.setItem("excel-bro.wps.chat.v1", JSON.stringify({
      activeConversationId: "old-chat",
      conversations: [{
        id: "old-chat", title: "修改 A1", createdAt: now, updatedAt: now,
        messages: [{
          id: "activity", role: "assistant", text: "", createdAt: now,
          activityLog: { totalMs: 120, steps: [{ id: "writing:0", label: "执行第 1 个写入动作", elapsedMs: 120, status: "progress" }] }
        }]
      }]
    }));
    const { container } = await mount(panel());
    const timeline = container.querySelector(".task-progress-timeline");
    expect(timeline?.getAttribute("data-status")).toBe("interrupted");
    expect(timeline?.textContent).toContain("此过程未取得终态");
    expect(timeline?.classList.contains("is-live")).toBe(false);
  });

  it("shows task progress beside the question that started it", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    let finishTurn: ((value: Awaited<ReturnType<typeof runWpsReadonlyTurn>>) => void) | undefined;
    runTurnMock.mockImplementation((input) => {
      input.onProgress?.({ stage: "intent", stepId: "intent", title: "理解需求", message: "正在判断任务类型" });
      return new Promise((resolve) => { finishTurn = resolve; });
    });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");

    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "统计当前表");
    await submit(textarea);
    const question = Array.from(container.querySelectorAll(".message-row.user"))
      .find((item) => item.textContent?.includes("统计当前表"));
    const processMessage = question?.nextElementSibling;
    expect(processMessage?.classList.contains("message-row")).toBe(true);
    expect(processMessage?.querySelector(".message-author")?.textContent?.trim()).toBe("Excel Bro");
    expect(processMessage?.querySelector(".task-progress-timeline.is-live")).not.toBeNull();
    expect(processMessage?.textContent).toContain("理解需求");

    await act(async () => {
      finishTurn?.({ kind: "answer", message: "统计完成", resultContext: null, provider: "model" });
    });
    expect(processMessage?.querySelector(".task-progress-timeline.is-live")).toBeNull();
    expect(processMessage?.querySelector(".task-progress-timeline.activity-log-card")).not.toBeNull();
    expect(processMessage?.textContent).toContain("已完成");
    const saved = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    expect(saved.conversations[0].messages.find((message: { activityLog?: { outcome?: string } }) =>
      message.activityLog?.outcome === "completed")).toBeDefined();
  });

  it("copies both sides of the conversation when the embedded clipboard API rejects", async () => {
    const clipboardDescriptor = Object.getOwnPropertyDescriptor(navigator, "clipboard");
    const commandDescriptor = Object.getOwnPropertyDescriptor(document, "execCommand");
    const writeText = vi.fn().mockRejectedValue(new Error("not allowed in WPS WebView"));
    const legacyCopy = vi.fn().mockReturnValue(true);
    Object.defineProperty(navigator, "clipboard", { configurable: true, value: { writeText } });
    Object.defineProperty(document, "execCommand", { configurable: true, value: legacyCopy });
    try {
      listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
      runTurnMock.mockResolvedValue({ kind: "answer", message: "这是一条回答", resultContext: null, provider: "model" });
      const { container } = await mount(panel());
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
      const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
      await type(textarea, "这是一条提问");
      await submit(textarea);
      await until(() => textOf(container).includes("这是一条回答"), "应收到回答");

      const userCopy = container.querySelector<HTMLButtonElement>(".message-row.user .message-copy-button");
      expect(userCopy?.getAttribute("aria-label")).toBe("复制这条消息");
      expect(userCopy?.title).toBe("复制这条消息");
      expect(userCopy?.querySelector("svg rect")).not.toBeNull();
      expect(userCopy?.textContent).toBe("");
      await click(userCopy);
      expect(legacyCopy).toHaveBeenCalledTimes(1);
      expect(container.querySelector(".message-row.user .message-copy-button")?.textContent).toContain("已复制");
      const answer = Array.from(container.querySelectorAll(".message-row.assistant"))
        .find((item) => item.textContent?.includes("这是一条回答"));
      const answerCopy = answer?.querySelector<HTMLButtonElement>(".message-copy-button");
      expect(answerCopy?.querySelector("svg rect")).not.toBeNull();
      expect(answerCopy?.textContent).toBe("");
      await click(answerCopy ?? null);
      expect(writeText).toHaveBeenCalledTimes(2);
      expect(legacyCopy).toHaveBeenCalledTimes(2);
      expect(answer?.querySelector(".message-copy-button")?.textContent).toContain("已复制");
    } finally {
      if (clipboardDescriptor) Object.defineProperty(navigator, "clipboard", clipboardDescriptor);
      else Reflect.deleteProperty(navigator, "clipboard");
      if (commandDescriptor) Object.defineProperty(document, "execCommand", commandDescriptor);
      else Reflect.deleteProperty(document, "execCommand");
    }
  });

  it("shows query source, row count and warnings instead of dropping result metadata", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({
      kind: "answer", provider: "local", message: "查询完成",
      resultContext: {
        kind: "table", title: "按类别统计", headers: ["类别", "数量"],
        rows: [["A", 2]], sourceSheets: ["Sheet1"], warnings: ["仅统计非空类别"]
      }
    });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "统计类别");
    await submit(textarea);
    await until(() => textOf(container).includes("查询完成"), "应显示查询结果");
    expect(container.querySelector(".wps-query-result-meta")?.textContent).toContain("按类别统计");
    expect(container.querySelector(".wps-query-result-meta")?.textContent).toContain("显示 1 / 1 行");
    expect(container.querySelector(".wps-query-result-meta")?.textContent).toContain("数据来源：Sheet1");
    expect(container.querySelector(".wps-query-result-warnings")?.textContent).toContain("仅统计非空类别");
  });

  it("sends the model id shown in the picker and migrates a stale local record", async () => {
    localStorage.setItem(MODEL_STORAGE_KEY, "local");
    listModelsMock.mockResolvedValue({
      defaultModelId: secondConnection.id,
      models: [firstConnection, secondConnection]
    });
    runTurnMock.mockResolvedValue({
      kind: "answer",
      message: "已统计完成。",
      resultContext: null,
      provider: "model"
    });

    const { container } = await mount(panel());
    await until(
      () => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第二个连接") ?? false,
      "应解析出服务端声明的默认真实连接"
    );

    const textarea = container.querySelector("textarea");
    await type(textarea, "按部门统计平均分");
    await submit(textarea);

    await until(
      () => runTurnMock.mock.calls.length === 1,
      "有可用模型时应发起只读查询"
    );
    // 请求携带的必须是目录里解析出来的真实连接：既不是旧 local，
    // 也不是列表首项（首项是 connection:first）。
    expect(runTurnMock.mock.calls[0][0].modelId).toBe(secondConnection.id);
    // 旧的 local 记录已被迁移出本地存储。
    expect(localStorage.getItem(MODEL_STORAGE_KEY)).toBe(secondConnection.id);
  });

  it("shows no card for ordinary chat and accepts free text only when clarification is needed", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    // 澄清复核要能证明"还是建卡时那个对象"，所以这里必须给一个真实宿主。
    getWpsApplicationMock.mockReturnValue(clarificationHost());
    runTurnMock
      .mockResolvedValueOnce({ kind: "answer", message: "你好！", resultContext: null, provider: "model" })
      .mockResolvedValueOnce({
        kind: "clarification",
        clarification: {
          response: {
            kind: "clarification",
            provider: "model",
            clarification: {
              id: "clarify-1", summary: "需要明确目标", question: "想怎样处理？",
              reason: "不同操作结果不同", scopeLabel: "当前表",
              options: [
                { id: "read", label: "查询", description: "只读分析", resolution: "查询数据" },
                { id: "write", label: "修改", description: "预览后修改", resolution: "修改数据" }
              ]
            }
          }
        }
      })
      .mockResolvedValueOnce({ kind: "answer", message: "我来按你的补充说明处理。", resultContext: null, provider: "model" });

    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
    const textarea = container.querySelector("textarea");
    await type(textarea, "你好");
    await submit(textarea);
    await until(() => textOf(container).includes("你好！"), "问候应直接得到回复");
    expect(container.querySelector(".clarification-card")).toBeNull();

    await type(textarea, "帮我处理表格");
    await submit(textarea);
    await until(() => Boolean(container.querySelector(".clarification-card")), "关键歧义才显示卡片");
    const custom = container.querySelector<HTMLTextAreaElement>('[aria-label="其他想法"]');
    await type(custom, "只统计记录数，不修改表格");
    await click(findButton(container, "发送"));
    await until(() => runTurnMock.mock.calls.length === 3, "补充说明应继续原任务");
    expect(runTurnMock.mock.calls[2][0].prompt).toContain("补充确认：只统计记录数，不修改表格");
    expect(runTurnMock.mock.calls[2][0].intentConfirmed).toBe(true);
    expect(runTurnMock.mock.calls[2][0].taskContext).toEqual({ schemaVersion: 1,
      originalPrompt: "帮我处理表格", confirmations: [
        { questionId: "clarify-1", label: "补充说明", value: "只统计记录数，不修改表格" }
      ] });
    expect(container.querySelector(".clarification-card.resolved")).not.toBeNull();
    expect(textOf(container)).toContain("已按「只统计记录数，不修改表格」继续");
    const stored = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    expect(stored.conversations[0].messages.some((message: { clarification?: { status: string } }) =>
      message.clarification?.status === "resolved"
    )).toBe(true);
    expect(JSON.stringify(stored)).not.toContain("修改数据");
  });

  it("replacing a task resets its persisted confirmations as well as the next request", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    getWpsApplicationMock.mockReturnValue(clarificationHost());
    const question = (id: string, label: string, resolution: string, mode: "append" | "replace" = "append") => ({
      kind: "clarification" as const, clarification: { response: {
        kind: "clarification" as const, provider: "model" as const, clarification: {
          id, summary: "确认口径", question: "选择口径", reason: "尚缺信息", scopeLabel: "Sheet1",
          options: [{ id: "yes", label, description: label, resolution, resolutionMode: mode }]
        }
      } }
    });
    runTurnMock.mockResolvedValueOnce(question("old-mode", "旧规则", "旧规则已确认"))
      .mockResolvedValueOnce(question("replacement", "改为排序", "按名称排序", "replace"))
      .mockResolvedValueOnce(question("sort-order", "升序", "升序排列"))
      .mockResolvedValueOnce({ kind: "answer", message: "新任务规则已收到", provider: "model" });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "模型应就绪");
    const optionButton = (label: string) => Array.from(container.querySelectorAll<HTMLButtonElement>(".clarification-options button"))
      .find((button) => button.querySelector("strong")?.textContent === label) ?? null;
    const composer = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(composer, "旧的填充任务");
    await submit(composer);
    await until(() => !!optionButton("旧规则"), "第一道追问应出现");
    await click(optionButton("旧规则"));
    await until(() => !!optionButton("改为排序"), "应允许显式替换需求");
    await click(optionButton("改为排序"));
    await until(() => !!optionButton("升序"), "新需求应只补自己的缺项");
    expect(runTurnMock.mock.calls[2][0].taskContext).toEqual({ schemaVersion: 1, originalPrompt: "按名称排序", confirmations: [] });
    await click(optionButton("升序"));
    await until(() => textOf(container).includes("新任务规则已收到"), "新需求应完成");
    expect(runTurnMock.mock.calls[3][0].taskContext).toEqual({ schemaVersion: 1, originalPrompt: "按名称排序", confirmations: [
      { questionId: "sort-order", label: "升序", value: "升序排列" }
    ] });
    const history = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1")!);
    const card = history.conversations.flatMap((item: { messages: unknown[] }) => item.messages)
      .find((item: { taskResume?: { confirmations: { questionId?: string }[] } }) =>
        item.taskResume?.confirmations.some((decision) => decision.questionId === "sort-order"));
    expect(card.taskResume.originalPrompt).toBe("按名称排序");
    expect(card.taskResume.confirmations).toEqual([{ questionId: "sort-order", label: "升序", value: "升序排列" }]);
    expect(stageComparisonMock).not.toHaveBeenCalled();
  });

  it("renders separate fields and continues the original request with all answers, without executing", async () => {
    listModelsMock.mockResolvedValue({ models: [firstConnection], defaultModelId: firstConnection.id });
    getWpsApplicationMock.mockReturnValue(clarificationHost());
    runTurnMock.mockResolvedValueOnce({ kind: "clarification", clarification: { response: {
      kind: "clarification", provider: "model", clarification: {
        id: "fields-1", summary: "补充口径", question: "请填写两项信息", reason: "尚缺口径",
        scopeLabel: "Sheet1", options: [], fields: [
          { id: "period", label: "时间范围", placeholder: "说明时间范围" }, { id: "unit", label: "统计单位" }
        ]
      }
    } } }).mockResolvedValueOnce({ kind: "answer", message: "收到补充，准备预览", provider: "model" });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "模型应就绪");
    const prompt = container.querySelector("textarea");
    await type(prompt, "按已选表统计并预览");
    await submit(prompt);
    await until(() => !!container.querySelector(".clarification-fields"), "规划追问应显示表单");
    const button = findButton(container, "提交补充并生成预览");
    if (!button) throw new Error("缺少表单提交按钮");
    expect(button.disabled).toBe(true);
    await type(container.querySelector('[aria-label="时间范围"]'), "本季度");
    expect(button.disabled).toBe(true);
    await type(container.querySelector('[aria-label="统计单位"]'), "件");
    expect(button.disabled).toBe(false);
    await click(button);
    await until(() => runTurnMock.mock.calls.length === 2, "提交应继续原任务");
    expect(runTurnMock.mock.calls[1][0]).toMatchObject({
      prompt: "按已选表统计并预览\n补充确认：时间范围：本季度\n统计单位：件", intentConfirmed: true,
      taskContext: { originalPrompt: "按已选表统计并预览", confirmations: [
        { questionId: "fields-1", label: "补充说明", value: "时间范围：本季度\n统计单位：件" }
      ] }
    });
    expect(stageComparisonMock).not.toHaveBeenCalled();
    expect(validatePlanMock).not.toHaveBeenCalled();
    expect(container.querySelector(".clarification-card.resolved")).not.toBeNull();
  });

  it("keeps a cancelled clarification in history without leaving live choices", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    getWpsApplicationMock.mockReturnValue(clarificationHost());
    runTurnMock.mockResolvedValue({
      kind: "clarification",
      clarification: {
        response: {
          kind: "clarification",
          provider: "model",
          clarification: {
            id: "clarify-cancel",
            summary: "需要明确目标",
            question: "要查询还是修改？",
            reason: "两种操作不同",
            scopeLabel: "当前表",
            options: [{ id: "read", label: "查询", description: "只读", resolution: "查询数据" }]
          }
        }
      }
    });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
    const composer = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(composer, "帮我处理表格");
    await submit(composer);
    await until(() => Boolean(container.querySelector(".clarification-card.pending")), "应显示澄清卡");
    await click(findButton(container, "取消本次澄清"));
    expect(container.querySelector(".clarification-card.cancelled")).not.toBeNull();
    expect(container.querySelector(".clarification-card button")).toBeNull();
    expect(runTurnMock).toHaveBeenCalledTimes(1);
    const stored = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    expect(stored.conversations[0].messages.some((message: { clarification?: { status: string } }) =>
      message.clarification?.status === "cancelled"
    )).toBe(true);
  });

  it("pauses a turn without a checkpoint and retains the original request honestly", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockImplementation(({ signal }) => new Promise((_resolve, reject) => {
      signal?.addEventListener("abort", () => reject(new DOMException("stopped", "AbortError")));
    }));
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
    const composer = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(composer, "统计当前表");
    await submit(composer);
    await until(() => runTurnMock.mock.calls.length === 1 && Boolean(container.querySelector('[aria-label="暂停任务"]')), "运行时应显示暂停入口");
    await click(container.querySelector('[aria-label="暂停任务"]'));
    await until(() => textOf(container).includes("尚未形成可续跑检查点"), "不能伪造续跑现场");
    expect(composer?.value).toBe("");
    expect(container.querySelector('[aria-label="暂停任务"]')).toBeNull();
  });

  it("resolves follow-current authorization again at send time", async () => {
    getWpsApplicationMock.mockReturnValue(clarificationHost({ sheetNames: ["Sheet1", "Sheet2"] }));
    const resolveSelectedSheetIndexes = vi.fn(() => [2] as const);
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    runTurnMock.mockResolvedValue({
      kind: "answer",
      message: "已完成。",
      resultContext: null,
      provider: "model"
    });

    const { container } = await mount(
      <WpsAgentPanel
        selectedSheetIndexes={[1]}
        resolveSelectedSheetIndexes={resolveSelectedSheetIndexes}
        scopeDock={<div>当前表 Sheet1</div>}
      />
    );
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
    const textarea = container.querySelector("textarea");
    await type(textarea, "统计当前表");
    await submit(textarea);
    await until(() => runTurnMock.mock.calls.length === 1, "应发起查询");

    expect(resolveSelectedSheetIndexes).toHaveBeenCalledTimes(1);
    expect(runTurnMock.mock.calls[0][0].selectedSheetIndexes).toEqual([2]);
  });

  it("keeps earlier conversations when a new chat is created", async () => {
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    runTurnMock.mockResolvedValue({
      kind: "answer",
      message: "历史回答",
      resultContext: null,
      provider: "model"
    });

    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
    const textarea = container.querySelector("textarea");
    await type(textarea, "历史问题");
    await submit(textarea);
    await until(() => textOf(container).includes("历史回答"), "应完成首个对话");
    await click(container.querySelector<HTMLButtonElement>('[aria-label="新对话"]'));

    expect(textOf(container)).toContain("已新建对话");
    expect(textOf(container)).not.toContain("新对话已开始");
    await until(() => document.activeElement === textarea, "新对话后应聚焦输入框");

    // 已经是空白对话时再次点击只给反馈，不制造重复空历史。
    await click(container.querySelector<HTMLButtonElement>('[aria-label="新对话"]'));
    expect(textOf(container)).toContain("当前已是新对话");
    await click(container.querySelector<HTMLButtonElement>('[aria-label="更多"]'));
    await click(container.querySelector<HTMLButtonElement>('[aria-label="历史对话"]'));

    expect(textOf(container)).toContain("历史问题");
    expect(textOf(container)).toContain("新对话");
    expect(container.querySelectorAll(".history-item")).toHaveLength(2);

    await act(async () => {
      await new Promise<void>((resolve) => setTimeout(resolve, 1550));
    });
    expect(container.querySelector('[role="status"]')).toBeNull();
  });

  it.each(["new", "history"])("clears stale progress hints on %s chat navigation, showing the idle anchor for a finished turn", async (navigation) => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    vi.mocked(understandMessage).mockResolvedValue({ kind: "conversation", message: "历史回答", provider: "model" });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型") ?? false, "应加载模型");
    const textarea = container.querySelector("textarea");
    const sendMore = async (question: string, answer: string) => {
      vi.mocked(understandMessage).mockResolvedValue({ kind: "conversation", message: answer, provider: "model" });
      await type(textarea, question);
      await submit(textarea);
      await until(() => textOf(container).includes(answer), "应完成当前回答");
    };
    await sendMore("历史问题", "历史回答");
    if (navigation === "history") {
      await click(container.querySelector('[aria-label="新对话"]'));
      await sendMore("另一段对话", "另一段回答");
    }
    const stream = container.querySelector<HTMLElement>(".message-stream")!;
    Object.defineProperties(stream, {
      scrollHeight: { configurable: true, value: 1000 },
      clientHeight: { configurable: true, value: 300 }
    });
    stream.scrollTop = 200;
    await act(async () => { stream.dispatchEvent(new Event("scroll", { bubbles: true })); });
    await sendMore("补充问题", "补充回答");
    const hint = container.querySelector<HTMLButtonElement>('.new-progress-button')!;
    expect(hint).not.toBeNull();
    // 这一轮**已经跑完**了，所以是"停下"那一态：锚，而不是三点。
    expect(hint.dataset.mode).toBe("idle");
    expect(hint.querySelector(".progress-anchor-icon")).not.toBeNull();
    expect(hint.querySelector(".progress-dots")).toBeNull();
    expect(hint.getAttribute("aria-label")).toBe("回到底部查看最新内容");
    expect(stream.scrollTop).toBe(200);

    if (navigation === "new") await click(container.querySelector('[aria-label="新对话"]'));
    else {
      await click(container.querySelector('[aria-label="更多"]'));
      await click(container.querySelector('[aria-label="历史对话"]'));
      const history = Array.from(container.querySelectorAll<HTMLButtonElement>(".history-open-button"))
        .find((button) => button.textContent?.includes("历史问题"))!;
      await click(history);
    }
    expect(container.querySelector(".new-progress-button")).toBeNull();
    expect(stream.scrollTop).toBe(1000);
    expect(runTurnMock).not.toHaveBeenCalled();
  });

  it("keeps the three-dot indicator while the turn is in flight and swaps to the anchor once it lands", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    // 把 turn 挂住：`running` 只有在**真的还跑着**的时候才为真，不挂住就只能看到终态。
    let finishTurn: ((value: Awaited<ReturnType<typeof runWpsReadonlyTurn>>) => void) | undefined;
    runTurnMock.mockImplementation((input) => {
      input.onProgress?.({ stage: "intent", stepId: "intent", title: "理解需求", message: "正在判断任务类型" });
      return new Promise((resolve) => { finishTurn = resolve; });
    });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");

    // 先把用户"滚走"：浮标只在滚离底部时渲染（判据是位置，不要求有新内容）。
    const stream = container.querySelector<HTMLElement>(".message-stream")!;
    Object.defineProperties(stream, {
      scrollHeight: { configurable: true, value: 1000 },
      clientHeight: { configurable: true, value: 300 }
    });
    stream.scrollTop = 200;
    await act(async () => { stream.dispatchEvent(new Event("scroll", { bubbles: true })); });

    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "统计当前表");
    await submit(textarea);

    // 运行时：三点动效。`until` 等到的是 running 那一态本身 —— 如果接线写成
    // `running={false}`，这里永远等不到，用例会在这里失败而不是悄悄放行。
    await until(
      () => container.querySelector<HTMLButtonElement>(".new-progress-button")?.dataset.mode === "running",
      "turn 还在跑，浮标应当显示运行中的三点"
    );
    const hint = container.querySelector<HTMLButtonElement>(".new-progress-button")!;
    expect(hint.querySelectorAll(".progress-dots > span")).toHaveLength(3);
    expect(hint.querySelector(".progress-anchor-icon")).toBeNull();
    expect(hint.getAttribute("aria-label")).toBe("查看最新执行进度");

    await act(async () => {
      finishTurn?.({ kind: "answer", message: "统计完成", resultContext: null, provider: "model" });
    });

    // 跑完：同一个按钮换成锚（不是消失再出现）。
    await until(
      () => container.querySelector<HTMLButtonElement>(".new-progress-button")?.dataset.mode === "idle",
      "turn 结束后浮标应当换成锚"
    );
    expect(container.querySelector<HTMLButtonElement>(".new-progress-button")).toBe(hint);
    expect(hint.querySelector(".progress-anchor-icon")).not.toBeNull();
    expect(hint.querySelector(".progress-dots")).toBeNull();
  });

  it("shows a normalized write plan without touching cells until explicit confirmation", async () => {
    let value = "旧值";
    let formula = "旧值";
    let writes = 0;
    const range = { Rows: { Count: 1 }, Columns: { Count: 1 }, MergeCells: false } as Record<string, unknown>;
    Object.defineProperties(range, {
      Value2: {
        get: () => value,
        set: (next: string) => { value = next; formula = next; writes += 1; },
        enumerable: true
      },
      Formula: {
        get: () => formula,
        set: (next: string) => { formula = next; value = next; },
        enumerable: true
      }
    });
    const sheet = { Name: "Sheet1", ProtectContents: false, Range: () => range };
    const storage = new Map<string, unknown>();
    getWpsApplicationMock.mockReturnValue({
      Version: "12.0",
      ActiveWorkbook: {
        Name: "测试.xlsx",
        FullName: "C:\\Data\\测试.xlsx",
        Path: "C:\\Data",
        ReadOnly: false,
        Worksheets: { Count: 1, Item: () => sheet },
        ActiveSheet: sheet
      },
      ActiveSheet: sheet,
      PluginStorage: {
        getItem: (key: string) => storage.get(key),
        setItem: (key: string, item: unknown) => { storage.set(key, item); }
      },
      CreateTaskPane: () => ({ ID: 1, Visible: true }),
      GetTaskPane: () => undefined
    } as never);
    const writePlan = {
      id: "write-1",
      title: "更新 A1",
      summary: "把 A1 更新为新值",
      assumptions: [],
      warnings: [],
      actions: [{ type: "writeValues" as const, sheet: "Sheet1", range: "A1", values: [["新值"]] }],
      acceptanceCriteria: []
    };
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({ kind: "plan", plan: writePlan });
    validatePlanMock.mockImplementation(async (submitted) => submitted);

    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
    const textarea = container.querySelector("textarea");
    await type(textarea, "把 A1 改成新值");
    await submit(textarea);
    await until(() => textOf(container).includes("写入前确认"), "应展示写入计划");

    expect(writes).toBe(0);
    expect(container.querySelector(".wps-plan-card")?.closest(".message-row.assistant")).not.toBeNull();
    const savedHistory = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    const savedPreview = savedHistory.conversations[0].messages.find((message: { planPreview?: unknown }) => message.planPreview);
    expect(savedPreview.planPreview).toEqual({ title: "更新 A1", status: "pending" });
    expect(savedPreview.text).toBe("");
    /*
     * 开发环境提示**只进诊断入口**，而且只在副本已生成的审阅卡里。
     * 写入前预览阶段一个排查材料都不摆 —— 正向锚点在下面两行（预览确实渲染了）。
     */
    expect(textOf(container)).not.toContain("当前 WPS JSAPI 版本仅在开发调试模式开放写入");
    expect(textOf(container)).toContain("Sheet1!A1");
    expect(textOf(container)).toContain("新值");
    await click(container.querySelector<HTMLButtonElement>('[aria-label="新对话"]'));
    expect(textOf(container)).toContain("放弃待确认的写入计划？");
    expect(textOf(container)).toContain("当前对话还有一份尚未执行的写入计划");
    await click(findButton(container, "继续当前对话"));
    expect(textOf(container)).toContain("写入前确认");
    await click(findButton(container, "编辑计划"));
    expect(findButton(container, "生成右侧副本并排审阅")).toBeNull();
    const editBox = container.querySelector<HTMLTextAreaElement>('[aria-label="编辑写入计划 JSON"]');
    await type(editBox, JSON.stringify({
      ...writePlan,
      actions: [{ type: "writeValues", sheet: "Sheet1", range: "A1", values: [["修订值"]] }]
    }));
    await click(findButton(container, "应用修改并重新预览"));
    await until(() => textOf(container).includes("修订值") && findButton(container, "生成右侧副本并排审阅") !== null, "编辑后应重新预览最终值");
    expect(writes).toBe(0);
    await stageAndSave(container);
    await until(
      () => Boolean(container.querySelector('button[aria-label="备份与恢复"] b')),
      "确认后应完成写入并提供恢复入口"
    );
    expect(writes).toBe(1);
    expect(value).toBe("修订值");
    expect(textOf(container)).toContain("计划已执行并回读确认");
    const resultMessage = container.querySelector(".wps-verification-card.verified")?.closest(".message-row.assistant");
    expect(resultMessage?.querySelector(".message-author")?.textContent?.trim()).toBe("Excel Bro");
    expect(container.querySelector(".wps-verification-card.verified")?.textContent).toContain("1 个动作 · 1 个单元格已回读核对");
    expect(container.querySelector(".wps-verification-card.verified")?.textContent).toContain("Sheet1!A1");
    expect(textOf(container)).toContain("保存写前恢复点");
    expect(textOf(container)).toContain("回读核对第 1 个动作");
    const completedTimeline = Array.from(container.querySelectorAll(".task-progress-timeline.activity-log-card"))
      .find((item) => item.textContent?.includes("执行第 1 个写入动作"));
    expect(completedTimeline).not.toBeNull();
    expect(completedTimeline?.textContent).toContain("Sheet1!A1");
    expect(completedTimeline?.textContent).toContain("writeValues");
  });

  it("本端算出的计划不送服务端校验", async () => {
    // 验收条件里带着完整值矩阵（RangeEqualsCriterion.expected），送校验等于把
    // 刚才刻意留在本地的源值上传 —— 单射映射可逆。这条用例把"不上传"钉住。
    let value = "旧值";
    const range = {
      Rows: { Count: 1 },
      Columns: { Count: 1 },
      MergeCells: false,
      Formula: "旧值"
    } as Record<string, unknown>;
    Object.defineProperty(range, "Value2", {
      get: () => value,
      set: (next: string) => { value = next; },
      enumerable: true
    });
    const sheet = { Name: "Sheet1", ProtectContents: false, Range: () => range };
    const storage = new Map<string, unknown>();
    getWpsApplicationMock.mockReturnValue({
      Version: "12.0",
      ActiveWorkbook: {
        Name: "测试.xlsx",
        FullName: "C:\\Data\\测试.xlsx",
        Path: "C:\\Data",
        ReadOnly: false,
        Worksheets: { Count: 1, Item: () => sheet },
        ActiveSheet: sheet
      },
      ActiveSheet: sheet,
      PluginStorage: {
        getItem: (key: string) => storage.get(key),
        setItem: (key: string, item: unknown) => { storage.set(key, item); }
      },
      CreateTaskPane: () => ({ ID: 1, Visible: true }),
      GetTaskPane: () => undefined
    } as never);
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    runTurnMock.mockResolvedValue({
      kind: "plan",
      computedLocally: true,
      sourceSnapshots: [{ sheetName: "Sheet1", range: "A1", values: [["旧值"]] }],
      plan: {
        id: "wps-map-column-abc",
        title: "按映射表填写 Sheet1!A1",
        summary: "按 1 条映射规则写入。",
        assumptions: [],
        warnings: [],
        actions: [
          { type: "writeValues" as const, sheet: "Sheet1", range: "A1", values: [["行政"]] }
        ]
      }
    });

    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
    const textarea = container.querySelector("textarea");
    await type(textarea, "把类别列翻译成归属");
    await submit(textarea);
    await until(() => textOf(container).includes("写入前确认"), "应展示写入计划");

    expect(validatePlanMock).not.toHaveBeenCalled();
    expect(textOf(container)).toContain("Sheet1!A1");
    await click(findButton(container, "编辑计划"));
    await click(findButton(container, "应用修改并重新预览"));
    await until(() => findButton(container, "生成右侧副本并排审阅") !== null, "本地计划编辑后应重新预览");
    expect(validatePlanMock).not.toHaveBeenCalled();
  });

  it("keeps a secret leaked by an upstream error out of the persisted chat history", async () => {
    // 这条断言必须**先注入**再检查：如果不往流程里塞一个密钥，
    // `not.toContain("sk-")` 会因为压根没有密钥而空过 —— 那是假绿灯。
    // 通道是真实的：`turnErrorMessage` 会把非连接类错误的 message 原样写进
    // 对话历史，而服务端错误 message 源自 `str(error)`，可能带上请求里的凭据。
    const secret = "sk-live-leaked-from-upstream-9f3a";
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockRejectedValue(
      new Error(`模型连接校验失败：上游拒绝 Authorization: Bearer ${secret}`)
    );
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "统计订单金额");
    await submit(textarea);
    await until(
      () => container.querySelector(".message-row.system") !== null,
      "非连接类错误应产生一条系统消息"
    );

    const persisted = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
    // 断言的是**密钥本体**不落盘。`Bearer` 这个词本身不是凭据 ——
    // 修复后的文本是「Authorization: Bearer 【已隐藏】」，方案名保留反而更可诊断，
    // 所以不能把「出现 Bearer 字面」当成泄漏。
    expect(persisted).not.toContain(secret);
    expect(persisted).not.toContain("sk-live");
    expect(persisted).not.toContain("sk-");
    expect(textOf(container)).not.toContain(secret);
    // 正向锚点：这条错误确实写进了历史（否则上面的负向断言可能空过），
    // 且密钥位置被换成了明确占位符 —— 是「被隐藏」而不是「整段消失」。
    expect(persisted).toContain("模型连接校验失败");
    expect(persisted).toContain("Bearer 【已隐藏】");
    // 用户仍应看到失败这件事与可诊断的部分，只是不带凭据。
    expect(textOf(container)).toContain("模型连接校验失败");
    expect(textOf(container)).toContain("已隐藏");
  });

  it("sanitizes an old record before it can be rendered, and blocks resuming it", async () => {
    // 读取路径的职责（本轮修正后）：原始存储 → **立刻脱敏** → 进入可渲染状态，
    // 同时登记一个持久的旧记录标记用于发送闸。旧密钥**不得**出现在界面、
    // 复制内容或内存里的可发送历史中。
    const secret = "sk-live-legacy-stored-2d77";
    localStorage.setItem(
      "excel-bro.wps.chat.v1",
      JSON.stringify({
        activeConversationId: "legacy",
        conversations: [
          {
            id: "legacy",
            title: `用 ${secret} 统计订单`,
            createdAt: "2026-09-27T00:00:00.000Z",
            updatedAt: "2026-09-27T00:00:00.000Z",
            messages: [
              {
                id: "m1",
                role: "user",
                text: `用 ${secret} 统计订单`,
                createdAt: "2026-09-27T00:00:00.000Z"
              }
            ]
          }
        ]
      })
    );
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });

    const { container } = await mount(panel());
    await until(
      () => container.querySelector(".stale-credentials-warning") !== null,
      "旧记录应给出「需新开会话」提示"
    );

    // 正向锚点：旧记录确实被读回来了、非敏感内容完整保留。
    expect(textOf(container)).toContain("统计订单");
    // 关键负向断言：界面里不得出现真密钥。
    expect(textOf(container)).not.toContain(secret);
    expect(textOf(container)).not.toContain("sk-live");

    // 落盘内容同样不含密钥，且带上了持久标记。
    const persisted = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
    expect(persisted).not.toContain(secret);
    expect(persisted).toContain("legacyCredentialConversationIds");
  });

  it("renders a restored placeholder and keeps an ordinary question untouched", async () => {
    // 已经落盘脱敏过的记录：读取时应把占位符还原成可读措辞；
    // 若它带着旧记录标记，则会继续提示，但文本里不含任何真凭据。
    localStorage.setItem(
      "excel-bro.wps.chat.v1",
      JSON.stringify({
        activeConversationId: "saved",
        conversations: [
          {
            id: "saved",
            title: "用 【已隐藏】 统计订单",
            createdAt: "2026-09-27T00:00:00.000Z",
            updatedAt: "2026-09-27T00:00:00.000Z",
            messages: [
              {
                id: "m1",
                role: "user",
                text: "用 【已隐藏】 统计订单",
                createdAt: "2026-09-27T00:00:00.000Z"
              },
              {
                id: "m2",
                role: "user",
                text: "Bearer 和 api key 有什么区别？",
                createdAt: "2026-09-27T00:00:01.000Z"
              }
            ]
          }
        ]
      })
    );
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });

    const { container } = await mount(panel());
    await until(
      () => container.querySelector(".message-row.user") !== null,
      "旧历史的用户消息应被读回并渲染"
    );

    const stream = textOf(container.querySelector(".message-stream") as HTMLElement);
    // 落盘标记被还原成更明确的措辞。
    expect(stream).toContain("此处原为凭据");
    expect(stream).not.toContain("【已隐藏】");
    // 用户正常提问一字不改 —— 读取时不做新的脱敏。
    expect(stream).toContain("Bearer 和 api key 有什么区别？");
    // 没有被标为旧记录的干净历史不触发提示（否则每次重开都会误锁续聊）。
    expect(container.querySelector(".stale-credentials-warning")).toBeNull();
  });

  it("blocks resuming a chat whose stored history still carries a real credential", async () => {
    // 核心行为：拦在**发送之前**，且拦的依据是持久标记而不是"还能不能匹配密钥"。
    const secret = "sk-live-resume-blocked-6c02";
    localStorage.setItem(
      "excel-bro.wps.chat.v1",
      JSON.stringify({
        activeConversationId: "legacy",
        conversations: [
          {
            id: "legacy",
            title: `用 ${secret} 统计订单`,
            createdAt: "2026-09-27T00:00:00.000Z",
            updatedAt: "2026-09-27T00:00:00.000Z",
            messages: [
              {
                id: "m1",
                role: "user",
                text: `用 ${secret} 统计订单`,
                createdAt: "2026-09-27T00:00:00.000Z"
              }
            ]
          }
        ]
      })
    );
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });

    const { container } = await mount(panel());
    await until(
      () => container.querySelector(".composer textarea") !== null,
      "面板应就绪"
    );

    await type(container.querySelector<HTMLTextAreaElement>(".composer textarea"), "继续统计");
    await submit(container.querySelector<HTMLTextAreaElement>(".composer textarea"));

    // 正向锚点：确实给出了反馈，而不是静默什么都不发生。
    await until(
      () => textOf(container).includes("已被拦下"),
      "应提示这段历史的凭据被拦下"
    );
    // 关键负向断言：旧记录没有被送给模型。
    expect(runTurnMock).not.toHaveBeenCalled();
    // 界面里也看不到真密钥。
    expect(textOf(container)).not.toContain(secret);
  });

  it("warns without leaking when the host storage refuses the history write", async () => {
    // 宿主禁止 localStorage（或配额用尽）时，落盘 effect 过去是**静默忽略**的：
    // 用户会以为"历史保存了"，关掉任务窗格才发现全没了。这条用例钉三件事：
    // 1. 界面给出**明确**提示；
    // 2. 提示**不含**错误对象原文，也不含任何凭据；
    // 3. 内存历史仍保持脱敏与陈旧标记 —— 落盘失败**不会**让发送闸松开。
    const secret = "sk-live-storage-refused-77c1";
    localStorage.setItem(
      "excel-bro.wps.chat.v1",
      JSON.stringify({
        activeConversationId: "legacy",
        conversations: [
          {
            id: "legacy",
            title: `用 ${secret} 统计订单`,
            createdAt: "2026-09-27T00:00:00.000Z",
            updatedAt: "2026-09-27T00:00:00.000Z",
            messages: [
              {
                id: "m1",
                role: "user",
                text: `用 ${secret} 统计订单`,
                createdAt: "2026-09-27T00:00:00.000Z"
              }
            ]
          }
        ]
      })
    );
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });

    // 只让**对话历史**那一笔写入失败；读取（getItem）与其他键不受影响。
    // 抛出的错误里刻意塞进凭据，用来证明提示不会把异常原文透出去。
    const originalSetItem = Storage.prototype.setItem;
    const spy = vi
      .spyOn(Storage.prototype, "setItem")
      .mockImplementation(function (this: Storage, key: string, value: string) {
        if (key === "excel-bro.wps.chat.v1") {
          throw new Error(`QuotaExceededError while writing ${secret}`);
        }
        originalSetItem.call(this, key, value);
      });

    try {
      const { container } = await mount(panel());
      await until(
        () => container.querySelector(".storage-write-warning") !== null,
        "落盘失败应给出明确提示"
      );

      const notice = container.querySelector(".storage-write-warning") as HTMLElement;
      // 正向锚点：提示确实说明了"不会被保留"这件事。
      expect(notice.textContent ?? "").toContain("不会被保留");
      // 关键负向断言 1：提示里不得出现错误对象原文，更不得出现凭据。
      expect(notice.textContent ?? "").not.toContain(secret);
      expect(notice.textContent ?? "").not.toContain("sk-live");
      expect(notice.textContent ?? "").not.toContain("QuotaExceededError");
      // 关键负向断言 2：整个界面都不能因为"写入失败"而把原文露出来。
      expect(textOf(container)).not.toContain(secret);
      expect(textOf(container)).not.toContain("sk-live");

      // 关键负向断言 3：落盘失败**不解除发送闸** —— 旧记录仍被拦在发送之前。
      await type(container.querySelector<HTMLTextAreaElement>(".composer textarea"), "继续统计");
      await submit(container.querySelector<HTMLTextAreaElement>(".composer textarea"));
      await until(
        () => textOf(container).includes("已被拦下"),
        "即使落盘失败，旧记录仍应被拦"
      );
      expect(runTurnMock).not.toHaveBeenCalled();
      // 陈旧提示也还在。
      expect(container.querySelector(".stale-credentials-warning")).not.toBeNull();
    } finally {
      spy.mockRestore();
    }
  });

  it("keeps the legacy flag across a reopen so the send gate stays closed", async () => {
    // 重开之后：文本已被脱敏，靠"匹配密钥"已无法判断，必须靠持久标记。
    const secret = "sk-live-reopen-wps-90af";
    localStorage.setItem(
      "excel-bro.wps.chat.v1",
      JSON.stringify({
        activeConversationId: "legacy",
        conversations: [
          {
            id: "legacy",
            title: `用 ${secret} 统计订单`,
            createdAt: "2026-09-27T00:00:00.000Z",
            updatedAt: "2026-09-27T00:00:00.000Z",
            messages: [
              {
                id: "m1",
                role: "user",
                text: `用 ${secret} 统计订单`,
                createdAt: "2026-09-27T00:00:00.000Z"
              }
            ]
          }
        ]
      })
    );
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });

    // 第一次挂载：脱敏 + 落盘（写入标记）。
    const first = await mount(panel());
    await until(
      () => first.container.querySelector(".stale-credentials-warning") !== null,
      "首次挂载应提示"
    );
    await settle();
    const persisted = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
    expect(persisted).not.toContain(secret);
    first.unmount();

    // 第二次挂载（等价于重开任务窗格）：文本已干净，但标记仍在。
    const second = await mount(panel());
    await until(
      () => second.container.querySelector(".composer textarea") !== null,
      "面板应就绪"
    );
    await until(
      () => second.container.querySelector(".stale-credentials-warning") !== null,
      "重开后仍应保持提示"
    );

    await type(
      second.container.querySelector<HTMLTextAreaElement>(".composer textarea"),
      "继续统计"
    );
    await submit(second.container.querySelector<HTMLTextAreaElement>(".composer textarea"));
    await until(
      () => textOf(second.container).includes("已被拦下"),
      "重开后发送仍应被拦"
    );
    expect(runTurnMock).not.toHaveBeenCalled();
  });

  it("keeps credentials and raw cell values out of the persisted chat history", async () => {
    // 持久历史进的是宿主 localStorage，用户可读可复制。
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({
      kind: "answer", message: "统计完成", provider: "model", resultContext: null
    });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "统计订单金额");
    await submit(textarea);
    await until(() => textOf(container).includes("统计完成"), "应完成一轮");

    const persisted = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
    // 正向锚点：历史确实写进去了（否则下面的负向断言又会空过）。
    expect(persisted).toContain("统计订单金额");
    expect(persisted).not.toContain("sk-");
    expect(persisted).not.toContain("apiKey");
    expect(persisted).not.toContain("Authorization");
    expect(persisted).not.toContain("bearer");
  });

  it("keeps a secret pasted into the prompt out of the persisted conversation title", async () => {
    // 标题 = 首条用户消息前 34 字。用户把密钥连同问题一起粘进来时，
    // 只脱敏 message.text 会让标题成为绕过通道。
    const secret = "sk-live-in-title-c41d";
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({
      kind: "answer", message: "已统计。", provider: "model", resultContext: null
    });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, `用 ${secret} 统计订单金额`);
    await submit(textarea);
    await until(() => textOf(container).includes("已统计。"), "应完成一轮");

    const persisted = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
    const parsed = JSON.parse(persisted) as {
      conversations: Array<{ title: string }>;
    };
    // 正向锚点：标题确实生成了（否则负向断言可能空过），且仍是可读句子。
    expect(parsed.conversations[0].title).toContain("统计订单金额");
    expect(parsed.conversations[0].title).not.toContain(secret);
    expect(parsed.conversations[0].title).not.toContain("sk-");
    expect(parsed.conversations[0].title).toContain("已隐藏");
    // 正文同样不能漏。持久化原文（含转义）整体也不该出现密钥。
    expect(persisted).not.toContain(secret);
  });

  it("keeps a JSON-shaped secret out of every persisted text field", async () => {
    // 用户可能整段粘贴配置片段。JSON / 类 JSON 形态走的是另一条正则分支，
    // 必须单独注入验证，不能因为「裸 apiKey= 已覆盖」就假定这里也覆盖。
    const secret = "testsecret123456";
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({
      kind: "answer", message: `收到配置 {"apiKey":"${secret}"}，已统计。`,
      provider: "model", resultContext: null
    });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, `按 {'api_key': '${secret}'} 统计订单金额`);
    await submit(textarea);
    await until(() => textOf(container).includes("已统计。"), "应完成一轮");

    const persisted = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
    const parsed = JSON.parse(persisted) as {
      conversations: Array<{ title: string; messages: Array<{ text: string }> }>;
    };
    const title = parsed.conversations[0].title;
    const answerText = parsed.conversations[0].messages
      .map((message) => message.text)
      .join("\n");

    expect(persisted).not.toContain(secret);
    // 键名与结构保留，说明是「脱敏」而不是「整条消息被丢弃」。
    expect(title).toContain("api_key");
    expect(answerText).toContain("apiKey");
    expect(title).toContain("已隐藏");
    expect(answerText).toContain("已隐藏");
  });

  it("does not persist the locally computed target matrix with the plan preview", async () => {
    // 本地映射计划的目标值矩阵与源值一一对应（单射可逆），既不上传也不落盘。
    // 源区域在预览时会被对着活宿主复核（prepareWpsPlan），所以替身单元格必须
    // 与快照值一致 —— 否则拒绝的原因是"源值变了"，测不到持久化这件事。
    let value = "办公用品";
    const range = {
      Rows: { Count: 1 }, Columns: { Count: 1 }, MergeCells: false, Formula: "办公用品"
    } as Record<string, unknown>;
    Object.defineProperty(range, "Value2", {
      get: () => value,
      set: (next: string) => { value = next; },
      enumerable: true
    });
    const sheet = { Name: "Sheet1", ProtectContents: false, Range: () => range };
    const storage = new Map<string, unknown>();
    getWpsApplicationMock.mockReturnValue({
      Version: "12.0",
      ActiveWorkbook: {
        Name: "测试.xlsx", FullName: "C:\\Data\\测试.xlsx", Path: "C:\\Data",
        ReadOnly: false, Worksheets: { Count: 1, Item: () => sheet }, ActiveSheet: sheet
      },
      ActiveSheet: sheet,
      PluginStorage: {
        getItem: (key: string) => storage.get(key),
        setItem: (key: string, item: unknown) => { storage.set(key, item); }
      },
      CreateTaskPane: () => ({ ID: 1, Visible: true }),
      GetTaskPane: () => undefined
    } as never);
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({
      kind: "plan",
      computedLocally: true,
      sourceSnapshots: [{ sheetName: "Sheet1", range: "A1", values: [["办公用品"]] }],
      plan: {
        id: "wps-map-column-xyz",
        title: "按映射表填写 Sheet1!A1",
        summary: "按 1 条映射规则写入。",
        assumptions: [], warnings: [],
        actions: [{ type: "writeValues" as const, sheet: "Sheet1", range: "A1", values: [["行政"]] }]
      }
    });

    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "把类别列翻译成归属");
    await submit(textarea);
    await until(() => textOf(container).includes("写入前确认"), "应展示写入计划");

    // 目标矩阵「行政」与源值「办公用品」都不能出现在持久历史里。
    const persisted = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
    expect(persisted).not.toContain("办公用品");
    expect(persisted).not.toContain("行政");
    expect(persisted).toContain("按映射表填写 Sheet1!A1");
    expect(validatePlanMock).not.toHaveBeenCalled();
  });

  it("keeps an initial catalog failure out of chat and offers contextual retry", async () => {
    listModelsMock.mockRejectedValueOnce(new Error("Failed to fetch"));
    const timeoutSpy = vi.spyOn(window, "setTimeout");

    const { container } = await mount(panel());
    await until(
      () => textOf(container).includes("点击添加模型开始"),
      "首次连接不可用时应保留中性的入口状态"
    );

    // 尚未配置与暂时无法取得目录是两种不同状态，首次打开不能抛出失败卡。
    expect(textOf(container)).not.toContain("尚未配置模型");
    expect(textOf(container)).not.toContain("Failed to fetch");
    expect(textOf(container)).toContain("点击上方模型图标开始连接");
    expect(container.querySelector(".message-row.system")).toBeNull();
    expect(timeoutSpy.mock.calls.some(([, delay]) => delay === 5000)).toBe(false);
    timeoutSpy.mockRestore();

    await act(async () => {
      window.dispatchEvent(new Event("focus"));
    });
    expect(listModelsMock).toHaveBeenCalledTimes(1);

    const textarea = container.querySelector("textarea");
    await type(textarea, "你好");
    await submit(textarea);
    await submit(textarea);
    expect(textarea?.value).toBe("你好");
    expect(container.querySelector(".message-row.system")).toBeNull();
    expect(runTurnMock).not.toHaveBeenCalled();
    expect(textOf(container)).toContain("暂时无法取得已保存的模型连接");

    listModelsMock.mockResolvedValueOnce({
      defaultModelId: secondConnection.id,
      models: [secondConnection]
    });
    await click(findButton(container, "重试读取模型配置"));

    await until(
      () => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第二个连接") ?? false,
      "重试成功后应显示真实连接"
    );
    expect(listModelsMock).toHaveBeenCalledTimes(2);
  });

  it("removes only the obsolete model-catalog notices from saved chat history", async () => {
    const obsolete = "暂时无法读取模型配置，请点击上方模型菜单里的「重试」。已输入的内容仍保留在输入框中。";
    localStorage.setItem("excel-bro.wps.chat.v1", JSON.stringify({
      activeConversationId: "saved-1",
      conversations: [{
        id: "saved-1", title: "你好", createdAt: "2026-09-26T00:00:00.000Z",
        updatedAt: "2026-09-26T00:00:00.000Z",
        messages: [
          { id: "user-1", role: "user", text: "你好", createdAt: "2026-09-26T00:00:00.000Z" },
          { id: "old-1", role: "system", text: obsolete, createdAt: "2026-09-26T00:00:01.000Z" }
        ]
      }]
    }));
    listModelsMock.mockResolvedValue({ defaultModelId: null, models: [] });

    const { container } = await mount(panel());
    await until(() => textOf(container).includes("尚未配置模型"), "应完成目录读取");
    expect(textOf(container)).toContain("你好");
    expect(textOf(container)).not.toContain(obsolete);
    expect(localStorage.getItem("excel-bro.wps.chat.v1")).not.toContain(obsolete);
  });

  it("opens the WPS add-model entry immediately while checking saved connections", async () => {
    listModelsMock.mockRejectedValueOnce(new Error("Failed to fetch"));
    listModelsMock.mockResolvedValue({ defaultModelId: null, models: [] });
    getModelSettingsMock.mockResolvedValue({
      baseUrl: null,
      defaultModel: null,
      apiKeyConfigured: false,
      apiKeyHint: null,
      formulaModelId: "",
      connections: []
    });

    const { container } = await mount(panel());
    await until(() => textOf(container).includes("点击添加模型开始"), "应完成首次模型目录读取");
    await click(container.querySelector(".model-picker-trigger"));
    await click(findButton(container, "＋ 添加模型连接"));

    expect(textOf(container)).toContain("添加模型连接");
    expect(textOf(container)).not.toContain("正在读取模型连接");
    expect(textOf(container)).not.toContain("后台暂时不可用");
    await until(() => getModelSettingsMock.mock.calls.length === 1, "新增表单应异步检查已有连接");
    expect(textOf(container)).toContain("已检查：这台电脑上暂无已保存的模型连接");
    expect(listModelsMock).toHaveBeenCalledTimes(2);

    await click(container.querySelector('.wps-model-settings-header [aria-label="关闭模型设置"]'));
    await click(container.querySelector(".model-picker-trigger"));
    await click(findButton(container, "管理模型连接"));
    await until(() => getModelSettingsMock.mock.calls.length === 2, "管理入口也应读取已有连接");
  });

  it("keeps the add form usable when saved-connection discovery fails", async () => {
    getModelSettingsMock.mockRejectedValue(new Error("Failed to fetch"));

    function Harness() {
      const management = useModelManagement({
        refreshServiceHealth: async () => null
      });
      const [open, setOpen] = useState(true);
      return (
        <WpsModelSettingsPanel
          open={open}
          models={[]}
          currentModelId=""
          management={management}
          onClose={() => setOpen(false)}
          onAddConnection={() => {
            management.openConnectionCreator();
            void management.checkExistingConnections();
          }}
        />
      );
    }

    const { container } = await mount(<Harness />);

    const addButton = findButton(container, "＋ 添加");
    expect(addButton?.hasAttribute("disabled")).toBe(false);
    await click(addButton);

    // 新建表单不等待后台；读取失败不是“没有已保存连接”。
    await until(
      () => textOf(container).includes("添加模型连接"),
      "后台不可读时也应能打开新增表单"
    );
    expect(
      findButton(container, "添加并使用")?.hasAttribute("disabled")
    ).toBe(true);
    await until(() => textOf(container).includes("暂时无法检查已有连接"), "应区分读取失败和空列表");
    expect(textOf(container)).not.toContain("已检查：这台电脑上暂无已保存的模型连接");
    expect(findButton(container, "重试检查")).not.toBeNull();
    expect(getModelSettingsMock).toHaveBeenCalledTimes(1);
  });

  it("renders the complete local model connection manager", () => {
    const noop = () => undefined;
    const management = {
      modelSettings: {
        baseUrl: null,
        defaultModel: null,
        apiKeyConfigured: false,
        apiKeyHint: null,
        formulaModelId: "",
        connections: []
      },
      apiKeyDraft: "",
      showApiKey: false,
      connectionDraft: {
        id: null,
        label: "",
        baseUrl: "",
        modelId: "",
        apiKey: "",
        clearApiKey: false,
        supportsVision: false
      },
      pendingDeleteConnectionId: null,
      settingsSaving: false,
      settingsTesting: false,
      settingsLoading: false,
      settingsFeedback: "",
      connectionTestFeedback: null,
      connectionVerified: false,
      saveApiKey: noop,
      editModelConnection: noop,
      verifyConnection: noop,
      saveConnection: noop,
      removeConnection: noop,
      saveFormulaModel: noop,
      setApiKeyDraft: noop,
      setShowApiKey: noop,
      setConnectionDraft: noop,
      setPendingDeleteConnectionId: noop,
      setSettingsFeedback: noop
    };

    const markup = renderToStaticMarkup(
      <WpsModelSettingsPanel
        open
        models={[]}
        currentModelId=""
        management={management as never}
        onClose={noop}
        onAddConnection={noop}
      />
    );

    expect(markup).toContain("模型设置");
    expect(markup).toContain("API Key 安全保存在这台电脑上");
    expect(markup).toContain("添加模型连接");
    expect(markup).toContain("测试连接");
    expect(markup).toContain("添加并使用");
    expect(markup).toContain("/function 公式模型");

    const existingConnectionMarkup = renderToStaticMarkup(
      <WpsModelSettingsPanel
        open
        models={[]}
        currentModelId=""
        management={{
          ...management,
          existingConnectionsStatus: "ready",
          connectionVerified: true,
          connectionDraft: {
            id: null,
            label: "再次添加",
            baseUrl: "https://api.example.com/",
            modelId: "GPT-4.1",
            apiKey: "",
            clearApiKey: false,
            supportsVision: false
          },
          modelSettings: {
            ...management.modelSettings,
            connections: [{
              id: "saved",
              catalogModelId: "connection:saved",
              label: "已有模型",
              baseUrl: "https://api.example.com",
              modelId: "gpt-4.1",
              supportsVision: false,
              apiKeyConfigured: true,
              apiKeyHint: "••••1234"
            }]
          },
          selectModel: noop
        } as never}
        onClose={noop}
        onAddConnection={noop}
      />
    );
    expect(existingConnectionMarkup).toContain("Excel 与 WPS 共用这些配置");
    expect(existingConnectionMarkup).not.toContain("无需重复添加");
    expect(existingConnectionMarkup.match(/<button class="model-primary-action"[^>]*>/)?.[0]).not.toContain("disabled");
    expect(existingConnectionMarkup.match(/<button type="button"[^>]*>测试连接<\/button>/)?.[0]).not.toContain("disabled");

    const unavailableMarkup = renderToStaticMarkup(
      <WpsModelSettingsPanel
        open
        models={[]}
        currentModelId=""
        management={{
          ...management,
          connectionDraft: null,
          modelSettings: {
            ...management.modelSettings,
            connections: [{
              id: "saved",
              catalogModelId: "connection:saved",
              label: "已有模型",
              baseUrl: "https://api.example.com",
              modelId: "gpt-4.1",
              supportsVision: false,
              apiKeyConfigured: true,
              apiKeyHint: "••••1234"
            }]
          }
        } as never}
        onClose={noop}
        onAddConnection={noop}
      />
    );
    expect(unavailableMarkup).toContain("部分连接暂不可切换");
    expect(unavailableMarkup).toContain("重试读取");
    expect(unavailableMarkup.match(/<button class="model-connection-select"[^>]*>/)?.[0]).toContain("disabled");
  });

  it("switches models in place and expands editing below the clicked card", async () => {
    const connections = [
      {
        id: "first",
        catalogModelId: firstConnection.id,
        label: "DS",
        baseUrl: "https://api.deepseek.com",
        modelId: "deepseek-chat",
        supportsVision: false,
        apiKeyConfigured: true,
        apiKeyHint: "••••1234"
      },
      {
        id: "second",
        catalogModelId: secondConnection.id,
        label: "vec",
        baseUrl: "https://api.example.com/v1",
        modelId: "claude-opus-4-6",
        supportsVision: false,
        apiKeyConfigured: true,
        apiKeyHint: "••••5678"
      }
    ];
    const onClose = vi.fn();
    function Harness() {
      const management = useModelManagement({ refreshServiceHealth: async () => null });
      useEffect(() => {
        management.setModelSettings({
          baseUrl: null,
          defaultModel: null,
          apiKeyConfigured: false,
          apiKeyHint: null,
          formulaModelId: "",
          connections
        });
      }, []);
      return (
        <WpsModelSettingsPanel
          open
          models={[firstConnection, secondConnection]}
          currentModelId={management.selectedModelId || firstConnection.id}
          management={management}
          onClose={onClose}
          onAddConnection={management.openConnectionCreator}
        />
      );
    }

    const { container } = await mount(<Harness />);
    await until(() => container.querySelectorAll(".model-connection-select").length === 2, "应列出两个模型");
    const cards = container.querySelectorAll<HTMLButtonElement>(".model-connection-select");
    expect(cards[0].disabled).toBe(true);
    expect(cards[1].disabled).toBe(false);
    expect(cards[1].textContent).toContain("点击切换使用");

    const editButtons = container.querySelectorAll<HTMLButtonElement>(".model-connection-actions button:first-child");
    await click(editButtons[1]);
    expect(textOf(container)).toContain("编辑 vec");
    expect(container.querySelector(".model-connection-list")).not.toBeNull();
    expect(container.querySelectorAll(".model-connection-card")[0].querySelector(".model-connection-editor")).toBeNull();
    expect(container.querySelectorAll(".model-connection-card")[1].querySelector(".model-connection-editor")).not.toBeNull();
    expect(textOf(container)).toContain("/function 公式模型");
    expect(textOf(container)).toContain("已保存 ••••5678");
    expect(textOf(container)).toContain("现有密钥不回显");
    const keyInput = container.querySelector<HTMLInputElement>('input[placeholder="输入新 Key 以替换；留空保持不变"]');
    expect(keyInput?.value).toBe("");
    expect(findButton(container, "显示新 Key")).toBeNull();
    await act(async () => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
      setter?.call(keyInput, "replacement-test-key");
      keyInput?.dispatchEvent(new Event("input", { bubbles: true }));
    });
    expect(findButton(container, "显示新 Key")).not.toBeNull();
    await click(findButton(container, "显示新 Key"));
    expect(container.querySelector<HTMLInputElement>('input[value="replacement-test-key"]')?.type).toBe("text");

    await click(findButton(container, "取消"));
    expect(container.querySelector(".model-connection-editor")).toBeNull();
    expect(container.querySelectorAll(".model-connection-select")).toHaveLength(2);
    await click(container.querySelectorAll(".model-connection-select")[1]);
    expect(localStorage.getItem(MODEL_STORAGE_KEY)).toBe(secondConnection.id);
    expect(container.querySelectorAll(".model-connection-card")[1].classList.contains("current")).toBe(true);
    expect(container.querySelector(".wps-model-settings")).not.toBeNull();
    expect(container.querySelector(".wps-model-settings.tool-drawer.settings-drawer")).not.toBeNull();
    expect(container.querySelector(".model-settings-backdrop")).toBeNull();
    expect(onClose).not.toHaveBeenCalled();
  });

  it("keeps the connection form usable after an explicit management read failure", () => {
    const noop = () => undefined;
    const management = {
      modelSettings: null,
      apiKeyDraft: "",
      showApiKey: false,
      connectionDraft: null,
      pendingDeleteConnectionId: null,
      settingsSaving: false,
      settingsTesting: false,
      settingsLoading: false,
      settingsFeedback: "暂时无法连接 Excel Bro 后台：Failed to fetch。",
      connectionTestFeedback: null,
      connectionVerified: false,
      saveApiKey: noop,
      editModelConnection: noop,
      verifyConnection: noop,
      saveConnection: noop,
      removeConnection: noop,
      saveFormulaModel: noop,
      setApiKeyDraft: noop,
      setShowApiKey: noop,
      setConnectionDraft: noop,
      setPendingDeleteConnectionId: noop,
      setSettingsFeedback: noop
    };

    const markup = renderToStaticMarkup(
      <WpsModelSettingsPanel
        open
        models={[]}
        currentModelId=""
        management={management as never}
        onClose={noop}
        onAddConnection={noop}
      />
    );
    const addButton =
      markup.match(/<button type="button"([^>]*)>＋ 添加<\/button>/)?.[1] ?? "";

    expect(addButton).not.toContain("disabled");
    expect(markup).toContain("暂时无法连接 Excel Bro 后台");

    const editorMarkup = renderToStaticMarkup(
      <WpsModelSettingsPanel
        open
        models={[]}
        currentModelId=""
        management={
          {
            ...management,
            connectionDraft: {
              id: null,
              label: "DeepSeek",
              baseUrl: "https://api.deepseek.com",
              modelId: "deepseek-chat",
              apiKey: "test-key",
              clearApiKey: false,
              supportsVision: false
            }
          } as never
        }
        onClose={noop}
        onAddConnection={noop}
      />
    );
    const testButton =
      editorMarkup.match(
        /<button type="button"([^>]*)>测试连接<\/button>/
      )?.[1] ?? "";
    const saveButton =
      editorMarkup.match(
        /<button class="model-primary-action" type="button"([^>]*)>添加并使用<\/button>/
      )?.[1] ?? "";

    expect(testButton).not.toContain("disabled");
    expect(saveButton).toContain("disabled");
  });

  it("keeps the model menu above the conversation cards", () => {
    expect(excelStyles).toMatch(/\.chat-header\s*\{[\s\S]*?z-index:\s*4;/);
    expect(excelStyles).toMatch(
      /\.chat-header\.menu-open\s*\{[\s\S]*?z-index:\s*50;/
    );
    expect(excelStyles).toMatch(/\.model-menu\s*\{[\s\S]*?z-index:\s*45;/);
  });

  it("positions WPS model settings as an in-panel drawer below the header", () => {
    expect(styles).toMatch(/\.wps-model-settings\s*\{[\s\S]*?position:\s*absolute;[\s\S]*?top:\s*53px;/);
  });

  it("keeps the service status readable before collapsing the narrow header", () => {
    expect(styles).toMatch(
      /@media \(max-width: 620px\)[\s\S]*?\.chat-header \.brand-copy span\s*\{[\s\S]*?max-width:\s*none;[\s\S]*?overflow:\s*visible;/
    );
    expect(styles).toMatch(
      /@media \(max-width: 620px\)[\s\S]*?\.chat-header \.new-chat-entry > span:not\(\[aria-hidden="true"\]\)\s*\{[\s\S]*?display:\s*none;/
    );
  });

  it("mirrors the Excel plan card hierarchy for the same pending task", async () => {
    // Excel 用「采用的信息」/「补充说明」两个区块承载服务端下发的口径；WPS
    // 拿到的是同一个 AnalysisPlan，这两个前提不该在预览里丢掉。
    writeRangeHost();
    const { container } = await mountPlanPreview({
      ...singleWritePlan,
      id: "plan-notes",
      assumptions: ["源列空白按「未分类」处理。"],
      warnings: ["汇总表已有数据，写入会覆盖原值。"]
    }, "按部门汇总平均分");

    const notes = Array.from(
      container.querySelector(".wps-plan-card")?.querySelectorAll(".wps-plan-notes") ?? []
    );
    expect(notes.map((note) => note.querySelector("strong")?.textContent?.trim()))
      .toEqual(["采用的信息", "补充说明"]);
    expect(notes[0]?.textContent).toContain("源列空白按「未分类」处理。");
    expect(notes[1]?.textContent).toContain("汇总表已有数据，写入会覆盖原值。");
    // 恢复能力必须写成"会话内"，不能给人永久可撤销的错觉。
    const meta = Array.from(container.querySelectorAll(".wps-plan-meta span"))
      .map((item) => item.textContent?.trim());
    expect(meta).toContain("会话内可撤销");
    // 还没确认，就不得出现任何执行结果。
    expect(container.querySelector(".wps-verification-card")).toBeNull();
    expect(textOf(container)).not.toContain("写入验证通过");
  });

  it("hides both note blocks when the plan states neither", async () => {
    writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");

    expect(container.querySelector(".wps-plan-assumptions")).toBeNull();
    expect(container.querySelector(".wps-plan-warnings")).toBeNull();
    expect(textOf(container)).not.toContain("采用的信息");
  });

  it("marks the preview as edited only after the user rewrote the plan", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");

    expect(textOf(container)).not.toContain("已按你的修改更新");
    await click(findButton(container, "编辑计划"));
    const editBox = container.querySelector<HTMLTextAreaElement>('[aria-label="编辑写入计划 JSON"]');
    await type(editBox, JSON.stringify({ ...singleWritePlan, summary: "改成最大值" }));
    await click(findButton(container, "应用修改并重新预览"));
    await until(() => textOf(container).includes("已按你的修改更新"), "编辑后重新预览必须有可见标识");

    expect(container.querySelector(".wps-plan-card")?.textContent).toContain("写入前请重新核对");
    expect(host.writes()).toBe(0);
    // 标识跟着消息持久化：重新载入面板后仍能知道这份计划被改过。
    const saved = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    const savedPreview = saved.conversations[0].messages.find(
      (message: { planPreview?: { revised?: boolean } }) => message.planPreview
    );
    expect(savedPreview.planPreview.revised).toBe(true);
  });

  it("keeps the original untouched during native review and discards the copy without committing", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); expect(host.cellValue()).toBe('90'); await click(findButton(container,'放弃变更')); expect(host.cellValue()).toBe('旧值'); expect(textOf(container)).toContain('已放弃并恢复'); expect(findButton(container,'确认变更')).toBeNull(); expect(discardWpsNativeComparison).toHaveBeenCalledOnce();
});

  it.each(["close_unconfirmed", "failed", "unsupported", "identity_lost", "ambiguous", "user_saved",
    "extra_edits", "edits_unverified"] as const)(
    "副本未收尾（%s）：放弃仍是终态、不恢复审阅状态，业务结论照旧「本次计划未写入原表」", async status => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); vi.mocked(discardWpsNativeComparison).mockReturnValueOnce(status); await click(findButton(container,'放弃变更')); expect(host.cellValue()).toBe('旧值'); expect(textOf(container)).toContain('已放弃并恢复'); expect(findButton(container,'确认变更')).toBeNull(); expect(findButton(container,'放弃变更')).toBeNull(); expect(textOf(container)).toContain('副本仍保留'); expect(textOf(container)).not.toContain('选择「不保存」'); expect(textOf(container)).not.toContain('正在询问是否保存');
});

  it('审阅按钮接通 WPS 30% 调宽（口径是整个可用宽度），放弃变更恢复普通宽度且原表零写入', async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); expect(findButton(container,'重新排列')).toBeNull(); expect(findButton(container,'双视图对比')).not.toBeNull(); expect(host.cellValue()).toBe('90'); const count=host.writes(); await click(findButton(container,'双视图对比')); expect(host.writes()).toBe(count); expect(findButton(container,'确认变更')).not.toBeNull();
});

  it('无法测量宿主时审阅卡提供真实手动调宽指引，不阻断副本确认入口', async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); await click(findButton(container,'双视图对比')); expect(findButton(container,'确认变更')?.disabled).toBe(false); expect(host.cellValue()).toBe('90'); await click(findButton(container,'放弃变更')); expect(host.cellValue()).toBe('旧值');
});

  it("公式副本计算值不符时重新预览固定值副本，原表仍待用户确认", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); expect(vi.mocked(stageWpsNativeComparison).mock.calls[0]?.[2]).toBe(true); expect(compareCommittedWpsResultToCopy).not.toHaveBeenCalled(); await stageAndSave(container); expect(host.cellValue()).toBe('90');
});

  it("回读不符时不谎报验证通过，但确有恢复点的部分写入仍必须给出卡与撤销入口", async () => {
    // 写入不落盘 → 回读与计划不同。执行器据此判定为「可恢复的部分写入」，
    // 界面要如实说"写入未完成、恢复资料已保留"，**绝不**说验证通过；
    // 也不能因为"不是干净成功"就把恢复点藏起来。
    const host = writeRangeHost({ staleWrite: true });
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");

    await stageAndSave(container);
    await until(() => container.querySelector(".message-row.system") !== null, "失败应给出系统反馈");

    // 负向断言必须先有正向锚点，否则只是"什么都没渲染"的空过。
    expect(textOf(container)).not.toContain("写入验证通过");
    expect(container.querySelector(".wps-verification-card.verified")).toBeNull();
    expect(container.querySelector(".wps-verification-card.unverified")).toBeNull();

    const card = container.querySelector(".wps-verification-card.partial");
    expect(card).not.toBeNull();
    expect(card?.textContent).toContain("恢复资料已保留");
    expect(cardUndoButton(container)).not.toBeNull();

    expect(container.querySelector(".message-row.system")?.textContent)
      .toContain("回读结果与计划不一致");
    // 计划仍在等待确认：可以修正后重试，界面不能假装已经写完了。
    expect(container.querySelector(".wps-plan-card")).not.toBeNull();
    expect(textOf(container)).toContain("写入前确认");
    expect(host.writes()).toBeGreaterThan(0);
  });

  it("写回后计算值与副本不同，不发验证通过卡但结果卡内保留撤销入口", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); expect(compareCommittedWpsResultToCopy).not.toHaveBeenCalled(); expect(container.querySelector('.wps-verification-card.verified')).not.toBeNull(); await stageAndSave(container); expect(cardUndoButton(container)?.disabled).toBe(false);
});

  it("switches the verified card to restored and keeps recovery session-scoped", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");

    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card.verified") !== null,
      "写入成功应给出回读核对结果"
    );
    expect(container.querySelector(".wps-verification-card")?.textContent)
      .toContain("1 个动作 · 1 个单元格已回读核对");

    await click(await undoButton(container));
    await until(() => textOf(container).includes("已撤销并核对恢复"), "撤销后应给出核对过的恢复结果");

    const undone = container.querySelector(".wps-verification-card.undone");
    expect(undone).not.toBeNull();
    expect(undone?.textContent).not.toContain("写入验证通过");
    expect(undone?.textContent).toContain("恢复点仅在当前加载项会话内可用");
    expect(container.querySelector(".wps-verification-card.verified")).toBeNull();
    expect(host.cellValue()).toBe("旧值");
  });

  it("uses native review before commit, then keeps only the verified result and undo summary", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "按计划更新 A1");

    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card.verified") !== null,
      "写入成功应给出回读核对结果"
    );

    // Successful WPS writes no longer reopen the obsolete narrow HTML grid.
    expect(container.querySelector(".change-review-summary")).toBeNull();
    expect(container.querySelector(".change-review-table")).toBeNull();
    expect(container.querySelector(".wps-chat-shell")).not.toBeNull();

    // 明细不进聊天历史：持久化下来的只有定位与计数，没有单元格内容。
    const saved = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    expect(saved).not.toBeNull();
    // "旧值"是单元格内容，不该出现在被持久化的对话里（时间戳之类的数字串不适合
    // 做这类断言，故另设哨兵用例，见下一条）。
    expect(JSON.stringify(saved)).not.toContain("旧值");

    await click(await undoButton(container));
    await until(() => textOf(container).includes("已撤销并核对恢复"), "撤销后应给出核对过的恢复结果");
    // 区域已回到执行前，不能继续挂着过期的变更明细。
    expect(container.querySelector(".change-review-summary")).toBeNull();
    expect(host.cellValue()).toBe("旧值");
    // 但"这次确实改过、后来撤销了"是操作记录，不该跟着明细一起被抹掉；
    // 留下来的形态只有定位与计数，不含任何单元格内容。
    const note = container.querySelector(".change-review-undo-note");
    expect(note).not.toBeNull();
    expect(note?.textContent).toContain("Sheet1!A1");
    expect(note?.textContent).toContain("改动 1 / 1");
    expect(note?.textContent).not.toContain("旧值");
    expect(note?.textContent).not.toContain("90");
  });

  it("keeps the 'changed then undone' note free of cell values in the UI and in history", async () => {
    const SENTINEL = "私密撤销值-Q7W2";
    writeRangeHost();
    const { container } = await mountPlanPreview({
      ...singleWritePlan,
      id: "plan-undo-sentinel",
      actions: [{ type: "writeValues", sheet: "Sheet1", range: "A1", values: [[SENTINEL]] }]
    }, "按计划更新 A1");

    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card.verified") !== null,
      "写入成功应给出回读核对结果"
    );
    expect(container.querySelector(".change-review-summary")).toBeNull();

    await click(await undoButton(container));
    await until(
      () => container.querySelector(".change-review-undo-note") !== null,
      "撤销后应留下「曾修改、现已撤销」的摘要"
    );
    // 明细已释放，摘要里也不该有单元格内容。
    expect(container.querySelector(".change-review-summary")).toBeNull();
    expect(container.querySelector(".change-review-undo-note")?.textContent).not.toContain(SENTINEL);
    const saved = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    expect(JSON.stringify(saved)).not.toContain(SENTINEL);
  });

  it("attaches the change review to the failure message and never claims success", async () => {
    writeRangeHost({ staleWrite: true });
    const { container } = await mountPlanPreview(singleWritePlan, "按计划更新 A1");

    await stageAndSave(container);
    await until(() => container.querySelector(".message-row.system") !== null, "失败应给出系统反馈");

    const summary = container.querySelector(".message-row.system .change-review-summary");
    expect(summary).not.toBeNull();
    expect(summary!.textContent).toContain("回读与预期不一致");
    expect(summary!.querySelector("button")).toBeNull();
    expect(container.textContent).not.toContain("打开全幅双视图");
    expect(container.querySelector(".wps-review-shell")).toBeNull();
    expect(summary!.textContent).not.toContain("写入验证通过");
  });

  it("keeps cell values out of the persisted conversation even when a review was delivered", async () => {
    // 哨兵只存在于工作簿单元格与回读矩阵里：用户提示、计划标题、活动日志都不含它。
    // 这样才能证明"明细不进历史"是真的，而不是碰巧没撞上子串。
    const SENTINEL = "单元格私密值-K9Z3";
    writeRangeHost();
    const { container } = await mountPlanPreview({
      ...singleWritePlan,
      id: "plan-sentinel",
      actions: [{ type: "writeValues", sheet: "Sheet1", range: "A1", values: [[SENTINEL]] }]
    }, "按计划更新 A1");

    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card.verified") !== null,
      "写入成功应给出回读核对结果"
    );
    // 成功结果不再提供旧的窄 HTML 对照；副本逐格审阅由严格原生宿主用例覆盖。
    expect(container.querySelector(".change-review-summary")).toBeNull();

    const saved = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    expect(saved).not.toBeNull();
    expect(JSON.stringify(saved)).not.toContain(SENTINEL);
  });

  describe("待答澄清的授权身份", () => {
    /**
     * 造一台可切换工作表身份的宿主。
     *
     * 澄清卡只有在**授权身份没变**时才允许继续回答。索引只是"第几张表"，
     * 重排/替换之后同一个索引会指向另一张表 —— 那时用户对着旧表回答的问题
     * 会被拿去算新表，授权边界形同虚设。所以这里的判据是**名称**，不是索引。
     */
    function ambiguityHost(sheetNames: string[]) {
      const sheets = sheetNames.map((name) => ({ Name: name, ProtectContents: false }));
      const storage = new Map<string, unknown>();
      const application = {
        Version: "12.0",
        ActiveWorkbook: {
          Name: "测试.xlsx",
          FullName: "C:\\Data\\测试.xlsx",
          Path: "C:\\Data",
          ReadOnly: false,
          Worksheets: {
            Count: sheets.length,
            Item: (index: number) => sheets[Number(index) - 1]
          },
          ActiveSheet: sheets[0]
        },
        ActiveSheet: sheets[0],
        PluginStorage: {
          getItem: (key: string) => storage.get(key),
          setItem: (key: string, item: unknown) => { storage.set(key, item); }
        },
        CreateTaskPane: () => ({ ID: 1, Visible: true }),
        GetTaskPane: () => undefined
      } as never;
      return {
        application,
        rename: (position: number, name: string) => { sheets[position].Name = name; },
        reorder: (next: string[]) => {
          const byName = new Map(sheets.map((sheet) => [sheet.Name, sheet]));
          next.forEach((name, position) => {
            const sheet = byName.get(name);
            if (sheet) sheets[position] = sheet;
          });
        }
      };
    }

    /** 让这一轮返回一张**需求澄清**卡（与列映射归因无关，专测范围失效判定）。 */
    function ambiguityCard() {
      runTurnMock.mockResolvedValue({
        kind: "clarification",
        clarification: {
          response: {
            kind: "clarification",
            provider: "model",
            clarification: {
              id: `clarify-${Math.random().toString(16).slice(2)}`,
              summary: "需要明确目标",
              question: "要查询还是修改？",
              reason: "两种操作不同",
              scopeLabel: "当前表",
              options: [
                { id: "read", label: "查询", description: "只读", resolution: "查询数据" }
              ]
            }
          }
        }
      });
    }

    /** 走到待答澄清卡为止。 */
    async function pendingClarification() {
      const { container } = await mount(panel());
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
      const composer = container.querySelector<HTMLTextAreaElement>(".composer textarea");
      await type(composer, "帮我处理表格");
      await submit(composer);
      await until(
        () => Boolean(container.querySelector(".clarification-card")),
        "应显示澄清卡"
      );
      await until(() => runTurnMock.mock.calls.length === 1, "应完成这一轮");
      return container;
    }

    /** 卡片上的选项按钮 —— 与"取消本次澄清"区分开。 */
    function clarificationOptions(container: HTMLElement): HTMLButtonElement[] {
      return Array.from(
        container.querySelectorAll<HTMLButtonElement>(
          ".clarification-options button"
        )
      );
    }

    it("同索引换成另一张表：待答澄清失效，不拿旧答案继续算", async () => {
      listModelsMock.mockResolvedValue({
        defaultModelId: firstConnection.id,
        models: [firstConnection]
      });
      ambiguityCard();
      const host = ambiguityHost(["订单", "明细"]);
      getWpsApplicationMock.mockReturnValue(host.application);

      const container = await pendingClarification();
      expect(clarificationOptions(container)).toHaveLength(1);
      // 授权身份变了：第 1 张表从「订单」换成「客户」。
      host.rename(0, "客户");

      await click(clarificationOptions(container)[0]);
      await until(
        () => container.querySelector('[role="dialog"]') !== null,
        "同索引换表后必须先确认操作目标"
      );
      // 没有第二轮：失效的澄清不能带着旧上下文继续跑。
      expect(runTurnMock).toHaveBeenCalledTimes(1);
      expect(textOf(container)).toContain("检测到文件或工作表变化");
      expect(container.querySelector(".clarification-card.invalidated")).toBeNull();
    });

    it("工作表重排：索引指向了另一张表，待答澄清失效", async () => {
      listModelsMock.mockResolvedValue({
        defaultModelId: firstConnection.id,
        models: [firstConnection]
      });
      ambiguityCard();
      const host = ambiguityHost(["订单", "明细"]);
      getWpsApplicationMock.mockReturnValue(host.application);

      const container = await pendingClarification();
      // 重排：原第 2 张「明细」挪到第 1 位，索引 1 现在指向另一张表。
      host.reorder(["明细", "订单"]);

      await click(clarificationOptions(container)[0]);
      await until(
        () => container.querySelector('[role="dialog"]') !== null,
        "重排后必须先确认操作目标"
      );
      expect(runTurnMock).toHaveBeenCalledTimes(1);
    });

    it("已授权表改名（顺序与集合都不变）：待答澄清失效并要求重新授权", async () => {
      listModelsMock.mockResolvedValue({
        defaultModelId: firstConnection.id,
        models: [firstConnection]
      });
      ambiguityCard();
      const host = ambiguityHost(["订单", "明细"]);
      getWpsApplicationMock.mockReturnValue(host.application);

      const container = await pendingClarification();
      host.rename(0, "订单（新）");

      await click(clarificationOptions(container)[0]);
      await until(
        () => container.querySelector('[role="dialog"]') !== null,
        "改名后必须先确认操作目标"
      );
      expect(runTurnMock).toHaveBeenCalledTimes(1);
    });

    it("授权范围与身份都没变：澄清可以正常继续", async () => {
      listModelsMock.mockResolvedValue({
        defaultModelId: firstConnection.id,
        models: [firstConnection]
      });
      runTurnMock
        .mockResolvedValueOnce({
          kind: "clarification",
          clarification: {
            response: {
              kind: "clarification",
              provider: "model",
              clarification: {
                id: "clarify-keep",
                summary: "需要明确目标",
                question: "要查询还是修改？",
                reason: "两种操作不同",
                scopeLabel: "当前表",
                options: [
                  { id: "read", label: "查询", description: "只读", resolution: "查询数据" }
                ]
              }
            }
          }
        })
        .mockResolvedValueOnce({
          kind: "answer",
          message: "已按补充说明处理。",
          resultContext: null,
          provider: "model"
        });
      const host = ambiguityHost(["订单", "明细"]);
      getWpsApplicationMock.mockReturnValue(host.application);

      const container = await pendingClarification();
      // 只做一次**无副作用**的读取类调用：身份不变，卡片必须保留。
      await click(clarificationOptions(container)[0]);
      await until(() => runTurnMock.mock.calls.length === 2, "身份未变时应继续原任务");
      expect(container.querySelector(".clarification-card.resolved")).not.toBeNull();
      expect(textOf(container)).not.toContain("数据范围已经变化");
    });
  });

  /**
   * 侧栏信息层级的**结构与状态**断言。
   *
   * 这里能覆盖的是"哪些元素在哪、什么时候出现、按钮叫什么"，**不是布局**：
   * jsdom 不计算真实排版，`.wps-native-review-actions` 里的 DOM 断言无法证明
   * 按钮在 350/430/700 CSS px 下没有被挤压。CSS 契约只锁住"声明本身还在"，
   * 真机视口表现必须另外用 WPS 截图验收（见
   * `docs/WPS_NATIVE_REVIEW_UI_GUIDE_2026-09-30.md` 第 4 节）。
   */
describe("撤销按钮归属：执行结果卡内，按 executionId 绑定", () => {
    /** 再跑一轮同样的计划（同一容器、同一宿主），用于"连续执行"的用例。 */
    async function runAnotherPlan(container: HTMLElement, task: string) {
      const textarea = container.querySelector("textarea");
      await type(textarea, task);
      await submit(textarea);
      await until(
        () => findButton(container, "生成右侧副本并排审阅") !== null,
        "应展示新一轮计划预览"
      );
      await stageAndSave(container);
    }

    it("写入成功后：按钮在对应结果卡内，消息流末尾不再有全局撤销节点", async () => {
      writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await stageAndSave(container);
      await until(
        () => container.querySelector(".wps-verification-card") !== null,
        "确认写入后应给出执行结果卡"
      );

      // 正向锚点：按钮确实在产生恢复点的那张卡里，并且卡上写明状态。
      const card = container.querySelector(".wps-verification-card");
      expect(cardUndoButton(container)).not.toBeNull();
      expect(card?.querySelector(".wps-undo-inline")?.textContent).toContain("可撤销");

      // 负向：那张无条件挂在消息流末尾的全局撤销卡不许回来……
      expect(container.querySelector(".wps-undo-card")).toBeNull();
      // ……而且消息流里任何同文案按钮都必须落在某张卡（结果卡/备份卡）里面，
      // 不允许存在"没有卡片容器的裸按钮"。
      const buttons = flowUndoButtons(container);
      expect(buttons.length).toBeGreaterThan(0);
      for (const button of buttons) {
        expect(button.closest(".wps-verification-card, .wps-backup-card")).not.toBeNull();
      }

      // 抽屉仍保留快捷入口，并标明它属于哪一次执行。
      const opener = container.querySelector<HTMLButtonElement>('button[aria-label="备份与恢复"]');
      await click(opener);
      expect(await undoButton(container)).not.toBeNull();
      expect(container.querySelector(".wps-backup-card")?.textContent).toContain("所属执行");
    });

    it("之后来一条无关回复，按钮仍留在旧结果卡里，不会被赶到最新回复下面", async () => {
      writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await stageAndSave(container);
      await until(
        () => container.querySelector(".wps-verification-card") !== null,
        "确认写入后应给出执行结果卡"
      );
      const card = container.querySelector(".wps-verification-card")!;

      runTurnMock.mockResolvedValue({
        kind: "answer",
        message: "好的，还有什么需要处理的？",
        provider: "model",
        resultContext: null
      });
      const textarea = container.querySelector("textarea");
      await type(textarea, "随便聊聊");
      await submit(textarea);
      await until(() => textOf(container).includes("还有什么需要处理的"), "应显示新的普通回复");

      // 旧卡（同一个 DOM 节点）连同按钮一起按时间顺序往上走，没有被重建或挪走。
      expect(container.contains(card)).toBe(true);
      expect(cardUndoButton(container)).not.toBeNull();

      // 最新那条普通回复里没有任何撤销入口。
      const rows = Array.from(container.querySelectorAll<HTMLElement>(".message-row"));
      const last = rows[rows.length - 1];
      expect(last.textContent).toContain("还有什么需要处理的");
      expect(last.querySelector(".wps-undo-inline")).toBeNull();
      expect(container.querySelector(".wps-undo-card")).toBeNull();
    });

    it("连续两次执行：只有当前活跃恢复点那一次可点，旧卡降级为「恢复点已不可用」", async () => {
      writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "第一次更新 A1");
      await stageAndSave(container);
      await until(
        () => container.querySelectorAll(".wps-verification-card").length === 1,
        "第一次执行应给出一张结果卡"
      );

      await runAnotherPlan(container, "第二次更新 A1");
      await until(
        () => container.querySelectorAll(".wps-verification-card").length === 2,
        "第二次执行后应有两张结果卡"
      );

      const cards = Array.from(container.querySelectorAll<HTMLElement>(".wps-verification-card"));
      // 旧卡在前、新卡在后：它是按时间顺序待在自己那条消息里的。
      expect(cards[0].querySelector(".wps-undo-inline")?.textContent).toContain("恢复点已不可用");
      expect(cards[0].querySelector('button[aria-label="撤销本次执行"]')).toBeNull();
      expect(cards[1].querySelector(".wps-undo-inline")?.textContent).toContain("可撤销");
      expect(cards[1].querySelector('button[aria-label="撤销本次执行"]')).not.toBeNull();
      // 全局节点依旧不存在。
      expect(container.querySelector(".wps-undo-card")).toBeNull();
    });

    it("撤销冲突：执行后区域又被改动时，卡片改为「撤销冲突」且不再可点", async () => {
      const host = writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await stageAndSave(container);
      await until(
        () => container.querySelector(".wps-verification-card") !== null,
        "确认写入后应给出执行结果卡"
      );
      expect(cardUndoButton(container)).not.toBeNull();

      // 模拟用户在执行之后手工改了这块区域。
      host.setCell("人工改过的值");

      await click(cardUndoButton(container));
      await until(() => textOf(container).includes("又被修改"), "应给出冲突提示");

      const card = container.querySelector(".wps-verification-card")!;
      expect(card.querySelector(".wps-undo-inline")?.textContent).toContain("撤销冲突");
      expect(card.querySelector('button[aria-label="撤销本次执行"]')).toBeNull();
      // 冲突不等于"已经被撤销"：区域里是用户后来写的内容，不能被覆盖。
      expect(card.className).not.toContain("undone");
      expect(host.cellValue()).toBe("人工改过的值");
    });

    it("面板重开（新的加载项会话）：旧结果卡只显示「恢复点已不可用」，没有可点的假按钮", async () => {
      writeRangeHost();
      const first = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await stageAndSave(first.container);
      await until(
        () => cardUndoButton(first.container) !== null,
        "第一次会话里应给出可点的撤销入口"
      );
      first.unmount();
      mounted = mounted.filter((entry) => entry !== first);

      // 重开：聊天历史（localStorage）还在，但当前加载项会话里已没有恢复点。
      writeRangeHost();
      const second = await mount(panel());
      await until(
        () => second.container.querySelector(".wps-verification-card") !== null,
        "历史里的执行结果卡仍应显示"
      );

      const card = second.container.querySelector(".wps-verification-card")!;
      expect(card.querySelector(".wps-undo-inline")?.textContent).toContain("恢复点已不可用");
      expect(cardUndoButton(second.container)).toBeNull();
      // 也不能靠抽屉里的残留入口变相提供一个可点的假按钮。
      expect(await undoButton(second.container)).toBeNull();
    });

    it("新建独立对话：旧执行卡留在旧对话，不会移到新对话的消息末尾", async () => {
      writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await stageAndSave(container);
      await until(
        () => container.querySelector(".wps-verification-card") !== null,
        "确认写入后应给出执行结果卡"
      );

      await click(container.querySelector('button[aria-label="新对话"]'));
      await until(() => textOf(container).includes("已新建对话"), "应新建对话");

      // 新对话里既没有那张卡，也没有任何撤销入口。
      expect(container.querySelector(".wps-verification-card")).toBeNull();
      expect(flowUndoButtons(container).length).toBe(0);
      expect(container.querySelector(".wps-undo-card")).toBeNull();

      // 旧执行卡留在旧对话：切回去还在，入口仍在卡里。
      await click(container.querySelector('[aria-label="更多"]'));
      await click(container.querySelector('button[aria-label="历史对话"]'));
      const entry = Array.from(container.querySelectorAll<HTMLElement>(".history-open-button")).find(
        (item) => item.textContent?.includes("把 A1 改成 90")
      );
      await click(entry ?? null);
      await until(
        () => container.querySelector(".wps-verification-card") !== null,
        "旧对话里的执行结果卡应保留"
      );
      expect(cardUndoButton(container)).not.toBeNull();
      expect(container.querySelector(".wps-undo-card")).toBeNull();
    });

    it("成功撤销后：所属执行卡变为「已撤销」，且不再残留可点按钮", async () => {
      const host = writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await stageAndSave(container);
      await until(
        () => container.querySelector(".wps-verification-card.verified") !== null,
        "写入成功应给出回读核对结果"
      );

      await click(cardUndoButton(container));
      await until(
        () => container.querySelector(".wps-verification-card.undone") !== null,
        "撤销后所属执行卡应变为已撤销"
      );

      const card = container.querySelector(".wps-verification-card")!;
      expect(card.querySelector(".wps-undo-inline")?.textContent).toContain("已撤销");
      expect(cardUndoButton(container)).toBeNull();
      expect(container.querySelector(".wps-verification-card.verified")).toBeNull();
      expect(host.cellValue()).toBe("旧值");
      // 抽屉也不该再提供入口。
      expect(await undoButton(container)).toBeNull();
    });
  });

  describe("审阅侧栏的信息层级", () => {
    it("keeps the long review status out of the action row so buttons never fight it for width", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); const bar=container.querySelector('.result-review-bar')!; expect(bar.closest('.message-row')).toBeNull(); expect(bar.querySelector('.result-review-actions')?.textContent).toBe('双视图对比放弃变更确认变更'); expect(bar.querySelector('[role="status"]')?.closest('.result-review-actions')).toBeNull();
});

    it("names the two commit actions unambiguously and hides them before staging", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); expect(findButton(container,'确认变更')).not.toBeNull(); expect(findButton(container,'放弃变更')).not.toBeNull(); expect(findButton(container,'确认写回原表')).toBeNull(); expect(findButton(container,'重新排列')).toBeNull();
});

    it("disables both commit actions while the commit is in flight", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); const button=findButton(container,'确认变更')!; const key=Object.keys(button).find(x=>x.startsWith('__reactProps$'))!; const props=(button as unknown as Record<string,{onClick:()=>void}>)[key]; const count=host.writes(); await act(async()=>{props.onClick();props.onClick();}); expect(host.writes()).toBe(count); expect(textOf(container)).toContain('已确认变更');
});

    it("shows the actual before/after change on the first screen, before staging", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); expect(host.cellValue()).toBe('90'); expect(textOf(container)).toContain('请查看主表中的执行结果'); expect(textOf(container)).toContain('Sheet1!A1');
});

    it("marks the before column as unknown instead of fabricating a value when the host exposes no values", async () => {
      // 宿主不交出 `Value2`（只有 `Formula`）：`beforeValuesCaptured` 为 false。
      // 这时"修改前"必须如实标未知 —— 绝不拿公式原文或空白去冒充原值。
      const value = "目标值";
      let formula = "目标值";
      const range = { Rows: { Count: 1 }, Columns: { Count: 1 }, MergeCells: false } as Record<string, unknown>;
      Object.defineProperties(range, {
        // `Value2` 读出来是 undefined —— 与"空单元格"不同，这里根本没有值通道。
        Value2: { get: undefined },
        Formula: {
          get: () => formula,
          set: (next: string) => { formula = next; },
          enumerable: true
        }
      });
      const sheet = { Name: "Sheet1", ProtectContents: false, Range: () => range };
      getWpsApplicationMock.mockReturnValue({
        Version: "12.0",
        ActiveWorkbook: {
          Name: "测试.xlsx",
          FullName: "C:\\Data\\测试.xlsx",
          Path: "C:\\Data",
          ReadOnly: false,
          Worksheets: { Count: 1, Item: () => sheet },
          ActiveSheet: sheet
        },
        ActiveSheet: sheet,
        PluginStorage: { getItem: () => undefined, setItem: () => undefined },
        CreateTaskPane: () => ({ ID: 1, Visible: true }),
        GetTaskPane: () => undefined
      } as never);

      listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
      runTurnMock.mockResolvedValue({ kind: "plan", plan: singleWritePlan, turnId: "turn-preview-1" });
      validatePlanMock.mockImplementation(async (submitted) => submitted);
      void value;
      const { container } = await mount(panel());
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型目录");
      const textarea = container.querySelector("textarea");
      await type(textarea, "把 A1 改成 90");
      await submit(textarea);
      await until(() => findButton(container, "确认变更") !== null || textOf(container).includes("本次执行需要核对"), "应展示执行结果或明确失败");

      const preview = container.querySelector(".wps-plan-change-preview")!;
      // 正向锚点：确实渲染了定位与"修改后"，否则下面的负向断言是空过。
      expect(preview.textContent).toContain("A1");
      expect(preview.textContent).toContain("90");
      // 负向：没有原值通道时，不许凭空写出"修改前" —— 只能标"原值未取得"。
      expect(preview.querySelectorAll(".wps-plan-change-unknown").length).toBeGreaterThan(0);
      expect(preview.textContent).toContain("原值未取得");
    });

    it("collapses the long body by default while keeping summary and actions on screen", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); const card=container.querySelector('.execution-result-card')!; expect(card.querySelector('details')?.open).toBe(false); expect(container.querySelector('.result-review-bar')).not.toBeNull();
});

    it("labels expected values and copy read-back as different things", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); expect(textOf(container)).toContain('双视图中的副本展示执行前状态'); expect(textOf(container)).not.toContain('原表尚未修改'); expect(host.writes()).toBe(1);
});

    it("states formula trial and fixed-value fallback with different wording", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); expect(stageWpsNativeComparison).toHaveBeenCalledOnce(); expect(vi.mocked(stageWpsNativeComparison).mock.calls[0]?.[2]).toBe(true); expect(compareCommittedWpsResultToCopy).not.toHaveBeenCalled();
});

    it("keeps the CSS contract that two buttons stay readable at narrow and wide widths", () => {
      // 再次强调：这些是**声明层面**的契约，不能替代真机视口验收。
      // 真机 350/430 px 下"按钮没有逐字折行"只能靠 WPS 截图证明。
      //
      // 正则必须用 `[^}]*`（**不许跨 `}`**）：用 `[\s\S]*?` 会一路匹配到
      // 后面别的规则里的同名声明，断言就变成恒真的假绿灯 —— 这一点用一个
      // 删掉 `white-space: nowrap` 的反向变异验过。
      expect(styles).toMatch(/\.wps-native-review-actions button\s*\{[^}]*white-space:\s*nowrap;/);
      expect(styles).toMatch(/\.wps-native-review-actions button\s*\{[^}]*min-height:\s*44px;/);
      // 窄窗格整行纵排。
      expect(styles).toMatch(/@media \(max-width:\s*559px\)\s*\{[^}]*flex-basis:\s*100%/);
      // 旧的单行 flex 容器不得复活：说明和按钮再同排就会重现真机竖排。
      expect(styles).not.toMatch(/\.wps-plan-buttons\s*\{/);
      // 长公式走等宽 + 横向滚动，不再任意折断。
      expect(styles).toMatch(
        /\.wps-plan-values th,\s*\.wps-plan-values td\s*\{[^}]*white-space:\s*nowrap;/
      );
      expect(styles).not.toMatch(
        /\.wps-plan-values th,\s*\.wps-plan-values td\s*\{[^}]*overflow-wrap:\s*anywhere/
      );
    });
  });
});

/**
 * 审阅卡的**四个阶段**逐阶段验收：审阅中 / 已放弃 / 写回成功 / 写回失败。
 *
 * 这一组用例盯的不是算法，而是**界面层级**：哪个阶段说什么话、确认框与按钮在不在、
 * 副本状态是不是独立显示、不同的结果有没有互相冒充。四条用例共用同一份"事实"入参，
 * 所以任何一处文案串档都会在这里变红。
 */
describe("审阅卡四阶段界面（审阅中 / 已放弃 / 写回成功 / 写回失败）", () => {
  /**
   * 读取 React 挂在 DOM 节点上的 props（`__reactProps$xxx`）。
   *
   * 「旧回调」没法从 DOM 上点得到 —— 放弃的那一刻按钮就没了。要真的拿到
   * **放弃之前那个闭包**，只能从节点上把当时的 props 取下来。取不到就返回 null，
   * 用例里会先断言"确实拿到了"，避免负向断言变成空过。
   */
  function reactProps(node: Element): Record<string, unknown> | null {
    const key = Object.keys(node).find((item) => item.startsWith("__reactProps$"));
    return key ? (node as unknown as Record<string, Record<string, unknown>>)[key] : null;
  }

  it("审阅中：首屏只留五项，确认框是真闸，排查材料全部下沉", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); const bar=container.querySelector('.result-review-bar')!; expect(bar.textContent).toContain('主表已更新'); expect(bar.querySelector('input[type="checkbox"]')).toBeNull(); expect(bar.querySelectorAll('button')).toHaveLength(3); expect(findButton(container,'确认变更')?.disabled).toBe(false); expect(findButton(container,'保存为工具')).toBeNull();
});

  it("已放弃：终态撤下确认框与原审阅按钮，旧回调也拒绝继续写回", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); const button=findButton(container,'确认变更')!; const props=reactProps(button); expect(typeof props?.onClick).toBe('function'); await click(findButton(container,'放弃变更')); const count=host.writes(); await act(async()=>{(props!.onClick as ()=>void)();}); expect(host.writes()).toBe(count); expect(host.cellValue()).toBe('旧值'); expect(findButton(container,'确认变更')).toBeNull(); expect(textOf(container)).toContain('已放弃并恢复');
});

  it("写回成功：业务结论在前，副本状态独立显示，不再保留任何审阅入口", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); const count=host.writes(); await stageAndSave(container); expect(host.writes()).toBe(count); expect(textOf(container)).toContain('已确认变更'); expect(findButton(container,'确认变更')).toBeNull(); expect(findButton(container,'放弃变更')).toBeNull(); expect(cardUndoButton(container)).not.toBeNull();
});

  it("写回失败：按实际状态展示，绝不套成功模板、也不混用已放弃口径", async () => {
    // 写入不落盘 ⇒ 回读与计划不符 ⇒ 执行器判为「可恢复的部分写入」。
    const host = writeRangeHost({ staleWrite: true });
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await click(findButton(container, "生成右侧副本并排审阅"));
    await acknowledgeCopy(container);
    await click(findButton(container, "确认写回原表"));
    await until(
      () => container.querySelector(".message-row.system") !== null,
      "失败应给出系统反馈"
    );

    // 正向锚点：失败被如实说出来，否则下面的负向断言都是空过。
    expect(textOf(container)).toContain("回读结果与计划不一致");
    // 绝不套成功模板。
    expect(container.querySelector(".wps-verification-card.verified")).toBeNull();
    expect(textOf(container)).not.toContain("写入验证通过");
    expect(textOf(container)).not.toContain("个单元格已回读核对");
    // 也绝不混用"已放弃"那一档：原表**可能已经被改过**，不能说成"未写入"。
    expect(textOf(container)).not.toContain("本次计划未写入原表。");
    expect(textOf(container)).not.toContain("已放弃变更");
    // 排查材料留在现场：结果卡与恢复点都还在，撤销入口可用。
    expect(container.querySelector(".wps-verification-card.partial")).not.toBeNull();
    expect(textOf(container)).toContain("写入未完成，恢复资料已保留");
    expect(cardUndoButton(container)).not.toBeNull();
    expect(host.writes()).toBeGreaterThan(0);
  });
});

/**
 * 异步写回闸门：授权检查回来之后，必须仍是**同一份**现场。
 *
 * 「当前有审阅」只说明"有"，不说明"还是那一次"。旧请求在途期间用户点了「放弃变更」、
 * 甚至放弃之后又生成了新副本并勾了新的确认框 —— 那时"计划与副本都在场 + 已勾选"
 * **全部成立**，旧请求若据此继续写回，就是拿新副本的确认状态去写原表：用户核对过的
 * 那一份从来没被写，没核对过的那一份反而被写了。
 */
describe("异步写回闸门：授权检查回来后必须仍是同一份现场", () => {
  /**
   * 从 React 挂在 DOM 上的 props 里取出**那一刻**的回调。
   *
   * 写回一旦在途，`busy` 就会把所有入口禁用 —— 按钮上的 onClick 也换成了新的。
   * 要拿到"用户在这个瞬间点的那一下"，只能从节点上把当时的 props 取下来
   * （已排队的点击、快捷键、程序直调在真机上走的就是这条闭包）。
   */
  function staleHandler(node: Element): (() => void) | null {
    const key = Object.keys(node).find((item) => item.startsWith("__reactProps$"));
    const props = key ? (node as unknown as Record<string, { onClick?: () => void }>)[key] : null;
    return typeof props?.onClick === "function" ? props.onClick : null;
  }

  it("判据逐项分开：换了副本 / 换了计划 / 换了原工作簿，各自都能单独拦下", () => {
    const reviewA = { copy: "Book1" }, reviewB = { copy: "Book2" };
    const planA = { revision: 1 }, planB = { revision: 2 };
    const book = "url:C:/data/格式验收.xlsx";
    const targetOf = (review: unknown, plan: unknown, workbookIdentity: string): ReviewWriteBackTarget =>
      ({ review, plan, workbookIdentity });
    const now = (overrides: Partial<{ acknowledged: boolean; plan: unknown; review: unknown; target: ReviewWriteBackTarget | null }> = {}) =>
      ({ acknowledged: true, plan: planA, review: reviewA, target: targetOf(reviewA, planA, book), ...overrides });

    // 正向锚点：同一份现场 + 已勾选 ⇒ 放行（否则下面几条"拦下"都是恒真的）。
    expect(wpsWriteBackStillCurrent(targetOf(reviewA, planA, book), now())).toBe(true);

    // ① 换了副本：放弃之后又生成了新副本，并且勾了新的确认框 —— 现场"在"，但那是别人的确认。
    expect(wpsWriteBackStillCurrent(targetOf(reviewA, planA, book),
      now({ review: reviewB, target: targetOf(reviewB, planA, book) }))).toBe(false);
    // ② 换了计划：计划被替换 / 编辑后重新校验。
    expect(wpsWriteBackStillCurrent(targetOf(reviewA, planA, book),
      now({ plan: planB, target: targetOf(reviewA, planB, book) }))).toBe(false);
    // ③ 换了原工作簿。
    expect(wpsWriteBackStillCurrent(targetOf(reviewA, planA, book),
      now({ target: targetOf(reviewA, planA, "url:C:/data/另一本.xlsx") }))).toBe(false);
    // ④ 放弃：计划与副本都不在场 ⇒ 现场身份拿不到，一律按已变化处理。
    expect(wpsWriteBackStillCurrent(targetOf(reviewA, planA, book), { target: null })).toBe(false);
    // ⑤ 请求发起时就没有现场 ⇒ 一律不继续（fail-closed）。
    expect(wpsWriteBackStillCurrent(null, now())).toBe(false);
  });

  /*
   * 关于"旧请求在途期间放弃并生成新副本"这条**复合**路径：
   *
   * 目前它在界面上**走不出来** —— `busy` 含 `targetGuard.checking`，写回等待期间
   * 放弃 / 生成副本 / 发新任务的入口全部禁用，`runTurn` 自己也挡 `busy`。所以这里
   * 能端到端证明的是"放弃 ⇒ 旧请求零写入"；"换了副本但现场仍在"那一条由上面的
   * 判据用例直接判定（`wpsWriteBackStillCurrent`），它是这道闸唯一管辖的事。
   */
  it("放弃是终态：授权检查在途期间点「放弃变更」，旧请求回来后原表零写入", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "先并排审阅");
    await click(findButton(container, "生成右侧副本并排审阅"));
    await acknowledgeCopy(container);

    const confirm = findButton(container, "确认写回原表");
    const staleDiscard = staleHandler(findButton(container, "放弃变更")!);
    expect(confirm).not.toBeNull();
    expect(staleDiscard).not.toBeNull();

    // 把授权检查挂起：请求就停在 `await targetGuard.check()` 那一行，窗口是真的打开着的。
    holdInspect();
    await act(async () => {
      confirm!.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });
    await until(() => inspectPending(), "授权检查应停在闸门里");

    // 在途期间用户点了「放弃变更」：计划与副本立刻撤场。
    await act(async () => { staleDiscard!(); });
    await settle();
    expect(textOf(container)).toContain("已放弃变更");
    expect(host.writes()).toBe(0);

    // 放行授权检查 ⇒ 复核必须拦下：旧请求不能借着"看起来还在"的现场继续写回。
    await releaseInspect();
    await settle();
    expect(host.writes()).toBe(0);
    expect(container.querySelector(".wps-verification-card")).toBeNull();
    expect(textOf(container)).toContain("本次计划未写入原表。");
  });

  it("同一条闸门不放行也不乱拦：没有放弃时，同样的挂起-放行能正常写回（正向锚点）", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "先并排审阅");
    await click(findButton(container, "生成右侧副本并排审阅"));
    await acknowledgeCopy(container);

    const confirm = findButton(container, "确认写回原表")!;
    holdInspect();
    await act(async () => {
      confirm.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });
    await until(() => inspectPending(), "授权检查应停在闸门里");
    expect(host.writes()).toBe(0);   // 挂在闸门里时还没写 —— 证明这个窗口是真的。

    await releaseInspect();
    await until(() => container.querySelector(".wps-verification-card.verified") !== null, "放行后应正常写回并验证");
    expect(host.writes()).toBeGreaterThan(0);
  });
});

/**
 * 输入框的「上拉」：与 Excel 端共用 `useComposerResize`。
 *
 * 这里能覆盖的是"把手在哪、叫什么、高度和滚动怎么变"；jsdom 不排版，所以
 * 真实的拖动手感与窄窗格表现仍须在 WPS 真机里验收。
 */
describe("输入框上拉：与 Excel 端同一套行为", () => {  it("把手挂在输入框**外面**的 region 上，文案来自共享常量", async () => {
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型目录");

    const handle = container.querySelector<HTMLElement>(".composer-resize-handle");
    expect(handle).not.toBeNull();
    expect(handle?.getAttribute("role")).toBe("separator");
    expect(handle?.getAttribute("aria-label")).toBe(COMPOSER_RESIZE_LABEL);
    expect(handle?.getAttribute("aria-orientation")).toBe("horizontal");
    expect(handle?.getAttribute("tabindex")).toBe("0");
    // 必须待在 region 的上内边距里：挪进 composer-box 会盖住文本区、抢走光标。
    expect(handle?.closest(".composer-input-region")).not.toBeNull();
    expect(handle?.closest(".composer-box")).toBeNull();
  });

  it("内容超过上限时改为可滚动查看，而不是把多余的行切掉", async () => {
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型目录");
    const textarea = container.querySelector<HTMLTextAreaElement>("textarea")!;

    // jsdom 不排版：直接给一个"内容很高"的宿主输入，等价于真机里输入了很多行。
    Object.defineProperty(textarea, "scrollHeight", { configurable: true, get: () => 900 });
    await type(textarea, "需要填很长的一段说明，长到超过输入框的自动高度上限为止……");

    // 正向前提：确实超过了上限，否则下面的断言就没意义。
    expect(textarea.scrollHeight).toBeGreaterThan(COMPOSER_AUTO_GROW_MAX_HEIGHT);
    expect(textarea.style.height).toBe(`${COMPOSER_AUTO_GROW_MAX_HEIGHT}px`);
    // 核心：溢出改为可滚动，而不是 `hidden` 把文字切掉。
    expect(textarea.style.overflowY).toBe("auto");
  });

  it("CSS 契约：不再依赖 WPS WebView 里拖不动的原生 resize 把手", () => {
    const block = styles.match(/\.wps-chat-shell \.composer-box textarea\s*\{[^}]*\}/)?.[0] ?? "";
    expect(block).not.toBe("");
    expect(block).not.toContain("resize: vertical");
    // 下限与 hook 的 `minHeight: 56` 一致，拖到底和实际渲染高度才对得上。
    expect(block).toContain("min-height: 56px");
  });

  it("恢复图片入口，并使用与 Excel 一致的三列工具条", async () => {
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型目录");

    // 正向锚点：工具条确实存在，否则下面的负向断言是空过。
    const toolbar = container.querySelector(".composer-toolbar");
    expect(toolbar).not.toBeNull();
    expect(toolbar?.querySelector(".composer-shortcut-hint")).not.toBeNull();
    // A legacy placeholder stays absent; the actual picker must be usable.
    expect(container.querySelector(".composer-image-button")).toBeNull();
    expect(container.querySelector<HTMLButtonElement>(".attach-image-button")?.disabled).toBe(false);
    expect(container.querySelector('input[accept="image/png,image/jpeg,image/webp"]')).not.toBeNull();

    // 三列分别是图片入口、可收缩提示、发送。
    const grid = styles.match(/\.wps-chat-shell \.composer-toolbar\s*\{[^}]*\}/)?.[0] ?? "";
    expect(grid).toContain("auto minmax(0, 1fr) auto");
  });

  it("「开发验收模式」只在开发构建里出现，而「写入为何没开放」任何构建都要说", () => {
    const source = readFileSync("src/WpsAgentPanel.tsx", "utf8");
    // 正向锚点：那条诊断文案确实还在，否则下面的负向断言是空过。
    expect(source).toContain("开发验收模式：仅在独立测试工作簿中确认写入与撤销。");
    // 负向：它必须被 `developmentWrite && import.meta.env.DEV` 守着 ——
    // 否则正式构建里普通用户也会看到"你不该用它"的诊断话术。
    const index = source.indexOf('"开发验收模式：');
    const guard = source.slice(Math.max(0, index - 90), index);
    expect(guard).toContain("import.meta.env.DEV");
    // 而"写入为何没开放"是用户该知道的真实状态，不能被同一个开关连坐关掉：
    // 它在三元表达式的另一分支，紧跟在诊断文案之后。
    const tail = source.slice(index, index + 260);
    expect(tail).toContain("当前仍可查询");
  });
});

/**
 * 任务恢复记录里"原表写入状态"的生命周期。
 *
 * 这一段最容易说假话：任务在写入途中被打断时，记录里如果没有留下任何痕迹，
 * 卡片就会对用户说"原表未写入" —— 而原表可能已经被改过了。所以判据是：
 *  - **写前**先记 `unknown`（宁可说未知，也不假装没写）；
 *  - 写成功 → `confirmed_result` + 终态（已完成的任务不再兜售"继续"）；
 *  - 部分写入/回读失败 → `needs_review`，下次只能核对；
 *  - 用户显式放弃 → 终态，且如实写"未写入"（确实一个格子都没碰）。
 */
/**
 * 从落盘的对话历史里取出所有任务恢复记录。
 *
 * 读的是 **localStorage**（真正的持久面），不是组件 state —— 这样断言的是
 * "重开之后还能看到什么"，而不是"内存里现在是什么"。
 */
function storedRecords(): Array<Record<string, unknown>> {
  const saved = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null") as {
    conversations?: Array<{ messages?: Array<{ taskResume?: Record<string, unknown> }> }>;
  } | null;
  return (saved?.conversations ?? [])
    .flatMap((conversation) => conversation.messages ?? [])
    .map((message) => message.taskResume)
    .filter(Boolean) as Array<Record<string, unknown>>;
}

describe("任务恢复记录：写前/写后状态与终态", () => {

  it("预览阶段就有记录：写前状态如实是「未写入」", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); const records=storedRecords(); expect(records.length).toBeGreaterThan(0); const record=records[records.length-1]; expect(record.writeState).toBe('confirmed_result'); expect(record.originalPrompt).toBe('把 A1 改成 90'); expect(record.turnId).toBe('turn-preview-1'); expect(record.resumability).toBe('unavailable');
});

  it("写入成功后：记录为「已确认写入结果」且是终态，不再出现可继续入口", async () => {
    writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card") !== null,
      "写入成功应给出结果卡"
    );

    const record = storedRecords().find((item) => item.writeState === "confirmed_result");
    expect(record).toBeDefined();
    expect(record?.resumability).toBe("unavailable");
    // 已完成的任务不能重新被摆成"可以继续"。
    expect(container.querySelector(".clarification-recovery-actions")).toBeNull();
    expect(findButton(container, "继续任务")).toBeNull();
  });

  it("部分写入失败后：记录为「需要核对原表」，不得谎报未写入", async () => {
    // 写入不落盘 → 回读与计划不符 → 执行器判定为可恢复的部分写入。
    const host = writeRangeHost({ staleWrite: true });
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => container.querySelector(".message-row.system") !== null,
      "失败应给出系统反馈"
    );
    expect(host.writes()).toBeGreaterThan(0);

    const record = storedRecords().find((item) => item.phase === "executing");
    expect(record).toBeDefined();
    expect(record?.writeState).toBe("needs_review");
    expect(record?.blocker).toBe("execution_unknown");
    expect(record?.resumability).toBe("reconcile");
    // 关键负向：绝不能是"未写入"——那时原表可能已经被改过。
    expect(record?.writeState).not.toBe("not_written");
  });

  it("用户显式放弃变更：终态，且如实记录「未写入」", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); await click(findButton(container,'放弃变更')); expect(host.cellValue()).toBe('旧值'); expect(textOf(container)).toContain('已放弃并恢复'); expect(findButton(container,'继续任务')).toBeNull(); expect(findButton(container,'确认变更')).toBeNull();
});

  it("第一次宿主写入发生的那一刻，落盘记录已经是「未知」", async () => {
    let atFirstWrite: Array<Record<string, unknown>> = [];
    writeRangeHost({
      onWrite: (count) => {
        // 在写入**发生的那一刻**取证：读的是 localStorage，不是组件 state。
        if (count === 1) atFirstWrite = storedRecords();
      }
    });
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card") !== null,
      "写入应完成"
    );

    expect(atFirstWrite.length).toBeGreaterThan(0);
    const record = atFirstWrite[atFirstWrite.length - 1];
    // 顺序保证：登记必须早于第一次写入，否则写入途中断线就没人知道原表动过。
    expect(record.writeState).toBe("unknown");
    expect(record.phase).toBe("executing");
  });

  it("执行前恢复标记写不进存储时：取消写入，原表写入次数为 0", async () => {
    const host = writeRangeHost();
    const original = Storage.prototype.setItem;
    const spy = vi
      .spyOn(Storage.prototype, "setItem")
      .mockImplementation(function (this: Storage, key: string, value: string) {
        if (key === "excel-bro.wps.chat.v1") throw new Error("QuotaExceededError");
        return original.call(this, key, value);
      });
    try {
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await click(findButton(container, "生成右侧副本并排审阅"));
      await acknowledgeCopy(container);
      await click(findButton(container, "确认写回原表"));
      await until(
        () => textOf(container).includes("已取消本次写入"),
        "存不下标记时应明确取消写入"
      );

      // 核心：没有持久证据就**一个单元格都不写**。
      expect(host.writes()).toBe(0);
      expect(host.cellValue()).toBe("旧值");
    } finally {
      spy.mockRestore();
    }
  });

  it("写入之后恢复点读取失败：不得显示「原表未写入」", async () => {
    const host = writeRangeHost({ staleWrite: true, breakStorageReadsAfterWrite: true });
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => container.querySelector(".message-row.system") !== null,
      "失败应给出系统反馈"
    );
    expect(host.writes()).toBeGreaterThan(0);

    const record = storedRecords().find((item) => item.phase === "executing");
    expect(record).toBeDefined();
    /*
     * 读取失败 ≠ 没写过。这里必须保持"未知"——把它折成 `not_written`，
     * 就是对着可能已被改过的原表说"原表未写入"。
     */
    expect(record?.writeState).toBe("unknown");
    expect(record?.writeState).not.toBe("not_written");
    expect(textOf(container)).not.toContain("原表未写入");
    expect(textOf(container)).toContain("无法确认");
  });
});

describe("WPS 写入耐久恢复（CAP-03）", () => {
  it("服务端登记失败：取消写入，原表写入次数为 0", async () => {
    const host = writeRangeHost();
    registerWpsExecutionMock.mockRejectedValueOnce(new Error("服务端不可用"));
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => textOf(container).includes("无法持久登记本次写入现场"),
      "登记失败应中止写入"
    );

    // 核心：拿不到持久证据就一个单元格都不写。
    expect(host.writes()).toBe(0);
    expect(host.cellValue()).toBe("旧值");
  });

  it("服务端未授予写入权（无 ownerToken）：取消写入", async () => {
    const host = writeRangeHost();
    registerWpsExecutionMock.mockResolvedValueOnce({
      operationId: "wps-execution-test",
      status: "running",
      revision: 1,
      workbookIdentity: "wps-path:c:\\data\\测试.xlsx",
      taskId: "task-test",
      descriptor: { actions: [] },
      authorizedSheetNames: ["Sheet1"],
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
      ownerToken: null
    });
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => textOf(container).includes("没有授予本次 WPS 写入权"),
      "无写入权应中止"
    );

    expect(host.writes()).toBe(0);
  });

  it("写成功并提交动作级检查点与完成回执", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card") !== null,
      "写入应完成"
    );

    expect(host.writes()).toBeGreaterThan(0);
    expect(startWpsActionMock).toHaveBeenCalled();
    expect(completeWpsActionMock).toHaveBeenCalled();
    expect(finishWpsExecutionMock).toHaveBeenCalled();
  });

  it("本地发现索引只存白名单字段，不含计划、授权或矩阵", async () => {
    const host = writeRangeHost();
    registerWpsExecutionMock.mockResolvedValueOnce({
      operationId: "wps-execution-test",
      status: "running",
      revision: 1,
      workbookIdentity: "wps-path:c:\\data\\测试.xlsx",
      taskId: "task-test",
      descriptor: { actions: [] },
      authorizedSheetNames: ["Sheet1"],
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
    });
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card") !== null,
      "写入应完成"
    );

    const raw = localStorage.getItem("excel-bro.pending-wps-execution.v1");
    // 写成功后索引已被清除（完成回执释放）。
    expect(raw).toBeNull();
    void host;
  });

  /*
   * 重开面板后的恢复卡：**点真实按钮、看真实 API 调用与宿主写入次数**。
   * 不只看徽标文案 —— 上一轮的缺陷正是"显示可以继续却没有按钮"。
   */
  const PENDING_KEY = "excel-bro.pending-wps-execution.v1";

  function recoveryState(overrides: Record<string, unknown> = {}) {
    const total = Number(overrides.totalActions ?? 1);
    const fingerprint = matrixFingerprintV2([["旧值"]], [["旧值"]]);
    return {
      operationId: "wps-execution-test",
      status: "running" as const,
      revision: 7,
      workbookIdentity: "wps-path:c:\\data\\测试.xlsx",
      taskId: "task-test",
      descriptor: { actions: Array.from({ length: total }, (_, actionIndex) => ({
        actionIndex, type: "writeValues", sheet: "Sheet1", range: `A${actionIndex + 1}`,
        rowCount: 1, columnCount: 1, cellCount: 1,
        previewFingerprint: fingerprint, expectedFingerprint: fingerprint
      })) },
      authorizedSheetNames: ["Sheet1"],
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
      ownerToken: null,
      ...overrides
    };
  }

  /** 带 `taskResume` 的历史消息：重新预览要靠它找回原始需求。 */
  function historyWithTaskResume(): unknown[] {
    return [
      {
        id: "msg-resume",
        role: "assistant",
        text: "写入未完成",
        createdAt: new Date().toISOString(),
        taskResume: {
          schemaVersion: 1,
          taskId: "task-test",
          phase: "executing",
          blocker: "execution_unknown",
          resumability: "reconcile",
          writeState: "unknown",
          originalPrompt: "把 A1 改成 90",
          confirmations: [],
          authorizationRevision: 1,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        }
      }
    ];
  }

  async function mountRecoveryCard(options: {
    state?: Record<string, unknown>;
    fullName?: string;
    withHistory?: boolean;
    /** 在挂载**之前**塞一条现场冲突的恢复点，让只读核对拿到 conflict。 */
    seedUndo?: "conflict";
  } = {}) {
    const host = writeRangeHost();
    if (options.seedUndo === "conflict") {
      seedConflictingUndoRecord("wps-execution-test");
      host.setCell("人工修改");
    }
    if (options.withHistory) {
      localStorage.setItem(
        "excel-bro.wps.chat.v1",
        JSON.stringify({
          activeConversationId: "conv-1",
          conversations: [
            {
              id: "conv-1",
              title: "历史",
              messages: historyWithTaskResume(),
              createdAt: new Date().toISOString(),
              updatedAt: new Date().toISOString()
            }
          ]
        })
      );
    }
    if (options.fullName !== undefined) {
      const application = getWpsApplicationMock() as unknown as {
        ActiveWorkbook: { FullName: string };
      };
      application.ActiveWorkbook.FullName = options.fullName;
    }
    localStorage.setItem(
      PENDING_KEY,
      JSON.stringify({
        schemaVersion: 1,
        operationId: "wps-execution-test",
        conversationId: "conv-1",
        activityMessageId: "msg-1",
        workbookIdentity: "wps-path:c:\\data\\测试.xlsx",
        taskId: "task-test",
        createdAt: new Date().toISOString()
      })
    );
    getWpsExecutionStatusMock.mockResolvedValue(recoveryState(options.state) as never);
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型目录");
    await until(
      () => container.querySelector(".wps-write-recovery") !== null,
      "重载后应出现恢复核对卡"
    );
    await until(
      () => !textOf(container).includes("正在查询服务端写入现场"),
      "应查完服务端状态"
    );
    return { container, host };
  }

  function recoveryButtons(container: HTMLElement): string[] {
    const card = container.querySelector(".wps-write-recovery");
    if (!card) return [];
    // Only mutation/dismiss actions; the separate refresh button is read-only.
    return Array.from(card.querySelectorAll(".wps-recovery-action")).map((item) =>
      (item.textContent ?? "").trim()
    );
  }

  it("零写入证据的锚点：纯测试替身计数包含 Value2 和 Formula 两种 setter", () => {
    const host = writeRangeHost();
    // getWpsApplicationMock was explicitly replaced by writeRangeHost; this is
    // only a JS fixture, never the real WPS application or a runtime write path.
    const fixture = getWpsApplicationMock();
    const range = fixture.ActiveWorkbook!.Worksheets.Item(1).Range!("A1");
    expect(host.writes()).toBe(0);
    range.Value2 = "TEST_VALUE";
    expect(host.writes()).toBe(1);
    range.Formula = "=1+1";
    expect(host.writes()).toBe(2);
  });

  it("全部动作已完成、只丢最终回执：点补记完成 → 领取执行权 + 提交完成回执，宿主写入 0", async () => {
    const { container, host } = await mountRecoveryCard({
      state: { nextActionIndex: 2, totalActions: 2 }
    });
    claimWpsExecutionResumeMock.mockResolvedValueOnce(recoveryState({ nextActionIndex: 2,
      totalActions: 2, ownerToken: "claimed-owner", revision: 8 }) as never);

    expect(recoveryButtons(container)).toEqual([
      "补记完成（不重写原表）",
      "清除这条恢复提示"
    ]);

    await click(findButton(container, "补记完成（不重写原表）"));
    await until(
      () => textOf(container).includes("已补记这次写入的完成回执"),
      "补记应成功"
    );

    expect(claimWpsExecutionResumeMock).toHaveBeenCalledWith("wps-execution-test", 7);
    expect(finishWpsExecutionMock).toHaveBeenCalledTimes(1);
    expect(finishWpsExecutionMock).toHaveBeenCalledWith("wps-execution-test", {
      expectedRevision: 8, ownerToken: "claimed-owner", result: { backfilled: true, actionCount: 2 }
    });
    // 补记只改账本：一次宿主写入都没有。
    expect(host.writes()).toBe(0);
    expect(localStorage.getItem(PENDING_KEY)).toBeNull();
    expect(container.querySelector(".wps-write-recovery")).toBeNull();
  });

  it("还有动作没写：只给「重新读取并生成预览」，不给任何继续/写入按钮", async () => {
    const { container } = await mountRecoveryCard({
      state: { nextActionIndex: 1, totalActions: 3 }
    });

    const buttons = recoveryButtons(container);
    expect(buttons).toEqual(["重新读取并生成预览", "清除这条恢复提示"]);
    // 关键：不许出现"可以继续"这类冒充断点续跑的入口。
    expect(buttons.some((name) => name.includes("继续"))).toBe(false);
    expect(textOf(container)).toContain("无法精确续跑剩余动作");
  });

  it("点「重新读取并生成预览」→ 用原需求重新发起一轮，不是续跑", async () => {
    const { container, host } = await mountRecoveryCard({
      state: { nextActionIndex: 1, totalActions: 3 },
      withHistory: true
    });
    runTurnMock.mockResolvedValue({ kind: "plan", plan: singleWritePlan, turnId: "turn-replan" });
    expect(findButton(container, "重新读取并生成预览")?.disabled).toBe(true);
    await click(container.querySelector('.wps-write-recovery input[type="checkbox"]') as HTMLInputElement);
    await click(findButton(container, "重新读取并生成预览"));
    await until(
      () => textOf(container).includes("重新读取并生成预览 —— 那是新的一轮"),
      "应说明这是新的一轮"
    );

    // 重新走一轮规划，且**没有**对原表做任何写入。
    expect(runTurnMock).toHaveBeenCalled();
    expect(host.writes()).toBe(0);
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
    expect(claimWpsExecutionResumeMock).toHaveBeenCalled();
    expect(suspendWpsExecutionMock).toHaveBeenCalled();
    expect(localStorage.getItem(PENDING_KEY)).not.toBeNull();
  });

  it.each([
    ["canResume 变 false", { canResume: false }],
    ["另一个执行者领取", { ownerActive: true }],
    ["剩余动作重新出现", { nextActionIndex: 0 }]
  ])("点补记前重新查询：%s 不调用领取或完成", async (_label, patch) => {
    const { container, host } = await mountRecoveryCard({ state: { nextActionIndex: 1 } });
    getWpsExecutionStatusMock.mockResolvedValue(recoveryState({ nextActionIndex: 1, ...patch }) as never);
    await click(findButton(container, "补记完成（不重写原表）"));
    await until(() => getWpsExecutionStatusMock.mock.calls.length >= 2, "必须重新查询");
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
    expect(host.writes()).toBe(0);
  });

  it("卡片显示后人工改值：点补记重新读现场，拦住回执", async () => {
    const { container, host } = await mountRecoveryCard({ state: { nextActionIndex: 1 } });
    host.setCell("人工修改后");
    await click(findButton(container, "补记完成（不重写原表）"));
    await until(() => textOf(container).includes("疑似人工修改"), "应识别现场冲突");
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
    expect(host.writes()).toBe(0);
  });

  it("领取期间人工改值：第二次现场核对仍拦住完成回执", async () => {
    const { container, host } = await mountRecoveryCard({ state: { nextActionIndex: 1 } });
    claimWpsExecutionResumeMock.mockImplementationOnce(async () => {
      host.setCell("领取期间修改");
      return recoveryState({ nextActionIndex: 1, ownerToken: "new-owner", revision: 8 }) as never;
    });
    await click(findButton(container, "补记完成（不重写原表）"));
    await until(() => textOf(container).includes("领取执行权后现场发生变化"), "第二次核对应拦截");
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
    expect(suspendWpsExecutionMock).toHaveBeenCalled();
    expect(host.writes()).toBe(0);
  });

  it("最后动作回执丢失：点击核对提交指纹后只补记完成，零宿主写入", async () => {
    const { container, host } = await mountRecoveryCard({ state: {
      status: "needs_reconciliation", requiresReconciliation: true, inFlightActionIndex: 0
    }});
    reconcileWpsActionReceiptMock.mockResolvedValueOnce(recoveryState({ nextActionIndex: 1,
      ownerToken: "new-owner", revision: 8 }) as never);
    await click(findButton(container, "核对在途动作"));
    await until(() => textOf(container).includes("已核对最后一个动作并补记完成"), "应只补记");
    expect(reconcileWpsActionReceiptMock).toHaveBeenCalledWith("wps-execution-test", expect.objectContaining({
      outcome: "applied", checkpoint: expect.objectContaining({ type: "writeValues",
        sheetName: "Sheet1", range: "A1", afterFingerprint: expect.any(String) })
    }));
    expect(finishWpsExecutionMock).toHaveBeenCalledTimes(1);
    expect(runTurnMock).not.toHaveBeenCalled();
    expect(host.writes()).toBe(0);
  });

  it("canResume=false：只给清除提示，不给补记或核对按钮", async () => {
    const { container } = await mountRecoveryCard({
      state: { canResume: false, nextActionIndex: 2, totalActions: 2 }
    });
    expect(recoveryButtons(container)).toEqual(["清除这条恢复提示"]);
  });

  it("执行权仍活跃时仅只读刷新，释放后才出现补记入口", async () => {
    const { container, host } = await mountRecoveryCard({ state: { nextActionIndex: 1, ownerActive: true } });
    expect(recoveryButtons(container)).toEqual(["清除这条恢复提示"]);
    getWpsExecutionStatusMock.mockResolvedValue(recoveryState({ nextActionIndex: 1 }) as never);
    await click(findButton(container, "重新核对（只读）"));
    await until(() => recoveryButtons(container).includes("补记完成（不重写原表）"), "刷新后才开放补记");
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
    expect(host.writes()).toBe(0);
  });

  it("身份不一致（另一个工作簿）：只给清除提示，不给写入按钮", async () => {
    const { container } = await mountRecoveryCard({
      state: { nextActionIndex: 2, totalActions: 2 },
      fullName: "C:\\other\\测试.xlsx"
    });
    expect(recoveryButtons(container)).toEqual(["清除这条恢复提示"]);
    expect(textOf(container)).toContain("不是同一个文件");
  });

  /**
   * 往 PluginStorage 塞一条**现场与写前、写后都不一致**的恢复点。
   * 指纹故意用不可能匹配上的值，让只读核对必然得出 `conflict`。
   */
  function seedConflictingUndoRecord(executionId: string) {
    const application = getWpsApplicationMock() as unknown as {
      PluginStorage: { setItem: (key: string, value: unknown) => void };
    };
    application.PluginStorage.setItem(
      "excel-bro.wps.undo.v1",
      JSON.stringify([
        {
          version: 2,
          executionId,
          planId: "plan-old",
          planTitle: "旧计划",
          createdAt: new Date().toISOString(),
          workbookIdentity: "c:\\data\\测试.xlsx",
          authorizedSheetNames: ["Sheet1"],
          state: "applying",
          actions: [
            {
              actionIndex: 0,
              type: "writeValues",
              sheetIndex: 1,
              sheetName: "Sheet1",
              range: "A1",
              rowCount: 1,
              columnCount: 1,
              cellCount: 1,
              values: [["90"]],
              preview: [["90"]],
              previewFingerprint: "1:00000000",
              rect: { startRow: 1, startColumn: 1, endRow: 1, endColumn: 1 },
              beforeValues: [["旧值"]],
              beforeFormulas: [["旧值"]],
              beforeFingerprint: "1:aaaaaaaa",
              afterFingerprint: "1:bbbbbbbb",
              appliedRows: 1
            }
          ],
          error: null
        }
      ])
    );
  }

  it("现场冲突（疑似人工修改）：不给补记/核对按钮，也不提交任何回执", async () => {
    const { container } = await mountRecoveryCard({
      state: {
        status: "needs_reconciliation",
        requiresReconciliation: true,
        inFlightActionIndex: 0
      },
      seedUndo: "conflict"
    });
    expect(textOf(container)).toContain("疑似人工修改或部分写入");
    expect(recoveryButtons(container)).toEqual(["清除这条恢复提示"]);

    await click(findButton(container, "清除这条恢复提示"));
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
  });

  it("全部动作已完成但本机现场是写前（两者矛盾）：不给补记按钮", async () => {
    const { container } = await mountRecoveryCard({
      state: { nextActionIndex: 2, totalActions: 2 },
      seedUndo: "conflict"
    });
    // 服务端说写完了、本机却与写后不一致 → 矛盾，挂起等人工核对。
    expect(recoveryButtons(container)).toEqual(["清除这条恢复提示"]);
  });

  it("旧描述缺少写后指纹：不给核对按钮，也不提交回执", async () => {
    const { container } = await mountRecoveryCard({
      state: {
        status: "needs_reconciliation",
        requiresReconciliation: true,
        inFlightActionIndex: 0,
        descriptor: { actions: [{ actionIndex: 0, type: "writeValues", sheet: "Sheet1",
          range: "A1", rowCount: 1, columnCount: 1, cellCount: 1, previewFingerprint: "legacy" }] }
      }
    });
    expect(recoveryButtons(container)).toEqual(["清除这条恢复提示"]);
    await click(findButton(container, "清除这条恢复提示"));
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();
    expect(container.querySelector(".wps-write-recovery")).toBeNull();
  });
});

/*
 * ===========================================================================
 * 完整链路（**真实面板**）：工具桥查询 → 结果计划预览 → 确认 → 写前复核 →
 * 执行 → 回读。
 *
 * 只换两处边界：
 *  1. **网络**：`fetch` 是替身，但断言真实发出的**请求体**（poll / receipt）；
 *  2. **宿主**：JSAPI 替身提供 `UsedRange` / `Range().Value2`，模拟真宿主
 *     "只接受两个参数"的签名限制与"公式返回值"的常量语义。
 *
 * 桥状态机、授权复核、`readAuthorizedWpsSheets`、`executeQueryTableData`
 * 与 `wpsPlanExecutor` 都是**生产代码**；读与写落在**同一份矩阵**上，
 * 所以"读到的"就是"写回的"那张表。
 * ===========================================================================
 */
const ROUND_TRIP_OP = "op_" + "9".repeat(32);
const ROUND_TRIP_FP = "b".repeat(64);
const ROUND_TRIP_RESULT_REF = "res_" + "4".repeat(32);

const ROUND_TRIP_QUERY = {
  mode: "aggregate" as const,
  scope: "selected" as const,
  groupBy: ["影片名称"],
  metrics: [{ operation: "sum" as const, field: "计数项:影厅", outputName: "影厅数" }],
  sortBy: "影厅数",
  sortDirection: "desc" as const
};

function columnIndex(letters: string): number {
  return letters
    .split("")
    .reduce((total, letter) => total * 26 + (letter.charCodeAt(0) - 64), 0);
}

function parseA1(address: string) {
  const match = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(address.trim());
  if (!match) throw new Error(`替身不支持的区域地址：${address}`);
  return {
    startRow: Number(match[2]),
    startColumn: columnIndex(match[1]),
    endRow: Number(match[4] ?? match[2]),
    endColumn: columnIndex(match[3] ?? match[1])
  };
}

/**
 * 既能被桥**真读真算**、也能被执行器**真写真回读**的宿主替身。
 *
 * 两张授权表内容不同（同一批列），所以"跨表合并后再分组求和"只能由生产计算
 * 得出：功夫女足 24、恐怖游轮 3。写入落在同一份矩阵上，回读才可能一致。
 */
function bridgeRoundTripHost(
  options: {
    renameSheet2?: () => boolean;
    /**
     * 覆盖默认矩阵（键 = 工作表名）。合成案例用它注入**同一批脏数据**
     * （带空格键、前导零、空键、重复键、未匹配记录）来跑同一条链路。
     */
    matrices?: Record<string, (string | number | null)[][]>;
  } = {}
) {
  const matrices: Record<string, (string | number | null)[][]> = options.matrices ?? {
    Sheet1: [
      ["影院名称", "影片名称", "计数项:影厅"],
      ["影院A", "功夫女足", 11],
      ["影院B", "恐怖游轮", 2]
    ],
    Sheet2: [
      ["影院名称", "影片名称", "计数项:影厅"],
      ["影院C", "功夫女足", 13],
      ["影院D", "恐怖游轮", 1]
    ]
  };
  let writes = 0;
  let readSheetCalls = 0;

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
      // 单格返回标量、多格返回矩阵 —— 与真宿主一致（执行器的两条路径都依赖它）。
      Value2: {
        get: () => {
          readSheetCalls += 1;
          const current = slice();
          return rows === 1 && columns === 1 ? current[0][0] : current;
        },
        set: store,
        enumerable: true
      },
      // 真宿主对常量单元格的 `Formula` 返回常量本身，只有 `=` 开头才是公式。
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

  /*
   * 工作表对象**必须每个名字只有一个**：真宿主里同一张表始终是同一个包装对象，
   * 而授权核对会比对对象身份。每次 `Item()` 都新建一个对象，会被判成"目标变了"
   * 并弹出范围确认 —— 那不是被测行为，是替身比真宿主更宽松/更不稳定。
   */
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
      Value2: {
        get: () => grid.map((row) => [...row]),
        enumerable: true
      },
      MergeCells: { value: false }
    });
    return {
      Name: name,
      CodeName: `Code_${name}`,
      ProtectContents: false,
      UsedRange: usedRange,
      // 读路径传两个对象、写路径传一个字符串地址；两种都要支持。
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
    const all = Object.keys(matrices);
    if (options.renameSheet2?.()) {
      // 用户把第二张授权表改名了 —— 授权范围**不再覆盖**绑定时的集合。
      return [all[0], "被改名的工作表"];
    }
    return all;
  };

  const storage = new Map<string, unknown>();
  /*
   * `FullName` 必须是**真实的 Windows 盘符路径**（反斜杠）。
   * 写入闸 `workbookIdentity` 要求 `/^[A-Za-z]:\\[^\r\n]+\.xlsx$/i` —— 这是产品
   * 规则本身：写入只对已保存到本地磁盘的 .xlsx 开放。替身用正斜杠路径会被这道闸
   * 挡下（症状是"计划预览失败、提示先保存文件"），那不是被测行为，而是替身不真。
   */
  const workbook = {
    Name: "影院.xlsx",
    FullName: "C:\\Data\\影院.xlsx",
    Path: "C:\\Data",
    ReadOnly: false,
    Worksheets: {
      get Count() {
        return names().length;
      },
      Item: (key: number | string) => {
        const list = names();
        if (typeof key === "string" && !/^\d+$/.test(key)) {
          const found = list.find((name) => name === key);
          return found ? sheet(found) : undefined;
        }
        const name = list[Number(key) - 1];
        return name ? sheet(name) : undefined;
      }
    },
    ActiveSheet: undefined as unknown
  };
  const activeSheet = sheet("Sheet1");
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
    reads: () => readSheetCalls,
    /** 模拟"等待期间用户改了源表里的一个数"：身份不变，只有内容变。 */
    mutate: (sheetName: string, row: number, column: number, value: string | number) => {
      matrices[sheetName][row - 1][column - 1] = value;
    },
    cell: (sheetName: string, row: number, column: number) =>
      matrices[sheetName]?.[row - 1]?.[column - 1] ?? null
  };
}

function stubBridgeEndpoints(
  options: {
    onPoll?: () => void;
    /** `/api/sdk/tool-bridge/result` 的响应体；不给则给一份最小可用的真值。 */
    result?: unknown;
    /** result 端点的 HTTP 状态（>=400 用来模拟"已过保留期/授权不符"）。 */
    resultStatus?: number;
    /** 轮询响应体；不给则给一条真实的查询调用。 */
    poll?: unknown;
  } = {}
) {
  const captured: Array<{ url: string; body: Record<string, unknown> }> = [];
  vi.stubGlobal(
    "fetch",
    vi.fn(async (url: string, init?: RequestInit) => {
      const target = String(url);
      captured.push({ url: target, body: JSON.parse(String(init?.body ?? "{}")) });
      if (target.includes("/receipt")) {
        return new Response(JSON.stringify({ status: "accepted", code: null }), {
          status: 200,
          headers: { "Content-Type": "application/json" }
        });
      }
      if (target.includes("/result")) {
        const status = options.resultStatus ?? 200;
        if (status >= 400) {
          return new Response(JSON.stringify({ detail: { code: "RESULT_EXPIRED" } }), {
            status,
            headers: { "Content-Type": "application/json" }
          });
        }
        return new Response(
          JSON.stringify(
            options.result ?? {
              resultRef: ROUND_TRIP_RESULT_REF,
              title: "影厅数",
              headers: ["影片名称", "影厅数"],
              rows: [
                ["功夫女足", 24],
                ["恐怖游轮", 3]
              ],
              sourceSheets: ["Sheet1", "Sheet2"],
              warnings: [],
              complete: true,
              calculation: "求和(计数项:影厅)",
              scannedRows: 4
            }
          ),
          { status: 200, headers: { "Content-Type": "application/json" } }
        );
      }
      // 轮询一发出就表示"绑定已完成、已经进入等待期" —— 此刻改现场最贴切。
      options.onPoll?.();
      const payload = options.poll ?? {
        calls: [
          {
            operationId: ROUND_TRIP_OP,
            tool: "query_table",
            arguments: ROUND_TRIP_QUERY,
            requestFingerprint: ROUND_TRIP_FP,
            expiresAt: 0
          }
        ],
        generation: 1
      };
      return new Response(JSON.stringify(payload), {
        status: 200,
        headers: { "Content-Type": "application/json" }
      });
    })
  );
  return { captured };
}

/** 面板本轮：先发桥指针并挂住，等我们放行后才把这一轮的结果交给面板。 */
function armBridgeTurn(result: WpsTurnResult, pythonResult?: typeof ROUND_TRIP_RESULT_POINTER) {
  let resolveTurn: (() => void) | null = null;
  let released = false;
  runTurnMock.mockImplementation(async (input: WpsReadonlyTurnInput) => {
    /*
     * 事件形状照抄生产实现（`wpsReadonlyAgent` 里规划流 onProgress 那一支）：
     * `stage:"answering"` + `turnId` + `bridgeTool`。少给 `stage`/`message`
     * 会让 `readonlyActivityUpdate` 退化成兜底文案 —— 那不是被测行为。
     */
    await input.onProgress?.({
      stage: "answering",
      stepId: "bridge.query_table",
      title: "模型请求查询授权数据",
      turnId: "turn-bridge-panel",
      message: "模型正在等待本机读取并计算授权数据。",
      bridgeTool: { operationId: ROUND_TRIP_OP, tool: "query_table" }
    });
    /*
     * 回执是由 `onProgress` 里 `void drain()` **并发**跑出来的：测试可能在
     * onProgress 还没返回时就已看到回执并放行 —— 那时这个 promise 还没建。
     * 所以"已放行"必须单独记账，不能假设放行一定发生在 promise 建好之后。
     */
    if (!released) {
      await new Promise<void>((resolve) => {
        resolveTurn = resolve;
      });
    }
    if (pythonResult) await input.onProgress?.({ stage: "answering", turnId: "turn-bridge-panel",
      stepId: "python.completed", message: "隔离分析完成", pythonResult });
    return result;
  });
  return () => {
    released = true;
    resolveTurn?.();
  };
}

describe("WPS 面板 · 工具桥查询到写入回读的完整链路", () => {
  it("查询→结果计划预览→确认→写前复核→执行→回读：真实请求体与真实读写", async () => {
    const host = bridgeRoundTripHost();
    const bridge = stubBridgeEndpoints();
    const releaseTurn = armBridgeTurn({
      kind: "plan",
      plan: singleWritePlan,
      turnId: "turn-bridge-panel"
    });
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    validatePlanMock.mockImplementation(async (submitted) => submitted);

    const { container } = await mount(
      <WpsAgentPanel
        selectedSheetIndexes={[1, 2]}
        scopeDock={<div className="eb-scope-dock">当前表 Sheet1、Sheet2</div>}
      />
    );
    await until(
      () =>
        container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型") ?? false,
      "应加载模型目录"
    );
    const textarea = container.querySelector("textarea");
    await type(textarea, "统计各影片的影厅数，并把结论写到 A1");
    await submit(textarea);

    // 1) 桥真的把回执发回去了，而且回执里的值是生产计算从真实宿主矩阵算出来的。
    await until(
      () => bridge.captured.some((item) => item.url.includes("/receipt")),
      "应把本地查询结果回执给服务端"
    );
    const receipt = bridge.captured.find((item) => item.url.includes("/receipt"))?.body;
    expect(receipt).toMatchObject({
      hostId: "wps-et",
      operationId: ROUND_TRIP_OP,
      requestFingerprint: ROUND_TRIP_FP
    });
    const payload = (receipt?.receipt as { result?: { headers: unknown[]; rows: unknown[][] } })
      .result;
    expect(payload?.headers).toEqual(["影片名称", "影厅数"]);
    expect(payload?.rows).toEqual([
      ["功夫女足", 24],
      ["恐怖游轮", 3]
    ]);
    // 轮询请求体带完整的归属（含工作簿身份），不是只报个名字。
    const poll = bridge.captured.find((item) => item.url.includes("/poll"))?.body;
    expect(poll).toMatchObject({
      taskId: "turn-bridge-panel",
      hostId: "wps-et",
      workbookName: "影院.xlsx",
      // 已保存工作簿 → 用磁盘路径做命名空间（不是只报个名字）。
      workbookIdentity: "wps-path:c:\\data\\影院.xlsx"
    });
    // 写之前一个单元格都没被改。
    expect(host.writes()).toBe(0);

    // 2) 放行：结果计划回到面板 → 预览 → 确认 → 执行 → 回读。
    releaseTurn();
    await until(() => findButton(container, "确认变更") !== null || textOf(container).includes("本次执行需要核对"), "应展示执行结果或明确失败");
    expect(host.writes()).toBe(0);

    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card.verified") !== null,
      "写入成功应给出回读核对结果"
    );
    expect(host.writes()).toBeGreaterThan(0);
    // 回读核对读的是**同一份矩阵**：A1 现在真的是 90。
    expect(host.cell("Sheet1", 1, 1)).toBe("90");
    expect(container.querySelector(".wps-verification-card")?.textContent).toContain(
      "回读核对"
    );
  });

  it("等待期间授权被收窄（第二张表改名）：拒读，只回 unauthorized_scope，零写入", async () => {
    let narrowed = false;
    const host = bridgeRoundTripHost({ renameSheet2: () => narrowed });
    // 绑定发生在授权完整时；轮询一发出就表示进了等待期 —— 此刻收窄。
    const bridge = stubBridgeEndpoints({ onPoll: () => { narrowed = true; } });
    const releaseTurn = armBridgeTurn({
      kind: "plan",
      plan: singleWritePlan,
      turnId: "turn-bridge-panel"
    });
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    validatePlanMock.mockImplementation(async (submitted) => submitted);

    const { container } = await mount(
      <WpsAgentPanel
        selectedSheetIndexes={[1, 2]}
        scopeDock={<div className="eb-scope-dock">当前表 Sheet1、Sheet2</div>}
      />
    );
    await until(
      () =>
        container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型") ?? false,
      "应加载模型目录"
    );
    const textarea = container.querySelector("textarea");
    await type(textarea, "统计各影片的影厅数");
    await submit(textarea);

    await until(
      () => bridge.captured.some((item) => item.url.includes("/receipt")),
      "复核失败也必须回一张回执"
    );
    const receipt = bridge.captured.find((item) => item.url.includes("/receipt"))?.body;
    // 严格相等：回执里不许夹带任何结果值。
    expect(receipt?.receipt).toEqual({ errorCode: "unauthorized_scope" });
    // 正向锚点：确实发出了轮询，说明复核发生在"已进入等待期"之后。
    expect(bridge.captured.some((item) => item.url.includes("/poll"))).toBe(true);
    expect(host.writes()).toBe(0);

    releaseTurn();
    await settle();
    expect(container.querySelector(".wps-verification-card")).toBeNull();
    expect(host.writes()).toBe(0);
  });

  it("等待期间源值被改（工作簿身份不变）：只回 source_changed，零写入", async () => {
    const host = bridgeRoundTripHost();
    /*
     * 绑定发生在源还是原样的时候；轮询一发出就说明进了等待期 —— 此刻改一个数。
     * 身份、授权、表名统统没变，变的只有**内容** —— 这正是指纹要比对的那一档，
     * 只看身份会放行一个"读到的是另一版源"的查询。
     */
    const bridge = stubBridgeEndpoints({
      onPoll: () => host.mutate("Sheet1", 2, 3, 99)
    });
    const releaseTurn = armBridgeTurn({
      kind: "plan",
      plan: singleWritePlan,
      turnId: "turn-bridge-panel"
    });
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    validatePlanMock.mockImplementation(async (submitted) => submitted);

    const { container } = await mount(
      <WpsAgentPanel
        selectedSheetIndexes={[1, 2]}
        scopeDock={<div className="eb-scope-dock">当前表 Sheet1、Sheet2</div>}
      />
    );
    await until(
      () =>
        container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型") ?? false,
      "应加载模型目录"
    );
    const textarea = container.querySelector("textarea");
    await type(textarea, "统计各影片的影厅数，并把结论写到 A1");
    await submit(textarea);

    await until(
      () => bridge.captured.some((item) => item.url.includes("/receipt")),
      "源变了也必须回一张回执，而不是静默读旧值"
    );
    const receipt = bridge.captured.find((item) => item.url.includes("/receipt"))?.body;
    // 严格相等：不许夹带任何结果值。
    expect(receipt?.receipt).toEqual({ errorCode: "source_changed" });
    // 正向锚点：轮询确实发出过（复核发生在进入等待期之后），且没写过任何单元格。
    expect(bridge.captured.some((item) => item.url.includes("/poll"))).toBe(true);
    expect(host.writes()).toBe(0);

    releaseTurn();
    await settle();
    expect(container.querySelector(".wps-verification-card")).toBeNull();
    expect(host.writes()).toBe(0);
  });
});

/**
 * 任务返回、桥已停轮询之后，结果卡仍要能按引用把值取回来。
 *
 * 这条通路走的是**授权归属**校验（服务端不再看活跃会话），所以请求体里必须回带
 * **产生该结果那一轮的**绑定。结果过了保留期或授权不符时，界面要如实说"不可用"，
 * 绝不渲染一张空表冒充成功。
 */
const ROUND_TRIP_RESULT_POINTER = {
  resultRef: ROUND_TRIP_RESULT_REF,
  title: "影厅数",
  headers: ["影片名称", "影厅数"],
  rowCount: 2,
  columnCount: 2,
  sourceSheets: ["Sheet1", "Sheet2"],
  complete: true
};

describe("WPS 面板 · 完成后按引用取回结果 / 结果不可用 / 面板重开", () => {
  async function completeQueryTurn(options: { resultStatus?: number; result?: unknown; progressOnly?: boolean } = {}) {
    const host = bridgeRoundTripHost();
    const bridge = stubBridgeEndpoints(options);
    const releaseTurn = armBridgeTurn({
      kind: "answer",
      provider: "model",
      message: "两部影片的影厅数已取回",
      outcome: "completed",
      turnId: "turn-bridge-panel",
      localResult: options.progressOnly ? undefined : ROUND_TRIP_RESULT_POINTER
    }, options.progressOnly ? ROUND_TRIP_RESULT_POINTER : undefined);
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    const view = await mount(
      <WpsAgentPanel
        selectedSheetIndexes={[1, 2]}
        scopeDock={<div className="eb-scope-dock">当前表 Sheet1、Sheet2</div>}
      />
    );
    await until(
      () =>
        view.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型") ?? false,
      "应加载模型目录"
    );
    const textarea = view.container.querySelector("textarea");
    await type(textarea, "统计各影片的影厅数");
    await submit(textarea);
    await until(
      () => bridge.captured.some((item) => item.url.includes("/receipt")),
      "查询应经桥回执"
    );
    releaseTurn();
    return { ...view, host, bridge };
  }

  it.each([false, true])("桥已停：按引用取回该轮绑定并渲染真实值（仅进度事件=%s）", async (progressOnly) => {
    const { container, bridge } = await completeQueryTurn({ progressOnly });

    await until(
      () => container.querySelector(".local-result-card-table") !== null,
      "结果卡应把真实值取回来渲染"
    );
    const table = container.querySelector(".local-result-card-table");
    expect(table?.textContent).toContain("影片名称");
    expect(table?.textContent).toContain("功夫女足");
    expect(table?.textContent).toContain("24");
    expect(container.querySelector(".local-result-card-calculation")?.textContent).toContain(
      "求和"
    );

    const body = bridge.captured.find((item) => item.url.includes("/result"))?.body;
    expect(body).toMatchObject({
      hostId: "wps-et",
      workbookName: "影院.xlsx",
      workbookIdentity: "wps-path:c:\\data\\影院.xlsx",
      resultRef: ROUND_TRIP_RESULT_REF
    });
    // 任务身份必须是**这一轮**的 turnId，不是随便一个占位。
    expect(body?.taskId).toBe("turn-bridge-panel");
    // 绑定是完整的：源指纹也回带了（服务端据此判断"这一轮读的是哪一版源"）。
    expect(typeof body?.sourceFingerprint).toBe("string");
  });

  it("结果已过保留期（410）：如实说不可用，不渲染空表冒充成功", async () => {
    const { container } = await completeQueryTurn({ resultStatus: 410 });

    await until(
      () => container.querySelector(".local-result-card-unavailable") !== null,
      "取不到值时应如实说明"
    );
    // 负向断言配正向锚点：卡片本身在（确实走到了这一步），只有表没渲染。
    expect(container.querySelector(".local-result-card")).not.toBeNull();
    expect(container.querySelector(".local-result-card-table")).toBeNull();
    expect(
      container.querySelector(".local-result-card-unavailable")?.textContent
    ).toContain("保留期");
  });

  it("面板重开（新的加载项会话）：结果卡靠**该轮**绑定按引用再取回，不依赖活跃会话", async () => {
    const first = await completeQueryTurn();
    await until(
      () => first.container.querySelector(".local-result-card-table") !== null,
      "第一次挂载应取回结果"
    );
    const firstRoundCalls = first.bridge.captured.filter((item) =>
      item.url.includes("/result")
    ).length;
    expect(firstRoundCalls).toBe(1);

    // 重开：当前加载项会话里已没有任何桥状态，只剩对话历史（localStorage）。
    first.unmount();
    const second = await mount(
      <WpsAgentPanel
        selectedSheetIndexes={[1, 2]}
        scopeDock={<div className="eb-scope-dock">当前表 Sheet1、Sheet2</div>}
      />
    );
    await until(
      () => second.container.querySelector(".local-result-card-table") !== null,
      "重开后结果卡应仍能按该轮绑定取回"
    );
    expect(second.container.querySelector(".local-result-card-table")?.textContent).toContain(
      "功夫女足"
    );
    // 正向证明"是**重新**取回来的"：/result 又发了一次，且带着同一轮的绑定。
    const resultCalls = first.bridge.captured.filter((item) => item.url.includes("/result"));
    expect(resultCalls.length).toBe(2);
    expect(resultCalls[1].body).toMatchObject({
      taskId: "turn-bridge-panel",
      workbookIdentity: "wps-path:c:\\data\\影院.xlsx",
      resultRef: ROUND_TRIP_RESULT_REF
    });
  });
});

/**
 * 通用合成案例（WPS 端）：**同一批脏数据 → 清洗/关联 → 汇总 + 异常明细 →
 * 结果卡 → 预览 → 确认写回 → 回读 → 撤销**。
 *
 * 与 Excel 端 `syntheticResultEndToEnd.test.tsx` 用**同一套规则、同一批数据、
 * 同一套期望值** —— 两端语义必须一致（不是各写各的）。这里走的是 WPS 面板的
 * 真实链路：桥把本地 `executeQueryTableData`（两端同源）算出的结果回执给服务端，
 * 服务端给出结果计划，面板预览 → 确认 → 写回 → 回读，最后撤销。
 *
 * 脏数据（用户点名的那几类）：
 * - 键含**前后空格**、**前导零**编号、**空键**、**重复键**、**未匹配记录**。
 */
describe("WPS 面板 · 通用合成案例：脏数据清洗/汇总 → 结果卡 → 确认写回 → 回读 → 撤销", () => {
  /** 会员表：带空格键、前导零键、空键、重复键。 */
  const SYNTH_MEMBERS = "会员表";
  const SYNTH_ORDERS = "订单表";
  /** 独立汇总表：**只做写入目标**，与两张来源表完全分开（不覆盖来源任何一格）。 */
  const SYNTH_TARGET = "汇总表";
  const SYNTH_MATRICES: Record<string, (string | number | null)[][]> = {
    [SYNTH_MEMBERS]: [
      ["会员编号", "会员", "等级"],
      ["  7  ", "带空格会员", "金卡"],
      ["7", "重复键会员", "银卡"],
      ["007", "前导零会员", "银卡"],
      ["", "空键会员", "铜卡"],
      ["9", "正常会员", "金卡"]
    ],
    [SYNTH_ORDERS]: [
      ["会员编号", "金额"],
      ["7", 100],
      ["007", 200],
      ["9", 300],
      ["404", 400]
    ],
    // 汇总表在执行前是**空表**：撤销后必须逐格回到"空"，这是可判别的执行前状态。
    [SYNTH_TARGET]: []
  };

  /**
   * 汇总结果写进**独立且已授权**的「汇总表」A1:C5。
   * 上一版把结果写进 `订单表!A4:C8` —— 那与 `订单表` 的 A1:B5 数据行**重叠**，
   * 等于边算边覆盖来源；这里改为独立目标表，来源表在整条链路上**一格都不许变**。
   */
  const SYNTH_SUMMARY: (string | number)[][] = [
    ["会员编号", "订单数", "金额合计"],
    ["7", 1, 100],
    ["007", 1, 200],
    ["9", 1, 300],
    ["404", 1, 400]
  ];
  const SYNTH_PLAN: AnalysisPlan = {
    id: "plan-synth-wps",
    title: "会员消费汇总",
    summary: "按会员编号汇总订单金额",
    assumptions: [],
    warnings: [
      "空键：会员表有 1 条记录键为空，未参与关联。",
      "重复键：会员表有 1 个键重复，保留第一次出现。",
      "未匹配：订单表有 1 笔订单未匹配到会员（键 404）。"
    ],
    actions: [
      { type: "writeValues", sheet: SYNTH_TARGET, range: "A1:C5", values: SYNTH_SUMMARY }
    ],
    acceptanceCriteria: [
      { type: "rangeEquals", sheet: SYNTH_TARGET, range: "A1:C5", expected: SYNTH_SUMMARY }
    ]
  };

  /** 两张来源表在**整条链路**里的逐格快照：预览前 / 执行后 / 撤销后都必须一模一样。 */
  function sourceCells(host: ReturnType<typeof bridgeRoundTripHost>) {
    const snap = (sheetName: string, rows: number, cols: number) =>
      Array.from({ length: rows }, (_, r) =>
        Array.from({ length: cols }, (_, c) => host.cell(sheetName, r + 1, c + 1))
      );
    return {
      members: snap(SYNTH_MEMBERS, 6, 3),
      orders: snap(SYNTH_ORDERS, 5, 2)
    };
  }

  it("同一批脏数据经桥真读真算：汇总与异常明细到达结果（两端同一套规则）", async () => {
    const host = bridgeRoundTripHost({ matrices: SYNTH_MATRICES });
    const bridge = stubBridgeEndpoints({
      poll: {
        calls: [
          {
            operationId: ROUND_TRIP_OP,
            tool: "query_table",
            arguments: {
              mode: "aggregate",
              scope: "selected",
              groupBy: ["会员编号"],
              metrics: [
                { operation: "countRows", outputName: "订单数" },
                { operation: "sum", field: "金额", outputName: "金额合计" }
              ]
            },
            requestFingerprint: ROUND_TRIP_FP,
            expiresAt: 0
          }
        ],
        generation: 1
      }
    });
    const releaseTurn = armBridgeTurn({
      kind: "plan",
      plan: SYNTH_PLAN,
      turnId: "turn-synth-panel"
    });
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    validatePlanMock.mockImplementation(async (submitted) => submitted);

    const { container } = await mount(
      <WpsAgentPanel
        selectedSheetIndexes={[1, 2, 3]}
        scopeDock={<div className="eb-scope-dock">当前表 会员表、订单表、汇总表</div>}
      />
    );
    await until(
      () =>
        container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型") ?? false,
      "应加载模型目录"
    );
    const textarea = container.querySelector("textarea");
    await type(textarea, "按会员编号汇总订单金额，结果写到汇总表");
    await submit(textarea);

    // 桥把本地真实计算的结果回执给服务端；收据里就是脏数据算出的汇总。
    await until(
      () => bridge.captured.some((item) => item.url.includes("/receipt")),
      "应把本地查询结果回执给服务端"
    );
    const receipt = bridge.captured.find((item) => item.url.includes("/receipt"))?.body;
    const payload = (receipt?.receipt as { result?: { headers: unknown[]; rows: unknown[] } })
      .result;
    // 引擎按原始键分组：键与金额逐行一致（前导零与 `7` 是不同键）。
    expect(payload?.headers).toEqual(["会员编号", "订单数", "金额合计"]);
    expect(payload?.rows).toEqual(SYNTH_SUMMARY.slice(1));
    // 写之前一个单元格都没被改。
    expect(host.writes()).toBe(0);

    // **预览阶段**：来源两表逐格与执行前一致（快照在预览之前取）。
    const beforePreview = sourceCells(host);

    // 放行：结果计划回面板 → 异常明细如实到达 → 预览 → 确认 → 写回 → 回读。
    releaseTurn();
    await until(() => findButton(container, "确认变更") !== null || textOf(container).includes("本次执行需要核对"), "应展示执行结果或明确失败");
    expect(textOf(container)).toContain("空键：会员表有 1 条记录键为空");
    expect(textOf(container)).toContain("未匹配：订单表有 1 笔订单未匹配到会员");
    expect(host.writes()).toBe(0);
    // 预览阶段来源两表逐格未变。
    expect(sourceCells(host)).toEqual(beforePreview);

    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card.verified") !== null,
      "写入成功应给出回读核对结果"
    );
    expect(host.writes()).toBeGreaterThan(0);
    // **执行后**：来源两表逐格仍然未变（写入只落在独立的汇总表）。
    expect(sourceCells(host)).toEqual(beforePreview);
    // 回读核对读的是同一份矩阵：目标表 A2 现在真的是 `7`、B2 是 1、C2 是 100。
    expect(host.cell(SYNTH_TARGET, 1, 1)).toBe("会员编号");
    expect(host.cell(SYNTH_TARGET, 2, 1)).toBe("7");
    expect(host.cell(SYNTH_TARGET, 2, 2)).toBe(1);
    expect(host.cell(SYNTH_TARGET, 2, 3)).toBe(100);
    expect(host.cell(SYNTH_TARGET, 5, 1)).toBe("404");

    // 通过**真实面板按钮**撤销，并核实目标表逐格回到执行前（空表）。
    await click(cardUndoButton(container));
    await until(() => textOf(container).includes("已撤销并核对恢复"), "撤销后应给出核对过的恢复结果");
    expect(container.querySelector(".wps-verification-card.verified")).toBeNull();
    expect(container.querySelector(".wps-verification-card.undone")).not.toBeNull();
    // **撤销后**：目标表逐格回到执行前（执行前是空表），来源两表依旧未变。
    for (let row = 1; row <= 5; row += 1) {
      for (let column = 1; column <= 3; column += 1) {
        expect(host.cell(SYNTH_TARGET, row, column)).toBe(null);
      }
    }
    expect(sourceCells(host)).toEqual(beforePreview);
  });
});

describe("WPS 面板 · 格式动作闭环（预览显示格式变化 → 副本审阅 → 确认写回 → 回读 → 撤销）", () => {
  const formatPlan: AnalysisPlan = {
    id: "plan-format-panel",
    title: "加粗 A1",
    summary: "把 A1 加粗",
    assumptions: [],
    warnings: [],
    actions: [{ type: "setFont", sheet: "Sheet1", range: "A1", bold: true }],
    acceptanceCriteria: []
  };

  it("格式动作预览明确显示格式变化（不冒充写入值），确认写回后格式生效、值不变", async () => {
    const host = writeRangeHost({ withFormat: true });
    try {
      const { container } = await mountPlanPreview(formatPlan, "把 A1 加粗");
      // 预览标题与正文都明确是"设置字体"，不是"写入值"。
      expect(container.querySelector(".wps-plan-card")?.textContent).toContain("设置字体");
      // 展开"查看预计结果"后，格式描述出现（含"加粗"）。
      expect(textOf(container)).toContain("设置字体（加粗）");
      expect(textOf(container)).not.toMatch(/写入值.*A1/);

      // 确认写回前：原表格式未改、值未改、零写入。
      expect(host.format().bold).toBe(false);
      expect(host.cellValue()).toBe("旧值");
      expect(host.writes()).toBe(0);

      // 走副本审阅 → 确认写回。
      await stageAndSave(container);
      await until(() => container.querySelector(".wps-verification-card.verified") !== null, "写入成功应给出回读核对结果");

      // 格式生效，值不变。
      expect(host.format().bold).toBe(true);
      expect(host.cellValue()).toBe("旧值");

      // 撤销：格式回到写前，值仍不变。
      await click(cardUndoButton(container));
      await until(() => textOf(container).includes("已撤销并核对恢复"), "撤销后应给出核对过的恢复结果");
      expect(host.format().bold).toBe(false);
      expect(host.cellValue()).toBe("旧值");
    } catch (error) {
      // 临时调试：打印面板文本。
      // eslint-disable-next-line no-console
      console.log("DEBUG panel text:", Array.from(document.body.querySelectorAll("div")).map((el) => el.textContent).join(" | ").slice(0, 3000));
      throw error;
    }
  });

  it("resizeRange(rowHeight) 预览明示整行影响", async () => {
    writeRangeHost({ withFormat: true });
    const rowHeightPlan: AnalysisPlan = {
      ...formatPlan,
      id: "plan-format-rowheight",
      actions: [{ type: "resizeRange", sheet: "Sheet1", range: "A1", rowHeight: 24 }]
    };
    const { container } = await mountPlanPreview(rowHeightPlan, "把 A1 行高改成 24");
    expect(textOf(container)).toContain("影响整行");
    expect(textOf(container)).toContain("调整尺寸");
  });
});

describe("WPS 面板 · 验证后保存为个人工具", () => {
  const formulaPlan: AnalysisPlan = {
    id: "plan-save-tool",
    title: "计算 A1",
    summary: "用公式计算 A1",
    assumptions: [],
    warnings: [],
    actions: [{ type: "writeFormulas", sheet: "Sheet1", range: "A1", formulas: [["=1+1"]] }],
    acceptanceCriteria: []
  };

  /** 固定内容计划（writeValues）—— 需要「以后仍写入这些固定内容」确认。 */
  const fixedContentPlan: AnalysisPlan = singleWritePlan;

  /** 破坏性计划（清空明确区域的内容）—— 需要「包含覆盖或删除操作」确认。 */
  const destructivePlan: AnalysisPlan = {
    id: "plan-save-tool-destructive",
    title: "清空 A1 内容",
    summary: "清空 A1 的内容",
    assumptions: [],
    warnings: [],
    actions: [{ type: "clearRange", sheet: "Sheet1", range: "A1", applyTo: "contents" }],
    acceptanceCriteria: []
  };

  function toolDialog(container: HTMLElement): HTMLElement {
    const dialog = container.querySelector<HTMLElement>(".tool-dialog");
    if (!dialog) throw new Error("应打开「保存为工具」对话框");
    return dialog;
  }

  function dialogButton(dialog: HTMLElement, name: string): HTMLButtonElement | null {
    return (
      Array.from(dialog.querySelectorAll<HTMLButtonElement>("button")).find(
        (item) => (item.textContent ?? "").trim() === name
      ) ?? null
    );
  }

  /** 走到结果卡并打开「保存为工具」对话框。 */
  async function openSaveDialog(container: HTMLElement): Promise<HTMLElement> {
    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card.verified") !== null,
      "写入成功应给出回读核对结果"
    );
    expect(loadTools()).toHaveLength(0);
    await click(findButton(container, "保存为工具"));
    await until(() => container.querySelector(".tool-dialog") !== null, "应打开保存为工具对话框");
    return toolDialog(container);
  }

  it("公式计划不需要额外确认：对话框无固化检查，确认后正常保存并支持编辑/删除", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(formulaPlan, "用公式计算 A1");

    await stageAndSave(container);
    await until(() => container.querySelector(".wps-verification-card.verified") !== null, "写入成功应给出回读核对结果");
    expect(host.writes()).toBeGreaterThan(0);

    // 完全通过 ⇒ 结果卡里出现「保存为工具」。
    const saveButton = findButton(container, "保存为工具");
    expect(saveButton).not.toBeNull();
    expect(loadTools()).toHaveLength(0);

    await click(saveButton);
    const dialog = toolDialog(container);
    // 公式计划既无固定内容也无破坏性动作 ⇒ 不出现固化检查，也就没有需要勾的确认项。
    expect(dialog.querySelector(".tool-save-review")).toBeNull();
    const confirm = dialogButton(dialog, "保存到我的工具");
    expect(confirm?.disabled).toBe(false);

    await click(confirm);
    await until(() => textOf(container).includes("已把「计算 A1」保存为个人工具"), "保存成功应给出系统反馈");

    // 真正落到 storage（复用 shared SavedTool / storage 语义）。
    const saved = loadTools();
    expect(saved).toHaveLength(1);
    expect(saved[0].name).toBe("计算 A1");

    // 打开工具抽屉：能看到它，且带「编辑计划」「删除」入口。
    await click(container.querySelector(".tools-entry"));
    await until(() => container.querySelector('aside[aria-label="我的工具"]') !== null, "应打开我的工具");
    expect(textOf(container)).toContain("计算 A1");
    expect(findButton(container, "编辑计划")).not.toBeNull();
    expect(findButton(container, "删除")).not.toBeNull();

    // 删除：从 storage 消失。
    await click(findButton(container, "删除"));
    await until(() => loadTools().length === 0, "删除后 storage 应清空");
  });

  it("固定内容计划（writeValues）：未勾选确认不能保存，明确勾选后保存成功", async () => {
    writeRangeHost();
    const { container } = await mountPlanPreview(fixedContentPlan, "把 A1 改成 90");
    const dialog = await openSaveDialog(container);

    // 固化检查如实点出「固定写入值」，并给出对应的确认项（默认不勾选）。
    const review = dialog.querySelector<HTMLElement>(".tool-save-review");
    expect(review?.textContent).toContain("固定写入值");
    const checkbox = review?.querySelector<HTMLInputElement>('input[type="checkbox"]');
    expect(checkbox).not.toBeNull();
    expect(checkbox!.checked).toBe(false);
    expect(review?.textContent).toContain("以后仍写入这些固定内容");

    // 未勾选 ⇒ 确认按钮禁用，点不动，也不能靠"点一下"绕过。
    const confirm = dialogButton(dialog, "保存到我的工具")!;
    expect(confirm.disabled).toBe(true);
    await click(confirm);
    expect(loadTools()).toHaveLength(0);

    // 明确勾选后才能保存。
    await click(checkbox!);
    expect(dialogButton(toolDialog(container), "保存到我的工具")!.disabled).toBe(false);
    await click(dialogButton(toolDialog(container), "保存到我的工具"));
    await until(() => textOf(container).includes("已把「更新 A1」保存为个人工具"), "勾选确认后应保存成功");

    const saved = loadTools();
    expect(saved).toHaveLength(1);
    expect(saved[0].name).toBe("更新 A1");
    // 保存的审批项如实记录，供后续运行时复述风险。
    expect(saved[0].approvals).toContain("fixedContent");
  });

  it("破坏性计划（清空区域）：未确认拒绝，确认后成功", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(destructivePlan, "清空 A1");
    const dialog = await openSaveDialog(container);

    // 清空明确区域的内容是 WPS 可执行且被判定为破坏性的动作：确认前拒绝保存。
    expect(host.writes()).toBeGreaterThan(0);
    const review = dialog.querySelector<HTMLElement>(".tool-save-review");
    expect(review?.textContent).toContain("删除、清空、覆盖或替换");
    const checkbox = review?.querySelector<HTMLInputElement>('input[type="checkbox"]');
    expect(checkbox).not.toBeNull();
    expect(checkbox!.checked).toBe(false);

    const confirm = dialogButton(dialog, "保存到我的工具")!;
    expect(confirm.disabled).toBe(true);
    await click(confirm);
    expect(loadTools()).toHaveLength(0);

    await click(checkbox!);
    await click(dialogButton(toolDialog(container), "保存到我的工具"));
    await until(() => textOf(container).includes("已把「清空 A1 内容」保存为个人工具"), "确认破坏性操作后应保存成功");

    const saved = loadTools();
    expect(saved).toHaveLength(1);
    expect(saved[0].approvals).toContain("destructive");
  });

  it("取消保存不留新工具、不留批准状态；重开对话框时确认项复位", async () => {
    writeRangeHost();
    const { container } = await mountPlanPreview(fixedContentPlan, "把 A1 改成 90");
    const dialog = await openSaveDialog(container);

    // 先勾上固定内容确认（此时确认按钮可用），再取消。
    await click(dialog.querySelector<HTMLInputElement>('.tool-save-review input[type="checkbox"]')!);
    expect(dialogButton(toolDialog(container), "保存到我的工具")!.disabled).toBe(false);
    await click(dialogButton(toolDialog(container), "取消"));

    // 取消 ⇒ 对话框关闭、storage 无新工具。
    await until(() => container.querySelector(".tool-dialog") === null, "取消后对话框应关闭");
    expect(loadTools()).toHaveLength(0);

    // 重新打开：确认项必须复位为未勾选，不能继承上一次的批准。
    await click(findButton(container, "保存为工具"));
    const reopened = toolDialog(container);
    const checkbox = reopened.querySelector<HTMLInputElement>('.tool-save-review input[type="checkbox"]')!;
    expect(checkbox.checked).toBe(false);
    expect(dialogButton(reopened, "保存到我的工具")!.disabled).toBe(true);
  });

  it("保存后从「我的工具」再次使用：按当前授权重新校验并进入写入前预览", async () => {
    writeRangeHost();
    const { container } = await mountPlanPreview(formulaPlan, "用公式计算 A1");
    const dialog = await openSaveDialog(container);
    await click(dialogButton(dialog, "保存到我的工具"));
    await until(() => loadTools().length === 1, "应先保存出一个工具");

    // 打开我的工具，点击该工具（不是编辑/删除）。
    await click(container.querySelector(".tools-entry"));
    await until(() => container.querySelector('aside[aria-label="我的工具"]') !== null, "应打开我的工具");
    const card = container.querySelector<HTMLButtonElement>(".tool-card");
    expect(card?.textContent).toContain("计算 A1");

    // 使用工具必须重新走当前授权下的校验（validatePlan），而不是直接改工作簿。
    validatePlanMock.mockClear();
    await click(card);
    await until(() => validatePlanMock.mock.calls.length === 1, "使用工具应先按当前授权校验计划");
    await until(() => textOf(container).includes("写入前确认"), "校验通过后应进入写入前预览");
    // 抽屉已收起，且没有任何写入发生。
    expect(container.querySelector('aside[aria-label="我的工具"]')).toBeNull();
  });

  it("编辑工具计划模板后重新分析保存资格：新引入的固定内容确认不被旧批准掩盖", async () => {
    writeRangeHost();
    const { container } = await mountPlanPreview(formulaPlan, "用公式计算 A1");
    const dialog = await openSaveDialog(container);
    await click(dialogButton(dialog, "保存到我的工具"));
    await until(() => loadTools().length === 1, "应先保存公式工具（无固定内容确认项）");
    expect(loadTools()[0].approvals).not.toContain("fixedContent");

    // 打开「我的工具」→ 编辑计划，把公式计划改成固定内容计划（writeValues）。
    await click(container.querySelector(".tools-entry"));
    await until(() => container.querySelector('aside[aria-label="我的工具"]') !== null, "应打开我的工具");
    await click(findButton(container, "编辑计划"));
    await until(
      () => container.querySelector('aside[aria-label="编辑工具计划模板"]') !== null,
      "应打开计划模板编辑器"
    );
    const edited: AnalysisPlan = {
      ...formulaPlan,
      actions: [{ type: "writeValues", sheet: "Sheet1", range: "A1", values: [["固定值"]] }]
    };
    await type(
      container.querySelector<HTMLTextAreaElement>(".plan-editor-textarea"),
      JSON.stringify(edited, null, 2)
    );
    await click(findButton(container, "保存"));

    // 资质重新分析：如实告知所需确认项有变化，并把新引入的 fixedContent 记进工具。
    await until(() => textOf(container).includes("重新分析了保存资格"), "应告知已重新分析保存资格");
    const updated = loadTools();
    expect(updated).toHaveLength(1);
    expect(updated[0].approvals).toContain("fixedContent");
  });

  it("撤销执行后旧保存候选不能继续保存（资格失效，确认被拦）", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(formulaPlan, "用公式计算 A1");
    const dialog = await openSaveDialog(container);
    expect(dialogButton(dialog, "保存到我的工具")!.disabled).toBe(false);

    // 对话框仍开着时撤销这次执行。
    await click(cardUndoButton(container));
    await until(() => textOf(container).includes("已撤销并核对恢复"), "撤销后应给出核对过的恢复结果");

    // 旧候选资格失效：如实给出原因，确认按钮回到禁用，不能再保存。
    await until(
      () => toolDialog(container).querySelector(".tool-save-blocked-reason") !== null,
      "撤销后旧保存候选应显示资格失效原因"
    );
    expect(toolDialog(container).textContent).toContain("可保存资格已失效");
    expect(dialogButton(toolDialog(container), "保存到我的工具")!.disabled).toBe(true);
    expect(loadTools()).toHaveLength(0);

    // 结果卡上的「保存为工具」入口随之消失（同一执行已被撤销）。
    expect(findButton(container, "保存为工具")).toBeNull();
    expect(host.writes()).toBeGreaterThan(0);
  });
});

/*
 * 两种执行方式的验收（A–F），与 Excel 端 `ExcelExecutionModeAcceptance.test.tsx`
 * **同一套判据、同一份共享语义**（`@excel-bro/core/executionMode`）。
 * 这里是 jsdom + 宿主替身：能证明守卫与状态机的判别力，不能替代实机窗口证据。
 */
describe("执行方式选择：A–F", () => {
  /** 计划预览卡上的执行方式单选。 */
  function modeRadio(container: HTMLElement, label: string): HTMLInputElement {
    const field = container.querySelector<HTMLElement>(".wps-execution-mode");
    if (!field) throw new Error("计划预览卡应有「执行方式」选择");
    const input = Array.from(field.querySelectorAll<HTMLInputElement>('input[type="radio"]'))
      .find((item) => item.closest("label")?.textContent?.includes(label));
    if (!input) throw new Error(`执行方式里应有「${label}」`);
    return input;
  }

  // A 直接执行
  it("A 直接执行：跳过副本试算，照常写原表并留下撤销入口", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); expect(host.writes()).toBe(1); expect(host.cellValue()).toBe('90'); expect(container.querySelector('.wps-execution-mode')).toBeNull(); await stageAndSave(container); expect(cardUndoButton(container)?.disabled).toBe(false);
});

  it("A 直接执行如实说明没有与副本比对，而不是让这句话缺席", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); expect(compareCommittedWpsResultToCopy).not.toHaveBeenCalled(); expect(textOf(container)).toContain('双视图中的副本展示执行前状态'); expect(textOf(container)).not.toContain('原表目标回读与审阅副本一致');
});

  // B 副本比对
  it("B 副本比对：确认前原表零写入，确认后写回并只处理本次副本", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); expect(host.cellValue()).toBe('90'); const count=host.writes(); await click(findButton(container,'双视图对比')); await stageAndSave(container); expect(host.writes()).toBe(count); expect(discardWpsNativeComparison).toHaveBeenCalledOnce();
});

  // C 放弃
  it("C 放弃：原表零写入、保持打开，且副本处理提示只针对副本", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); await click(findButton(container,'放弃变更')); expect(host.cellValue()).toBe('旧值'); expect(textOf(container)).toContain('已放弃并恢复'); expect(discardWpsNativeComparison).toHaveBeenCalledOnce();
});

  // D 切换与并发
  it("D 副本在场时切到直接执行：旧确认与写回入口失效，原表仍零写入", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); await click(findButton(container,'双视图对比')); expect(container.querySelector('.wps-execution-mode')).toBeNull(); expect(findButton(container,'确认并直接执行到原表')).toBeNull(); expect(findButton(container,'确认变更')).not.toBeNull();
});

  it("D 切换后走直接执行，且旧副本确认不能借来写表", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); await click(findButton(container,'双视图对比')); await click(findButton(container,'放弃变更')); expect(host.cellValue()).toBe('旧值'); expect(findButton(container,'确认写回原表')).toBeNull(); expect(findButton(container,'确认变更')).toBeNull();
});

  it("D 重复点击写回入口只执行一次", async () => {
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); const confirm=findButton(container,'确认变更')!; const count=host.writes(); await click(confirm); await click(confirm); expect(host.writes()).toBe(count); expect(textOf(container)).toContain('已确认变更');
});
});

/** 副本比对路径的完整写回（生成 → 勾选 → 确认）。 */
async function stageAndSaveRemounted(container: HTMLElement, host: ReturnType<typeof writeRangeHost>): Promise<void> {
  const confirm = findButton(container, "确认写回原表");
  expect(confirm).not.toBeNull();
  await acknowledgeCopy(container);
  expect(findButton(container, "确认写回原表")?.disabled).toBe(false);
  await click(findButton(container, "确认写回原表"));
  await until(() => host.writes() > 0, "确认后应写回原表");
}
