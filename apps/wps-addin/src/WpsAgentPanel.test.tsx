import {fixture as resumeHost, operationId as resumeOperationId, recoveryChannel} from "./wpsResume.testHost";
import {backendRecovery, resetBackendRecovery} from "./wpsRecoveryStore.testHost";
import {pendingDiscoveryBackend, resetPendingDiscoveryBackend} from "./wpsPendingDiscovery.testHost";
import {executeWpsPlan} from "./wpsPlanExecutor";
import {armRecoveryProbe, readRecoveryProbe} from "@excel-bro/core/recoveryProbe";
import {wpsWorkbookIdentity} from "./wpsExecutionLedger";
import { executeWorksheetDeletion, openWorksheetCheckpoint, confirmWorksheetCheckpointBackup, restoreWorksheetCheckpoint } from "../../excel-addin/src/api";
import { readFileSync } from "node:fs";
import { act, useEffect, useState, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { renderToStaticMarkup } from "react-dom/server";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { MODEL_STORAGE_KEY } from "../../excel-addin/src/conversation";
import { useModelManagement } from "../../excel-addin/src/hooks/useModelManagement";
import { ApiRequestError, getModelSettings, listModels, understandMessage, validatePlan, registerWpsExecution, getWpsExecutionStatus, startWpsAction, completeWpsAction, finishWpsExecution, suspendWpsExecution, claimWpsExecutionResume, reconcileWpsActionReceipt, saveExecutionRecoveryPackage, listPendingWpsExecutions, dismissWpsExecutionDiscovery, recordExecutionDisposition } from "../../excel-addin/src/api";
import WpsAgentPanel, { wpsWriteBackBlockReason, wpsWriteBackStillCurrent } from "./WpsAgentPanel";
import type { ReviewWriteBackTarget } from "@excel-bro/core/reviewWriteBackGate";
import { matrixFingerprintV2 } from "./wpsPlanExecutor";
import {
  resetWpsRecoveryDiagnostics,
  wpsRecoveryDiagnosticEvents
} from "./wpsRecoveryDiagnostics";
import { capabilityHost } from "./wpsCapabilityTestHost";
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
import { compareCommittedWpsResultToCopy, discardWpsNativeComparison, stageWpsBeforeComparison } from "./wpsNativeComparison";
import { clearReviewMarksFromCopySheet } from "./wpsResultReviewMarkers";
import type { AnalysisPlan } from "../../excel-addin/src/contracts";
import { loadTools } from "../../excel-addin/src/storage";

vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual =
    await importOriginal<typeof import("../../excel-addin/src/api")>();
  const recoveryBackend = await import("./wpsRecoveryStore.testHost");
  /*
   * 发现入口的**服务端那一侧**：索引只是缓存，权威发现走账本。
   * 这个替身让"完整退出 WPS 后索引丢了"能被真正测到 —— 用例不必再手工塞 operationId。
   */
  const discoveryBackend = await import("./wpsPendingDiscovery.testHost");
  const audit = await import("@excel-bro/core/hostWriteAudit");
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
    pollSdkBridgeTool: (...args: Parameters<typeof actual.pollSdkBridgeTool>) =>
      planningTransport.openBridge
        ? Promise.resolve({ calls: [], generation: 1 }) : actual.pollSdkBridgeTool(...args),
    understandMessage: vi.fn(async () => ({ kind: "needs_workbook_context", summary: "需要字段结构", provider: "model" })),
    listModels: vi.fn(),
    getModelSettings: vi.fn(),
    validatePlan: vi.fn(),
    executeWorksheetDeletion: vi.fn(),
    confirmWorksheetCheckpointBackup: vi.fn(),
    openWorksheetCheckpoint: vi.fn(),
    restoreWorksheetCheckpoint: vi.fn(),
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
    })),
    // 本机恢复包存储：面板测试不走真实网络，用同一个本机后端替身（语义与服务端一致）。
    saveExecutionRecoveryPackage: vi.fn(async (id: string, text: string) => recoveryBackend.backendSave(id, text)),
    loadExecutionRecoveryPackage: vi.fn(async (id: string) => recoveryBackend.backendLoad(id)),
    deleteExecutionRecoveryPackage: vi.fn(async (id: string) => recoveryBackend.backendRemove(id)),
    /*
     * 未完成执行的发现（只读）。以前它没被替身，于是面板每挂一次就可能真的向
     * `http://127.0.0.1:8765` 发一条查询 —— 跑测试等于查用户正在用的服务。
     */
    listPendingWpsExecutions: vi.fn(async (workbookIdentity: string) =>
      discoveryBackend.pendingDiscoveryBackend().list(workbookIdentity)),
    listPendingOfficeExecutions: vi.fn(async (workbookIdentity: string) =>
      discoveryBackend.pendingDiscoveryBackend().list(workbookIdentity)),
    dismissWpsExecutionDiscovery: vi.fn(async (operationId: string) => {
      discoveryBackend.pendingDiscoveryBackend().dismiss(operationId);
    }),
    /*
     * 用户处置留痕（放弃/撤销）：也是真实网络出口（`POST /api/wps/executions/{id}/disposition`）。
     * 调用方 `void` 掉、失败被吞 —— 与 `reportWpsHostWrites` 同一类 best-effort，
     * 漏掉替身就会让测试静默地往用户服务端发请求。
     */
    recordExecutionDisposition: vi.fn(
      async (_namespace: string, _operationId: string, _disposition: string) => undefined
    ),
    /*
     * 宿主写入次数补报也是真实网络出口（`POST /api/wps/executions/{id}/host-writes`）。
     * 调用方 `void` 掉它、失败被吞 —— 正是这种 best-effort 最容易被漏掉替身，
     * 于是"测试往用户账本里写请求"能安静很久。这里给它一个纯内存的实现。
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
    /*
     * 恢复点列表同样有真实网络出口（`POST /api/worksheet-checkpoints/list`）。
     * 面板在「备份与恢复」入口会查它 —— 以前没替身，于是一次全量测试就会按
     * **当前工作簿路径**去查用户那台 8765。默认没有任何恢复点，需要恢复点的
     * 用例自己给返回值。
     */
    listWorksheetCheckpoints: vi.fn(async () => ({ checkpoints: [] }))
  };
});

/*
 * 只替掉"那一轮规划怎么跑"，**不**替掉整个模块。
 *
 * 桥宿主适配器还要用同模块里的 `sheetsForRequest`（把读到的矩阵交给确定性计算）。
 * 早先这里是整体替换，于是桥的执行路径在本文件里永远跑不起来 —— 这正是
 * "面板接线测试"要覆盖的东西，被 mock 顺手藏掉了。
 */
const planningTransport = vi.hoisted(() => ({ run: vi.fn(), openBridge: true }));
vi.mock("./wpsReadonlyAgent", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./wpsReadonlyAgent")>()),
  runWpsReadonlyTurn: async (input: WpsReadonlyTurnInput) => {
    const result = await planningTransport.run(input) as WpsTurnResult;
    if (!planningTransport.openBridge || result?.kind !== "plan" || result.computedLocally ||
      !result.plan.actions.some(action => action.type === "writeValues" || action.type === "writeTable")) return result;
    const turnId = result.turnId ?? "turn-panel-transport";
    // Immediate plan fixtures stand in for the SDK transport. Their source is
    // read through the real host/bridge before the panel receives the plan.
    // Barrier/round-trip fixtures disable this and emit their own read events.
    if (await input.preparePlanningBridge?.(turnId)) {
      await input.onProgress?.({stage: "answering", message: "读取授权来源", turnId,
        status: "completed", stepId: "transport-read",
        bridgeTool: {operationId: "op_11111111111111111111111111111111", tool: "query_table"}});
    }
    return { ...result, turnId };
  }
}));

// 面板测试只验证“两次确认”状态机；原生复制、着色及失效检测由
// wpsNativeComparison.test.ts 的严格宿主替身独立覆盖。
// 注意：替身必须给全 `WpsNativeComparison` 的**每个**字段。少给字段时
// `vi.fn` 的类型很宽松、tsc 也放过它，但侧栏一旦真的用到就会在运行期崩
// —— 这跟宿主替身少给参数是同一类"替身比真实现宽容"的缺口。
vi.mock("./wpsNativeComparison", () => ({
  // 旧"先复制再写入"那条路（`stageWpsNativeComparison` / `WpsFormulaTrialError`）已随
  // `stagePendingPlan` 一起删掉：面板不再 import 它们，替身也就不再提供 —— 留着会让
  // 后来人以为侧栏还会走那条路。函数本体仍有自己的直测（`wpsNativeComparison.test.ts`）。
  // 执行前快照：真实实现只读原表、不建副本。替身按同一形状给（值取自计划动作，
  // 保证"快照可用"这一前提成立），但**不**假装做过宿主复制 —— 复制只发生在 showBefore。
  /*
   * 执行前副本：真实实现按**执行前快照**重建（先复制、再把快照写回副本）。
   * 替身只保证形状与"调用过"，真正的重建逻辑由 `wpsNativeComparison` 单测覆盖。
   */
  stageWpsBeforeComparison: vi.fn((_application: unknown, before: unknown) => ({
    plan: (before as { prepared: unknown }).prepared,
    original: { Name: "测试.xlsx" },
    copy: { Name: "Book1" },
    copyIdentity: { name: "Book1", fullName: null, path: null },
    /*
     * `writeSheet` 是真实返回的一部分（`wpsNativeComparison.stageWpsBeforeComparison`）：
     * 本次审阅的目标绑定靠它记下"写入的是哪一张表"。替身少了它，绑定那一句会直接
     * 抛 TypeError，于是双视图打开失败、副本从未登记 —— 症状是"放弃之后没有副本
     * 收尾提示"，看上去像产品缺陷，其实是替身不完整。
     */
    writeSheet: { index: 1, name: "Sheet1" },
    originalBefore: [],
    copyFingerprints: [],
    copyAfter: [],
    reviews: [],
    changedCells: []
  })),
  captureWpsBeforeResult: vi.fn((_application: unknown, prepared: unknown) => {
    const plan = prepared as { actions: { values: unknown[][] }[] };
    return {
      prepared,
      values: plan.actions.map(() => ({ values: [[]], formulas: [[]], valuesCaptured: true })),
      workbookFullName: "C:////Data////测试.xlsx"
    };
  }),
  // 双视图排列：真实实现要宿主窗口能力；替身只证明"调用过、且没有再次执行计划"。
  rearrangeWpsComparison: vi.fn(() => "arranged"),
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

/*
 * 附件接缝：只替换 `prepareImageFile` 这**一步**（内部 `new Image()` + canvas 解码，
 * jsdom 做不到），其余限额、去重、顺序、`removeImage` 全跑真实实现 —— 于是
 * "只贴了图片、没打字"这条分支在 WPS 端也能被真正断言，而不是靠"两端读同一个
 * hook"这种同源结构推断。路径解析到与 `hooks/useImageAttachments` 同一个模块。
 */
vi.mock("../../excel-addin/src/imageAttachments", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../excel-addin/src/imageAttachments")>();
  return {
    ...actual,
    prepareImageFile: vi.fn(async (file: File) => ({
      id: "image-test-1",
      name: file.name,
      mediaType: "image/png" as const,
      data: "data:image/png;base64,AAAA",
      previewUrl: "blob:image-test-1"
    }))
  };
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
const listPendingWpsMock = vi.mocked(listPendingWpsExecutions);
const dismissDiscoveryMock = vi.mocked(dismissWpsExecutionDiscovery);
const runTurnMock = planningTransport.run as ReturnType<typeof vi.fn<typeof runWpsReadonlyTurn>>;
const getWpsApplicationMock = vi.mocked(getWpsApplication);
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
    rerender: async (element: ReactElement) => { await act(async () => { root.render(element); }); },
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
      /*
       * 按**可访问名**找，不只看 textContent：主操作按钮（右下角那一颗）只画字形
       * （↑/■/▶/◎/…），动作名在 `aria-label` 上。按钮从卡面搬到输入框之后，
       * 查找方式要跟着改，断言本身（零写入、身份、授权、只跑剩余动作）一个字没动。
       */
      (item) => (item.textContent ?? "").trim() === name || item.getAttribute("aria-label") === name
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

/**
 * 选中第 index 个快捷答案并**提交**。
 *
 * 新交互下点选项只选中（想改口径还有机会），提交是右下角那颗「发送回答」。
 * 旧用例写的是"点一下就等于作答"，语义等价于这里的两次点击。
 */
async function selectOption(container: HTMLElement, index: number) {
  const options = Array.from(container.querySelectorAll<HTMLButtonElement>(".clarification-options button"));
  const option = options[index];
  if (!option) throw new Error(`没有第 ${index} 个快捷答案`);
  await click(option);
  const submitButton = container.querySelector<HTMLButtonElement>('button[aria-label="发送回答"]');
  if (!submitButton) throw new Error("没有找到「发送回答」按钮");
  await click(submitButton);
}


/**
 * 点某个快捷答案并提交（新交互：点只选中，提交是右下角那颗「发送回答」）。
 */
async function answerWith(container: HTMLElement, option: HTMLElement | null) {
  await click(option);
  await click(container.querySelector<HTMLButtonElement>('button[aria-label="发送回答"]'));
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

/**
 * 往文件输入里放一张图。jsdom 的 `input.files` 只读，用 `defineProperty` 顶替；
 * React 对 `input[type=file]` 就是监听原生 `change`，所以派发一个冒泡的 change 即可。
 * 图片**解码**那一步由模块级 `prepareImageFile` 替身接管，本函数只负责"用户选了文件"。
 */
async function attachFile(input: HTMLInputElement | null, file: File): Promise<void> {
  if (!input) throw new Error("没有找到图片附件输入框");
  Object.defineProperty(input, "files", { value: [file], configurable: true });
  await act(async () => {
    input.dispatchEvent(new Event("change", { bubbles: true }));
  });
  await settle();
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
  /**
   * 由用例随时打开：此后每一次 `getItem` 都抛错。
   * 用于"写前阶段就已经读不动恢复点"——那时一次写入都还没发生，
   * `breakStorageReadsAfterWrite`（按写入次数触发）覆盖不到。
   */
  shouldBreakReads?: () => boolean;
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
  range.Item = () => range;
  const sheet = { Name: "Sheet1", Index: 1, UsedRange: range, ProtectContents: false, Range: () => range };
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
        if (options.shouldBreakReads?.()) {
          throw new Error("PluginStorage 读取失败");
        }
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

/** 走到"本次执行计划"这一步：两端对照用的同一任务状态。 */
async function mountPlanPreview(plan: AnalysisPlan, task: string, planOrigin?: "literal_range_edit") {
  listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
  runTurnMock.mockResolvedValue({ kind: "plan", plan, planOrigin, turnId: "turn-preview-1" });
  validatePlanMock.mockImplementation(async (submitted) => submitted);
  const { container } = await mount(panel());
  await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型目录");
  const textarea = container.querySelector("textarea");
  await type(textarea, task);
  await submit(textarea);
  await until(() => findButton(container, "确认变更") !== null || (textOf(container).includes("本次执行需要核对") || textOf(container).includes("未执行")), "应展示执行结果或明确失败");
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
  // 本机后端恢复包存储按用例隔离：上一例留下的材料绝不能变成这一例的"可续跑证据"。
  resetBackendRecovery();
  // 账本侧的发现也按用例隔离：上一条用例登记的执行不许被这一条"发现"出来。
  resetPendingDiscoveryBackend();
  planningTransport.openBridge = true;
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
  it("history alignment: cancel keeps records; confirm removes only the selected record and preserves other guards", async () => {
    const host = writeRangeHost();
    const key = "excel-bro.wps.chat.v1";
    const conversations = ["alpha", "beta", "guarded"].map((id) => ({ id, title: id,
      createdAt: "2026-10-01T00:00:00Z", updatedAt: "2026-10-01T00:00:00Z",
      messages: [{ id: `m-${id}`, role: "user", text: `普通聊天 ${id}`, createdAt: "2026-10-01T00:00:00Z" }]
    }));
    localStorage.setItem(key, JSON.stringify({ activeConversationId: "alpha", conversations,
      legacyCredentialConversationIds: ["guarded"] }));
    const view = await mount(panel());
    const saved = () => JSON.parse(localStorage.getItem(key)!);
    const history = async () => {
      await click(view.container.querySelector('[aria-label="更多"]'));
      expect(textOf(view.container)).toContain("查看并打开已保存的对话");
      await click(view.container.querySelector('[aria-label="历史对话"]'));
    };
    await history();
    await click(view.container.querySelector('[aria-label="删除「beta」"]'));
    expect(view.container.querySelector('[aria-label="删除历史对话"]')).not.toBeNull();
    await click(findButton(view.container, "取消"));
    expect(saved().conversations.map((c: {id: string}) => c.id)).toEqual(["alpha", "beta", "guarded"]);
    await click(view.container.querySelector('[aria-label="删除「beta」"]'));
    await click(findButton(view.container, "确认删除"));
    expect(saved().conversations.map((c: {id: string}) => c.id)).toEqual(["alpha", "guarded"]);
    expect(saved().activeConversationId).toBe("alpha");
    expect(saved().legacyCredentialConversationIds).toEqual(["guarded"]);
    await click([...view.container.querySelectorAll<HTMLButtonElement>(".history-open-button")]
      .find((b) => b.textContent?.includes("guarded"))!);
    expect(view.container.querySelector(".message-stream")?.textContent).toContain("普通聊天 guarded");
    expect(saved().legacyCredentialConversationIds).toEqual(["guarded"]);
    expect(view.container.querySelector(".stale-credentials-warning")).not.toBeNull();
    expect(host.writes()).toBe(0);
    for (const api of [registerWpsExecution, startWpsAction, completeWpsAction, finishWpsExecution,
      claimWpsExecutionResume, reconcileWpsActionReceipt, recordExecutionDisposition]) expect(api).not.toHaveBeenCalled();
    expect(runTurnMock).not.toHaveBeenCalled();
  });

  it("history alignment: existing 30 by 120 histories survive load, switch, save and reopen", async () => {
    const host = writeRangeHost();
    const key = "excel-bro.wps.chat.v1";
    const conversations = Array.from({length: 30}, (_, i) => ({ id: `c${i}`, title: `普通对话${i}`,
      createdAt: "2026-10-01T00:00:00Z", updatedAt: "2026-10-01T00:00:00Z",
      messages: Array.from({length: 120}, (_, j) => ({ id: `m${i}-${j}`, role: j % 2 ? "assistant" : "user",
        text: `合成消息${i}-${j}`, createdAt: "2026-10-01T00:00:00Z" })) }));
    localStorage.setItem(key, JSON.stringify({ activeConversationId: "c0", conversations }));
    const view = await mount(panel());
    const check = () => {
      const stored = JSON.parse(localStorage.getItem(key)!);
      expect(stored.conversations).toHaveLength(30);
      for (const c of stored.conversations) {
        expect(c.messages).toHaveLength(120);
        expect(c.messages[0].text).toBe(`合成消息${c.id.slice(1)}-0`);
      }
    };
    check();
    await click(view.container.querySelector('[aria-label="更多"]'));
    await click(view.container.querySelector('[aria-label="历史对话"]'));
    await click([...view.container.querySelectorAll<HTMLButtonElement>(".history-open-button")]
      .find((b) => b.textContent?.includes("普通对话29"))!);
    expect(view.container.querySelector(".message-stream")?.textContent).toContain("合成消息29-0");
    check(); view.unmount(); mounted = mounted.filter((entry) => entry !== view);
    const reopened = await mount(panel());
    expect(reopened.container.querySelector(".message-stream")?.textContent).toContain("合成消息29-119");
    check(); expect(host.writes()).toBe(0);
    for (const api of [registerWpsExecution, startWpsAction, claimWpsExecutionResume, reconcileWpsActionReceipt]) expect(api).not.toHaveBeenCalled();
  });

  it("history alignment: new chats retain messages and a reopened old chat supports ordinary follow-up", async () => {
    const host = writeRangeHost();
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    vi.mocked(understandMessage).mockResolvedValue({ kind: "conversation", message: "普通回答甲", provider: "model" });
    const view = await mount(panel());
    await until(() => view.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型") ?? false, "模型就绪");
    const send = async (text: string) => { await type(view.container.querySelector("textarea"), text); await submit(view.container.querySelector("textarea")); };
    await send("普通问题甲"); await until(() => textOf(view.container).includes("普通回答甲"), "甲回复");
    await click(view.container.querySelector('[aria-label="新对话"]'));
    vi.mocked(understandMessage).mockResolvedValue({ kind: "conversation", message: "普通回答乙", provider: "model" });
    await send("普通问题乙"); await until(() => textOf(view.container).includes("普通回答乙"), "乙回复");
    view.unmount(); mounted = mounted.filter((entry) => entry !== view);
    const reopened = await mount(panel());
    await click(reopened.container.querySelector('[aria-label="更多"]'));
    await click(reopened.container.querySelector('[aria-label="历史对话"]'));
    await click([...reopened.container.querySelectorAll<HTMLButtonElement>(".history-open-button")].find((b) => b.textContent?.includes("普通问题甲"))!);
    expect(reopened.container.querySelector(".message-stream")?.textContent).toContain("普通回答甲");
    expect(reopened.container.querySelector(".message-stream")?.textContent).not.toContain("普通回答乙");
    vi.mocked(understandMessage).mockResolvedValue({ kind: "conversation", message: "继续普通回答", provider: "model" });
    await type(reopened.container.querySelector("textarea"), "继续普通聊天"); await submit(reopened.container.querySelector("textarea"));
    await until(() => textOf(reopened.container).includes("继续普通回答"), "历史可以续聊");
    expect(runTurnMock).not.toHaveBeenCalled(); expect(host.writes()).toBe(0);
    expect(registerWpsExecution).not.toHaveBeenCalled();
  });

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
    expect(textOf(container)).not.toContain("先判断是否需要工作簿上下文");
    expect(textOf(container)).not.toContain("本次没有需要你处理的后续操作");
    expect(container.querySelector(".task-progress-timeline")?.closest(".message-row")?.querySelector(".message-author")).toBeNull();
    expect(textOf(container)).not.toContain("读取已授权表结构");
    expect(vi.mocked(understandMessage).mock.calls[0][0]).not.toHaveProperty("scope");
  });

  it("message-only clarification replaces composer and an answer never authorizes workbook access", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    vi.mocked(understandMessage).mockResolvedValueOnce({ kind: "clarification", provider: "model", clarification: {
      id: "rules", question: "怎样分级？", summary: "规则待定", reason: "需要规则", scopeLabel: "当前要求",
      options: [{ id: "range", label: "区间分级", description: "按区间", resolution: "按区间" }, { id: "map", label: "逐项对应", description: "按对应", resolution: "按对应" }]
    }}).mockResolvedValue({ kind: "needs_workbook_context", summary: "需要结构", provider: "model" });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型") ?? false, "模型就绪");
    await type(container.querySelector("textarea"), "先问规则，不要读取或修改工作表");
    await submit(container.querySelector("textarea"));
    await until(() => !!container.querySelector(".composer-box.is-answering"), "问卷替换输入框");
    const option = container.querySelector<HTMLButtonElement>(".clarification-options button");
    const other = container.querySelector<HTMLTextAreaElement>('textarea[data-clarification-other]');
    expect(other?.placeholder).toBe("其他");
    expect(container.querySelector(".message-stream .clarification-card.pending")).toBeNull();
    expect(textOf(container).split("怎样分级？")).toHaveLength(2);
    await click(option);
    expect(option?.getAttribute("aria-pressed")).toBe("true");
    await type(other, "我自己的规则");
    expect(option?.getAttribute("aria-pressed")).toBe("false");
    await click(option);
    expect(other?.value).toBe("");
    expect(understandMessage).toHaveBeenCalledTimes(1);
    await click(container.querySelector<HTMLButtonElement>(".composer-primary"));
    await until(() => textOf(container).includes("回答已记录，尚未读取或修改工作表"), "回答不授权读写");
    expect(understandMessage).toHaveBeenCalledTimes(2);
    expect(vi.mocked(understandMessage).mock.calls[1][0].prompt).toContain("不要读取或修改工作表");
    expect(runTurnMock).not.toHaveBeenCalled();
    expect(registerWpsExecution).not.toHaveBeenCalled();
    expect(container.querySelector(".composer-box.is-answering")).toBeNull();
    expect(container.querySelector(".message-stream .clarification-card.resolved")?.textContent).toContain("怎样分级？");
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
    await until(() => Boolean(container.querySelector(".composer-box .clarification-composer")), "关键歧义才显示卡片");
    // 自由输入就是输入框本身（不再有第二个光标），提交是右下角「发送回答」。
    await type(container.querySelector<HTMLTextAreaElement>(".composer textarea"), "只统计记录数，不修改表格");
    await click(container.querySelector<HTMLButtonElement>('button[aria-label="发送回答"]'));
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
    expect(optionButton("旧规则")?.closest(".composer-box")).not.toBeNull();
    expect(container.querySelector(".clarification-bypass-toggle")?.closest(".composer-box")).not.toBeNull();
    expect(container.querySelectorAll(".clarification-composer")).toHaveLength(1);
    await answerWith(container, optionButton("旧规则"));
    await until(() => !!optionButton("改为排序"), "应允许显式替换需求");
    await answerWith(container, optionButton("改为排序"));
    await until(() => !!optionButton("升序"), "新需求应只补自己的缺项");
    expect(runTurnMock.mock.calls[2][0].taskContext).toEqual({ schemaVersion: 1, originalPrompt: "按名称排序", confirmations: [] });
    await answerWith(container, optionButton("升序"));
    await until(() => textOf(container).includes("新任务规则已收到"), "新需求应完成");
    expect(container.querySelector(".composer-box.is-answering")).toBeNull();
    expect(runTurnMock.mock.calls[3][0].taskContext).toEqual({ schemaVersion: 1, originalPrompt: "按名称排序", confirmations: [
      { questionId: "sort-order", label: "升序", value: "升序排列" }
    ] });
    const history = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1")!);
    const card = history.conversations.flatMap((item: { messages: unknown[] }) => item.messages)
      .find((item: { taskResume?: { confirmations: { questionId?: string }[] } }) =>
        item.taskResume?.confirmations.some((decision) => decision.questionId === "sort-order"));
    expect(card.taskResume.originalPrompt).toBe("按名称排序");
    expect(card.taskResume.confirmations).toEqual([{ questionId: "sort-order", label: "升序", value: "升序排列" }]);
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
    // 字段表单的提交按钮已并入右下角那颗主按钮（必填没填完它就是禁用的）。
    const button = container.querySelector<HTMLButtonElement>('button[aria-label="发送回答"]');
    if (!button) throw new Error("缺少「发送回答」按钮");
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
    await until(() => Boolean(container.querySelector(".composer-box .clarification-composer")), "应显示澄清卡");
    await click(findButton(container, "取消本次任务"));
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
    await until(() => Boolean(container.querySelector('[aria-label="继续未完成的需求分析"]')), "应保留原需求的继续入口");
    expect(textOf(container)).not.toContain("尚未生成规划检查点");
    expect(container.querySelector('[aria-label="继续未完成的需求分析"]')).not.toBeNull();
    expect(container.querySelector(".wps-planning-pause-card")).toBeNull();
    expect(runTurnMock).toHaveBeenCalledTimes(1);
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

  it("executes plan immediately and shows result review", async () => {
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
    range.Item = () => range;
  const sheet = { Name: "Sheet1", Index: 1, UsedRange: range, ProtectContents: false, Range: () => range };
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
    // New flow: plan is executed immediately, showing result review bar
    await until(() => container.querySelector(".result-review-bar") !== null, "应展示执行结果审阅条");
    await until(() => textOf(container).includes("请查看主表中的执行结果"), "应提示查看主表结果");

    // Execution has completed
    expect(writes).toBe(1);
    expect(value).toBe("新值");

    // Result review bar shows review phase
    const reviewBar = container.querySelector(".result-review-bar");
    expect(reviewBar).not.toBeNull();
    expect((reviewBar as HTMLElement).textContent).toContain("更新 A1");

    // Verify buttons are present
    expect(findButton(container, "双视图对比")).not.toBeNull();
    expect(findButton(container, "放弃变更")).not.toBeNull();
    expect(findButton(container, "确认变更")).not.toBeNull();

    // Confirm the changes
    await click(findButton(container, "确认变更"));
    await until(() => textOf(container).includes("已确认变更"), "确认后应显示已确认状态");

    // Verify final state
    expect(writes).toBe(1);
    expect(value).toBe("新值");
    const savedHistory = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    const savedMessage = savedHistory.conversations[0].messages.find((message: { planPreview?: unknown }) => message.planPreview);
    expect(savedMessage).toBeDefined();
  });

  /*
   * 悬浮「回到底部」浮标的运行状态误判（用户点名的回归）。
   *
   * 根因：`running` 直接接了 `busy`，而 `busy` 含 `phase === 'reviewing'`
   * （"结果已就绪、等用户确认/放弃"）。于是执行早就结束了，浮标还在跳三点。
   *
   * 这里**不**只测"共享按钮收到 false 会画锚"——那证明不了接线对。要验的是
   * 面板自己算出来的 `running`：
   *   ① 结果已生成、reviewing、pending=false ⇒ 锚；
   *   ② 确认/放弃这类实际操作在途（pending=true）⇒ 三点；
   *   ③ 操作完成 ⇒ 回锚。
   * 同时 `busy` 的防重复执行/互斥边界必须仍在（确认按钮在 reviewing 且非 pending 时可点）。
   */
  it("review bar is waiting for the user, so the scroll hint shows the anchor - not the running dots", async () => {
    let value = "旧值";
    let formula = "旧值";
    const range = { Rows: { Count: 1 }, Columns: { Count: 1 }, MergeCells: false } as Record<string, unknown>;
    Object.defineProperties(range, {
      Value2: {
        get: () => value,
        set: (next: string) => { value = next; formula = next; },
        enumerable: true
      },
      Formula: {
        get: () => formula,
        set: (next: string) => { formula = next; value = next; },
        enumerable: true
      }
    });
    range.Item = () => range;
  const sheet = { Name: "Sheet1", Index: 1, UsedRange: range, ProtectContents: false, Range: () => range };
    const storage = new Map<string, unknown>();
    getWpsApplicationMock.mockReturnValue({
      Version: "12.0",
      ActiveWorkbook: {
        Name: "测试.xlsx", FullName: "C:\\Data\\测试.xlsx", Path: "C:\\Data", ReadOnly: false,
        Worksheets: { Count: 1, Item: () => sheet }, ActiveSheet: sheet
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
      id: "write-hint", title: "更新 A1", summary: "把 A1 更新为新值",
      assumptions: [], warnings: [],
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
    await until(() => container.querySelector(".result-review-bar") !== null, "应展示执行结果审阅条");

    // 滚离底部：浮标只在离开底部时渲染（判据是位置，不要求有新内容）。
    const stream = container.querySelector<HTMLElement>(".message-stream")!;
    Object.defineProperties(stream, {
      scrollHeight: { configurable: true, value: 1000 },
      clientHeight: { configurable: true, value: 300 }
    });
    stream.scrollTop = 200;
    await act(async () => { stream.dispatchEvent(new Event("scroll", { bubbles: true })); });

    /*
     * ① 结果已就绪、等人：`aria-busy=false` 证明审阅通道此刻没有操作在途
     *    （不是"按钮替身画了个锚"），此时浮标必须是**锚**。
     */
    const bar = container.querySelector<HTMLElement>(".result-review-bar")!;
    expect(bar.getAttribute("aria-busy")).toBe("false");
    expect((bar as HTMLElement).textContent).toContain("请查看主表中的执行结果");
    await until(
      () => container.querySelector<HTMLButtonElement>(".new-progress-button")?.dataset.mode === "idle",
      "只是等用户查看结果时，浮标应当显示锚而不是三点"
    );
    const hint = container.querySelector<HTMLButtonElement>(".new-progress-button")!;
    expect(hint.querySelector(".progress-anchor-icon")).not.toBeNull();
    expect(hint.querySelector(".progress-dots")).toBeNull();

    /*
     * ② 实际操作在途：点「放弃变更」→ 真实恢复流程里的
     *    `await targetGuard.check()`（→ `inspectWpsAuthorizedTarget`，可挂起）。
     *    这是**真的在跑**：`resultReviewState.pending` 会一直是 true，
     *    而 `phase` 仍是 `reviewing` —— 正好证明"在途"与"等人"是两件事。
     *    `until` 等的是三点那一态本身：接线若错，这里等不到而失败，不会悄悄放行。
     */
    holdInspect();
    await click(findButton(container, "放弃变更"));
    await until(
      () => container.querySelector<HTMLButtonElement>(".new-progress-button")?.dataset.mode === "running",
      "恢复流程还在跑，浮标应当显示三点"
    );
    expect(container.querySelector<HTMLButtonElement>(".new-progress-button")!.querySelectorAll(".progress-dots > span")).toHaveLength(3);
    // 同一时刻审阅通道确实在途：按钮被禁用、`aria-busy` 为真。
    expect(bar.getAttribute("aria-busy")).toBe("true");
    expect(findButton(container, "放弃变更")!.disabled).toBe(true);
    expect(inspectPending()).toBe(true);

    /*
     * ③ 操作完成：回到锚。同时 `busy` 的边界没被破坏 —— 放弃成功后固定条收起，
     *    而恢复本身真的发生了（`busy` 的互斥保护一直都在，只是不再冒充"在跑"）。
     */
    await releaseInspect();
    await until(() => textOf(container).includes("已放弃并恢复"), "放弃后应显示已放弃并恢复");
    await until(
      () => container.querySelector<HTMLButtonElement>(".new-progress-button")?.dataset.mode === "idle",
      "放弃恢复完成后浮标应当回到锚"
    );
    expect(container.querySelector<HTMLButtonElement>(".new-progress-button")).toBe(hint);
    // 固定条此刻应收起（成功终态、没有待办）；若还挂着 aria-busy=true，
    // 说明恢复没收尾完 —— 那才是真的"还在跑"。
    const settledBar = container.querySelector<HTMLElement>(".result-review-bar");
    if (settledBar) expect(settledBar.getAttribute("aria-busy")).toBe("false");
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
    range.Item = () => range;
  const sheet = { Name: "Sheet1", Index: 1, UsedRange: range, ProtectContents: false, Range: () => range };
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

    // New flow: plan executes immediately
    await until(() => container.querySelector(".result-review-bar") !== null, "应展示执行结果审阅条");

    // Core assertion: local plan never sent to server for validation
    expect(validatePlanMock).not.toHaveBeenCalled();

    // Source data should not appear in persisted history
    const persisted = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
    expect(persisted).not.toContain("旧值");
    expect(persisted).not.toContain("办公用品");
  });

  it("keeps a secret leaked by an upstream error out of the persisted chat history", async () => {
    // 这条断言必须**先注入**再检查：如果不往流程里塞一个密钥，
    // `not.toContain("sk-")` 会因为压根没有密钥而空过 —— 那是假绿灯。
    // 通道是真实的：`turnErrorMessage` 会把非连接类错误的 message 原样写进
    // 对话历史，而服务端错误 message 源自 `str(error)`，可能带上请求里的凭据。
    const secret = "sk-test1234567890";
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
    const secret = "sk-test1234567890";
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
    const secret = "sk-test1234567890";
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
    const secret = "sk-test1234567890";
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
    const secret = "sk-test1234567890";
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
    const secret = "sk-test1234567890";
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
    range.Item = () => range;
  const sheet = { Name: "Sheet1", Index: 1, UsedRange: range, ProtectContents: false, Range: () => range };
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
    await until(() => textOf(container).includes("本次执行计划"), "应展示写入计划");

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
      findButton(container, "添加连接")?.hasAttribute("disabled")
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
    expect(markup).toContain("添加连接");
    expect(markup).toContain("/function 公式模型");
    expect(markup).toContain("接口协议");
    expect(markup).toContain("测试用途");
    expect(markup).toContain("Anthropic 兼容（SDK Agent）");
    expect(markup).toContain("独立公式测试通过不代表 Agent 可用");

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
        /<button class="model-primary-action" type="button"([^>]*)>添加连接<\/button>/
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
    // 新流程是"先执行、后审阅"：执行结果必须已经在场，且与计划摘要**同时**可见。
    // 这两块信息不能互相顶替 —— 一个说"改成了什么"，一个说"按什么口径改的"。
    expect(container.querySelector(".wps-verification-card")).not.toBeNull();
    expect(notes.map((note) => note.textContent).join(" ")).toContain("源列空白按「未分类」处理。");
  });

  it("hides both note blocks when the plan states neither", async () => {
    writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");

    expect(container.querySelector(".wps-plan-assumptions")).toBeNull();
    expect(container.querySelector(".wps-plan-warnings")).toBeNull();
    expect(textOf(container)).not.toContain("采用的信息");
  });

  /*
   * 原「marks the preview as edited only after the user rewrote the plan」随
   * 「编辑计划」一起取消：新流程下计划一经校验即执行，不存在"改写后再预览"这一步。
   * 它当时钉的信息面（"这份计划被改过 / 到底按哪份动作写的"）仍然必要，但载体换了：
   * 执行摘要卡保留假设、警告、格式变化说明，并提供只读的完整计划 JSON。
   * 那条判据在「mirrors the Excel plan card hierarchy」与本地映射用例里继续成立。
   */

  it("放弃变更走真实恢复点撤销；从未打开双视图时不创建也不处理任何副本", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    // 新流程：执行已经发生，主表是新值 —— 这是"放弃"要撤销的对象。
    expect(host.cellValue()).toBe("90");
    await click(findButton(container, "放弃变更"));
    expect(host.cellValue()).toBe("旧值");
    expect(textOf(container)).toContain("已放弃并恢复");
    expect(findButton(container, "确认变更")).toBeNull();
    // 没点过「双视图对比」就没有副本可处理：不能凭空调用副本清理，
    // 那会把「关闭某个窗口」做成一次没有对象、却可能打错目标的操作。
    expect(discardWpsNativeComparison).not.toHaveBeenCalled();
  });

  it.each(["close_unconfirmed", "failed", "unsupported", "identity_lost", "ambiguous", "user_saved",
    "extra_edits", "edits_unverified"] as const)(
    "双视图后放弃，副本未收尾（%s）：业务结论仍是「已放弃并恢复」，副本问题单列", async status => {
      const host = writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      // 先真的打开一次双视图：只有这时才存在"本次拥有的副本"。
      await click(findButton(container, "双视图对比"));
      vi.mocked(discardWpsNativeComparison).mockReturnValueOnce(status);
      await click(findButton(container, "放弃变更"));
      // 原表恢复是业务结论，与副本收尾成败无关。
      expect(host.cellValue()).toBe("旧值");
      expect(textOf(container)).toContain("已放弃并恢复");
      expect(findButton(container, "确认变更")).toBeNull();
      expect(findButton(container, "放弃变更")).toBeNull();
      expect(textOf(container)).toContain("执行前副本仍保留");
      // 不推断宿主在问什么；也不教用户无脑选"不保存"。
      expect(textOf(container)).not.toContain("选择「不保存」");
      expect(textOf(container)).not.toContain("正在询问是否保存");
    });

  it("双视图只展示执行前状态：打开对比不会再次执行计划", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    const writes = host.writes();
    await click(findButton(container, "双视图对比"));
    expect(host.writes()).toBe(writes);
    /*
     * 副本是"执行前"的：它由 `stageWpsBeforeComparison` 按**执行前快照**重建，
     * 而不是把执行后的表复制一份（那样对比毫无意义）。
     * 正向锚点在上面：`host.writes()` 一格没变 + 只调了执行前快照那一条。
     */
    expect(stageWpsBeforeComparison).toHaveBeenCalledOnce();
    expect(findButton(container, "确认变更")).not.toBeNull();
  });

  it("双视图打开失败不吞掉已执行的结果，确认入口仍可用", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    vi.mocked(stageWpsBeforeComparison).mockImplementationOnce(() => {
      throw new Error("当前宿主未提供工作表副本接口，未打开双视图。");
    });
    await click(findButton(container, "双视图对比"));
    expect(host.writes()).toBe(1);
    expect(textOf(container)).toContain("未提供工作表副本接口");
    expect(findButton(container, "确认变更")?.disabled).toBe(false);
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
    expect(textOf(container)).toContain("本次执行计划");
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
    await until(() => textOf(container).includes("已撤销并恢复"), "撤销后应给出核对过的恢复结果");

    const undone = container.querySelector(".wps-verification-card.undone");
    expect(undone).not.toBeNull();
    expect(undone?.textContent).not.toContain("写入验证通过");
    expect(undone?.textContent).toContain("恢复点仅在当前加载项会话内可用");
    expect(container.querySelector(".wps-verification-card.verified")).toBeNull();
    expect(host.cellValue()).toBe("旧值");
  });

  it("撤销本次执行 → 服务端留痕记「撤销」，不是「放弃」", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card.verified") !== null,
      "写入成功应给出回读核对结果"
    );

    await click(await undoButton(container));
    await until(() => textOf(container).includes("已撤销并恢复"), "撤销后应给出核对过的恢复结果");

    /*
     * 留痕是 best-effort（`void` 掉的），等的是"端点被调用过"。
     * 这一格以前完全不存在：账本只记进度，用户点了「撤销本次执行」服务端一个字都不留。
     */
    await until(
      () => vi.mocked(recordExecutionDisposition).mock.calls.length > 0,
      "撤销成功后必须上报处置留痕"
    );
    const [namespace, executionId, disposition] = vi.mocked(recordExecutionDisposition).mock.calls[0];
    expect(namespace).toBe("wps");
    expect(typeof executionId).toBe("string");
    expect((executionId as string).length).toBeGreaterThan(0);
    expect(disposition).toBe("undone"); // 不是 "abandoned"
    // 留痕不改结论：单元格确实回到执行前。
    expect(host.cellValue()).toBe("旧值");
  });

  it("留痕上报失败：撤销照样成功、不重试、不再撤一次", async () => {
    /*
     * 留痕是观测，不是恢复的前置条件：上报抛错不许把"已撤销并恢复"改判成失败，
     * 也不许自动重试（重试若落在错的位置就变成"再撤一次"）。
     */
    vi.mocked(recordExecutionDisposition).mockRejectedValueOnce(new Error("网络不可用"));
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card.verified") !== null,
      "写入成功应给出回读核对结果"
    );

    await click(await undoButton(container));
    await until(() => textOf(container).includes("已撤销并恢复"), "撤销后应给出核对过的恢复结果");
    await until(
      () => vi.mocked(recordExecutionDisposition).mock.calls.length > 0,
      "撤销成功后必须上报处置留痕"
    );
    expect(vi.mocked(recordExecutionDisposition)).toHaveBeenCalledTimes(1);
    expect(textOf(container)).not.toContain("无法记录这次执行的处置");
    expect(host.cellValue()).toBe("旧值");
  });

  it("撤销被指纹复核挡下：宿主一个单元格都没回退，就不许留痕", async () => {
    /*
     * 留痕的语义是"用户**确实**把这次执行撤回了"。恢复没发生（冲突被挡下）时写
     * `undone_at` 就是造假 —— 这条与 Excel 侧的「恢复失败不留痕」是同一条判据的两端。
     *
     * WPS 的判据位置在 `undoExecution` 的成功分支：`await undoWpsExecution(...)`
     * 抛错就直接跳出，`recordExecutionDisposition` 根本走不到。
     */
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card.verified") !== null,
      "写入成功应给出回读核对结果"
    );

    // 执行之后用户自己动了同一格：撤销必须停下，不能把用户的编辑盖掉。
    host.setCell("我手动改的");
    const writes = host.writes();
    await click(await undoButton(container));
    await until(
      () => textOf(container).includes("在执行后又被修改，不能自动覆盖"),
      "撤销应被执行后指纹复核挡下"
    );

    expect(vi.mocked(recordExecutionDisposition)).not.toHaveBeenCalled();
    expect(host.writes()).toBe(writes);
    expect(host.cellValue()).toBe("我手动改的");
    expect(textOf(container)).not.toContain("已撤销并恢复");
  });

  it("放弃变更 → 服务端留痕记「放弃」，不是「撤销」", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");

    await click(findButton(container, "放弃变更"));
    await until(() => textOf(container).includes("已放弃并恢复"), "放弃后应显示已放弃并恢复");

    await until(
      () => vi.mocked(recordExecutionDisposition).mock.calls.length > 0,
      "放弃成功后必须上报处置留痕"
    );
    const [namespace, , disposition] = vi.mocked(recordExecutionDisposition).mock.calls[0];
    expect(namespace).toBe("wps");
    // 两条入口共用同一个恢复实现，区别**只**由 origin 给出：放弃不是撤销。
    expect(disposition).toBe("abandoned");
    expect(host.cellValue()).toBe("旧值");
  });

  it("确认变更不是处置：不写「放弃/撤销」留痕", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");

    await click(findButton(container, "确认变更"));
    await until(() => textOf(container).includes("已确认变更"), "确认应成功");
    expect(vi.mocked(recordExecutionDisposition)).not.toHaveBeenCalled();
    expect(textOf(container)).not.toContain("已放弃并恢复");
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
    await until(() => textOf(container).includes("已撤销并恢复"), "撤销后应给出核对过的恢复结果");
    // 区域已回到执行前，不能继续挂着过期的变更明细。
    expect(container.querySelector(".change-review-summary")).toBeNull();
    expect(host.cellValue()).toBe("旧值");
    // 但"这次确实改过、后来撤销了"是操作记录，不该跟着明细一起被抹掉；
    // 留下来的形态只有定位与计数，不含任何单元格内容。
    const note = container.querySelector(".change-review-undo-note");
    expect(note).not.toBeNull();
    expect(note?.textContent).toContain("Sheet1!A1");
    // 计数口径：只说**回读确认过**的改动数，不再报一个"改动 N / M"的分数
    // （分子可能混进"写后快照缺失"的假改动，分母是动作覆盖的总格数）。
    expect(note?.textContent).toContain("共 1 个单元格");
    expect(note?.textContent).toContain("经回读确认改动 1 个");
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
        () => Boolean(container.querySelector(".composer-box .clarification-composer")),
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

      await selectOption(container, 0);
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

      await selectOption(container, 0);
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

      await selectOption(container, 0);
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
      await selectOption(container, 0);
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
        () => findButton(container, "确认变更") !== null || textOf(container).includes("本次执行需要核对"),
        "新一轮计划应已执行并进入结果审阅"
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
      range.Item = () => range;
  const sheet = { Name: "Sheet1", Index: 1, UsedRange: range, ProtectContents: false, Range: () => range };
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

    it("执行阶段不建副本；只有点「双视图对比」才建**执行前**副本", async () => {
      // 双视图是可选的：用户没点，就不该为它付出任何代价（不复制工作表、
      // 不要求宿主具备复制接口）。
      writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      expect(stageWpsBeforeComparison).not.toHaveBeenCalled();
      expect(compareCommittedWpsResultToCopy).not.toHaveBeenCalled();

      await click(findButton(container, "双视图对比"));
      // 只走"按执行前快照重建"这一条。
      expect(stageWpsBeforeComparison).toHaveBeenCalledOnce();
      // 传给它的必须是**执行前快照**（含 prepared 与逐动作 before 值）。
      const passed = vi.mocked(stageWpsBeforeComparison).mock.calls[0]?.[1] as { values?: unknown[] };
      expect(Array.isArray(passed?.values)).toBe(true);
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
    // 新流程：计划就绪即执行。`staleWrite` 让回读与计划不符，执行器判为部分写入。
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
  /*
   * 「写回在途期间放弃」这条**界面路径**在新流程里不存在：执行不再由用户点击发起，
   * 授权复核窗口发生在"计划就绪之后、写入之前"，此时操作条只有"正在执行"，
   * 没有放弃/确认按钮。所以这里不再摆一个只能靠直接调用闭包才走得通的测试。
   *
   * 这道闸本身仍然在位，并且由**纯函数判据**直接覆盖（同一 describe 上方）：
   *   - `wpsWriteBackBlockReason`：没勾选/无副本/忙/编辑中一律拦；
   *   - `wpsWriteBackStillCurrent`：三维身份（审阅/计划/原工作簿）任一变化即拦；
   *   - `reviewWriteBackTargetChanged`：任一侧缺席 fail-closed。
   * 它们才是"旧请求不许借新现场写表"的实际判据；界面层只是调用它们的地方。
   */
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
    expect(findButton(container, "继续分析")).toBeNull();
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
const host = writeRangeHost(); const { container } = await mountPlanPreview(singleWritePlan, '把 A1 改成 90'); await click(findButton(container,'放弃变更')); expect(host.cellValue()).toBe('旧值'); expect(textOf(container)).toContain('已放弃并恢复'); expect(findButton(container,'继续分析')).toBeNull(); expect(findButton(container,'确认变更')).toBeNull();
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
      // 执行在计划就绪时自动开始：恢复标记写不进存储，就必须在动原表之前停下。
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

  it("写入回执中断后不依赖窗格重挂载，立即在最新消息之后显示恢复入口", async () => {
    const host = writeRangeHost();
    completeWpsActionMock.mockRejectedValueOnce(new Error("lost action receipt"));
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(() => container.querySelector(".wps-write-recovery") !== null,
      "保留原页面的 WPS 窗格也必须发现刚中断的任务");
    const card = container.querySelector(".wps-write-recovery")!;
    const rows = container.querySelectorAll(".message-row");
    expect(rows.length).toBeGreaterThan(0);
    expect(rows[rows.length - 1].compareDocumentPosition(card) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(suspendWpsExecutionMock).toHaveBeenCalledTimes(1);
    expect(completeWpsActionMock).toHaveBeenCalledTimes(1);
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();
    expect(host.cellValue()).toBe("90");
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
      () => container.querySelector(".wps-write-recovery") !== null ||
        container.querySelector("[data-recovery-terminal]") !== null,
      "重载后应出现恢复核对卡或终态摘要"
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

  it.each(["确认变更", "放弃变更"])("重开后点击真实续跑入口，仅写剩余动作，随后%s不重放计划", async outcome => {
    const f = resumeHost();
    await expect(executeWpsPlan(f.application, f.prepared, undefined, undefined, async (phase, c) => {
      if (phase === "after" && c.actionIndex === 1) throw new Error("lost receipt");
    }, resumeOperationId, {taskId: "task-1", ...recoveryChannel})).rejects.toThrow("lost receipt");
    f.writes.length = 0;
    getWpsApplicationMock.mockReturnValue(f.application);
    localStorage.setItem(PENDING_KEY, JSON.stringify({schemaVersion:1, operationId:resumeOperationId,
      conversationId:"conv-1",activityMessageId:"old-message",workbookIdentity:f.state.workbookIdentity,
      taskId:f.state.taskId,createdAt:new Date().toISOString()}));
    getWpsExecutionStatusMock.mockResolvedValue(f.state);
    let owned = {...f.state, ownerToken:"new-owner", ownerActive:true, revision:6};
    claimWpsExecutionResumeMock.mockImplementation(async () => owned);
    startWpsActionMock.mockImplementation(async (_id, request) => {
      owned={...owned,inFlightActionIndex:request.actionIndex,revision:owned.revision+1};return owned;
    });
    completeWpsActionMock.mockImplementation(async (_id, request) => {
      owned={...owned,inFlightActionIndex:null,nextActionIndex:request.actionIndex+1,revision:owned.revision+1};return owned;
    });
    finishWpsExecutionMock.mockImplementation(async () => ({...owned,status:"completed" as const}));
    listModelsMock.mockResolvedValue({defaultModelId:firstConnection.id,models:[firstConnection]});
    const {container}=await mount(<WpsAgentPanel selectedSheetIndexes={[1,2]} scopeDock={<div>固定两表</div>}/>);
    await until(()=>findButton(container,"继续剩余动作") !== null,"恢复卡必须接入生产续跑判据");
    await click(findButton(container,"继续剩余动作"));
    await until(()=>findButton(container,"确认变更") !== null,"续跑后必须进入审阅");
    expect(f.writes).toEqual(["Target!A3"]);
    expect(startWpsActionMock).toHaveBeenCalledWith(resumeOperationId,expect.objectContaining({actionIndex:2}));
    expect(runTurnMock).not.toHaveBeenCalled();
    expect(registerWpsExecutionMock).not.toHaveBeenCalled();
    await click(findButton(container,outcome));
    await until(()=>container.querySelector(".result-review-bar")===null,"审阅应收尾");
    if(outcome === "确认变更") {
      expect(f.writes).toEqual(["Target!A3"]);
      expect([1,2,3].map(i=>f.cells.get(`Target!A${i}`))).toEqual([10,20,30]);
    } else {
      expect([1,2,3].map(i=>f.cells.get(`Target!A${i}`))).toEqual(["old1","old2","old3"]);
    }
    expect(localStorage.getItem(PENDING_KEY)).toBeNull();
  });

  /*
   * 真实缺陷回归（2026-10-27 真机）：**完整退出 WPS 再打开**，`PluginStorage` 里读不到
   * 原执行记录 —— 旧实现据此就说"找不到原执行的本机恢复记录"，于"已核对第二步、仍可继续
   * 第三步"断开。现在材料在本机后端，宿主缓存丢了必须能补回来，且仍然只写第三步。
   *
   * 这里模拟的正是那条路：`storage.clear()` + 换一个新的应用对象；后端那一份不动。
   */
  it.each(["确认变更", "放弃变更"])("完整退出 WPS 清空宿主缓存后重开，从本机恢复包续写第三步，随后%s不重放", async outcome => {
    const f = resumeHost();
    await expect(executeWpsPlan(f.application, f.prepared, undefined, undefined, async (phase, c) => {
      if (phase === "after" && c.actionIndex === 1) throw new Error("lost receipt");
    }, resumeOperationId, {taskId: "task-1", ...recoveryChannel})).rejects.toThrow("lost receipt");
    // WPS 整个进程退出：宿主缓存没了，换一个新宿主对象；后端材料留下（另一个进程）。
    f.storage.clear();
    const reopened = {...f.application};
    expect(backendRecovery().ids()).toContain(resumeOperationId);
    f.writes.length = 0;
    getWpsApplicationMock.mockReturnValue(reopened);
    localStorage.setItem(PENDING_KEY, JSON.stringify({schemaVersion:1, operationId:resumeOperationId,
      conversationId:"conv-1",activityMessageId:"old-message",workbookIdentity:f.state.workbookIdentity,
      taskId:f.state.taskId,createdAt:new Date().toISOString()}));
    getWpsExecutionStatusMock.mockResolvedValue(f.state);
    let owned = {...f.state, ownerToken:"new-owner", ownerActive:true, revision:6};
    claimWpsExecutionResumeMock.mockImplementation(async () => owned);
    startWpsActionMock.mockImplementation(async (_id, request) => {
      owned={...owned,inFlightActionIndex:request.actionIndex,revision:owned.revision+1};return owned;
    });
    completeWpsActionMock.mockImplementation(async (_id, request) => {
      owned={...owned,inFlightActionIndex:null,nextActionIndex:request.actionIndex+1,revision:owned.revision+1};return owned;
    });
    finishWpsExecutionMock.mockImplementation(async () => ({...owned,status:"completed" as const}));
    listModelsMock.mockResolvedValue({defaultModelId:firstConnection.id,models:[firstConnection]});
    const {container}=await mount(<WpsAgentPanel selectedSheetIndexes={[1,2]} scopeDock={<div>固定两表</div>}/>);
    // 宿主缓存是空的，但恢复材料在盘上 ⇒ 必须给出"继续剩余动作"，不许报"找不到恢复记录"。
    await until(()=>findButton(container,"继续剩余动作") !== null,"完整退出后必须能从本机恢复包续跑");
    await click(findButton(container,"继续剩余动作"));
    await until(()=>findButton(container,"确认变更") !== null,"续跑后必须进入审阅");
    expect(f.writes).toEqual(["Target!A3"]);
    // 前两步不许被重写：账本只推进第三步，宿主也只被写了这一格。
    expect(startWpsActionMock.mock.calls.map(call => call[1].actionIndex)).toEqual([2]);
    expect(registerWpsExecutionMock).not.toHaveBeenCalled();
    expect(runTurnMock).not.toHaveBeenCalled();
    await click(findButton(container,outcome));
    await until(()=>container.querySelector(".result-review-bar")===null,"审阅应收尾");
    if(outcome === "确认变更") {
      expect(f.writes).toEqual(["Target!A3"]);
      expect([1,2,3].map(i=>f.cells.get(`Target!A${i}`))).toEqual([10,20,30]);
      // 确认之后仍然可以撤销 ⇒ 材料必须还在（不许在执行/确认成功时就删）。
      expect(backendRecovery().ids()).toContain(resumeOperationId);
    } else {
      expect([1,2,3].map(i=>f.cells.get(`Target!A${i}`))).toEqual(["old1","old2","old3"]);
      // 原表已回到执行前 ⇒ 终态回收，材料不再需要。
      await until(()=>backendRecovery().ids().length === 0,"撤销成功后应回收本机恢复材料");
    }
    expect(localStorage.getItem(PENDING_KEY)).toBeNull();
  });

  /*
   * ==========================================================================================
   * 2026-10-27 真机缺口：**完整退出后没有"重新发现未完成执行"的入口**
   * ==========================================================================================
   *
   * 现场只读取证：账本 `needs_reconciliation / nextActionIndex=1 / inFlightActionIndex=1`、
   * 原表仍是 111/222/-303、授权还是两张表 —— 进度与现场都在，缺的只是入口。旧实现把浏览器
   * `localStorage` 的单条索引当成**唯一**入口，完整退出后它不在了，于是"待核对在途动作"再也
   * 摆不出来，只剩一条历史任务的"已撤销"。
   *
   * 下面这组用例的前提是：浏览器索引**真的**没有了（不手工塞 operationId），面板必须经后端
   * 按当前工作簿稳定身份把它找回来。上面那条"清宿主缓存但保留索引"的写法恰好绕过了这个缺口，
   * 所以不能顶替这一组。
   */

  /** 一次"第二步已写、回执丢了"的真实执行：宿主值、账本进度、持久材料都按真机来。 */
  async function interruptedAfterSecondWrite(options: { extraSheet?: boolean } = {}) {
    const f = resumeHost(options);
    await expect(executeWpsPlan(f.application, f.prepared, undefined, undefined, async (phase, c) => {
      if (phase === "after" && c.actionIndex === 1) throw new Error("lost receipt");
    }, resumeOperationId, {taskId: "task-1", ...recoveryChannel})).rejects.toThrow("lost receipt");
    return f;
  }

  /** 账本侧的真机形状：第二步在途、现场已写、执行权没人持有。 */
  function inFlightJob(f: ReturnType<typeof resumeHost>, overrides: Record<string, unknown> = {}) {
    return {
      operationId: resumeOperationId,
      taskId: f.state.taskId,
      workbookIdentity: f.state.workbookIdentity,
      status: "needs_reconciliation" as const,
      nextActionIndex: 1,
      totalActions: 3,
      inFlightActionIndex: 1,
      ownerActive: false,
      requiresReconciliation: true,
      canResume: true,
      ...overrides
    };
  }

  /** 服务端状态：与账本同形（在途动作 = 第二步）。 */
  function inFlightState(f: ReturnType<typeof resumeHost>, overrides: Record<string, unknown> = {}) {
    return {...f.state, status: "needs_reconciliation" as const, nextActionIndex: 1,
      inFlightActionIndex: 1, requiresReconciliation: true, ...overrides};
  }

  /** 挂载面板并等模型目录就绪（恢复卡由各用例自己等）。 */
  async function mountPanel() {
    listModelsMock.mockResolvedValue({defaultModelId: firstConnection.id, models: [firstConnection]});
    return mount(<WpsAgentPanel selectedSheetIndexes={[1, 2]} scopeDock={<div>固定两表</div>}/>);
  }

  it("完整退出 WPS 丢掉浏览器索引与宿主缓存后，面板经后端发现找回在途执行，且零自动动作", async () => {
    const f = await interruptedAfterSecondWrite();
    // 完整退出：宿主缓存与浏览器索引都没了。**不**手工塞 operationId —— 那等于跳过被测路径。
    f.storage.clear();
    localStorage.removeItem(PENDING_KEY);
    expect(localStorage.getItem(PENDING_KEY)).toBeNull();
    f.writes.length = 0;
    // 账本与恢复材料留在"另一个进程"里，一条不动。
    const jobs = pendingDiscoveryBackend();
    jobs.publish(inFlightJob(f));
    const ledgerBefore = jobs.snapshot(resumeOperationId);
    expect(backendRecovery().ids()).toContain(resumeOperationId);
    getWpsApplicationMock.mockReturnValue({...f.application});
    getWpsExecutionStatusMock.mockResolvedValue(inFlightState(f) as never);

    const {container} = await mountPanel();
    await until(() => textOf(container).includes("待核对在途动作"),
      "索引丢了也必须经后端发现重新摆出待核对入口");

    // 发现用的是**当前工作簿的稳定身份**，不是文件名、不是"最近一条"。
    expect(listPendingWpsMock).toHaveBeenCalledWith("wps-path:c:\\data\\test.xlsx");
    // 只读：一个自动动作都没有（认领、补记、记账、写入全为 0）。
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();
    expect(startWpsActionMock).not.toHaveBeenCalled();
    expect(completeWpsActionMock).not.toHaveBeenCalled();
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
    expect(suspendWpsExecutionMock).not.toHaveBeenCalled();
    // 核对入口已搬到输入框右下角（卡面只剩次要动作）；"只读、零写入"的含义不变。
    expect(recoveryButtons(container)).toEqual(["清除这条恢复提示"]);
    const reconcileEntry = findButton(container, "核对在途动作");
    expect(reconcileEntry, "核对入口必须存在").not.toBeNull();
    // 唯一主入口：它就是输入框右下角那颗主按钮，不是卡面上另画的一颗。
    expect(reconcileEntry!.closest(".composer")).not.toBeNull();
    expect(reconcileEntry!.classList.contains("composer-primary")).toBe(true);
    expect(container.querySelectorAll('[aria-label="核对在途动作"]')).toHaveLength(1);
    /*
     * 有内容 ⇒ 主按钮变「发送」，但**必须说明为什么核对入口不在了**（不许静默顶掉）。
     * 两个方向都断言：有字时给说明、清空后入口回来。
     */
    const composerTextarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(composerTextarea, "再补一列金额");
    await until(() => findButton(container, "发送") !== null, "有字时主按钮应变成发送");
    expect(textOf(container)).toContain("输入框有内容：发送会作为新需求");
    expect(textOf(container)).toContain("请先清空输入框");
    await type(composerTextarea, "");
    await until(() => findButton(container, "核对在途动作") !== null, "清空后核对入口应回来");
    /*
     * 只贴图片、文字为空：`hasInput` 把待发图片也算内容，同一个入口同样被 `send`
     * 顶掉 —— 提示必须说明"还得移除图片"，否则用户清空文字之后仍然点不到核对，
     * 却看不出为什么。断言看**主按钮的实际状态**（可访问名 + 它是输入框里那颗
     * `.composer-primary`），不只看文案。
     */
    await attachFile(
      container.querySelector<HTMLInputElement>('input[type="file"]'),
      new File([new Uint8Array([1, 2, 3])], "u.png", { type: "image/png" })
    );
    await until(() => container.querySelector(".composer-attachments") !== null, "附件应出现在输入区");
    await until(() => findButton(container, "发送") !== null, "只贴图时主按钮也应变发送");
    const imageOnlyPrimary = container.querySelector<HTMLElement>(".composer .composer-primary");
    expect(imageOnlyPrimary!.getAttribute("aria-label")).toBe("发送");
    expect(textOf(container)).toContain("输入框有内容：发送会作为新需求");
    expect(textOf(container)).toContain("并移除待发图片");
    // 移除图片（文字仍为空）⇒ 核对入口回来，且是同一颗主按钮，不是卡面新画一颗。
    await click(container.querySelector('[aria-label="移除图片 u.png"]'));
    await until(() => findButton(container, "核对在途动作") !== null, "移除图片后核对入口应回来");
    expect(container.querySelector(".composer .composer-primary")!.getAttribute("aria-label"))
      .toBe("核对在途动作");
    expect(container.querySelectorAll('[aria-label="核对在途动作"]')).toHaveLength(1);
    // 贴图/移除全程只动内存：宿主一格没写，账本一条没动。
    expect(f.writes).toEqual([]);
    // 宿主一格没写，账本一条没动（发现不得推进 revision 或写处置元数据）。
    expect(f.writes).toEqual([]);
    expect(jobs.snapshot(resumeOperationId)).toEqual(ledgerBefore);
    // 卡面不含单元格内容。
    expect(textOf(container)).not.toContain("old1");
  });

  /*
   * 终态（撤销确认 / 放弃 / 收尾）退出固定区域：恢复卡已经没有任何动作，不该再占
   * 输入区上方的主视图或"还有 N 项"；结论、进度、边界说明改放进**所属任务的消息**里，
   * 一条可展开的摘要 —— 记录一条不删，账本与恢复材料不动。
   *
   * 反向边界由 core 判据钉住（`executionRecoveryView.test.ts`「未知、阻断、证据不足…不是终态」），
   * 这里只验"两端接线后，界面上确实退出、且摘要落在消息里"。
   */
  it("旧任务已撤销 → 发普通聊天 → 固定区域不再显示旧恢复卡，只剩消息里的可展开终态摘要", async () => {
    const f = await interruptedAfterSecondWrite();
    // 中断已经写了两格；清空计数，只观察"终态摆卡 / 发聊天"这段**不再新增写入**。
    f.writes.length = 0;
    const now = new Date().toISOString();
    const terminalState = {
      ...inFlightState(f),
      disposition: { undoneAt: Date.now(), abandonedAt: null }
    };
    // 历史里有一条**属于这次执行**的写入结果卡（executionId 对齐）。
    localStorage.setItem("excel-bro.wps.chat.v1", JSON.stringify({
      activeConversationId: "conv-1",
      conversations: [{
        id: "conv-1", title: "旧任务", createdAt: now, updatedAt: now,
        messages: [{
          id: "q", role: "user", text: "旧任务", createdAt: now
        }, {
          id: "m-ver", role: "assistant", text: "", createdAt: now,
          verification: {
            executionId: resumeOperationId,
            actionCount: 3,
            cellCount: 3,
            status: "undone",
            outcome: "readback_passed",
            targets: [
              { type: "writeValues", sheetName: "Target", range: "A1:A1" },
              { type: "writeValues", sheetName: "Target", range: "A2:A2" },
              { type: "writeValues", sheetName: "Target", range: "A3:A3" }
            ]
          }
        }]
      }]
    }));
    // 后端发现仍会按当前工作簿身份把它找回来，但结论是终态（已撤销）。
    const jobs = pendingDiscoveryBackend();
    jobs.publish(inFlightJob(f));
    getWpsApplicationMock.mockReturnValue({...f.application});
    getWpsExecutionStatusMock.mockResolvedValue(terminalState as never);

    const {container} = await mountPanel();
    // 终态摘要出现在**消息里**（可展开），而不是固定区域那张可点的恢复卡。
    await until(() => container.querySelector('[data-recovery-terminal="undone"]') !== null,
      "已撤销后消息里应留下可展开的终态摘要");
    const summary = container.querySelector('[data-recovery-terminal="undone"]');
    expect(summary!.textContent).toContain("已撤销");
    expect(summary!.textContent).toContain("展开");

    // 固定区域**没有**可点的恢复卡：主视图不是它，主按钮也不承接核对/继续。
    expect(container.querySelector(".wps-write-recovery:not(.is-terminal)")).toBeNull();
    expect(container.querySelector('[data-recovery-primary-hint]')).toBeNull();
    expect(findButton(container, "核对在途动作")).toBeNull();
    expect(findButton(container, "继续剩余动作")).toBeNull();

    // 发一句普通聊天：固定区域**仍然**不显示旧恢复卡（终态不是待办）。
    runTurnMock.mockResolvedValue({
      kind: "answer",
      message: "这是一句普通回复。",
      provider: "model",
      resultContext: null
    });
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "随便聊聊");
    await submit(textarea);
    await until(() => textOf(container).includes("这是一句普通回复"), "普通回复应显示");
    // 旧卡仍在（消息里、可展开），但没有被重新摆回固定区域。
    expect(container.querySelectorAll('[data-recovery-terminal="undone"]').length).toBeGreaterThan(0);
    expect(container.querySelector(".wps-write-recovery:not(.is-terminal)")).toBeNull();
    expect(findButton(container, "继续剩余动作")).toBeNull();
    // 一条账本、一个恢复材料都不许删：发现是只读的，终态只是"退出固定区域"。
    expect(f.writes).toEqual([]);
    expect(backendRecovery().ids()).toContain(resumeOperationId);
  });

  it.each(["确认变更", "放弃变更"])("经发现重新找到后：显式核对只补记、显式继续只写第三步，随后%s不重放", async outcome => {
    const f = await interruptedAfterSecondWrite();
    f.storage.clear();
    localStorage.removeItem(PENDING_KEY);
    f.writes.length = 0;
    const jobs = pendingDiscoveryBackend();
    jobs.publish(inFlightJob(f));
    getWpsApplicationMock.mockReturnValue({...f.application});
    /*
     * 账本游标按真实推进走：补记只把第二步记为已写入（不碰宿主），
     * 续跑只认领第三步。用同一个可变状态串起来，避免各 mock 各说一套。
     */
    let current: Record<string, unknown> = inFlightState(f);
    getWpsExecutionStatusMock.mockImplementation(async () => current as never);
    reconcileWpsActionReceiptMock.mockImplementation(async () => {
      current = {...current, status: "running", nextActionIndex: 2, inFlightActionIndex: null,
        requiresReconciliation: false, revision: 8, ownerToken: "reconciled-owner"};
      return current as never;
    });
    suspendWpsExecutionMock.mockImplementation(async () => {
      current = {...current, ownerToken: null, ownerActive: false};
      return current as never;
    });
    let owned: Record<string, unknown> = {...current};
    claimWpsExecutionResumeMock.mockImplementation(async () => {
      owned = {...current, ownerToken: "owner-2", ownerActive: true, revision: 9};
      return owned as never;
    });
    startWpsActionMock.mockImplementation(async (_id, request) => {
      owned = {...owned, inFlightActionIndex: request.actionIndex, revision: Number(owned.revision) + 1};
      current = owned;
      return owned as never;
    });
    completeWpsActionMock.mockImplementation(async (_id, request) => {
      owned = {...owned, inFlightActionIndex: null, nextActionIndex: request.actionIndex + 1,
        revision: Number(owned.revision) + 1};
      current = owned;
      return owned as never;
    });
    finishWpsExecutionMock.mockImplementation(async () => {
      current = {...owned, status: "completed"};
      return current as never;
    });

    const {container} = await mountPanel();
    await until(() => findButton(container, "核对在途动作") !== null, "必须先摆出待核对入口");
    // 摆卡本身不许补记。
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();

    await click(findButton(container, "核对在途动作"));
    await until(() => reconcileWpsActionReceiptMock.mock.calls.length === 1, "点击后应补记在途动作");
    expect(reconcileWpsActionReceiptMock).toHaveBeenCalledWith(resumeOperationId, expect.objectContaining({
      actionIndex: 1, outcome: "applied"
    }));
    // 补记只改账本：原表一格都没再写。
    expect(f.writes).toEqual([]);

    await until(() => findButton(container, "继续剩余动作") !== null, "补记后应可继续剩余动作");
    await click(findButton(container, "继续剩余动作"));
    await until(() => findButton(container, "确认变更") !== null, "续跑后必须进入审阅");
    // 只写第三步；前两步没有重放，也没有重新登记或重新规划。
    expect(f.writes).toEqual(["Target!A3"]);
    expect(startWpsActionMock.mock.calls.map(call => call[1].actionIndex)).toEqual([2]);
    expect(registerWpsExecutionMock).not.toHaveBeenCalled();
    expect(runTurnMock).not.toHaveBeenCalled();

    await click(findButton(container, outcome));
    await until(() => container.querySelector(".result-review-bar") === null, "审阅应收尾");
    if (outcome === "确认变更") {
      // 确认不再执行一次：写入次数仍是 1。
      expect(f.writes).toEqual(["Target!A3"]);
      expect([1, 2, 3].map(i => f.cells.get(`Target!A${i}`))).toEqual([10, 20, 30]);
    } else {
      // 放弃把整份写前状态恢复回来 —— 包括前两步。
      expect([1, 2, 3].map(i => f.cells.get(`Target!A${i}`))).toEqual(["old1", "old2", "old3"]);
    }
  });

  it("同名不同路径的工作簿互不发现（只按身份完整相等）", async () => {
    const f = await interruptedAfterSecondWrite();
    f.storage.clear();
    localStorage.removeItem(PENDING_KEY);
    f.writes.length = 0;
    // 账本里有一条**同名、不同目录**的未完成执行。
    pendingDiscoveryBackend().publish(inFlightJob(f, {
      operationId: "wps-resume-other-path",
      taskId: "task-other",
      workbookIdentity: "wps-path:d:\\other\\test.xlsx"
    }));
    getWpsApplicationMock.mockReturnValue({...f.application});
    getWpsExecutionStatusMock.mockResolvedValue(inFlightState(f) as never);

    const {container} = await mountPanel();
    await until(() => listPendingWpsMock.mock.calls.length > 0, "应当按当前身份查一次");
    expect(listPendingWpsMock.mock.calls.every((call) =>
      call[0] === "wps-path:c:\\data\\test.xlsx")).toBe(true);
    // 那个簿的记录一条都不许被显示，也不许顺手查它的状态。
    expect(container.querySelector(".wps-write-recovery")).toBeNull();
    expect(getWpsExecutionStatusMock).not.toHaveBeenCalled();
    expect(f.writes).toEqual([]);
  });

  it("先打开另一个工作簿再切回原簿：靠身份观察重跑发现，仍能找到在途执行", async () => {
    const f = await interruptedAfterSecondWrite();
    f.storage.clear();
    localStorage.removeItem(PENDING_KEY);
    f.writes.length = 0;
    pendingDiscoveryBackend().publish(inFlightJob(f));
    const other = {...f.application,
      ActiveWorkbook: {...f.application.ActiveWorkbook!, FullName: "C:\\Data\\other.xlsx"}};
    getWpsApplicationMock.mockReturnValue(other);
    getWpsExecutionStatusMock.mockResolvedValue(inFlightState(f) as never);

    const {container} = await mountPanel();
    await until(() => listPendingWpsMock.mock.calls.some((call) =>
      call[0] === "wps-path:c:\\data\\other.xlsx"), "另一个簿也必须按自己的身份查");
    expect(container.querySelector(".wps-write-recovery")).toBeNull();

    // 切回原簿：宿主没有"文档已切换"事件，靠焦点/可见性/间隔的身份观察重跑发现。
    getWpsApplicationMock.mockReturnValue({...f.application});
    await act(async () => { window.dispatchEvent(new Event("focus")); });
    await until(() => textOf(container).includes("待核对在途动作"),
      "切回原簿后必须重新发现，而不是停在\"没有未完成任务\"");
    expect(f.writes).toEqual([]);
  });

  it("工作簿换了（选表序号相同）必须重新核查：旧簿的在途卡不许留在新簿上", async () => {
    const f = await interruptedAfterSecondWrite();
    f.storage.clear();
    localStorage.removeItem(PENDING_KEY);
    f.writes.length = 0;
    pendingDiscoveryBackend().publish(inFlightJob(f));
    getWpsApplicationMock.mockReturnValue({...f.application});
    getWpsExecutionStatusMock.mockResolvedValue(inFlightState(f) as never);

    const {container} = await mountPanel();
    await until(() => textOf(container).includes("待核对在途动作"), "原簿应能发现");

    // 换成另一本簿：选表序号仍是 [1,2]（同一个序号在另一个簿上指向另一张表）。
    getWpsApplicationMock.mockReturnValue({...f.application,
      ActiveWorkbook: {...f.application.ActiveWorkbook!, FullName: "C:\\Data\\switched.xlsx"}});
    await act(async () => { window.dispatchEvent(new Event("focus")); });
    await until(() => container.querySelector(".wps-write-recovery") === null,
      "换簿后旧簿的恢复卡必须撤掉，不能拿它当当前簿的现场");
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();
    expect(f.writes).toEqual([]);
  });

  it("身份读不出来（未保存）时一个请求都不发，也不说成\"没有未完成任务\"；保存后才发现", async () => {
    const f = await interruptedAfterSecondWrite();
    f.storage.clear();
    localStorage.removeItem(PENDING_KEY);
    f.writes.length = 0;
    pendingDiscoveryBackend().publish(inFlightJob(f));
    getWpsApplicationMock.mockReturnValue({...f.application,
      ActiveWorkbook: {...f.application.ActiveWorkbook!, FullName: "工作簿1"}});
    getWpsExecutionStatusMock.mockResolvedValue(inFlightState(f) as never);

    const {container} = await mountPanel();
    await settle();
    await settle();
    // 未保存簿没有稳定身份 ⇒ 不许退回文件名猜，一个发现请求都不发。
    expect(listPendingWpsMock).not.toHaveBeenCalled();
    expect(container.querySelector(".wps-write-recovery")).toBeNull();
    expect(textOf(container)).not.toContain("没有未完成");

    // 用户保存成 .xlsx：身份变得可读，发现必须自己再跑一次（元数据晚到不是"永远没有"）。
    getWpsApplicationMock.mockReturnValue({...f.application});
    await act(async () => { window.dispatchEvent(new Event("focus")); });
    await until(() => textOf(container).includes("待核对在途动作"), "身份可读后应发现未完成执行");
    expect(listPendingWpsMock).toHaveBeenCalledWith("wps-path:c:\\data\\test.xlsx");
  });

  it("旧查询迟到不覆盖新工作簿：先前那次发现的结果不许画到当前簿上", async () => {
    const f = await interruptedAfterSecondWrite();
    f.storage.clear();
    localStorage.removeItem(PENDING_KEY);
    f.writes.length = 0;
    pendingDiscoveryBackend().publish(inFlightJob(f));
    const jobs = pendingDiscoveryBackend();
    // 扣住原簿那一次查询的响应，模拟"网络慢、用户已经切走了"。
    jobs.holdNextList();
    getWpsApplicationMock.mockReturnValue({...f.application});
    getWpsExecutionStatusMock.mockResolvedValue(inFlightState(f) as never);

    const {container} = await mountPanel();
    await until(() => listPendingWpsMock.mock.calls.length > 0, "原簿的查询应已发出");

    // 切到另一本簿，再让原簿那次迟到响应回来。
    getWpsApplicationMock.mockReturnValue({...f.application,
      ActiveWorkbook: {...f.application.ActiveWorkbook!, FullName: "C:\\Data\\switched.xlsx"}});
    await act(async () => { window.dispatchEvent(new Event("focus")); });
    const released = jobs.releaseHeldLists();
    expect(released).toBeGreaterThan(0);
    await settle();
    await settle();

    // 判别点：迟到的旧响应**不许**画到新簿上 —— 连状态都不许为它再查一次。
    expect(container.querySelector(".wps-write-recovery"),
      "换到新簿后迟到的旧响应不许画出恢复卡").toBeNull();
    expect(textOf(container)).not.toContain("待核对在途动作");
    expect(getWpsExecutionStatusMock).not.toHaveBeenCalled();
    expect(f.writes).toEqual([]);
  });

  it("工作簿有多条未完成写入时只给选择入口：不自动挑最新、也不自动查状态或写入", async () => {
    const f = await interruptedAfterSecondWrite();
    f.storage.clear();
    localStorage.removeItem(PENDING_KEY);
    f.writes.length = 0;
    const jobs = pendingDiscoveryBackend();
    jobs.publish(inFlightJob(f));
    // 第二条：一个动作都没推进的写前失败，且更新时间更新 —— 它**不该**被当成"最新的"直接选中。
    jobs.publish({
      operationId: "wps-resume-0002",
      taskId: "task-2",
      workbookIdentity: f.state.workbookIdentity,
      status: "failed",
      nextActionIndex: 0,
      totalActions: 2,
      inFlightActionIndex: null,
      requiresReconciliation: false,
      canResume: true,
      updatedAt: Date.now() + 5_000
    });
    getWpsApplicationMock.mockReturnValue({...f.application});
    getWpsExecutionStatusMock.mockResolvedValue(inFlightState(f) as never);

    const {container} = await mountPanel();
    await until(() => container.querySelectorAll(".wps-recovery-candidate").length === 2,
      "两条候选都必须摆出来给用户选");
    expect(textOf(container)).toContain("有多条未完成的写入");
    // 排序把"在途动作"放最前，写前失败放最后 —— 但一条都不自动采用。
    const labels = Array.from(container.querySelectorAll(".wps-recovery-candidate"))
      .map(item => (item.textContent ?? "").trim());
    expect(labels[0]).toContain("第 2 步待核对");
    expect(labels[1]).toContain("写前失败");
    expect(getWpsExecutionStatusMock).not.toHaveBeenCalled();
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();
    expect(f.writes).toEqual([]);
    // 候选行不带单元格内容。
    expect(labels.join("|")).not.toContain("old1");
    // 另一条任务不许被顺手清掉。
    expect(jobs.entries().map(entry => entry.operationId).sort()).toEqual(
      [resumeOperationId, "wps-resume-0002"].sort());
  });

  /*
   * ==========================================================================================
   * 2026-10-28 真机阻断：**用户选的是"在途"那条，界面给出的却是另一条的"需重新规划"**
   * ==========================================================================================
   *
   * 只读取证（隔离实例复刻 + 8765 只读查询，两边同一份账本）：
   *   - `/api/wps/executions/<0f5bffdc…>`：`needs_reconciliation / next=1 / inFlight=1 /
   *     total=3 / canResume / ownerActive=false`；
   *   - `/api/execution-recovery/<0f5bffdc…>`：材料在盘上（1726 字节，摘要一致），到期 2026-11-03；
   *   - 包与账本逐字段相同：计划指纹 `plan:v1:3:fccc0533`、三个动作的写前/写后指纹全部重算一致。
   * 这份材料只可能得出"待核对在途动作"。
   *
   * 而界面显示"还有 **1** 个动作没写"——`total - next = 1` ⇒ `next=2`。账本里这个工作簿
   * 唯一那条 `next=2 / total=3` 的是**另一条**失败执行（`de702e1b…`，无材料、无宿主缓存）。
   * 所以卡片属于另一条执行，不是用户选的那条。
   *
   * 机制就在下面两条用例里：浏览器里残留的旧索引先被"缓存优先"采纳，而"需重新规划"被当成
   * **能做事**，于是后端发现被拦在门外 —— 真正的在途执行永远不出现，用户无论怎么选看到
   * 的都是那张 replan 卡（"找不到原执行的本机恢复记录"对它自己是真的，对用户的选择不是）。
   */

  it("浏览器里留着一条\"需重新规划\"的旧索引时，后端发现的在途执行仍必须摆出来", async () => {
    const f = await interruptedAfterSecondWrite();
    f.storage.clear();                        // 完整退出：宿主缓存没了
    const jobs = pendingDiscoveryBackend();
    jobs.publish(inFlightJob(f));             // 账本里真正的在途执行（第 2 步待核对）
    /*
     * 旧索引还在，但指向**另一条**执行：同一工作簿、第三个动作从没写过、材料早就没有。
     * 它会得出 `replan + remaining=1 + 找不到原执行的本机恢复记录` —— 与真机界面逐字相同。
     */
    localStorage.setItem(PENDING_KEY, JSON.stringify({
      schemaVersion: 1, operationId: "wps-resume-stale", conversationId: "conv-1",
      activityMessageId: "old-message", workbookIdentity: f.state.workbookIdentity,
      taskId: "task-stale", createdAt: new Date().toISOString()
    }));
    getWpsApplicationMock.mockReturnValue({...f.application});
    const stale = {...f.state, operationId: "wps-resume-stale", taskId: "task-stale",
      status: "failed" as const, nextActionIndex: 2, inFlightActionIndex: null,
      requiresReconciliation: false};
    getWpsExecutionStatusMock.mockImplementation(async (operationId: string) =>
      (operationId === "wps-resume-stale" ? stale : inFlightState(f)) as never);
    f.writes.length = 0;

    const {container} = await mountPanel();
    await until(() => textOf(container).includes("待核对在途动作"),
      "旧索引判成\"需重新规划\"不等于这个工作簿没有在途执行");
    // 摆出来的是**在途**那条：真机现场就在这里，不该被一条无材料的旧键挡住。
    expect(getWpsExecutionStatusMock).toHaveBeenCalledWith(resumeOperationId);
    expect(textOf(container)).not.toContain("需重新规划");
    // 核对入口已搬到输入框右下角（卡面只剩次要动作）；"只读、零写入"的含义不变。
    expect(recoveryButtons(container)).toEqual(["清除这条恢复提示"]);
    const reconcileEntry = findButton(container, "核对在途动作");
    expect(reconcileEntry, "核对入口必须存在").not.toBeNull();
    // 唯一主入口：它就是输入框右下角那颗主按钮，不是卡面上另画的一颗。
    expect(reconcileEntry!.closest(".composer")).not.toBeNull();
    expect(reconcileEntry!.classList.contains("composer-primary")).toBe(true);
    expect(container.querySelectorAll('[aria-label="核对在途动作"]')).toHaveLength(1);
    // 仍旧只读：不认领、不补记、不写宿主。
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();
    expect(f.writes).toEqual([]);
  });

  it("宿主身份读不到时不许采纳浏览器索引：不猜是哪个工作簿", async () => {
    const f = await interruptedAfterSecondWrite();
    f.storage.clear();
    localStorage.setItem(PENDING_KEY, JSON.stringify({
      schemaVersion: 1, operationId: resumeOperationId, conversationId: "conv-1",
      activityMessageId: "msg-1", workbookIdentity: f.state.workbookIdentity,
      taskId: f.state.taskId, createdAt: new Date().toISOString()
    }));
    // 未保存 / 宿主还没就绪：`FullName` 读不到稳定身份。
    getWpsApplicationMock.mockReturnValue({
      ...f.application,
      ActiveWorkbook: {...f.application.ActiveWorkbook!, FullName: ""}
    });
    getWpsExecutionStatusMock.mockResolvedValue(inFlightState(f) as never);

    const {container} = await mountPanel();
    await settle();
    await settle();
    /*
     * 身份是"这个索引属于哪个工作簿"的唯一依据。读不到就不采纳 —— 否则同名异路径的
     * 旧键会在身份读不出来的那一瞬间被当成当前簿的记录（`downloads\wps恢复验收.xlsx`
     * 与 `.codex\…\wps恢复验收.xlsx` 正是这种一对）。
     */
    expect(container.querySelector(".wps-write-recovery"),
      "身份不可读时不许凭缓存猜出恢复卡").toBeNull();
    expect(getWpsExecutionStatusMock).not.toHaveBeenCalled();
    // 身份不可读 ⇒ 连发现请求都不发（不拿空身份去问服务端）。
    expect(listPendingWpsMock).not.toHaveBeenCalled();
  });

  it("用户点选的那条不许被后来的自动发现改判", async () => {
    const f = await interruptedAfterSecondWrite();
    f.storage.clear();
    localStorage.removeItem(PENDING_KEY);
    const jobs = pendingDiscoveryBackend();
    jobs.publish(inFlightJob(f));
    jobs.publish({
      operationId: "wps-resume-0002", taskId: "task-2",
      workbookIdentity: f.state.workbookIdentity, status: "failed",
      nextActionIndex: 2, totalActions: 3, inFlightActionIndex: null,
      requiresReconciliation: false, canResume: true, updatedAt: Date.now() + 5_000
    });
    getWpsApplicationMock.mockReturnValue({...f.application});
    getWpsExecutionStatusMock.mockImplementation(async (operationId: string) =>
      (operationId === resumeOperationId
        ? inFlightState(f)
        : {...f.state, operationId: "wps-resume-0002", taskId: "task-2", status: "failed" as const,
          nextActionIndex: 2, inFlightActionIndex: null, requiresReconciliation: false}) as never);
    f.writes.length = 0;

    const {container} = await mountPanel();
    const rows = () => Array.from(container.querySelectorAll<HTMLButtonElement>(".wps-recovery-candidate"));
    await until(() => rows().length === 2, "两条候选都必须摆出来");
    await until(() => rows().every(row => !row.disabled), "候选行必须可点（选择入口不能因为面板忙而被禁掉）");
    await click(rows()[0]);
    await until(() => textOf(container).includes("待核对在途动作"), "点选后应摆出在途核对入口");
    /*
     * 真机上这一步之后，授权元数据/身份观察会再跑一次发现。那次**不许**把用户已经选定的
     * 卡片换成缓存里那条：`选择`是用户做的决定，自动发现只能补，不能改判。
     *
     * 用真实触发源：宿主短暂读不到身份（`focus` 里 `FullName` 为空），随后又恢复。
     */
    getWpsApplicationMock.mockReturnValue({
      ...f.application,
      ActiveWorkbook: {...f.application.ActiveWorkbook!, FullName: ""}
    });
    await act(async () => { window.dispatchEvent(new Event("focus")); });
    await settle();
    getWpsApplicationMock.mockReturnValue({...f.application});
    await act(async () => { window.dispatchEvent(new Event("focus")); });
    await settle();
    await settle();

    expect(textOf(container)).toContain("待核对在途动作");
    expect(textOf(container)).not.toContain("需重新规划");
    // 也不许退回"再选一次"的候选列表 —— 那是把用户的选择丢掉。
    expect(container.querySelectorAll(".wps-recovery-candidate").length).toBe(0);
    expect(f.writes).toEqual([]);
  });

  /*
   * ==========================================================================================
   * 2026-10-28 真机（第二轮）：**授权范围变化后，恢复卡不重新核验**
   * ==========================================================================================
   *
   * 现场：`wps-execution-1fd938fd…` 第二步已补记（`next=2`、无在途动作），用户点过
   * "核对在途动作"；随后把误勾的第三张表取消，底部只剩目标表 + 来源表，卡面却仍停在
   * 按三张表算出的"需重新规划（工作簿、宿主版本或授权已变化）"。
   *
   * 机制：`discoverWpsRecovery` 里那段 `selected` 保护直接 `return` —— 它的本意是
   * "不许被**别的候选**改判"（2026-10-28 上一轮修的），但同时把"对**同一条**重算结论"
   * 也一起挡掉了。授权键是发现 effect 的依赖，可依赖变化只换来了"什么都不做"。
   *
   * 下面这组用例的前提：登记范围是 Source+Target 两张表，现场多勾了一张 Extra。
   * 判据本身一个字没改 —— 授权集合与原登记不同就必须阻断；改的只是"范围一变就重新算"。
   */

  /**
   * 真机形状的现场 + 已点选"在途那条"的面板。
   *
   * 三张表（Source/Target/Extra）时授权集合与原登记（两张）不一致 ⇒ 阻断；收回 Extra 后
   * 同一条执行必须重新核验。点选走**真实候选行**，所以卡片 `source` 是 `selected`。
   */
  async function selectInFlightUnderScope(indexes: number[]) {
    const f = await interruptedAfterSecondWrite({extraSheet: true});
    const jobs = pendingDiscoveryBackend();
    const reconciled = {
      status: "running" as const,
      nextActionIndex: 2,
      inFlightActionIndex: null,
      requiresReconciliation: false
    };
    jobs.publish(inFlightJob(f, reconciled));
    // 第二条（写前失败）只是为了让"候选列表"出现 —— 只有 ≥2 条时才有可点的选择入口。
    jobs.publish({
      operationId: "wps-resume-0002", taskId: "task-2",
      workbookIdentity: f.state.workbookIdentity, status: "failed",
      nextActionIndex: 0, totalActions: 2, updatedAt: Date.now() + 5_000
    });
    getWpsExecutionStatusMock.mockImplementation(async (operationId: string) =>
      (operationId === resumeOperationId
        ? inFlightState(f, reconciled)
        : {...f.state, operationId: "wps-resume-0002", taskId: "task-2", status: "failed" as const,
          nextActionIndex: 0, inFlightActionIndex: null, requiresReconciliation: false}) as never);
    getWpsApplicationMock.mockReturnValue({...f.application});
    f.writes.length = 0;
    listModelsMock.mockResolvedValue({defaultModelId: firstConnection.id, models: [firstConnection]});
    const view = await mount(
      <WpsAgentPanel selectedSheetIndexes={indexes} scopeDock={<div>固定范围</div>} />
    );
    const rows = () =>
      Array.from(view.container.querySelectorAll<HTMLButtonElement>(".wps-recovery-candidate"));
    await until(() => rows().length === 2, "两条候选都必须摆出来");
    await until(() => rows().every(row => !row.disabled), "候选行必须可点");
    await click(rows()[0]);
    return {f, jobs, view};
  }

  it("授权从三张表收回成原两张后，同一条已选执行重新核验并改判", async () => {
    const {f, view} = await selectInFlightUnderScope([1, 2, 3]);

    // 三张表 ≠ 原执行登记的两张：准确阻断，且原因点名**授权**（不是材料、不是宿主）。
    await until(() => textOf(view.container).includes("需重新规划"), "授权集合不一致必须阻断");
    expect(textOf(view.container)).toContain("当前授权 3 张表，原执行登记 2 张表");
    expect(findButton(view.container, "继续剩余动作")).toBeNull();
    expect(f.writes).toEqual([]);

    const before = getWpsExecutionStatusMock.mock.calls.length;
    await view.rerender(
      <WpsAgentPanel selectedSheetIndexes={[1, 2]} scopeDock={<div>固定范围</div>} />
    );
    await until(() => textOf(view.container).includes("可继续剩余动作"),
      "收回第三张表后必须对同一条执行重新核验，而不是停在按旧范围算出的结论上");
    expect(findButton(view.container, "继续剩余动作")).not.toBeNull();

    // 重新核验只查用户选定的那一条（不是"最近一条"，也不是缓存里那条）。
    const calls = getWpsExecutionStatusMock.mock.calls.slice(before);
    expect(calls.length).toBeGreaterThan(0);
    expect(calls.every((call) => call[0] === resumeOperationId)).toBe(true);
    // 仍然不退回候选列表：用户的选择没有被丢掉。
    expect(view.container.querySelectorAll(".wps-recovery-candidate").length).toBe(0);
  });

  it("授权收窄成不符合条件的范围时，旧的可续跑结论必须立刻失效", async () => {
    const {f, view} = await selectInFlightUnderScope([1, 2]);
    await until(() => textOf(view.container).includes("可继续剩余动作"), "两张表下应可继续");
    expect(findButton(view.container, "继续剩余动作")).not.toBeNull();

    // 只勾目标表：来源表不在授权里 ⇒ 来源证据无法复核，续跑入口必须消失。
    const before = getWpsExecutionStatusMock.mock.calls.length;
    await view.rerender(
      <WpsAgentPanel selectedSheetIndexes={[2]} scopeDock={<div>固定范围</div>} />
    );
    await until(() => findButton(view.container, "继续剩余动作") === null,
      "范围不再满足原执行要求时，旧的续跑按钮不许留在界面上");
    expect(textOf(view.container)).not.toContain("可继续剩余动作");
    expect(getWpsExecutionStatusMock.mock.calls.length).toBeGreaterThan(before);
    // 只读：撤掉入口不等于动现场。
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(f.writes).toEqual([]);
  });

  it("授权变化后的刷新不采纳旧索引或别的候选：只重算用户选的那一条", async () => {
    const {f, view} = await selectInFlightUnderScope([1, 2, 3]);
    await until(() => textOf(view.container).includes("需重新规划"), "先摆出三张表的阻断");

    /*
     * 浏览器里残留一条**别的执行**的索引（同一工作簿、材料早就没有）。刷新如果改走
     * "缓存优先"，卡面就会变成那条的结论 —— 用户选的是在途这条，界面给的却会是另一条。
     */
    localStorage.setItem(PENDING_KEY, JSON.stringify({
      schemaVersion: 1, operationId: "wps-resume-stale", conversationId: "conv-1",
      activityMessageId: "old-message", workbookIdentity: f.state.workbookIdentity,
      taskId: "task-stale", createdAt: new Date().toISOString()
    }));

    const before = getWpsExecutionStatusMock.mock.calls.length;
    await view.rerender(
      <WpsAgentPanel selectedSheetIndexes={[1, 2]} scopeDock={<div>固定范围</div>} />
    );
    await until(() => textOf(view.container).includes("可继续剩余动作"), "收窄授权后应改判为可续跑");
    const calls = getWpsExecutionStatusMock.mock.calls.slice(before).map((call) => call[0]);
    expect(calls).toEqual([resumeOperationId]);
    expect(calls).not.toContain("wps-resume-stale");
    expect(view.container.querySelectorAll(".wps-recovery-candidate").length).toBe(0);
  });

  it("连续切换授权时迟到的核验结果不许覆盖新现场；换工作簿后也不许再画回来", async () => {
    const {f, view} = await selectInFlightUnderScope([1, 2, 3]);
    await until(() => textOf(view.container).includes("需重新规划"), "先摆出三张表的阻断");

    // 扣住接下来的状态响应：两次授权切换各发一次，由用例决定放行顺序。
    const held: Array<(state: unknown) => void> = [];
    getWpsExecutionStatusMock.mockImplementation(
      () => new Promise((resolve) => { held.push(resolve); }) as never
    );

    await view.rerender(
      <WpsAgentPanel selectedSheetIndexes={[1, 2]} scopeDock={<div>固定范围</div>} />
    );
    await until(() => held.length === 1, "第一次切换应发出一次只读查询");
    await view.rerender(
      <WpsAgentPanel selectedSheetIndexes={[1, 2, 3]} scopeDock={<div>固定范围</div>} />
    );
    await until(() => held.length === 2, "第二次切换应再发一次");

    // 先放行**后发**的那次（三张表 ⇒ 仍阻断），再放行**先发**的那次（两张表 ⇒ 可继续）。
    await act(async () => {
      held[1](inFlightState(f, {status: "running", nextActionIndex: 2,
        inFlightActionIndex: null, requiresReconciliation: false}));
    });
    await settle();
    await act(async () => {
      held[0](inFlightState(f, {status: "running", nextActionIndex: 2,
        inFlightActionIndex: null, requiresReconciliation: false}));
    });
    await settle();
    await settle();

    // 判别点：最终画面是"最后一次切换"的结论；先发那次的迟到结果不许翻案。
    expect(textOf(view.container), "先发那次的迟到结论不许翻案").toContain("需重新规划");
    expect(textOf(view.container), "迟到的可续跑结论不许盖掉新的阻断").not.toContain("可继续剩余动作");

    // 换工作簿：旧簿的重新核验（即使还在飞）也不许画到新簿上。
    getWpsApplicationMock.mockReturnValue({...f.application,
      ActiveWorkbook: {...f.application.ActiveWorkbook!, FullName: "C:\\Data\\switched.xlsx"}});
    await act(async () => { window.dispatchEvent(new Event("focus")); });
    await settle();
    await settle();
    expect(view.container.querySelector(".wps-write-recovery"), "换簿后旧簿的核验结果不许画回来").toBeNull();
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();
    expect(f.writes).toEqual([]);
  });

  it("授权变化后的重新核验全程只读：不认领、不补记、不动账本、不写宿主", async () => {
    const {f, jobs, view} = await selectInFlightUnderScope([1, 2, 3]);
    await until(() => textOf(view.container).includes("需重新规划"), "三张表时应先报阻断");
    const ledgerBefore = jobs.snapshot(resumeOperationId);
    const storageBefore = JSON.stringify([...f.storage.entries()].sort());
    f.writes.length = 0;

    await view.rerender(
      <WpsAgentPanel selectedSheetIndexes={[1, 2]} scopeDock={<div>固定范围</div>} />
    );
    await until(() => textOf(view.container).includes("可继续剩余动作"), "收回第三张表后应改判");

    // 账本动作一个都没发生：没有认领、没有补记、没有动作级回执、没有收尾/挂起。
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();
    expect(startWpsActionMock).not.toHaveBeenCalled();
    expect(completeWpsActionMock).not.toHaveBeenCalled();
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
    expect(suspendWpsExecutionMock).not.toHaveBeenCalled();
    // 宿主一格没写；宿主缓存（PluginStorage）也没被改写 —— 材料本来就在，不需要水合。
    expect(f.writes).toEqual([]);
    expect(JSON.stringify([...f.storage.entries()].sort())).toBe(storageBefore);
    // 账本这一条原封不动（连处置时间戳都没写）。
    expect(jobs.snapshot(resumeOperationId)).toEqual(ledgerBefore);
  });

  it("持久包与宿主缓存都没有时，原因必须说\"没有材料\"，不许冒充\"找不到原执行的本机恢复记录\"", async () => {
    const f = await interruptedAfterSecondWrite();
    f.storage.clear();
    localStorage.removeItem(PENDING_KEY);
    pendingDiscoveryBackend().publish(inFlightJob(f, {
      status: "running", nextActionIndex: 2, inFlightActionIndex: null, requiresReconciliation: false
    }));
    // 这次执行确实没有恢复材料：宿主缓存没了，后端也没有它的包。
    backendRecovery().remove(resumeOperationId);
    expect(backendRecovery().load(resumeOperationId)).toBeNull();
    getWpsApplicationMock.mockReturnValue({...f.application});
    getWpsExecutionStatusMock.mockResolvedValue({...f.state, nextActionIndex: 2,
      inFlightActionIndex: null, requiresReconciliation: false} as never);

    const {container} = await mountPanel();
    await until(() => container.querySelector(".wps-write-recovery") !== null, "应摆出恢复卡");
    await until(() => !textOf(container).includes("正在查询服务端写入现场"), "应查完状态");
    expect(textOf(container)).toContain("需重新规划");
    /*
     * "材料缺失"与"宿主缓存里没有这条记录"是两件事。以前两者都落到同一句话，于是
     * 一个有持久包、只是缓存没了（真机 2026-10-28 的现场）的任务看起来也像"材料没了"。
     */
    expect(textOf(container)).toContain("本机没有这次执行的恢复材料");
    expect(textOf(container)).not.toContain("找不到原执行的本机恢复记录");
  });

  it("服务不可达时说\"暂时读不到\"，绝不说成\"没有未完成任务\"", async () => {
    const f = await interruptedAfterSecondWrite();
    f.storage.clear();
    localStorage.removeItem(PENDING_KEY);
    f.writes.length = 0;
    pendingDiscoveryBackend().failLists(new Error("本机服务不可达"));
    getWpsApplicationMock.mockReturnValue({...f.application});

    const {container} = await mountPanel();
    await until(() => textOf(container).includes("暂时读不到"), "查不到必须如实说读不到");
    // 徽标必须落在"暂时读不到"，而不是"没有未完成任务"。
    expect(container.querySelector(".wps-write-recovery-badge")?.textContent).toBe("暂时读不到");
    expect(textOf(container)).not.toContain("没有未完成的任务");
    expect(textOf(container)).not.toContain("暂无未完成的写入");
    // 真的查过（不是"根本没发请求，于是什么都没有"）。
    expect(listPendingWpsMock).toHaveBeenCalled();
    expect(container.querySelector(".wps-write-recovery")).not.toBeNull();
    expect(getWpsExecutionStatusMock).not.toHaveBeenCalled();
    expect(f.writes).toEqual([]);
    // 一恢复可用就重新核对：这张卡不是终态。
    pendingDiscoveryBackend().clearFailure();
    pendingDiscoveryBackend().publish(inFlightJob(f));
    getWpsExecutionStatusMock.mockResolvedValue(inFlightState(f) as never);
    await click(findButton(container, "重新核对（只读）"));
    await until(() => textOf(container).includes("待核对在途动作"), "恢复后应重新发现");
  });

  it("恢复材料损坏时仍按既有判据拒绝续写，只允许重新规划", async () => {
    const f = await interruptedAfterSecondWrite();
    f.storage.clear();
    localStorage.removeItem(PENDING_KEY);
    f.writes.length = 0;
    // 账本这一条没有在途动作（前两步已确认）⇒ 能否续跑完全取决于材料。
    pendingDiscoveryBackend().publish(inFlightJob(f, {
      status: "running", nextActionIndex: 2, inFlightActionIndex: null, requiresReconciliation: false
    }));
    backendRecovery().corrupt(resumeOperationId, "{\"schemaVersion\":1, broken");
    getWpsApplicationMock.mockReturnValue({...f.application});
    getWpsExecutionStatusMock.mockResolvedValue({...f.state, nextActionIndex: 2,
      inFlightActionIndex: null, requiresReconciliation: false} as never);

    const {container} = await mountPanel();
    await until(() => textOf(container).includes("本机恢复材料已损坏"), "损坏必须如实说损坏");
    expect(findButton(container, "继续剩余动作")).toBeNull();
    expect(findButton(container, "重新读取并规划")).not.toBeNull();
    expect(findButton(container, "重新读取并规划")?.disabled).toBe(true);
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(f.writes).toEqual([]);
  });

  it.each(["expired","missing"])("真实恢复材料 %s 时禁用旧计划续写，重开仍不自动执行",async issue=>{
    const f=resumeHost();
    await expect(executeWpsPlan(f.application,f.prepared,undefined,undefined,async(phase,c)=>{
      if(phase==="after" && c.actionIndex===1) throw new Error("lost receipt");
    },resumeOperationId,{taskId:"task-1",...recoveryChannel})).rejects.toThrow("lost receipt");
    f.writes.length=0;
    if(issue==="missing") {
      // 主机缓存与本机后端那一份**都不在**了，这才是"材料真的不存在"。
      // 只删宿主缓存不再等于丢材料 —— 完整退出 WPS 走的就是那条路，而它能续跑。
      f.storage.delete("excel-bro.wps.undo.v1");
      backendRecovery().clear();
    } else {
      const records=JSON.parse(String(f.storage.get("excel-bro.wps.undo.v1")));
      records[0].createdAt="2000-01-01T00:00:00Z";
      f.storage.set("excel-bro.wps.undo.v1",JSON.stringify(records));
      // 保留期到了：后端读不出正文（线上 410），界面必须说"缺失/过期"而不是硬续。
      backendRecovery().forceExpire(resumeOperationId);
    }
    getWpsApplicationMock.mockReturnValue(f.application);
    localStorage.setItem(PENDING_KEY,JSON.stringify({schemaVersion:1,operationId:resumeOperationId,
      conversationId:"conv-1",activityMessageId:"old-message",workbookIdentity:f.state.workbookIdentity,
      taskId:f.state.taskId,createdAt:new Date().toISOString()}));
    getWpsExecutionStatusMock.mockResolvedValue(f.state);
    listModelsMock.mockResolvedValue({defaultModelId:firstConnection.id,models:[firstConnection]});
    for(let count=0;count<2;count++) {
      const view=await mount(<WpsAgentPanel selectedSheetIndexes={[1,2]} scopeDock={null}/>);
      await until(()=>findButton(view.container,"重新读取并规划")!==null,"missing material must require a new plan");
      expect(findButton(view.container,"继续剩余动作")).toBeNull();
      expect(findButton(view.container,"重新读取并规划")?.disabled).toBe(true);
      expect(textOf(view.container)).toContain("本机材料缺失、过期");
      /*
       * 原因要分得开：过期说保留期，缺失说没有这份材料。
       * 以前两者都落到"找不到恢复记录"，用户没法判断是等一等（其实不行）还是重新规划。
       */
      if(issue==="expired") {
        expect(textOf(view.container)).toContain("本机恢复包已超过保留期");
        expect(textOf(view.container)).not.toContain("找不到原执行的本机恢复记录");
      } else {
        // 缺失 = 宿主缓存与持久恢复包都不在，措辞要说"没有材料"，不能说"找不到恢复记录"。
        expect(textOf(view.container)).toContain("本机没有这次执行的恢复材料");
        expect(textOf(view.container)).not.toContain("找不到原执行的本机恢复记录");
      }
      unmountMounted(view.container);
    }
    expect(f.writes).toEqual([]);
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(startWpsActionMock).not.toHaveBeenCalled();
    expect(runTurnMock).not.toHaveBeenCalled();
  });

  it("真实过期材料重新规划全链：新执行完成并确认后重开，不复活旧卡或重复写入", async () => {
    const f=resumeHost();
    await expect(executeWpsPlan(f.application,f.prepared,undefined,undefined,async(phase,c)=>{
      if(phase==="after" && c.actionIndex===1) throw new Error("lost receipt");
    },resumeOperationId,{taskId:"task-1",...recoveryChannel})).rejects.toThrow("lost receipt");
    /*
     * "材料过期"必须让**持久材料的保留期**到点，改宿主缓存里的 `createdAt` 不再算数：
     * 持久包有服务端首次保存时定的 `expiresAt`，它才是权威期限（改缓存时间点既
     * 不延长也不缩短那份材料）。这里让本机后端真的读不出正文（线上 410），
     * 界面才该给出"需要重新规划"。
     */
    backendRecovery().forceExpire(resumeOperationId);
    f.writes.length=0;
    getWpsApplicationMock.mockReturnValue(f.application);
    const history=JSON.parse(JSON.stringify(historyWithTaskResume()));
    history[0].taskResume.taskId="task-1";
    history[0].taskResume.originalPrompt="核对已完成内容，仅将 Target A3 写成 30";
    localStorage.setItem("excel-bro.wps.chat.v1",JSON.stringify({activeConversationId:"conv-1",conversations:[{
      id:"conv-1",title:"过期恢复",messages:history,createdAt:new Date().toISOString(),updatedAt:new Date().toISOString()
    }]}));
    localStorage.setItem(PENDING_KEY,JSON.stringify({schemaVersion:1,operationId:resumeOperationId,
      conversationId:"conv-1",activityMessageId:"old-message",workbookIdentity:f.state.workbookIdentity,
      taskId:f.state.taskId,createdAt:new Date().toISOString()}));
    getWpsExecutionStatusMock.mockResolvedValue(f.state);
    claimWpsExecutionResumeMock.mockResolvedValueOnce({...f.state,ownerToken:"old-owner",revision:6});
    listModelsMock.mockResolvedValue({defaultModelId:firstConnection.id,models:[firstConnection]});
    const newPlan: AnalysisPlan={...f.prepared.plan,id:"new-replanned-plan",actions:[{type:"writeValues",sheet:"Target",range:"A3",values:[[30]]}]};
    validatePlanMock.mockResolvedValueOnce(newPlan);
    runTurnMock.mockResolvedValue({kind:"plan",plan:newPlan,turnId:"new-replanned-turn"});
    const view=await mount(<WpsAgentPanel selectedSheetIndexes={[1,2]} scopeDock={null}/>);
    await until(()=>findButton(view.container,"重新读取并规划")!==null,"expired recovery entry");
    expect(f.writes).toEqual([]);
    expect(runTurnMock).not.toHaveBeenCalled();
    await click(view.container.querySelector('.wps-write-recovery input[type="checkbox"]') as HTMLInputElement);
    await click(findButton(view.container,"重新读取并规划"));
    await until(()=>findButton(view.container,"确认变更")!==null,"new execution must reach review: "+textOf(view.container));
    expect(runTurnMock).toHaveBeenCalledTimes(1);
    expect(registerWpsExecutionMock).toHaveBeenCalledTimes(1);
    const newId=registerWpsExecutionMock.mock.calls[0][0].operationId;
    expect(newId).not.toBe(resumeOperationId);
    expect(startWpsActionMock).toHaveBeenCalledTimes(1);
    expect(startWpsActionMock.mock.calls[0][0]).toBe(newId);
    expect(finishWpsExecutionMock).toHaveBeenCalledTimes(1);
    expect(finishWpsExecutionMock.mock.calls[0][0]).toBe(newId);
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();
    expect(f.writes).toEqual(["Target!A3"]);
    await click(findButton(view.container,"确认变更"));
    await until(()=>view.container.querySelector(".result-review-bar")===null,"review settles");
    expect(localStorage.getItem(PENDING_KEY)).toBeNull();
    const statusCalls=getWpsExecutionStatusMock.mock.calls.length;
    unmountMounted(view.container);
    const reopened=await mount(<WpsAgentPanel selectedSheetIndexes={[1,2]} scopeDock={null}/>);
    await until(()=>Boolean(reopened.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型")),"reopened ready");
    expect(reopened.container.querySelector(".wps-write-recovery")).toBeNull();
    expect(getWpsExecutionStatusMock).toHaveBeenCalledTimes(statusCalls);
    expect(registerWpsExecutionMock).toHaveBeenCalledTimes(1);
    expect(runTurnMock).toHaveBeenCalledTimes(1);
    expect(f.writes).toEqual(["Target!A3"]);
    expect([1,2,3].map(i=>f.cells.get(`Target!A${i}`))).toEqual([10,20,30]);
  });

  it("scope changes recompute recovery read-only and discard the older scope response", async () => {
    const f = resumeHost();
    await executeWpsPlan(f.application,f.prepared,undefined,undefined,undefined,resumeOperationId,{taskId:"task-1",...recoveryChannel});
    f.state.nextActionIndex=3; f.writes.length=0;
    getWpsApplicationMock.mockReturnValue(f.application);
    localStorage.setItem(PENDING_KEY,JSON.stringify({schemaVersion:1,operationId:resumeOperationId,
      conversationId:"conv-1",activityMessageId:"old-message",workbookIdentity:f.state.workbookIdentity,
      taskId:f.state.taskId,createdAt:new Date().toISOString()}));
    getWpsExecutionStatusMock.mockResolvedValue(f.state);
    listModelsMock.mockResolvedValue({defaultModelId:firstConnection.id,models:[firstConnection]});
    const view=await mount(<WpsAgentPanel selectedSheetIndexes={[2]} scopeDock={null}/>);
    await until(()=>textOf(view.container).includes("证据不足"),"narrow scope cannot prove original authorization");
    await view.rerender(<WpsAgentPanel selectedSheetIndexes={[1,2]} scopeDock={null}/>);
    await until(()=>findButton(view.container,"补记完成（不重写原表）")!==null,"restored scope must refresh without clicking recheck");
    let late!: (s: typeof f.state) => void;
    getWpsExecutionStatusMock.mockImplementationOnce(()=>new Promise(resolve=>{late=resolve;}));
    await view.rerender(<WpsAgentPanel selectedSheetIndexes={[2]} scopeDock={null}/>);
    await until(()=>Boolean(late),"old scope request must start");
    await view.rerender(<WpsAgentPanel selectedSheetIndexes={[1,2]} scopeDock={null}/>);
    await until(()=>findButton(view.container,"补记完成（不重写原表）")!==null,"new response applied");
    await act(async()=>late({...f.state,workbookIdentity:"wps-path:c:\\other.xlsx"}));
    expect(findButton(view.container,"补记完成（不重写原表）")).not.toBeNull();
    expect(f.writes).toEqual([]);
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
    expect(runTurnMock).not.toHaveBeenCalled();
  });

  it.each(["确认变更", "放弃变更"])("最终回执丢失后补记并审阅，%s不重写任何动作", async outcome => {
    const f = resumeHost();
    await executeWpsPlan(f.application, f.prepared, undefined, undefined, undefined, resumeOperationId, {taskId:"task-1", ...recoveryChannel});
    f.state.nextActionIndex = 3;
    f.writes.length = 0;
    getWpsApplicationMock.mockReturnValue(f.application);
    localStorage.setItem(PENDING_KEY, JSON.stringify({schemaVersion:1, operationId:resumeOperationId,
      conversationId:"conv-1",activityMessageId:"old-message",workbookIdentity:f.state.workbookIdentity,
      taskId:f.state.taskId,createdAt:new Date().toISOString()}));
    getWpsExecutionStatusMock.mockResolvedValue(f.state);
    let owned = {...f.state, ownerToken:"new-owner", ownerActive:true, revision:6};
    claimWpsExecutionResumeMock.mockImplementation(async () => owned);
    startWpsActionMock.mockImplementation(async (_id, request) => {
      owned={...owned,inFlightActionIndex:request.actionIndex,revision:owned.revision+1};return owned;
    });
    completeWpsActionMock.mockImplementation(async (_id, request) => {
      owned={...owned,inFlightActionIndex:null,nextActionIndex:request.actionIndex+1,revision:owned.revision+1};return owned;
    });
    finishWpsExecutionMock.mockImplementation(async () => ({...owned,status:"completed" as const}));
    listModelsMock.mockResolvedValue({defaultModelId:firstConnection.id,models:[firstConnection]});
    const {container}=await mount(<WpsAgentPanel selectedSheetIndexes={[1,2]} scopeDock={<div>固定两表</div>}/>);
    await until(()=>findButton(container,"补记完成（不重写原表）") !== null,"恢复卡必须接入生产续跑判据");
    await click(findButton(container,"补记完成（不重写原表）"));
    await until(()=>findButton(container,"确认变更") !== null,"续跑后必须进入审阅");
    expect(f.writes).toEqual([]);
    expect(startWpsActionMock).not.toHaveBeenCalled();
    expect(runTurnMock).not.toHaveBeenCalled();
    expect(registerWpsExecutionMock).not.toHaveBeenCalled();
    await click(findButton(container,outcome));
    await until(()=>container.querySelector(".result-review-bar")===null,"审阅应收尾");
    if(outcome === "确认变更") {
      expect(f.writes).toEqual([]);
      expect([1,2,3].map(i=>f.cells.get(`Target!A${i}`))).toEqual([10,20,30]);
    } else {
      expect([1,2,3].map(i=>f.cells.get(`Target!A${i}`))).toEqual(["old1","old2","old3"]);
    }
    expect(localStorage.getItem(PENDING_KEY)).toBeNull();
  });

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

  it("已完成任务连续重开两次仅查询，不认领、补记或写入", async () => {
    const {container,host}=await mountRecoveryCard({state:{status:"completed",nextActionIndex:1,totalActions:1}});
    // 已完成 = 收尾终态：退出可点的恢复卡，只剩一条只读终态摘要（没有核对/补记/清除按钮）。
    expect(container.querySelector(".wps-write-recovery:not(.is-terminal)")).toBeNull();
    expect(container.querySelector('[data-recovery-terminal="settled"]')).not.toBeNull();
    expect(recoveryButtons(container)).toEqual([]);
    // 无所属消息时，终态摘要**只出现在消息流里一次**，固定区域（dock）一个像素都不画。
    const terminalInStream = container.querySelectorAll("main.message-stream [data-recovery-terminal]");
    expect(terminalInStream.length).toBe(1);
    expect(container.querySelector(".composer-action-dock [data-recovery-terminal]")).toBeNull();
    expect(container.querySelectorAll(".composer-action-dock-chip").length).toBe(0);
    unmountMounted(container);
    const reopened=await mount(panel());
    await until(()=>getWpsExecutionStatusMock.mock.calls.length===2,"second mount queries persisted index");
    await until(()=>!textOf(reopened.container).includes("正在查询服务端写入现场"),"second query settles");
    expect(reopened.container.querySelector(".wps-write-recovery:not(.is-terminal)")).toBeNull();
    expect(reopened.container.querySelector('[data-recovery-terminal="settled"]')).not.toBeNull();
    expect(recoveryButtons(reopened.container)).toEqual([]);
    // 重开（重复发现）不重复追加：终态摘要仍只在消息流出现一次。
    expect(reopened.container.querySelectorAll("main.message-stream [data-recovery-terminal]").length).toBe(1);
    expect(reopened.container.querySelector(".composer-action-dock [data-recovery-terminal]")).toBeNull();
    expect(host.writes()).toBe(0);
    for(const mutation of [claimWpsExecutionResumeMock,reconcileWpsActionReceiptMock,
      registerWpsExecutionMock,startWpsActionMock,completeWpsActionMock,finishWpsExecutionMock,runTurnMock]) {
      expect(mutation).not.toHaveBeenCalled();
    }
  });

  it("恢复记录收尾退出 dock 后，新任务进入待审阅仍给确认/放弃（终态退出不误伤审阅入口）", async () => {
    // 先有：一条已完成的恢复记录（settled 终态，退出固定区域，只留消息流摘要）。
    const {container} = await mountRecoveryCard({state:{status:"completed",nextActionIndex:1,totalActions:1}});
    expect(container.querySelector(".composer-action-dock [data-recovery-terminal]")).toBeNull();
    expect(container.querySelector("main.message-stream [data-recovery-terminal]")).not.toBeNull();

    // 再发起一个新任务，执行到待审阅阶段。
    const writePlan = {
      id: "write-2", title: "更新 A1", summary: "把 A1 更新为新值",
      assumptions: [], warnings: [],
      actions: [{ type: "writeValues" as const, sheet: "Sheet1", range: "A1", values: [["新值"]] }],
      acceptanceCriteria: []
    };
    runTurnMock.mockResolvedValue({ kind: "plan", plan: writePlan });
    validatePlanMock.mockImplementation(async (submitted) => submitted);
    const textarea = container.querySelector<HTMLTextAreaElement>("textarea");
    await type(textarea, "把 A1 改成新值");
    await submit(textarea);
    await until(() => container.querySelector(".result-review-bar") !== null, "应进入待审阅");

    // 待审阅的确认/放弃入口仍在 —— 恢复记录的终态退出没有把它一起清掉。
    expect(findButton(container, "确认变更")).not.toBeNull();
    expect(findButton(container, "放弃变更")).not.toBeNull();
    // 恢复终态摘要仍只在消息流里，没有因为新任务被重新摆回固定区域。
    expect(container.querySelectorAll("main.message-stream [data-recovery-terminal]").length).toBe(1);
    expect(container.querySelector(".composer-action-dock [data-recovery-terminal]")).toBeNull();
  });

  it("连续点击补记入口只领取一次执行权且只提交一次回执", async () => {
    const {container, host}=await mountRecoveryCard({state:{nextActionIndex:2,totalActions:2}});
    let release!: (state: ReturnType<typeof recoveryState>) => void;
    claimWpsExecutionResumeMock.mockImplementationOnce(()=>new Promise(resolve=>{release=resolve;}));
    const button=findButton(container,"补记完成（不重写原表）")!;
    await act(async()=>{
      button.dispatchEvent(new MouseEvent("click",{bubbles:true}));
      button.dispatchEvent(new MouseEvent("click",{bubbles:true}));
    });
    await until(()=>Boolean(release),"claim must be pending");
    expect(claimWpsExecutionResumeMock).toHaveBeenCalledTimes(1);
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
    await act(async()=>release(recoveryState({nextActionIndex:2,totalActions:2,ownerToken:"owner",revision:8})));
    await until(()=>localStorage.getItem(PENDING_KEY)===null,"backfill must settle");
    expect(finishWpsExecutionMock).toHaveBeenCalledTimes(1);
    expect(host.writes()).toBe(0);
    expect(startWpsActionMock).not.toHaveBeenCalled();
    expect(runTurnMock).not.toHaveBeenCalled();
  });

  it("还有动作没写：只给「重新读取并规划」，不给任何继续/写入按钮", async () => {
    const { container } = await mountRecoveryCard({
      state: { nextActionIndex: 1, totalActions: 3 }
    });

    const buttons = recoveryButtons(container);
    expect(buttons).toEqual(["重新读取并规划", "清除这条恢复提示"]);
    // 关键：不许出现"可以继续"这类冒充断点续跑的入口。
    expect(buttons.some((name) => name.includes("继续"))).toBe(false);
    expect(textOf(container)).toContain("本机材料缺失、过期或不符合精确续跑条件");
  });

  it("点「重新读取并规划」→ 用原需求重新发起一轮，不是续跑", async () => {
    const { container, host } = await mountRecoveryCard({
      state: { nextActionIndex: 1, totalActions: 3 },
      withHistory: true
    });
    runTurnMock.mockResolvedValue({ kind: "plan", plan: singleWritePlan, turnId: "turn-replan" });
    /*
     * 校验替身必须在本例里自己给：它以前靠上一条用例残留的 `mockImplementation`
     * 活着，单独跑这一条就会拿到 `undefined`，走进 `waitForPlanning` 的
     * "reading 'then'" 而看不出真实意图。用例之间不许靠泄漏的实现接续。
     */
    validatePlanMock.mockImplementation(async (submitted) => submitted);
    expect(findButton(container, "重新读取并规划")?.disabled).toBe(true);
    await click(container.querySelector('.wps-write-recovery input[type="checkbox"]') as HTMLInputElement);
    await click(findButton(container, "重新读取并规划"));
    await until(
      () => textOf(container).includes("不是从上次进度接着写"),
      "应说明这是新的一轮"
    );

    /*
     * "新的一轮"在新流程里会**真的执行**（计划就绪即写），所以不能再拿"零写入"
     * 当判据 —— 那会把正常执行误判成续跑。真正区分"续跑"与"新一轮"的是：
     * 恢复端点（补记 / 完成回执）一次都没被调用，而领取与挂起按新轮次走。
     */
    expect(runTurnMock).toHaveBeenCalled();
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();
    /*
     * 新一轮会**自己执行**并提交它自己的完成回执 —— 那不是"续跑"。
     * 判据因此从"finish 一次都没调用"收紧为"**没有一次是针对旧执行**"：
     * 拿旧 id 去补记/完成，才是把上一轮的现场当成这一轮的。
     */
    const oldOperationId = "wps-execution-test";
    expect(
      finishWpsExecutionMock.mock.calls.some(([id]) => String(id) === oldOperationId)
    ).toBe(false);
    expect(claimWpsExecutionResumeMock).toHaveBeenCalled();
    expect(suspendWpsExecutionMock).toHaveBeenCalled();
    /*
     * 挂起索引：旧执行被领取并挂起后不该再被当成"可续跑"；
     * 新一轮若已完成，索引会被正常清空。两种都算对，**指向旧执行**才算错。
     */
    const pendingIndex = localStorage.getItem(PENDING_KEY);
    expect(pendingIndex === null || !pendingIndex.includes(oldOperationId)).toBe(true);
    await until(() => container.querySelector(".wps-write-recovery") === null,
      "新计划已接管发现索引后，不得刷新并复活旧任务恢复卡");
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

  /*
   * 2026-10-28 夜间真机现象的最小复现：**第一次读不到现场 → 点一次只读重试就好**。
   *
   * 那次现场只有「证据不足」四个字，`unknown` 的十几条产生路径一条都分不出来。
   * 这里钉住三件事：
   * ① 两次核验各自留一条诊断，`stage` 能分开（discovery / recheck），`code` 能说清原因；
   * ② 两条记录能用同一个操作 ID 尾部串起来（否则"两次对照"无从谈起）；
   * ③ 只读核验**一次宿主写入都没有**，且诊断里没有表名、单元格值、路径。
   */
  it("首次发现与只读重试各留一条诊断：前一次说得出原因，后一次不再 unknown，且零写入", async () => {
    resetWpsRecoveryDiagnostics();
    // 第一次：动作目标表不在本次授权里 —— 真机上「读不到现场」的一支。
    const { container, host } = await mountRecoveryCard({ state: {
      totalActions: 1,
      descriptor: { actions: [{
        actionIndex: 0, type: "writeValues", sheet: "未授权表", range: "A1",
        rowCount: 1, columnCount: 1, cellCount: 1,
        previewFingerprint: matrixFingerprintV2([["旧值"]], [["旧值"]]),
        expectedFingerprint: matrixFingerprintV2([["旧值"]], [["旧值"]])
      }] },
      authorizedSheetNames: ["Sheet1"]
    } });
    expect(recoveryButtons(container)).toEqual(["清除这条恢复提示"]);
    expect(textOf(container)).toContain("无法读取或证明当前目标现场");

    const first = wpsRecoveryDiagnosticEvents();
    const discovery = first.find((event) => event.stage === "discovery");
    expect(discovery, "首次发现必须留一条诊断").toBeTruthy();
    expect(discovery!.outcome).toBe("unknown");
    expect(discovery!.code).toBe("action_sheet_not_allowed");
    expect(discovery!.verdict).toBe("insufficient");
    // 原因码落到具体动作上，而不是"某个动作"。
    expect(discovery!.actionIndex).toBe(0);

    // 第二次：现场恢复正常（这里只改服务端状态与授权表名），点只读重试。
    getWpsExecutionStatusMock.mockResolvedValue(recoveryState({ nextActionIndex: 1 }) as never);
    await click(findButton(container, "重新核对（只读）"));
    await until(() => recoveryButtons(container).includes("补记完成（不重写原表）"), "重试后应给出可操作结论");

    const events = wpsRecoveryDiagnosticEvents();
    const recheck = events.filter((event) => event.stage === "recheck").at(-1);
    expect(recheck, "只读重试必须留一条诊断").toBeTruthy();
    expect(recheck!.outcome).not.toBe("unknown");
    expect(recheck!.verdict).not.toBe("insufficient");
    // 两条记录指向同一次执行：靠 ID 尾部对照，不靠时间猜。
    expect(recheck!.operationIdTail).toBe(discovery!.operationIdTail);
    expect(recheck!.taskIdTail).toBe(discovery!.taskIdTail);
    // 序号递增：现场按顺序看就知道"先这次、后那次"。
    expect(recheck!.seq).toBeGreaterThan(discovery!.seq);

    // 只读核验全程零宿主写入。
    expect(host.writes()).toBe(0);
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();

    /*
     * 诊断里不许出现业务数据。这三种值在本次现场里都真实存在：
     * 表名「未授权表」、工作簿路径 `c:\data\测试.xlsx`、单元格值「旧值」。
     */
    const serialized = JSON.stringify(events);
    for (const secret of ["未授权表", "Sheet1", "测试.xlsx", "旧值", "c:\\"]) {
      expect(serialized).not.toContain(secret);
    }
  });

  it("账本登记的身份与当前文件不同（文件被移动/改名）：只给清除提示，不给写入按钮", async () => {
    // 缓存索引属于当前簿，但**账本那一条**登记的是另一个路径 —— 这才是真正的身份不一致。
    const { container } = await mountRecoveryCard({
      state: { nextActionIndex: 2, totalActions: 2, workbookIdentity: "wps-path:c:\\moved\\测试.xlsx" }
    });
    expect(recoveryButtons(container)).toEqual(["清除这条恢复提示"]);
    expect(textOf(container)).toContain("不是同一个文件");
  });

  it("打开另一个工作簿时：属于原簿的本地索引不在这里摆卡（发现只认当前身份）", async () => {
    // 本地索引是缓存，不是权威：它记的是**原簿**。当前簿要靠自己的身份去发现，
    // 不许把另一本簿的执行画成"这个工作簿上次写入的恢复核对"。
    const host = writeRangeHost();
    const application = getWpsApplicationMock() as unknown as {
      ActiveWorkbook: { FullName: string };
    };
    application.ActiveWorkbook.FullName = "C:\\other\\测试.xlsx";
    localStorage.setItem(PENDING_KEY, JSON.stringify({
      schemaVersion: 1, operationId: "wps-execution-test", conversationId: "conv-1",
      activityMessageId: "msg-1", workbookIdentity: "wps-path:c:\\data\\测试.xlsx",
      taskId: "task-test", createdAt: new Date().toISOString()
    }));
    getWpsExecutionStatusMock.mockResolvedValue(recoveryState({ nextActionIndex: 2, totalActions: 2 }) as never);
    listModelsMock.mockResolvedValue({defaultModelId: firstConnection.id, models: [firstConnection]});
    const { container } = await mount(panel());
    await until(() => listPendingWpsMock.mock.calls.some((call) =>
      call[0] === "wps-path:c:\\other\\测试.xlsx"), "应当按当前簿的身份去发现");
    expect(container.querySelector(".wps-write-recovery")).toBeNull();
    // 连状态都不许去查另一本簿的执行。
    expect(getWpsExecutionStatusMock).not.toHaveBeenCalled();
    expect(host.writes()).toBe(0);
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
  planningTransport.openBridge = false;
  let resolveTurn: (() => void) | null = null;
  let released = false;
  runTurnMock.mockImplementation(async (input: WpsReadonlyTurnInput) => {
    expect(await input.preparePlanningBridge?.("turn-bridge-panel")).toBe(true);
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
    // 这条链路的服务端会正常下发一份基于本次读取的计划（源没有变化）。
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
    // 计划到达**之前**：一个单元格都没被改（这是"执行前零写入"的正确观察点）。
    expect(host.writes()).toBe(0);

    // 2) 放行：结果计划回到面板 → 预览 → 确认 → 执行 → 回读。
    releaseTurn();
    await settle();
    await until(() => findButton(container, "确认变更") !== null || textOf(container).includes("本次执行需要核对"), "应展示执行结果或明确失败");
    // 计划到达即执行：此处已写入一次。零写入的观察点在上一段（计划到达之前）。
    expect(host.writes()).toBe(1);

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
    /*
     * 服务端已经判定本次读取不可用（`source_changed` / `unauthorized_scope`），
     * 就不会再下发一份"基于那份读取"的计划。替身若仍放出一份计划，面板只能照单
     * 执行 —— 测到的是替身不真实，而不是"面板无视来源/授权变化"。所以放行一个
     * 普通回答，与真实服务端行为一致。
     */
    const releaseTurn = armBridgeTurn({
      kind: "answer",
      message: "本次读取的源数据已不可用，未生成写入计划。",
      provider: "model"
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
    /*
     * 服务端已经判定本次读取不可用（`source_changed` / `unauthorized_scope`），
     * 就不会再下发一份"基于那份读取"的计划。替身若仍放出一份计划，面板只能照单
     * 执行 —— 测到的是替身不真实，而不是"面板无视来源/授权变化"。所以放行一个
     * 普通回答，与真实服务端行为一致。
     */
    const releaseTurn = armBridgeTurn({
      kind: "answer",
      message: "本次读取的源数据已不可用，未生成写入计划。",
      provider: "model"
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

  it("结果先可读后过期404：连续重开不复用缓存或自动重算", async () => {
    const options={resultStatus:200};
    const first=await completeQueryTurn(options);
    await until(()=>first.container.querySelector(".local-result-card-table")!==null,"initial result");
    const calls=runTurnMock.mock.calls.length;
    unmountMounted(first.container); options.resultStatus=404;
    for(let index=0;index<2;index++) {
      const reopened=await mount(<WpsAgentPanel selectedSheetIndexes={[1,2]} scopeDock={null}/>);
      await until(()=>reopened.container.querySelector(".local-result-card-unavailable")!==null,"expired result");
      expect(reopened.container.querySelector(".local-result-card-table")).toBeNull();
      expect(textOf(reopened.container)).toContain("请重新发起查询");
      expect(runTurnMock).toHaveBeenCalledTimes(calls);
      unmountMounted(reopened.container);
    }
    expect(first.bridge.captured.filter(c=>c.url.includes("/result"))).toHaveLength(3);
    expect(first.host.writes()).toBe(0);
    expect(registerWpsExecutionMock).not.toHaveBeenCalled();
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
      turnId: "turn-bridge-panel"
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
    /*
     * 计划到达即执行（先执行、后审阅），所以这里已经是"已写入"状态。
     * "执行前零写入"的正确观察点在 `releaseTurn()` **之前** —— 上面已经断言过
     * （`expect(host.writes()).toBe(0)` 在计划到达前、来源快照仍未变）。
     */
    expect(host.writes()).toBeGreaterThan(0);
    // 写入只落在目标表：两张来源表的逐格内容不得被这次执行改动。
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
    await until(() => textOf(container).includes("已撤销并恢复"), "撤销后应给出核对过的恢复结果");
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

      /*
       * 执行已在计划就绪时完成（先执行、后审阅）。"执行前格式未改"这个状态
       * 在面板里已经过去，不应再在这里断言 —— 它属于执行前的观察点。
       * 这里核对的是业务结果：格式真的生效、值没有被格式动作带跑。
       */
      await stageAndSave(container);
      await until(() => container.querySelector(".wps-verification-card.verified") !== null, "写入成功应给出回读核对结果");

      // 格式生效，值不变。
      expect(host.format().bold).toBe(true);
      expect(host.cellValue()).toBe("旧值");

      // 撤销：格式回到写前，值仍不变。
      await click(cardUndoButton(container));
      await until(() => textOf(container).includes("已撤销并恢复"), "撤销后应给出核对过的恢复结果");
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

  it("保存动态复制工具保留原动作，复用按当前授权重新读取来源", async () => {
    const host = capabilityHost(["Sheet1"]);
    getWpsApplicationMock.mockReturnValue(host.application);
    const copyPlan: AnalysisPlan = {
      ...formulaPlan, id: "save-dynamic-copy", title: "复制当前 A1 到 B1",
      actions: [{ type: "copyRange", sheet: "Sheet1", sourceSheet: "Sheet1",
        sourceRange: "A1", targetRange: "B1", copyType: "values", skipBlanks: false, transpose: false }]
    };
    const { container } = await mountPlanPreview(copyPlan, "复制当前 A1 到 B1");
    const dialog = await openSaveDialog(container);
    expect(host.grids[0][0][1].value).toBe("old-0-0");
    expect(textOf(dialog)).not.toContain("以后仍写入这些固定内容");
    for (const checkbox of dialog.querySelectorAll<HTMLInputElement>('input[type="checkbox"]')) {
      if (!checkbox.checked) await click(checkbox);
    }
    const save = dialogButton(dialog, "保存到我的工具");
    expect(save?.disabled).toBe(false);
    await click(save);
    await until(() => loadTools().length === 1, "复制工具应保存成功");
    expect(loadTools()[0].planTemplate.actions[0].type).toBe("copyRange");
    expect(JSON.stringify(loadTools()[0].planTemplate)).not.toContain("old-0-0");
    host.grids[0][0][0].value = "最新来源";
    host.grids[0][0][0].formula = "最新来源";
    validatePlanMock.mockClear();
    await click(container.querySelector(".tools-entry"));
    await click(container.querySelector<HTMLButtonElement>(".tool-card"));
    await until(() => host.grids[0][0][1].value === "最新来源", "复用必须重新读取当前来源");
    expect(validatePlanMock).toHaveBeenCalledTimes(1);
    expect(validatePlanMock.mock.calls[0][0].actions[0].type).toBe("copyRange");
  });

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
    await until(() => textOf(container).includes("本次执行计划"), "校验通过后应进入写入前预览");
    // 抽屉已收起，且没有任何写入发生。
    expect(container.querySelector('aside[aria-label="我的工具"]')).toBeNull();
  });

  it("WPS 工具编辑器的受控 JS 按当前勾选表生成计划，越界不覆盖草稿，保存仍重算资格", async () => {
    writeRangeHost();
    const { container } = await mountPlanPreview(formulaPlan, "用公式计算 A1");
    const dialog = await openSaveDialog(container);
    await click(dialogButton(dialog, "保存到我的工具"));
    await until(() => loadTools().length === 1, "应先保存工具");
    await click(container.querySelector(".tools-entry"));
    await click(findButton(container, "编辑计划"));
    const editor = container.querySelector<HTMLElement>('aside[aria-label="编辑工具计划模板"]')!;
    const input = editor.querySelector<HTMLTextAreaElement>('textarea[aria-label="受控 JS 代码"]');
    expect(input).not.toBeNull();
    const draft = editor.querySelector<HTMLTextAreaElement>(".plan-editor-textarea")!;
    const before = draft.value;
    await type(input, 'emit({type:"writeValues",sheet:"NotSelected",range:"A1",values:[["bad"]]});');
    await click(findButton(container, "生成待核对计划"));
    expect(textOf(editor)).toContain("未勾选");
    expect(draft.value).toBe(before);
    await type(input, 'for(const sheet of authorizedSheets) emit({type:"writeValues",sheet,range:"A1",values:[["JS值"]]});');
    await click(findButton(container, "生成待核对计划"));
    expect(JSON.parse(draft.value).actions).toEqual([{type:"writeValues",sheet:"Sheet1",range:"A1",values:[["JS值"]]}]);
    await click(findButton(container, "保存"));
    await until(() => textOf(container).includes("重新分析了保存资格"), "JS 草稿应沿现有保存路径重新分析");
    expect(loadTools()[0].approvals).toContain("fixedContent");
    expect(loadTools()[0].planTemplate.actions[0].type).toBe("writeValues");
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
    await until(() => textOf(container).includes("已撤销并恢复"), "撤销后应给出核对过的恢复结果");

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
describe("执行与审阅生命周期（先执行、后审阅、可选双视图）", () => {
  it("执行 → 单视图结果 → 确认：只写一次，确认不重放计划", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    expect(host.writes()).toBe(1);
    expect(findButton(container, "双视图对比")).not.toBeNull();
    const writes = host.writes();
    await click(findButton(container, "确认变更"));
    expect(host.writes()).toBe(writes);
    expect(textOf(container)).toContain("已确认变更");
    // 确认之后不再有审阅入口，但撤销入口仍在。
    expect(findButton(container, "确认变更")).toBeNull();
    expect(cardUndoButton(container)).not.toBeNull();
  });

  it("执行 → 双视图 → 确认 → 撤销：全程只有一次写入，撤销回到执行前", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await click(findButton(container, "双视图对比"));
    expect(host.writes()).toBe(1);
    await click(findButton(container, "确认变更"));
    expect(host.writes()).toBe(1);
    await click(cardUndoButton(container));
    await until(() => textOf(container).includes("已撤销并恢复"), "撤销应回读确认恢复");
    expect(host.cellValue()).toBe("旧值");
  });

  it("放弃 → 主表恢复 → 不再有任何确认入口", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await click(findButton(container, "放弃变更"));
    expect(host.cellValue()).toBe("旧值");
    expect(textOf(container)).toContain("已放弃并恢复");
    expect(findButton(container, "确认变更")).toBeNull();
    expect(findButton(container, "放弃变更")).toBeNull();
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

/*
 * 标记生命周期与可选双视图的**面板级**判据（宿主替身说明见各用例）。
 * 这里刻意覆盖"宿主没有 Copy / 没有 FormatConditions"的真实形态：本地计划与
 * WPS 测试替身都不提供这两样，正好证明"可选能力缺失不阻断主表执行"。
 */
describe("标记生命周期与可选双视图（面板级）", () => {
  it("宿主没有 Copy 接口：主表仍可执行并确认，双视图按钮只是不可用而不是写入前置", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await until(() => textOf(container).includes("本次执行计划"), "应生成执行计划");
    // 执行发生（替身没有 Worksheet.Copy）。
    expect(host.writes()).toBe(1);
    expect(host.cellValue()).toBe("90");
    // 标记接口也缺失：如实说明，但不改写执行结论。
    await until(() => container.querySelector(".result-review-bar") !== null, "应进入结果审阅");
    expect(textOf(container)).toContain("颜色标记未完整建立");
    // 确认仍然可用，且不会第二次写入。
    await click(findButton(container, "确认变更"));
    await until(() => textOf(container).includes("已确认变更"), "确认应完成");
    expect(host.writes()).toBe(1);
  });

  it("刷新后认出上次遗留的标记记录：读不到规则就保留记录、如实说明，不按区域清空", async () => {
    /*
     * 宿主替身没有 FormatConditions：无法证明规则是否还在，就不许声称"已清理"。
     * 这正是"归属不确定时不得按区域批量删"的可观测后果。
     */
    localStorage.setItem("excel-bro.wps.reviewMarks.v1", JSON.stringify({
      v: 1,
      // 与替身宿主同一个磁盘路径（用 join 拼，避免转义把判据写错）。
      workbookFullName: ["C:", "Data", "测试.xlsx"].join(String.fromCharCode(92)),
      formula: '=N("EB_REVIEW_legacy")=0',
      ranges: [{ sheet: "Sheet1", range: "A1" }]
    }));
    const { container } = await mount(panel());
    await until(() => textOf(container).includes("上次审阅"), "应识别并说明遗留标记");
    expect(textOf(container)).toContain("未清理");
    // 记录保留：下一次（有能力时）还能认领它。
    expect(localStorage.getItem("excel-bro.wps.reviewMarks.v1")).not.toBeNull();
  });
});

/*
 * 双视图把副本激活成 ActiveWorkbook 之后的真实形态。
 *
 * 模拟范围：宿主替身（值、工作表、条件格式、Workbooks 集合、ActiveWorkbook 切换）
 * + `stageWpsBeforeComparison`（只负责"建副本并激活"，其内容重建由
 * `wpsBeforeComparison.test.ts` 用真实模块覆盖）；模型、网络、账本仍为替身。
 * 这里刻意让 `CompareSideBySideWith`/`Copy` 之后**活动簿变成副本** —— 那正是实机行为。
 */
/*
 * 数据范围面板的**呈现层**判据必须跑在**真实** `WpsAgentPanel` 上。
 *
 * 为什么不放在 `wpsScopeDraft.test.tsx`：那一份把面板本体换成替身，替身自己在
 * JSX 里就写了 `.scope-layer`，于是"把面板塞回输入框那一层"的变异在那里永远
 * 变不了红 —— 判据看着存在，其实没有判别力。所以这里另挂一条真实组件的判据。
 */
describe("真实面板：数据范围面板本体住在独立呈现层", () => {
  it("面板本体不是 .composer 的后代；触发器那一行仍在输入框上方", async () => {
    const { container } = await mount(
      <WpsAgentPanel selectedSheetIndexes={[1]}
        scopeDock={<span>当前表 Sheet1</span>}
        scopeOverlay={<section className="scope-popover" aria-label="选择数据范围">
          <div className="scope-popover-scroll">正文</div>
        </section>} />
    );
    const popover = container.querySelector(".scope-popover");
    expect(popover).not.toBeNull();
    // 负向锚点：它不在输入框那一层里，而是外壳下 `.scope-layer` 的内容。
    expect(popover!.closest(".composer")).toBeNull();
    expect(popover!.closest(".scope-layer")).not.toBeNull();
    // 触发器仍留在 composer（那一行的位置没变）。
    expect(container.querySelector(".composer .scope-dock, .composer span")?.textContent).toBe("当前表 Sheet1");
  });
});

describe("双视图激活副本后仍按原表确认/放弃/核对", () => {
  function reviewHostWithCopy(faults: { borderless?: boolean; refuseDelete?: boolean; copyCountUnreadable?: boolean } = {}) {
    const base = writeRangeHost();
    /*
     * `writeRangeHost` 只回工具方法；宿主对象是它塞进 `getWpsApplication` 替身的
     * 那一个 —— 取回来才能把 `ActiveWorkbook` 改成"可切换"。
     */
    const app = getWpsApplicationMock() as unknown as {
      ActiveWorkbook: Record<string, unknown>;
      Workbooks?: { Count: number; Item(index: number): unknown };
      [key: string]: unknown;
    };
    const original = app.ActiveWorkbook as unknown as Record<string, unknown> & {
      Name: string; FullName: string; Path: string;
      Worksheets: { Count: number; Item(index: number): Record<string, unknown> };
    };
    const writeSheet = original.Worksheets.Item(1) as Record<string, unknown> & {
      Range: (...args: unknown[]) => Record<string, unknown>;
    };
    // 条件格式：标记真的能建、能删、能按随机公式认领。
    const rules: Array<{ formula: string; deleted: boolean }> = [];
    const originalRange = writeSheet.Range.bind(writeSheet);
    writeSheet.Range = (...args: unknown[]) => {
      const range = originalRange(...args) as Record<string, unknown>;
      if (!range.FormatConditions) {
        range.FormatConditions = {
          get Count() { return rules.filter(rule => !rule.deleted).length; },
          Add: (_type: number, _operator: unknown, formula: string) => {
            const state = { formula, deleted: false };
            rules.push(state);
            const remove = () => {
              if (faults.refuseDelete) throw new Error("宿主拒绝删除条件格式");
              state.deleted = true;
            };
            return {
              get Formula1() { return state.formula; },
              // `borderless`：规则建出来了、但没有边框接口 —— 标记在 "Add 之后" 失败。
              ...(faults.borderless ? {} : { Borders: { Item: () => ({}) } }),
              Interior: {},
              Delete: remove
            };
          },
          Item: (index: number) => {
            const live = rules.filter(rule => !rule.deleted)[index - 1];
            return {
              get Formula1() { return live?.formula; },
              Delete: () => {
                if (faults.refuseDelete) throw new Error("宿主拒绝删除条件格式");
                if (live) live.deleted = true;
              }
            };
          }
        };
      }
      return range;
    };

    /*
     * 副本表上的条件格式：`Worksheet.Copy()` 会把原表的规则一起带过来，所以副本上
     * 同时有"本次标记"和用户自建规则。`copyCountUnreadable` 复现宿主偶发的
     * "集合读得出 Item、却读不出条数" —— 清理函数必须把它报成"无法核实"。
     */
    const copyRules: Array<{ formula: string; deleted: boolean }> = [];
    const copySheet = {
      Name: "Sheet1", ProtectContents: false,
      Range: () => ({ FormatConditions: {
        ...(faults.copyCountUnreadable ? {} : {
          get Count() { return copyRules.filter(rule => !rule.deleted).length; }
        }),
        Item: (index: number) => {
          const live = copyRules.filter(rule => !rule.deleted)[index - 1];
          return { get Formula1() { return live?.formula; }, Delete: () => { if (live) live.deleted = true; } };
        }
      } })
    };
    const copy = {
      Name: "工作簿1", FullName: "工作簿1", Path: "", ReadOnly: false,
      Worksheets: { Count: 1, Item: () => copySheet },
      Close: () => true
    };
    let active: unknown = original;
    const activate = (book: unknown) => { active = book; };
    Object.defineProperty(app, "ActiveWorkbook", { get: () => active, configurable: true });
    Object.defineProperty(app, "ActiveSheet", { get: () => (active as { ActiveSheet?: unknown }).ActiveSheet ?? original.ActiveSheet, configurable: true });
    app.Workbooks = { Count: 2, Item: (index: number) => (index === 1 ? original : copy) };
    let copiesClosed = 0;
    const activateCopy = () => { active = copy; };
    const activateOriginal = () => { active = original; };
    return {
      base, app, original, copy, rules,
      copyRules, copySheet,
      activateCopy, activateOriginal, activate,
      closeCopy: () => { copiesClosed += 1; },
      copiesClosed: () => copiesClosed,
      liveRules: () => rules.filter(rule => !rule.deleted)
    };
  }

  it("并排后副本获得焦点：确认成功、只写一次、固定条收起、标记清掉", async () => {
    const host = reviewHostWithCopy();
    const original = host.original as unknown as { Name: string; FullName: string; Path: string; Worksheets: { Count: number; Item(i: number): { Name: string } } };
    vi.mocked(stageWpsBeforeComparison).mockImplementationOnce((() => {
      host.activateCopy();
      return {
        original,
        copy: host.copy,
        copyIdentity: { name: "工作簿1", fullName: "工作簿1", path: null },
        writeSheet: { index: 1, name: "Sheet1" },
        plan: {} as never,
        reviews: [], changedCells: [], originalBefore: [], copyFingerprints: [], copyAfter: []
      } as never;
    }) as never);
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    // 执行一次，走双视图（副本被激活）。
    await click(findButton(container, "双视图对比"));
    expect(host.liveRules().length).toBeGreaterThan(0);
    // 执行已经发生：状态文案不能再停在"尚未写入"。
    expect(textOf(container)).not.toContain("尚未写入");

    // 副本仍是活动簿时点确认：不得报"工作簿或授权发生变化"。
    await click(findButton(container, "确认变更"));
    await until(() => textOf(container).includes("已确认变更"), "确认应成功");
    expect(textOf(container)).not.toContain("工作簿或授权");
    expect(host.base.writes()).toBe(1);
    expect(host.liveRules()).toHaveLength(0);
    // 固定条已收起：既没有大框，也没有审阅按钮。
    expect(container.querySelector(".result-review-bar")).toBeNull();
    expect(findButton(container, "确认变更")).toBeNull();
    expect(findButton(container, "双视图对比")).toBeNull();
    // 业务状态留在消息里，保存工具与撤销入口仍在。
    expect(textOf(container)).toContain("已确认变更");
    // 结果卡里的撤销入口（带 aria-label 的那颗）仍在，且属于本次执行。
    expect(container.querySelector('[aria-label="撤销本次执行"]')).not.toBeNull();
  });

  /*
   * A1-F1 残留①：标记**部分失败**（规则建出来了、删不掉）时，本次公式必须仍被交给
   * 双视图，让副本能精确清掉本次残留；清不掉时要如实提示，而不是让"执行前副本"也黄成一片。
   */
  it("标记部分失败后打开双视图：把本次公式交给副本，并在无法核实清理时如实提示", async () => {
    const host = reviewHostWithCopy({ borderless: true, refuseDelete: true });
    const original = host.original as never;
    const handed: Array<string | undefined> = [];
    vi.mocked(stageWpsBeforeComparison).mockImplementationOnce(((
      _application: unknown, _before: unknown, formula?: string
    ) => {
      handed.push(formula);
      host.activateCopy();
      return {
        original, copy: host.copy,
        copyIdentity: { name: "工作簿1", fullName: "工作簿1", path: null },
        writeSheet: { index: 1, name: "Sheet1" },
        plan: {} as never, reviews: [], changedCells: [], originalBefore: [], copyFingerprints: [], copyAfter: [],
        // 副本侧集合读不出来 —— 面板必须据此提示"可能仍然同色"。
        reviewMarks: { cleared: 0, unconfirmed: 1 }
      } as never;
    }) as never);
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    // 业务写入已完成、审阅照常进入，只是标记没建全。
    await until(() => textOf(container).includes("颜色标记未完整建立"), "标记部分失败应如实说明");
    const residual = host.liveRules();
    expect(residual).toHaveLength(1);
    expect(host.base.writes()).toBe(1);

    await click(findButton(container, "双视图对比"));
    // 交给副本的正是**残留那一条规则的完整随机公式**，不是前缀、不是区域。
    expect(handed).toEqual([residual[0].formula]);
    expect(handed[0]).toMatch(/^=N\("EB_REVIEW_[0-9a-f]{32}"\)=0$/);
    await until(() => textOf(container).includes("未能核实清理"), "无法确认清理结果时必须明确提示");
    // 提示不改变业务结论：没有第二次写入，确认入口仍在。
    expect(host.base.writes()).toBe(1);
    expect(findButton(container, "确认变更")).not.toBeNull();
  });

  /*
   * A1-F1-a 判据：副本集合**读不出条数**时，清理必须报 unconfirmed，而这条结论
   * 要经**真实的** `comparisonNotice` 显示出来。上面那条用例给的是现成的
   * `reviewMarks` 字面量 —— 这里让结论由真实的清理函数算出来，面板只负责显示。
   */
  it("副本 Count 读不出来：真实清理结论经 comparisonNotice 显示在面板", async () => {
    const host = reviewHostWithCopy({ borderless: true, refuseDelete: true, copyCountUnreadable: true });
    const original = host.original as never;
    let copyResult: { cleared: number; unconfirmed: number } | undefined;
    vi.mocked(stageWpsBeforeComparison).mockImplementationOnce(((
      _application: unknown, _before: unknown, formula?: string
    ) => {
      host.activateCopy();
      // Copy() 的后果：副本上带着本次残留与用户规则，但集合读不出条数。
      host.copyRules.push(
        { formula: String(formula ?? ""), deleted: false },
        { formula: "=A1>0", deleted: false }
      );
      copyResult = clearReviewMarksFromCopySheet(host.copySheet as never, ["A1:B1"], String(formula ?? ""));
      return {
        original, copy: host.copy,
        copyIdentity: { name: "工作簿1", fullName: "工作簿1", path: null },
        writeSheet: { index: 1, name: "Sheet1" },
        plan: {} as never, reviews: [], changedCells: [], originalBefore: [], copyFingerprints: [], copyAfter: [],
        reviewMarks: copyResult
      } as never;
    }) as never);
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await until(() => textOf(container).includes("颜色标记未完整建立"), "标记部分失败应如实说明");
    await click(findButton(container, "双视图对比"));

    // 读不出条数就一条都不删：本次残留与用户规则都还在（禁止按区域批量删）。
    expect(copyResult).toEqual({ cleared: 0, unconfirmed: 1 });
    expect(host.copyRules.filter(rule => !rule.deleted)).toHaveLength(2);
    await until(() => textOf(container).includes("未能核实清理"), "无法核实清理必须显示在面板上");
    expect(textOf(container)).toContain("1 处区域");
    expect(host.base.writes()).toBe(1);
    expect(findButton(container, "确认变更")).not.toBeNull();
  });

  it("副本获得焦点时放弃：原表恢复、只清理本次副本", async () => {
    const host = reviewHostWithCopy();
    const original = host.original as never;
    vi.mocked(stageWpsBeforeComparison).mockImplementationOnce((() => {
      host.activateCopy();
      return {
        original, copy: host.copy,
        copyIdentity: { name: "工作簿1", fullName: "工作簿1", path: null },
        writeSheet: { index: 1, name: "Sheet1" },
        plan: {} as never, reviews: [], changedCells: [], originalBefore: [], copyFingerprints: [], copyAfter: []
      } as never;
    }) as never);
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await click(findButton(container, "双视图对比"));
    await click(findButton(container, "放弃变更"));
    await until(() => textOf(container).includes("已放弃并恢复"), "放弃应成功并恢复");
    expect(host.base.cellValue()).toBe("旧值");
    // 只清理本次副本，原表不被关闭；固定条收起。
    expect(discardWpsNativeComparison).toHaveBeenCalled();
    expect(container.querySelector(".result-review-bar")).toBeNull();
    expect(textOf(container)).not.toContain("工作簿或授权");
  });

  it("副本获得焦点时读数据范围：仍是原表与它的授权工作表", async () => {
    const host = reviewHostWithCopy();
    const original = host.original as never;
    vi.mocked(stageWpsBeforeComparison).mockImplementationOnce((() => {
      host.activateCopy();
      return {
        original, copy: host.copy,
        copyIdentity: { name: "工作簿1", fullName: "工作簿1", path: null },
        writeSheet: { index: 1, name: "Sheet1" },
        plan: {} as never, reviews: [], changedCells: [], originalBefore: [], copyFingerprints: [], copyAfter: []
      } as never;
    }) as never);
    await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await click(findButton(document.body, "双视图对比"));
    // 数据范围核对走真实 inspect：副本活动时解析为原表。
    const target = await inspectWpsAuthorizedTarget([1]);
    expect(target.name).toBe("测试.xlsx");
    expect(target.sheets.map(sheet => sheet.name)).toEqual(["Sheet1"]);
  });

  /*
   * 反向锚点：把副本认回原表，不等于"永远按原表办事"。
   * 切到一本**无关**工作簿时该拦的还是必须拦住；而且这是"当前性不成立"，
   * 不是"标记没清掉"——界面不许出现「重试清理标记」，那会把原因说错。
   */
  it("切到无关工作簿：确认仍被拦下，且不谎称是标记清理失败", async () => {
    const host = reviewHostWithCopy();
    const original = host.original as never;
    vi.mocked(stageWpsBeforeComparison).mockImplementationOnce((() => {
      host.activateCopy();
      return {
        original, copy: host.copy,
        copyIdentity: { name: "工作簿1", fullName: "工作簿1", path: null },
        writeSheet: { index: 1, name: "Sheet1" },
        plan: {} as never, reviews: [], changedCells: [], originalBefore: [], copyFingerprints: [], copyAfter: []
      } as never;
    }) as never);
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await click(findButton(container, "双视图对比"));
    expect(host.liveRules().length).toBeGreaterThan(0);
    // 用户自己切到另一本工作簿（不是本次副本）：这本不能被当成原表。
    host.activate({ Name: "别的.xlsx", FullName: ["C:", "Data", "别的.xlsx"].join(String.fromCharCode(92)),
      Path: ["C:", "Data"].join(String.fromCharCode(92)),
      Worksheets: { Count: 1, Item: () => ({ Name: "Sheet1" }) } });
    await click(findButton(container, "确认变更"));
    await until(() => textOf(container).includes("确认未完成"), "无关工作簿上确认必须被拦下");
    const bar = container.querySelector(".result-review-bar")?.textContent ?? "";
    expect(bar).toContain("工作簿、授权或计划已变化");
    expect(bar).not.toContain("可重试清理标记");
    expect(container.querySelector('[aria-label="重试清理标记"]')).toBeNull();
    // 拦下的是确认本身：没有第二次写入，标记也仍在（等回到原表再收尾）。
    expect(host.base.writes()).toBe(1);
    expect(host.liveRules().length).toBeGreaterThan(0);
    expect(textOf(container)).not.toContain("已确认变更");
  });
});


describe("WPS native worksheet checkpoint routing", () => {
  /*
   * 宿主侧的最小接线：`onAuthorizeTarget` 就是"把收窄后的范围写回**读取方**"。
   *
   * 生产里是 `WpsProbePanel` 的 `setSelectedIndexes`；测试里也必须接上 ——
   * 面板的 `inspect` 读的是 `selectedSheetIndexes` 这个受控属性，不接就永远读旧范围，
   * "收窄"只会是自欺：守卫基线动了、读取方没动，下一轮照样判成范围变化。
   */
  function PanelWithHostScope({ initial }: { initial: number[] }) {
    const [indexes, setIndexes] = useState(initial);
    return (
      <WpsAgentPanel
        selectedSheetIndexes={indexes}
        onAuthorizeTarget={(next) => setIndexes(next)}
        scopeDock={<span>全部表</span>}
      />
    );
  }

  it.each([false, true])("SDK delete plan uses one checkpoint batch, backup consent=%s", async needsConsent => {
    const app = clarificationHost({ sheetNames: ["Sheet1", "Sheet2"] }) as unknown as import("./wpsHost").WpsApplication;
    app.Hwnd = 123;
    getWpsApplicationMock.mockReturnValue(app);
    const plan: AnalysisPlan = { id: "delete-structure", title: "保留 Sheet1", summary: "删除 Sheet2", assumptions: [], warnings: [], actions: [{ type: "deleteWorksheet", sheet: "Sheet2" }] };
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({ kind: "plan", plan, turnId: "delete-turn" });
    validatePlanMock.mockImplementation(async submitted => submitted);
    /*
     * 目录**可变**的宿主：删表真的把表摘掉，撤销真的把表放回**原来那个位置**。
     * 同一个表名始终返回同一个包装对象（真宿主就是这样），否则撤销后的
     * "身份 + 顺序 + 内容"核对读到的只是刚造出来的一堆新对象，断言等于空转。
     */
    let names = ["Sheet1", "Sheet2"];
    // 执行前的目录顺序 —— 撤销要按它把表放回**原位**（服务端 `_anchor` 的同一口径）。
    const originalOrder = [...names];
    const sheetObjects = new Map<string, Record<string, unknown>>();
    const sheetFor = (name: string, index: number): Record<string, unknown> => {
      const cached = sheetObjects.get(name);
      if (cached) { cached.Index = index; return cached; }
      const created = { Name: name, Index: index, CodeName: `SheetCode${index}`, ProtectContents: false };
      sheetObjects.set(name, created);
      return created;
    };
    Object.defineProperty(app.ActiveWorkbook!, "Worksheets", {
      configurable: true,
      get: () => {
        const items = names.map((name, position) => sheetFor(name, position + 1));
        return { Count: items.length, Item: (index: number) => items[Number(index) - 1] };
      }
    });
    const executedOperationIds: string[] = [];
    vi.mocked(executeWorksheetDeletion).mockImplementation(async req => {
      expect(req.authorizedSheets).toEqual(["Sheet1", "Sheet2"]);
      expect(req.hostWindowId).toBe(123);
      expect(req.host).toBe("wps-et");
      // operationId 是执行时生成的（32 位 hex），不是计划 id。
      expect(req.operationId).toMatch(/^[a-f0-9]{32}$/);
      executedOperationIds.push(req.operationId);
      if (needsConsent) return { operationId: req.operationId, status: "needs_backup_consent", backupSheets: ["#EB_RULES", "#EB_RULES_BACKUP"], deletedSheets: [], remainingSheets: [], recoveryAvailable: false, message: "请确认额外备份许可" };
      names = names.filter(name => name !== "Sheet2");
      return { operationId: req.operationId, status: "completed", deletedSheets: ["Sheet2"], remainingSheets: ["Sheet1"], recoveryAvailable: true, message: "已核对剩余工作表。" };
    });
    vi.mocked(confirmWorksheetCheckpointBackup).mockImplementation(async req => {
      names = names.filter(name => name !== "Sheet2");
      return { operationId: req.operationId, status: "completed", deletedSheets: ["Sheet2"], remainingSheets: ["Sheet1", "#EB_RULES", "#EB_RULES_BACKUP"], recoveryAvailable: true, message: "已核对剩余工作表。" };
    });
    vi.mocked(restoreWorksheetCheckpoint).mockImplementation(async (operationId, deleted) => {
      // 原地撤销：按原顺序插回原位，原工作簿对象不变、不打开任何副本。
      const next = names.filter(name => !deleted.includes(name));
      for (const name of originalOrder.filter(name => deleted.includes(name))) {
        const predecessors = originalOrder.slice(0, originalOrder.indexOf(name)).filter(prev => next.includes(prev));
        const anchor = predecessors.at(-1);
        next.splice(anchor === undefined ? 0 : next.indexOf(anchor) + 1, 0, name);
      }
      names = next;
      return {
        operationId, status: "completed", restoredSheets: [...deleted],
        failedSheets: [], unverifiedSheets: [], repairedFormulas: 0, repairedNames: 0,
        limitations: [], message: `已把 ${deleted.length} 张工作表放回原工作簿。`
      } as never;
    });
    const { container } = await mount(<PanelWithHostScope initial={[1, 2]} />);
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型") ?? false, "模型加载");
    await type(container.querySelector("textarea"), "保留 Sheet1 删除其余表");
    await submit(container.querySelector("textarea"));
    if (needsConsent) {
      await until(() => findButton(container, "允许仅本地备份并继续") !== null, "需要主动询问额外备份许可");
      expect(app.ActiveWorkbook!.Worksheets.Count).toBe(2);
      expect(confirmWorksheetCheckpointBackup).not.toHaveBeenCalled();
      await click(findButton(container, "允许仅本地备份并继续"));
      expect(vi.mocked(confirmWorksheetCheckpointBackup).mock.calls[0][0].backupSheets).toEqual(["#EB_RULES", "#EB_RULES_BACKUP"]);
    }
    await until(() => findButton(container, "↶ 撤销本次执行") !== null, "删除结果必须有检查点恢复入口");
    expect(executeWorksheetDeletion).toHaveBeenCalledOnce();
    expect(registerWpsExecution).not.toHaveBeenCalled();
    expect(app.ActiveWorkbook!.Worksheets.Count).toBe(1);

    /*
     * 结果卡住在**发起这次删除的那条消息**里，不是聊天底部的常驻列表。
     * 旧实现在全部消息之后统一渲染检查点，换会话也清不掉。
     */
    const card = container.querySelector(".worksheet-checkpoint-card") as HTMLElement;
    expect(card).not.toBeNull();
    expect(card.closest(".message-row")).not.toBeNull();
    // 卡片与这一轮的过程时间线同属一条消息：它不是被追加到消息末尾的独立块。
    expect(card.closest(".message-row")?.querySelector(".task-progress-timeline")).not.toBeNull();
    /*
     * 入口的名字和位置与普通写入结果卡一致（右下角「↶ 撤销本次执行」），
     * 但这颗按钮**真的**在原工作簿里撤销，所以说明句写的是"保留后续编辑"，
     * 不再说"打开执行前副本" —— 副本入口在「备份与恢复」，两者不能互相冒充。
     */
    const undoFooter = card.querySelector(".execution-recovery-footer") as HTMLElement;
    expect(undoFooter).not.toBeNull();
    expect(undoFooter.querySelector("button")?.textContent?.trim()).toBe("↶ 撤销本次执行");
    expect(undoFooter.textContent).toContain("在原工作簿撤回本次删除，保留后续编辑");
    expect(undoFooter.textContent).not.toContain("打开执行前副本");
    expect(card.textContent).toContain("已删除 1 张工作表");
    expect(card.querySelector("details")?.open).toBe(false);
    /*
     * 顺序是「① 执行过程 → ② 执行结果 → ③ 卡底部的撤销按钮」。
     * 删表这一轮曾经把结果卡排在时间线**之前**，于是先看到"已完成"、再看到
     * 一步一步的过程，像结果倒着长回来。撤销按钮在卡内部，不可能跑到过程之上。
     */
    const timeline = card.closest(".message-row")!.querySelector(".task-progress-timeline") as HTMLElement;
    expect(timeline.compareDocumentPosition(card) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(card.parentElement!.querySelectorAll(".task-progress-timeline").length).toBe(1);

    // 原地撤销：不打开副本，表回到原簿的**原位置**，工作簿对象与文件名都不变。
    await act(async () => { findButton(container, "↶ 撤销本次执行")!.click(); });
    await until(() => textOf(container).includes("已撤销本次删除"), "撤销成功要落在同一张卡上");
    expect(vi.mocked(restoreWorksheetCheckpoint).mock.calls[0][0]).toBe(executedOperationIds.at(-1));
    expect(vi.mocked(restoreWorksheetCheckpoint).mock.calls[0][1]).toEqual(["Sheet2"]);
    expect(openWorksheetCheckpoint).not.toHaveBeenCalled();
    expect(names).toEqual(["Sheet1", "Sheet2"]);
    expect(app.ActiveWorkbook!.Name).toBe("测试.xlsx");
    expect(executeWorksheetDeletion).toHaveBeenCalledOnce();
    expect(findButton(container, "确认变更")).toBeNull();
    // 成功是终态：入口收起、结论写"已恢复 1 张"、不新增第二张卡。
    expect(findButton(container, "↶ 撤销本次执行")).toBeNull();
    expect(card.textContent).toContain("已恢复 1 张工作表");
    expect(container.querySelectorAll(".worksheet-checkpoint-card").length).toBe(1);
    expect(textOf(container)).not.toContain("已打开恢复副本");

    /*
     * 撤销把目录从 1 张变回 2 张。下一轮提问也不该被读成"用户改了授权范围" ——
     * 那会弹「工作表范围发生变化 / 尚未写入」，与刚撤销成功的事实直接冲突。
     * 放行必须有据：同一工作簿、目录里多出来的**只有本次放回的那张表**，
     * 授权集合本身没被扩大。
     */
    runTurnMock.mockResolvedValue({ kind: "answer", provider: "system", message: "现在有两张表。" });
    await until(() => container.querySelector<HTMLTextAreaElement>("textarea")?.disabled === false, "等待面板空闲");
    await type(container.querySelector("textarea"), "现在还剩几张表");
    await submit(container.querySelector("textarea"));
    await until(() => textOf(container).includes("现在有两张表。"), "追问照常回答");
    expect(container.querySelector(".target-scope-notice")).toBeNull();
    expect(textOf(container)).not.toContain("工作表范围发生变化");
    expect(container.querySelectorAll(".worksheet-checkpoint-card").length).toBe(1);
    expect(executeWorksheetDeletion).toHaveBeenCalledOnce();

    await until(() => container.querySelector<HTMLButtonElement>('[aria-label="新对话"]')?.disabled === false, "等待面板空闲");
    await act(async () => { container.querySelector<HTMLButtonElement>('[aria-label="新对话"]')!.click(); });
    await until(() => container.querySelector(".worksheet-checkpoint-card") === null, "新会话不继承旧删除卡");
    expect(executeWorksheetDeletion).toHaveBeenCalledOnce();
  });
  it.each(["99.0", ""])("does not send a native delete for blocked WPS version %s", async version => {
    const app = clarificationHost({ sheetNames: ["Sheet1", "Sheet2"] }) as unknown as import("./wpsHost").WpsApplication;
    app.Hwnd = 123; app.Version = version;
    getWpsApplicationMock.mockReturnValue(app);
    const plan: AnalysisPlan = { id: "blocked-delete", title: "保留 Sheet1", summary: "删除 Sheet2", assumptions: [], warnings: [], actions: [{ type: "deleteWorksheet", sheet: "Sheet2" }] };
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({ kind: "plan", plan, turnId: "blocked-delete-turn" });
    validatePlanMock.mockImplementation(async submitted => submitted);
    const { container } = await mount(<WpsAgentPanel selectedSheetIndexes={[1, 2]} scopeDock={<span>全部表</span>} />);
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型") ?? false, "模型加载");
    await type(container.querySelector("textarea"), "保留 Sheet1 删除其余表");
    await submit(container.querySelector("textarea"));
    await until(() => textOf(container).includes("未删除工作表"), "未开放的版本不得删除");
    expect(executeWorksheetDeletion).not.toHaveBeenCalled();
    expect(app.ActiveWorkbook!.Worksheets.Count).toBe(2);
  });

});

describe("A1-E1 结果引用过期验收入口（WPS 面板接线）", () => {
  /*
   * 这一端只负责"把 WPS 的宿主名与工作簿身份交给共享控件"。失效本身在服务端
   * （结果正文只在服务端登记器里），所以这里断言的是**接线存在且是失败关闭的**：
   * 读不到现场时入口仍出现，但启用按钮不可按 —— 不做"看不见就等于没事"的降级。
   */
  it("开发构建下渲染入口；读不到工作簿身份时不可启用", () => {
    getWpsApplicationMock.mockReturnValue({} as never);
    const markup = renderToStaticMarkup(panel());
    expect(markup).toContain("result-ref-fault-control");
    expect(markup).toContain("结果引用过期（开发专用）");
    const armButton = markup.match(/<button[^>]*>启用下一次引用过期<\/button>/)?.[0] ?? "";
    expect(armButton).toContain("disabled");
  });

  it("拿到真实工作簿时启用按钮可用，且只按工作簿身份构造请求", () => {
    const app = clarificationHost({ sheetNames: ["Sheet1"] }) as unknown as import("./wpsHost").WpsApplication;
    getWpsApplicationMock.mockReturnValue(app);
    const { container } = { container: { innerHTML: renderToStaticMarkup(panel()) } };
    const control = container.innerHTML;
    expect(control).toContain("result-ref-fault-control");
    // 有身份 ⇒ 不再禁用。
    const armButton = control.match(/<button[^>]*>启用下一次引用过期<\/button>/)?.[0] ?? "";
    expect(armButton).not.toContain("disabled");
    /*
     * 收起状态下**不播报**状态：`role="status"` 在面板里是"还有没有进度提示"的判据，
     * 一个常驻活区会让它失效。展开后才出现（见共享控件用例）。
     */
    expect(control).not.toContain('role="status"');
  });
});


/*
 * A1-E2 消费侧（WPS）：消费真实链的"引用过期恢复"工件，在真实面板上执行一次。
 *
 * 工件由 `scripts/drive_real_sdk_result_expiry.py` 在**真实模型 + 真实服务端 +
 * 真实登记器 + 真实 Docker** 上跑完后写出（`outputs/a1-e2-recovery.json`）。这里
 * 只消费它：旧引用（已被 A1-E1 精确失效）必须取不回来、新引用可读、计划引用的是
 * 新引用；计划在面板上**执行一次 → 回读 → 审阅**，确认不重执行、放弃恢复执行前。
 *
 * 门控：`EB_A1_E2=1`（缺工件即抛错，绝不拿旧工件/替身回放）。默认跑套件时跳过。
 */
type A1E2Cell = string | number | null;
type A1E2Artifact = {
  turnId: string;
  fault: { timeline: Array<{ phase: string; taskBound: boolean }>; final: { phase: string } };
  bridge: { prepareCount: number; refusals: unknown[] };
  registeredResultRefs: string[];
  probes: Array<{ resultRef: string; ok: boolean; status: number }>;
  repeatOldRef: { ok: boolean } | null;
  planActionCount: number;
  plan: {
    actions: Array<{ type: string; sheet: string; range: string; values: A1E2Cell[][] }>;
    localResultRef: string;
    acceptanceCriteria: Array<{ type: string; expected?: A1E2Cell[][] }>;
  };
};

function readA1E2Artifact(): A1E2Artifact {
  // 相对 vitest 的工作目录（`apps/wps-addin`）解析：`import.meta.url` 在 jsdom 里
  // 不是 file: URL，用它会抛错。
  return JSON.parse(readFileSync("../../outputs/a1-e2-recovery.json", "utf8")) as A1E2Artifact;
}

/** 按 A1 地址读写真实矩阵的 WPS 宿主替身（`Range(address)` 与面板同一调用形式）。 */
function a1e2Host(initial: Record<string, A1E2Cell[][]>) {
  const books: Record<string, A1E2Cell[][]> = {};
  for (const [name, rows] of Object.entries(initial)) books[name] = rows.map((row) => [...row]);
  /*
   * 公式矩阵必须可**写**：撤销/放弃时的逐动作恢复会写 `Range.Formula`，
   * 只给读的替身会让「放弃变更」永远返回 false（"本次结果未恢复"），
   * 而真实宿主是能写公式的 —— 那是替身缺能力，不是产品缺陷。
   */
  const formulas: Record<string, string[][]> = {};
  for (const [name, rows] of Object.entries(books)) {
    formulas[name] = rows.map((row) => row.map((cell) => String(cell ?? "")));
  }
  const writes: Array<{ sheet: string; address: string }> = [];
  /*
   * 宿主插件存储必须是真的：撤销记录写在这里，`inspectWpsUndoRecord` 也要从这里
   * 读回来核对。恒返回 null 的存储会让"确认变更"永远停在"无法核对"—— 那是替身
   * 缺能力，不是产品缺陷。
   */
  const storage = new Map<string, unknown>();
  const columnIndex = (letters: string) => {
    let value = 0;
    for (const char of letters) value = value * 26 + (char.charCodeAt(0) - 64);
    return value - 1;
  };
  const parse = (address: string) => {
    const [start, end = start] = address.replace(/\$/g, "").toUpperCase().split(":");
    const [, fromColumn, fromRow] = /^([A-Z]+)(\d+)$/.exec(start) ?? [];
    const [, toColumn, toRow] = /^([A-Z]+)(\d+)$/.exec(end) ?? [];
    const c1 = columnIndex(String(fromColumn));
    const r1 = Number(fromRow) - 1;
    const c2 = columnIndex(String(toColumn));
    const r2 = Number(toRow) - 1;
    return { r1, c1, rows: r2 - r1 + 1, columns: c2 - c1 + 1 };
  };
  const cellAddress = (row: number, column: number) => {
    let letters = "";
    for (let n = column; n > 0; n = Math.floor((n - 1) / 26)) {
      letters = String.fromCharCode(65 + (n - 1) % 26) + letters;
    }
    return `${letters}${row}`;
  };
  const rectOf = (sheetName: string, address: string) => {
    const { r1, c1, rows, columns } = parse(address);
    const range: Record<string, unknown> = {};
    Object.defineProperties(range, {
      Row: { value: r1 + 1 },
      Column: { value: c1 + 1 },
      Rows: { value: { Count: rows } },
      Columns: { value: { Count: columns } },
      MergeCells: { value: false },
      Item: { value: (row: number, column: number) =>
        rectOf(sheetName, cellAddress(r1 + row, c1 + column)) },
      Value2: {
        enumerable: true,
        get: () =>
          Array.from({ length: rows }, (_, row) =>
            Array.from({ length: columns }, (_, column) =>
              books[sheetName][r1 + row]?.[c1 + column] ?? null
            )
          ),
        set: (next: A1E2Cell[][]) => {
          writes.push({ sheet: sheetName, address });
          for (let row = 0; row < rows; row += 1) {
            books[sheetName][r1 + row] = books[sheetName][r1 + row] ?? [];
            formulas[sheetName][r1 + row] = formulas[sheetName][r1 + row] ?? [];
            for (let column = 0; column < columns; column += 1) {
              const value = next[row]?.[column] ?? null;
              books[sheetName][r1 + row][c1 + column] = value;
              formulas[sheetName][r1 + row][c1 + column] = String(value ?? "");
            }
          }
        }
      }
    });
    Object.defineProperty(range, "Formula", {
      enumerable: true,
      get: () =>
        Array.from({ length: rows }, (_, row) =>
          Array.from({ length: columns }, (_, column) =>
            formulas[sheetName][r1 + row]?.[c1 + column] ?? ""
          )
        ),
      /*
       * 写公式 = 写这一格的"内容"：真实宿主里 `Formula` 一写，显示值就跟着变。
       * 空公式（`""`）意味着**清空单元格**，读回来是空单元（`null`），不是空串 ——
       * 撤销靠 `setRangeFormulas(restorationMatrix(...))` 一次把值恢复回去，
       * 替身若把空公式记成空串，恢复后按值核对就会差一格。
       */
      set: (next: string[][]) => {
        for (let row = 0; row < rows; row += 1) {
          books[sheetName][r1 + row] = books[sheetName][r1 + row] ?? [];
          formulas[sheetName][r1 + row] = formulas[sheetName][r1 + row] ?? [];
          for (let column = 0; column < columns; column += 1) {
            const text = String(next[row]?.[column] ?? "");
            formulas[sheetName][r1 + row][c1 + column] = text;
            books[sheetName][r1 + row][c1 + column] = text === "" ? null : text;
          }
        }
      }
    });
    return range;
  };
  const sheets = Object.keys(books).map((name, position) => ({
    Name: name,
    Index: position + 1,
    CodeName: `SheetCode${position + 1}`,
    ProtectContents: false,
    UsedRange: rectOf(name, "A1:F9"),
    Range: (address: string | { Row: number; Column: number }, end?: { Row: number; Column: number }) =>
      rectOf(name, typeof address === "string" ? address :
        `${cellAddress(address.Row, address.Column)}:${cellAddress(end!.Row, end!.Column)}`)
  }));
  const application = {
    Version: "12.0",
    ActiveWorkbook: {
      Name: "会员.xlsx",
      FullName: "C:\\Data\\会员.xlsx",
      Path: "C:\\Data",
      ReadOnly: false,
      Worksheets: {
        Count: sheets.length,
        Item: (key: number | string) =>
          typeof key === "number" ? sheets[key - 1] : sheets.find((sheet) => sheet.Name === key)
      },
      ActiveSheet: sheets[0]
    },
    ActiveSheet: sheets[0],
    Worksheets: {
      Count: sheets.length,
      Item: (key: number | string) =>
        typeof key === "number" ? sheets[key - 1] : sheets.find((sheet) => sheet.Name === key)
    },
    PluginStorage: {
      getItem: (key: string) => storage.get(key) ?? null,
      setItem: (key: string, item: unknown) => {
        storage.set(key, item);
      },
      removeItem: (key: string) => {
        storage.delete(key);
      }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  };
  return { application, books, writes, storage };
}

describe.skipIf(process.env.EB_A1_E2 !== "1")("A1-E2 消费侧：恢复后的计划在真实 WPS 面板上执行一次", () => {
  /*
   * 工件**必须惰性读**：`describe.skipIf` 的工厂在"跳过"时仍会被调用来收集用例，
   * 在工厂里读文件会让"没有工件"的正常开发/CI 直接把整套用例带崩 —— 那是环境缺失
   * 污染了构建，不是产品缺陷。所以读完放在每个 `it` 里（那时才真的会跑）。
   */
  const artifactOf = () => readA1E2Artifact();

  function threeSheetPanel() {
    return (
      <WpsAgentPanel
        selectedSheetIndexes={[1, 2, 3]}
        scopeDock={<div className="eb-scope-dock">三张表已授权</div>}
      />
    );
  }

  async function runRecoveredPlan() {
    const artifact = artifactOf();
    const host = a1e2Host({
      会员表: [["会员编号", "会员", "等级"]],
      订单表: [["会员编号", "金额"]],
      汇总表: [[null]]
    });
    getWpsApplicationMock.mockReturnValue(host.application as never);
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    runTurnMock.mockResolvedValue({ kind: "plan", plan: artifact.plan as never });
    validatePlanMock.mockImplementation(async (submitted) => submitted);
    const mounted = await mount(threeSheetPanel());
    await until(
      () =>
        mounted.container
          .querySelector<HTMLButtonElement>(".model-picker-trigger")
          ?.title.includes("当前模型") ?? false,
      "应加载模型"
    );
    const textarea = mounted.container.querySelector("textarea");
    await type(textarea, "按会员等级汇总订单金额，结果写到汇总表");
    await submit(textarea);
    await until(
      () => mounted.container.querySelector(".result-review-bar") !== null,
      "计划到达后应自动执行一次并进入审阅"
    );
    return { ...mounted, host, artifact };
  }

  it("工件自洽：旧引用已失效、新引用可读、计划引用的是新引用", () => {
    const artifact = artifactOf();
    const plan = artifact.plan;
    const oldRef = artifact.registeredResultRefs[0];
    const newRef = artifact.registeredResultRefs[1];
    expect(artifact.registeredResultRefs).toHaveLength(2);
    expect(oldRef).toMatch(/^res_[0-9a-f]{32}$/);
    expect(newRef).toMatch(/^res_[0-9a-f]{32}$/);
    expect(oldRef).not.toBe(newRef);
    expect(plan.localResultRef).toBe(newRef);
    const phases = artifact.fault.timeline.map((item) => item.phase);
    expect(phases[0]).toBe("armed");
    expect(phases).toContain("bound");
    expect(phases[phases.length - 1]).toBe("consumed");
    const oldProbe = artifact.probes.find((probe) => probe.resultRef === oldRef);
    const newProbe = artifact.probes.find((probe) => probe.resultRef === newRef);
    expect(oldProbe?.ok).toBe(false);
    expect(oldProbe?.status).toBe(404);
    expect(newProbe?.ok).toBe(true);
    expect(artifact.repeatOldRef?.ok).toBe(false);
    // 过期之后是**重算**而不是重放读表：桥只读过一次表，也没有被拒的交值。
    expect(artifact.bridge.prepareCount).toBe(1);
    expect(artifact.bridge.refusals).toHaveLength(0);
    expect(plan.actions).toHaveLength(artifact.planActionCount);
  });

  it("执行一次 → 逐格回读 → 确认变更不重执行", async () => {
    const { container, host, artifact } = await runRecoveredPlan();
    const action = artifact.plan.actions[0];
    const plan = artifact.plan;
    // 执行一次：写入动作数 = 计划动作数（不是"有写入"就算）。
    expect(host.writes).toHaveLength(plan.actions.length);
    expect(
      host.books[action.sheet].map((row) => row.map((cell) => (cell === "" ? null : cell)))
    ).toEqual(action.values);
    // 来源表一格未变。
    expect(host.books["会员表"]).toEqual([["会员编号", "会员", "等级"]]);
    expect(host.books["订单表"]).toEqual([["会员编号", "金额"]]);

    const writesAfterExecution = host.writes.length;
    await click(findButton(container, "确认变更"));
    await until(() => textOf(container).includes("已确认变更"), "确认后应显示已确认状态");
    expect(host.writes).toHaveLength(writesAfterExecution);
  });

  it("放弃变更：目标区回到执行前，来源表全程未变", async () => {
    const { container, host, artifact } = await runRecoveredPlan();
    const action = artifact.plan.actions[0];
    await click(findButton(container, "放弃变更"));
    await until(
      () => textOf(container).includes("已放弃"),
      "放弃后应由产品自己声明已放弃（不是只看单元格变空）"
    );
    await until(
      () =>
        host.books[action.sheet].every((row) =>
          row.every((cell) => cell === null || cell === "")
        ),
      "放弃后目标区应回到执行前（空表）"
    );
    expect(host.books["会员表"]).toEqual([["会员编号", "会员", "等级"]]);
    expect(host.books["订单表"]).toEqual([["会员编号", "金额"]]);
  });
});

/*
 * A1-E2 反向分支之「来源在分析期间变化」（WPS 端）。
 *
 * 与「等待期间源值被改 ⇒ 桥只回 source_changed」那一条**不是同一层**：
 * 那一条是"模型还要再读一次"时被桥拦住；这一条是模型**已经不读了** ——
 * 它按旧来源算完、拿到新 resultRef、计划也编译成功了，只能靠**写前复核**拦住。
 *
 * 时序由屏障决定（`armBridgeTurn` 把计划扣住，放行前才改源），不用 sleep 抢。
 */
describe("A1-E2 来源变化分支：模型没有再次读取时，旧结果不得写入", () => {
  /** 旧来源算出来的固定值：金额 111/222/333 各按会员列成一行，合计 666。 */
  const SOURCE_MATRICES = {
    会员表: [
      ["会员编号", "会员", "等级"],
      ["7", "甲", "金卡"],
      ["8", "乙", "银卡"],
      ["9", "丙", "铜卡"]
    ],
    订单表: [
      ["会员编号", "金额"],
      ["7", 111],
      ["8", 222],
      ["9", 333]
    ],
    汇总表: [] as (string | number | null)[][]
  };
  const OLD_PLAN: AnalysisPlan = {
    id: "plan-source-change",
    title: "按会员汇总金额",
    summary: "按会员汇总订单金额",
    assumptions: [],
    warnings: [],
    actions: [
      {
        type: "writeValues",
        sheet: "汇总表",
        range: "A1:B4",
        values: [
          ["会员", "金额"],
          ["甲", 222],
          ["乙", 444],
          ["丙", 666]
        ]
      }
    ],
    acceptanceCriteria: []
  };

  it("计划按旧来源算完、来源在放行前被改：写入 0 次，目标区保持空", async () => {
    const host = bridgeRoundTripHost({ matrices: SOURCE_MATRICES });
    stubBridgeEndpoints();
    const releaseTurn = armBridgeTurn({
      kind: "plan",
      plan: OLD_PLAN,
      // 必须与桥绑定的 taskId 同一个轮次：锚点只认"这一轮"的现场。
      turnId: "turn-bridge-panel"
    });
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    validatePlanMock.mockImplementation(async (submitted) => submitted);

    const { container } = await mount(
      <WpsAgentPanel
        selectedSheetIndexes={[1, 2, 3]}
        scopeDock={<div className="eb-scope-dock">三张表已授权</div>}
      />
    );
    await until(
      () =>
        container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型") ?? false,
      "应加载模型目录"
    );
    const textarea = container.querySelector("textarea");
    await type(textarea, "按会员汇总订单金额，结果写到汇总表");
    await submit(textarea);

    // 首次读取完成（回执到齐）—— 模型手里已经是 111/222/333 那一版。
    await until(
      () => container.querySelector(".activity-log-card") !== null ||
        container.textContent?.includes("读取") === true,
      "应进入读取阶段"
    );
    const writesBefore = host.writes();

    // ── 屏障 ── 放行计划之前改源：222 → 250（身份、表名、授权都不变）。
    host.mutate("订单表", 3, 2, 250);
    expect(host.cell("订单表", 3, 2)).toBe(250);

    releaseTurn();
    await settle();

    /*
     * 判据顺序很重要：先等"产品自己说出拒绝"，再查写入次数 —— 反过来查
     * `writes === 0` 会在执行还没开始时就成立，变成假绿。
     */
    await until(
      () => container.textContent?.includes("数据来源在预览后发生了变化") === true,
      `应判来源变化并拒绝写入（当前新增写入 ${host.writes() - writesBefore} 次）`
    );
    expect(host.writes()).toBe(writesBefore);
    expect(container.textContent).toContain("未写入");
    // 目标区一格没动，来源表只保留那次外部修改。
    for (let row = 1; row <= 4; row += 1) {
      expect(host.cell("汇总表", row, 1)).toBeNull();
      expect(host.cell("汇总表", row, 2)).toBeNull();
    }
    expect(host.cell("会员表", 2, 2)).toBe("甲");
    expect(host.cell("订单表", 2, 2)).toBe(111);
  }, 30_000);
});


/*
 * A1-E2 B 路（WPS 端）：模型**确实**再次读取来源，被生产桥拒交值。
 *
 * 与上面「模型没有再次读取」那一条**不是同一层**：那一条只能靠写前复核；
 * 这一条要看的是"再读一次时，这次读取走生产桥的授权/来源复核并被 source_changed
 * 拒绝"。判定来自共享核心 `verifyBridgeAuthorization`（驱动通过
 * `scripts/bridge_authorization_gate.mjs` 直接跑那份 TS），不是驱动自己拼的字符串。
 *
 * 工件：`drive_real_sdk_result_expiry.py --scenario source-change-reread --host wps-et`。
 * 该轮**不启用**引用过期，于是旧引用仍有效、计划仍可能被编译 —— 迟到的旧计划
 * 有没有被写前复核拦下才有东西可验。
 *
 * 门控：`EB_A1_E2=1`；工件惰性读取（缺工件是环境问题，不能让收集阶段崩掉）。
 */
type A1E2ReReadArtifact = {
  scenario: string;
  turnId: string;
  promptLabel: string;
  fault: { armed: boolean };
  branch: string;
  bridge: {
    served: Array<{
      index: number;
      operationId: string;
      stepId: string;
      tool: string;
      refused: boolean;
      gate: { ok: boolean; errorCode?: string };
    }>;
    prepareCount: number;
    refusals: Array<{ index: number; code: string; operationId: string }>;
  };
  sourceChange: { afterServedCall: number; sheet: string; cell: string; before: number; after: number };
  secondRead: {
    attempted: boolean;
    refused: boolean;
    code: string | null;
    operationId: string | null;
    stepId: string | null;
    sameScope: boolean;
  };
  planSourceBasis: string;
  planActionCount: number;
  kind: string | null;
  hostWritesDuringPlanning: number;
  events: Array<{ stepId: string | null; status: string | null }>;
  plan: {
    actions: Array<{ type: string; sheet: string; range: string; values: A1E2Cell[][] }>;
    localResultRef: string;
    acceptanceCriteria: Array<{ type: string; expected?: A1E2Cell[][] }>;
  } | null;
};

function readA1E2ReReadArtifact(): A1E2ReReadArtifact {
  /*
   * 相对 vitest 的工作目录（`apps/wps-addin`）解析。`EB_A1_E2_REREAD_ARTIFACT`
   * 只用于**反向对照**：拿 A 路工件跑这组用例必须变红。
   */
  const path =
    process.env.EB_A1_E2_REREAD_ARTIFACT ?? "../../outputs/a1-e2-source-change-reread.json";
  return JSON.parse(readFileSync(path, "utf8")) as A1E2ReReadArtifact;
}

describe.skipIf(process.env.EB_A1_E2 !== "1")(
  "A1-E2 B 路（WPS）：模型再次读取被 source_changed 拒绝",
  () => {
    /** 与驱动同一份夹具（订单表 B2 = 100，改后 250）——改源只由测试驱动做。 */
    const SOURCE_ROWS = {
      会员表: [
        ["会员编号", "会员", "等级"],
        ["  7  ", "带空格会员", "金卡"],
        ["7", "重复键会员", "银卡"],
        ["007", "前导零会员", "银卡"],
        ["", "空键会员", "铜卡"],
        ["9", "正常会员", "金卡"]
      ],
      订单表: [
        ["会员编号", "金额"],
        ["7", 100],
        ["007", 200],
        ["9", 300],
        ["404", 400]
      ],
      汇总表: [] as A1E2Cell[][]
    };

    it("工件自洽：第二次读取是真实工具调用、被生产复核拒绝、拒绝前后写入 0 次", () => {
      const a = readA1E2ReReadArtifact();
      expect(a.scenario).toBe("source-change-reread");
      expect(a.promptLabel).toBe("two-read-verification");
      // 单独隔离"来源变化"：这一轮不启用引用过期。
      expect(a.fault.armed).toBe(false);
      expect(a.branch).toBe("B");

      const firstServed = a.bridge.served.find(
        (item) => item.tool === "prepare_python_data" && !item.refused
      );
      expect(firstServed).toBeDefined();
      expect(a.sourceChange.afterServedCall).toBe(firstServed!.index);

      const second = a.secondRead;
      expect(second.attempted).toBe(true);
      expect(second.refused).toBe(true);
      expect(second.code).toBe("source_changed");
      expect(second.sameScope).toBe(true);
      expect(second.operationId).toMatch(/^op_[0-9a-f]{32}$/);
      expect(second.operationId).not.toBe(firstServed!.operationId);
      expect(second.stepId).toBe(
        `bridge.prepare_python_data:${String(second.operationId).slice(3, 15)}`
      );
      const refusedCall = a.bridge.served.find((item) => item.operationId === second.operationId);
      expect(refusedCall?.gate).toEqual({ ok: false, errorCode: "source_changed" });

      const ledger = a.events.filter((event) => event.stepId === second.stepId);
      expect(ledger.map((event) => event.status)).toContain("warning");
      const firstLedger = a.events.filter((event) => event.stepId === firstServed!.stepId);
      expect(firstLedger.map((event) => event.status)).toContain("completed");

      expect(a.planSourceBasis).toBe("old");
      expect(a.hostWritesDuringPlanning).toBe(0);
    });

    it("迟到的旧计划：真实面板写前复核拒绝，写入 0 次，目标区保持空", async () => {
      const a = readA1E2ReReadArtifact();
      /*
       * 模型可能在第二次读取被拒后**仍然**按旧来源提交计划（这正是要验的），也可能
       * 如实停在"来源已变化"。两种都记下来：没提交就断言"没有计划"，不假装验过写入。
       */
      if (!a.plan) {
        expect(a.planActionCount).toBe(0);
        expect(a.kind).not.toBe("plan");
        return;
      }
      const host = bridgeRoundTripHost({ matrices: { ...SOURCE_ROWS } });
      stubBridgeEndpoints();
      const releaseTurn = armBridgeTurn({
        kind: "plan",
        plan: a.plan as unknown as AnalysisPlan,
        // 锚点只认"这一轮"的现场：与桥绑定的 taskId 必须一致。
        turnId: "turn-bridge-panel"
      });
      listModelsMock.mockResolvedValue({
        defaultModelId: firstConnection.id,
        models: [firstConnection]
      });
      validatePlanMock.mockImplementation(async (submitted) => submitted);

      const { container } = await mount(
        <WpsAgentPanel
          selectedSheetIndexes={[1, 2, 3]}
          scopeDock={<div className="eb-scope-dock">三张表已授权</div>}
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

      // 首次读取完成（锚点在那一刻固定）——模型手里已是 100 那一版。
      await until(
        () =>
          container.querySelector(".activity-log-card") !== null ||
          container.textContent?.includes("读取") === true,
        "应进入读取阶段"
      );
      const writesBefore = host.writes();

      // ── 屏障 ── 放行计划之前改源：订单表 B2 100 → 250（与工件同一格、同一个值）。
      host.mutate(a.sourceChange.sheet, 2, 2, a.sourceChange.after);
      expect(host.cell(a.sourceChange.sheet, 2, 2)).toBe(a.sourceChange.after);

      releaseTurn();
      await settle();

      // 先等产品自己说出拒绝，再查写入次数 —— 反过来查会假绿。
      await until(
        () => container.textContent?.includes("数据来源在预览后发生了变化") === true,
        `应判来源变化并拒绝写入（当前新增写入 ${host.writes() - writesBefore} 次）`
      );
      expect(host.writes()).toBe(writesBefore);
      expect(container.textContent).toContain("未写入");
      // 目标区一格没动（整块留空）：来源表只保留那次外部修改。
      for (let row = 1; row <= 12; row += 1) {
        for (let column = 1; column <= 8; column += 1) {
          expect(host.cell("汇总表", row, column)).toBeNull();
        }
      }
      expect(host.cell("会员表", 2, 2)).toBe("带空格会员");
      expect(host.cell("订单表", 2, 2)).toBe(a.sourceChange.after);
    }, 30_000);
  }
);

describe("SDK response provenance cannot replace a trusted bridge", () => {
  it("a response fingerprint without a source read cannot authorize fixed values", async () => {
    const host = writeRangeHost();
    planningTransport.openBridge = false;
    const { wpsAuthorizedSourceFingerprint } = await import("./wpsWorkbookReader");
    const {container} = await mountPlanPreview({...singleWritePlan,
      sourceFingerprint: wpsAuthorizedSourceFingerprint(getWpsApplication(), [1]),
      sourceFingerprintSheets: ["Sheet1"]}, "写入计算结果");
    expect(textOf(container)).toContain("没有记录它是按哪一版数据算出来的");
    expect(host.writes()).toBe(0);
    expect(findButton(container, "确认变更")).toBeNull();
  });

  it("a result from another turn cannot borrow this bridge's source anchor", async () => {
    const host = bridgeRoundTripHost();
    const bridge = stubBridgeEndpoints();
    const release = armBridgeTurn({kind:"plan", plan:singleWritePlan, turnId:"foreign-turn"});
    listModelsMock.mockResolvedValue({defaultModelId:firstConnection.id, models:[firstConnection]});
    validatePlanMock.mockImplementation(async submitted => submitted);
    const {container} = await mount(<WpsAgentPanel selectedSheetIndexes={[1,2]} scopeDock={<div>授权两张表</div>} />);
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型") ?? false, "model ready");
    await type(container.querySelector("textarea"), "读取后写入结果");
    await submit(container.querySelector("textarea"));
    await until(() => bridge.captured.some(item => item.url.includes("/receipt")), "source read receipt");
    release();
    await until(() => textOf(container).includes("没有记录它是按哪一版数据算出来的"), "foreign turn refused");
    expect(host.writes()).toBe(0);
    expect(findButton(container, "确认变更")).toBeNull();
  });
});

describe("WPS 面板：来源变化后的停止（A1-E3）", () => {
  /**
   * 服务端把 `source_changed` 判成**本次规划的终态**（409 / SOURCE_CHANGED /
   * retryable=false）。面板这侧的职责只有三件事，且都不许"重新解释"：
   *
   *  1. 把服务端那唯一一份文案**原文**给用户（不折成"任务处理失败"，也不写成
   *     "用户暂停"）；
   *  2. 本次执行**以失败收尾**，不留任何继续入口（旧任务已失效，继续它等于
   *     拿旧来源再算一次）；
   *  3. 不产生任何写入路径 —— 没有计划、没有预览卡、宿主写入 0 次。
   */
  const STOP_MESSAGE =
    "来源已变化，当前任务已停止：没有继续读取来源、没有继续计算，也没有写入。请按最新数据重新发起需求。";

  it("显示停止原因、不产生继续入口、不写入；重新发起仍可正常完成", async () => {
    const host = writeRangeHost();
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockRejectedValueOnce(
      new ApiRequestError(409, "SOURCE_CHANGED", STOP_MESSAGE, false)
    );

    const { container } = await mount(panel());
    await until(
      () => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false,
      "应加载模型目录"
    );
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "按部门统计平均分");
    await submit(textarea);

    await until(() => textOf(container).includes("来源已变化"), "应显示来源变化导致的停止");
    // ① 原文照转：文案只有服务端一份，面板不另写一句、也不降级成通用失败。
    expect(textOf(container)).toContain(STOP_MESSAGE);
    expect(textOf(container)).not.toContain("任务处理失败；本次未修改工作表。");
    // ② 不是"正常完成"、也不是"用户暂停"：以失败收尾，且没有继续入口。
    expect(container.querySelector(".task-progress-timeline")?.getAttribute("data-status")).toBe("failed");
    expect(findButton(container, "继续分析")).toBeNull();
    expect(container.querySelector(".wps-planning-pause-card")).toBeNull();
    // ③ 没有可被确认的写入：既没有计划预览，也没有宿主写入。
    expect(findButton(container, "确认变更")).toBeNull();
    expect(host.writes()).toBe(0);
    expect(host.cellValue()).toBe("旧值");
    // 本项目不存在"再执行一次"这种入口，确认也没有。
    expect(container.querySelector(".inline-plan")).toBeNull();

    // ④ 用户明确重新发起：新一轮按当前数据重新规划，正常完成。
    runTurnMock.mockResolvedValueOnce({
      kind: "answer",
      provider: "model",
      message: "已按最新数据重新统计",
      turnId: "turn-restarted"
    });
    const again = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(again, "按最新数据重新统计一次");
    await submit(again);
    await until(() => textOf(container).includes("已按最新数据重新统计"), "重新发起应正常完成");
    expect(runTurnMock).toHaveBeenCalledTimes(2);
    expect(host.writes()).toBe(0);
  });
});


describe("literal range edit without a source read", () => {
  /*
   * 反向对照：`noActionStarted` 是**正向证明**，"未执行"必须由四条判据一起支撑。
   * 少任何一条都只能保留本机待办与恢复入口 —— 把"可能已经动过原表"的执行说成
   * "未执行"，用户会照着界面去手工补数据，真正的现场就彻底乱了。
   *
   * 三条分别打在账本游标（已推进）、在途回执（第 0 步正在写）、执行权（挂起没成功）上；
   * 第四条打在宿主证据本身读不出来上。四条读的是不同事实，不能互相代替。
   */
  it.each([
    {label:"账本已推进到第二步，宿主记录却说回到了写前", suspend:{status:"failed" as const, nextActionIndex:1}},
    {label:"第 0 步在途、回执还没回来", suspend:{status:"failed" as const, inFlightActionIndex:0}},
    {label:"挂起失败：账本仍 running，本端仍握着执行权", suspend:{status:"running" as const, ownerActive:true}}
  ])("写前恢复准备失败但$label：保留本机待办，不许清成「未执行」",async({suspend})=>{
    const host=writeRangeHost(); planningTransport.openBridge=false;
    vi.mocked(saveExecutionRecoveryPackage).mockRejectedValueOnce(new ApiRequestError(503,"TASK_STATE_UNAVAILABLE","无法安全保存或读取任务状态，已停止执行以避免重复写入。",false));
    // 挂起回执如实反映这三种局面：账本游标、在途动作、执行权三处各有一处不干净。
    suspendWpsExecutionMock.mockImplementationOnce(async operationId=>({operationId,revision:2,
      workbookIdentity:"wps-path:c:\\data\\测试.xlsx",taskId:"task-test",descriptor:{actions:[]},
      authorizedSheetNames:["Sheet1"],totalActions:1,beforeSnapshot:null,checkpoint:null,result:null,
      error:null,canResume:true,requiresReconciliation:false,
      ownerToken:"owner-token-0123456789abcdef0123456789abcdef",nextActionIndex:0,
      inFlightActionIndex:null,ownerActive:false,...suspend} as never));
    const plan={...singleWritePlan,actions:[{type:"writeValues" as const,sheet:"Sheet1",range:"A1",values:[["完成"]]}]};
    const {container}=await mountPlanPreview(plan,"A1写入完成","literal_range_edit");
    // 前置事实：这次确实一格都没写（否则下面测的就不是"清不清"而是"写没写"）。
    expect(host.writes()).toBe(0); expect(host.cellValue()).toBe("旧值");
    expect(startWpsActionMock).not.toHaveBeenCalled();
    // 判据：本机待办还指向这次执行，恢复入口没被当成"未执行"抹掉。
    expect(localStorage.getItem("excel-bro.pending-wps-execution.v1")).not.toBeNull();
    expect(textOf(container)).not.toContain("原表未被修改");
  });

  it("宿主恢复点读不出来时不得当成「没写过」：保留本机待办",async()=>{
    let breakReads=false;
    const host=writeRangeHost({shouldBreakReads:()=>breakReads}); planningTransport.openBridge=false;
    vi.mocked(saveExecutionRecoveryPackage).mockImplementationOnce(async()=>{
      // 包没保存成功；同一刻宿主存储也读不动了 —— "有没有这条记录"根本无从得知。
      // 这与"读到了但没有记录"是两件事，不能合并成"未写入"。
      breakReads=true;
      throw new ApiRequestError(503,"TASK_STATE_UNAVAILABLE","无法安全保存或读取任务状态，已停止执行以避免重复写入。",false);
    });
    suspendWpsExecutionMock.mockImplementationOnce(async operationId=>({operationId,status:"failed",
      nextActionIndex:0,inFlightActionIndex:null,ownerActive:false} as never));
    const plan={...singleWritePlan,actions:[{type:"writeValues" as const,sheet:"Sheet1",range:"A1",values:[["完成"]]}]};
    const {container}=await mountPlanPreview(plan,"A1写入完成","literal_range_edit");
    expect(host.writes()).toBe(0); expect(host.cellValue()).toBe("旧值");
    // 证据读失败 ⇒ 结论只能是"未知"，不许折成"未执行"。
    expect(localStorage.getItem("excel-bro.pending-wps-execution.v1")).not.toBeNull();
    expect(textOf(container)).not.toContain("原表未被修改");
  });

  it("a failed prewrite package leaves zero host writes and no pending recovery card on reopen",async()=>{
    const host=writeRangeHost(); planningTransport.openBridge=false;
    vi.mocked(saveExecutionRecoveryPackage).mockRejectedValueOnce(new ApiRequestError(503,"TASK_STATE_UNAVAILABLE","无法安全保存或读取任务状态，已停止执行以避免重复写入。",false));
    suspendWpsExecutionMock.mockImplementationOnce(async operationId=>({operationId,status:"failed",nextActionIndex:0,
      inFlightActionIndex:null,ownerActive:false} as never));
    const plan={...singleWritePlan,actions:[{type:"writeValues" as const,sheet:"Sheet1",range:"A1",values:[["完成"]]}]};
    const {container}=await mountPlanPreview(plan,"A1写入完成","literal_range_edit");
    expect(host.writes()).toBe(0); expect(host.cellValue()).toBe("旧值");
    expect(startWpsActionMock).not.toHaveBeenCalled();
    expect(textOf(container)).toContain("无法安全保存或读取任务状态");
    expect(container.querySelector(".wps-write-recovery")).toBeNull();
    expect(localStorage.getItem("excel-bro.pending-wps-execution.v1")).toBeNull();
    await unmountMounted(container);
    const reopened=await mount(panel()); await settle();
    expect(reopened.container.querySelector(".wps-write-recovery")).toBeNull();
    expect(host.writes()).toBe(0);
  });
  it("the armed production panel stops after the second literal readback before its receipt", async () => {
    planningTransport.openBridge = false;
    const host = bridgeRoundTripHost({matrices:{Sheet1:[[null,null],[null,-101],[null,-202],[null,-303]]}});
    const identity = wpsWorkbookIdentity(getWpsApplication().ActiveWorkbook?.FullName);
    expect(identity).toBeTruthy();
    armRecoveryProbe(true, identity!);
    const register = registerWpsExecutionMock.getMockImplementation()!;
    registerWpsExecutionMock.mockImplementationOnce(async input => ({...await register(input),
      descriptor: input.descriptor, totalActions: input.descriptor.actions.length}));
    const plan: AnalysisPlan = {...singleWritePlan,id:"literal-probe-three",
      actions:[111,222,333].map((value,index)=>({type:"writeValues",sheet:"Sheet1",range:`B${index+2}`,values:[[value]]}))};
    const {container} = await mountPlanPreview(plan,"三个独立常量写入，不合并动作","literal_range_edit");
    expect([2,3,4].map(row=>host.cell("Sheet1",row,2))).toEqual([111,222,-303]);
    expect(host.writes()).toBe(2);
    expect(startWpsActionMock.mock.calls.map(call=>call[1].actionIndex)).toEqual([0,1]);
    expect(completeWpsActionMock.mock.calls.map(call=>call[1].actionIndex)).toEqual([0]);
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
    expect(readRecoveryProbe()?.phase).toBe("interrupted");
    expect(textOf(container)).toContain("第二步已写入并回读");
    expect(findButton(container,"确认变更")).toBeNull();
  });
  /*
   * 真机 2026-10-28 的收尾缺陷（同会话内）：
   * 中断 → 核对在途动作 → 继续剩余动作 → 确认变更 之后，`pendingPlan` 一直留着，
   * 新任务被"请先确认或取消当前写入计划，再开始新的任务。"挡住。
   * 用户当时靠「＋ → 放弃计划并新建」绕过去 —— 那是绕行，不是修复。
   *
   * 这里走的是同一条生产链路：真实开发探针在第二步回读之后、回执之前中断。
   */
  it.each(["确认变更", "放弃变更"])("同一会话中断→核对→续跑→%s之后可以开始新任务，且不重放已完成动作", async outcome => {
    planningTransport.openBridge = false;
    const host = bridgeRoundTripHost({matrices:{Sheet1:[[null,null],[null,-101],[null,-202],[null,-303]]}});
    const identity = wpsWorkbookIdentity(getWpsApplication().ActiveWorkbook?.FullName)!;
    armRecoveryProbe(true, identity);
    /*
     * 账本从登记那一刻起就按真机形状走：把登记返回的身份留下，之后的查询、补记、
     * 认领、动作回执都回同一份可变状态 —— 各 mock 各说一套会让判据失去意义。
     */
    let execution: Record<string, unknown> = {};
    let ledger: Record<string, unknown> = {};
    const snapshot = () => ({...execution, ...ledger});
    const register = registerWpsExecutionMock.getMockImplementation()!;
    registerWpsExecutionMock.mockImplementationOnce(async input => {
      const state = {...await register(input), descriptor: input.descriptor, totalActions: 3,
        workbookIdentity: input.workbookIdentity, taskId: input.taskId,
        authorizedSheetNames: input.authorizedSheetNames};
      execution = state as unknown as Record<string, unknown>;
      // 第二步已写、回执未完成：这就是 interrupt 之后账本该有的样子。
      ledger = {status: "needs_reconciliation", nextActionIndex: 1, inFlightActionIndex: 1,
        requiresReconciliation: true, ownerActive: false, ownerToken: null, revision: 6};
      return state as never;
    });
    startWpsActionMock.mockImplementation(async (_id, request) => {
      ledger = {...ledger, status: "running", inFlightActionIndex: request.actionIndex,
        ownerToken: "live-owner", revision: Number(ledger.revision) + 1};
      return snapshot() as never;
    });
    completeWpsActionMock.mockImplementation(async (_id, request) => {
      ledger = {...ledger, inFlightActionIndex: null, nextActionIndex: request.actionIndex + 1,
        revision: Number(ledger.revision) + 1};
      return snapshot() as never;
    });
    suspendWpsExecutionMock.mockImplementation(async () => {
      ledger = {...ledger, ownerToken: null, ownerActive: false};
      return snapshot() as never;
    });
    getWpsExecutionStatusMock.mockImplementation(async () => snapshot() as never);

    const plan: AnalysisPlan = {...singleWritePlan, id: "literal-after-interrupt",
      actions: [111,222,333].map((value,index)=>({type:"writeValues",sheet:"Sheet1",range:`B${index+2}`,values:[[value]]}))};
    const {container} = await mountPlanPreview(plan, "B2/B3/B4 依次写入 111/222/333", "literal_range_edit");
    // 中断现场：第二步已写入并回读，第三步没动。
    expect([2,3,4].map(row=>host.cell("Sheet1",row,2))).toEqual([111,222,-303]);
    expect(findButton(container,"确认变更")).toBeNull();

    await until(()=>findButton(container,"核对在途动作") !== null, "中断后必须在同一会话里摆出待核对入口");
    reconcileWpsActionReceiptMock.mockImplementation(async () => {
      ledger = {...ledger, status: "running", nextActionIndex: 2, inFlightActionIndex: null,
        requiresReconciliation: false, ownerToken: "reconciled-owner", revision: 8};
      return snapshot() as never;
    });
    await click(findButton(container,"核对在途动作"));
    await until(()=>findButton(container,"继续剩余动作") !== null, "补记第二步后应可继续第三步");
    // 补记只改账本，不写宿主。
    expect(host.writes()).toBe(2);

    claimWpsExecutionResumeMock.mockImplementation(async () => {
      ledger = {...ledger, ownerToken: "resume-owner", ownerActive: true, revision: 9};
      return snapshot() as never;
    });
    finishWpsExecutionMock.mockImplementation(async () => {
      ledger = {...ledger, status: "completed", ownerToken: null, ownerActive: false};
      return snapshot() as never;
    });
    await click(findButton(container,"继续剩余动作"));
    await until(()=>findButton(container,"确认变更") !== null, "续跑后必须进入审阅");
    // 只写第三步：前两步没有重放，也没有重新登记或重新规划。
    expect(host.writes()).toBe(3);
    expect([2,3,4].map(row=>host.cell("Sheet1",row,2))).toEqual([111,222,333]);
    expect(startWpsActionMock.mock.calls.map(call=>call[1].actionIndex)).toEqual([0,1,2]);
    expect(registerWpsExecutionMock).toHaveBeenCalledTimes(1);
    expect(runTurnMock).toHaveBeenCalledTimes(1);

    await click(findButton(container,outcome));
    await until(()=>textOf(container).includes(outcome === "确认变更" ? "已确认变更" : "已放弃并恢复"),
      "审阅应完成");
    /*
     * "不再执行一次"的判据落在**动作**上，不落在写入次数上：放弃要把整份写前值
     * 逐格恢复回去，宿主写入次数本来就会增加（那是撤销，不是重放计划）。
     */
    expect(startWpsActionMock.mock.calls.map(call=>call[1].actionIndex)).toEqual([0,1,2]);
    expect([2,3,4].map(row=>host.cell("Sheet1",row,2))).toEqual(
      outcome === "确认变更" ? [111,222,333] : [-101,-202,-303]);
    if (outcome === "确认变更") expect(host.writes()).toBe(3);

    /*
     * 收尾判据：这次执行已经做完，它自己的计划卡不许再挡新任务。
     * 关键不是"出现了某句话"，而是**新任务真的进了规划**（`runTurn` 第二次被调用）。
     */
    runTurnMock.mockResolvedValue({kind:"answer", message:"好的，换一件事。", provider:"model"});
    const textarea = container.querySelector("textarea");
    await type(textarea, "换一件事：把 C1 写成一个标题");
    await submit(textarea);
    await settle();
    expect(textOf(container)).not.toContain("请先确认或取消当前写入计划");
    expect(runTurnMock).toHaveBeenCalledTimes(2);
    // 计划卡仍留在对话里供事后复核，但不再是"待执行"（执行摘要卡，不是待执行卡）。
    expect(container.querySelector('[aria-label="本次执行的计划摘要"]')).not.toBeNull();
  });

  /*
   * 反面判据：修"收尾没清"绝不能顺手把闸门放宽。中断还没解决时，新任务必须
   * 仍然被挡住 —— 那时原表可能处于"写了一半"的现场，放行等于让它被下一份计划覆盖。
   */
  it("中断还没解决时，新任务仍被挡下（闸门不因修复而放宽）", async () => {
    planningTransport.openBridge = false;
    const host = bridgeRoundTripHost({matrices:{Sheet1:[[null,null],[null,-101],[null,-202],[null,-303]]}});
    armRecoveryProbe(true, wpsWorkbookIdentity(getWpsApplication().ActiveWorkbook?.FullName)!);
    const register = registerWpsExecutionMock.getMockImplementation()!;
    registerWpsExecutionMock.mockImplementationOnce(async input => ({
      ...await register(input), descriptor: input.descriptor, totalActions: 3,
      workbookIdentity: input.workbookIdentity, taskId: input.taskId,
      authorizedSheetNames: input.authorizedSheetNames}));
    const plan: AnalysisPlan = {...singleWritePlan, id: "literal-unresolved",
      actions: [111,222,333].map((value,index)=>({type:"writeValues",sheet:"Sheet1",range:`B${index+2}`,values:[[value]]}))};
    const {container} = await mountPlanPreview(plan, "B2/B3/B4 依次写入 111/222/333", "literal_range_edit");
    expect([2,3,4].map(row=>host.cell("Sheet1",row,2))).toEqual([111,222,-303]);
    await until(()=>container.querySelector(".wps-write-recovery") !== null, "中断后必须摆出恢复入口");

    runTurnMock.mockResolvedValue({kind:"answer", message:"不该走到这里", provider:"model"});
    const textarea = container.querySelector("textarea");
    await type(textarea, "再发一个新任务");
    await submit(textarea);
    await settle();
    expect(textOf(container)).toContain("请先确认或取消当前写入计划");
    // 没有进入规划、也没有再写一格。
    expect(runTurnMock).toHaveBeenCalledTimes(1);
    expect(host.writes()).toBe(2);
  });
  it("a failed prewrite attempt releases its probe binding without arming another task", async () => {
    planningTransport.openBridge = false;
    const host = bridgeRoundTripHost({matrices:{Sheet1:[[null,null],[null,-101],[null,-202],[null,-303]]}});
    armRecoveryProbe(true, wpsWorkbookIdentity(getWpsApplication().ActiveWorkbook?.FullName)!);
    const register = registerWpsExecutionMock.getMockImplementation()!;
    registerWpsExecutionMock.mockImplementationOnce(async input => ({...await register(input),
      descriptor: input.descriptor, totalActions: input.descriptor.actions.length}));
    vi.mocked(saveExecutionRecoveryPackage).mockRejectedValueOnce(new ApiRequestError(503,"TASK_STATE_UNAVAILABLE","恢复材料保存失败",false));
    const plan: AnalysisPlan = {...singleWritePlan,id:"literal-probe-failed",
      actions:[111,222,333].map((value,index)=>({type:"writeValues",sheet:"Sheet1",range:`B${index+2}`,values:[[value]]}))};
    const {container} = await mountPlanPreview(plan,"三个独立常量写入，不合并动作","literal_range_edit");
    expect(host.writes()).toBe(0);
    expect(readRecoveryProbe()?.phase).toBe("not_interrupted");
    expect(textOf(container)).toContain("验收中断未触发");
    expect(textOf(container)).not.toContain("已绑定本次执行，等待第二步回读");
    expect(findButton(container,"启用下一次第二步中断")?.disabled).toBe(false);
  });
  it.each(["confirm", "discard"])("executes three trusted literal actions in order without a source read before %s", async outcome => {
    planningTransport.openBridge = false;
    const host = bridgeRoundTripHost({matrices:{Sheet1:[[null,null],[null,-101],[null,-202],[null,-303]]}});
    const plan: AnalysisPlan = {...singleWritePlan,id:"literal-ordered-three",
      actions:[111,222,333].map((value,index)=>({type:"writeValues",sheet:"Sheet1",range:`B${index+2}`,values:[[value]]}))};
    const {container} = await mountPlanPreview(plan,"B2/B3/B4依次写入111/222/333，不合并动作","literal_range_edit");
    expect(findButton(container,"确认变更"),textOf(container)).not.toBeNull();
    expect([2,3,4].map(row=>host.cell("Sheet1",row,2))).toEqual([111,222,333]);
    expect(host.writes()).toBe(3);
    await click(findButton(container,outcome==="confirm"?"确认变更":"放弃变更"));
    await until(()=>textOf(container).includes(outcome==="confirm"?"已确认变更":"已放弃并恢复"),"审阅应完成");
    if(outcome==="confirm") expect(host.writes()).toBe(3);
    else expect([2,3,4].map(row=>host.cell("Sheet1",row,2))).toEqual([-101,-202,-303]);
  });
  it("executes the user literal once, then confirmation never rewrites", async () => {
    const host = writeRangeHost(); planningTransport.openBridge = false;
    const plan = {...singleWritePlan, actions: [{type: "writeValues" as const, sheet: "Sheet1", range: "A1", values: [["完成"]]}]};
    const {container} = await mountPlanPreview(plan, "仅在 Sheet1!A1 写入完成", "literal_range_edit");
    expect(host.cellValue()).toBe("完成"); expect(host.writes()).toBe(1);
    expect(textOf(container)).not.toContain("没有记录它是按哪一版");
    await click(findButton(container, "确认变更"));
    await until(() => textOf(container).includes("已确认变更"), "confirmation finishes");
    expect(host.writes()).toBe(1);
  });
  it("missing computed provenance refuses before execution without inventing recovery", async () => {
    const host = writeRangeHost(); planningTransport.openBridge = false;
    const {container} = await mountPlanPreview(singleWritePlan, "写入计算结果");
    expect(host.writes()).toBe(0); expect(host.cellValue()).toBe("旧值");
    expect(textOf(container)).toContain("未执行");
    expect(textOf(container)).not.toContain("结果无法确认");
    expect(textOf(container)).not.toContain("上次写入的恢复核对");
    const records = storedRecords();
    expect(records.length).toBeGreaterThan(0);
    expect(records.every(record => record.writeState === "not_written")).toBe(true);
    await unmountMounted(container);
    const reopened = await mount(<WpsAgentPanel selectedSheetIndexes={[1]} scopeDock={null} />);
    await settle();
    expect(textOf(reopened.container)).not.toContain("上次写入的恢复核对");
    expect(host.writes()).toBe(0);
  });
});
