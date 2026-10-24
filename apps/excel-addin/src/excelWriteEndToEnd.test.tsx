// @vitest-environment jsdom
/**
 * Excel 写侧端到端：从真实面板点击，到**真实宿主写入与回读**。
 *
 * 本文件用 `officeTestHost`（有真实状态、原表与审阅副本分账的 Office.js 替身）
 * 驱动真实 `<App/>`，覆盖两段此前从未被断言过的链路：
 *
 * 1. **原生并排审阅**：SDK 结果计划 → 生成副本 → 用户确认 → 写前复核 → 写入 →
 *    回读比对；
 * 2. **断点恢复登记的真实可达性**：授权范围／实际来源依赖／写入目标三者分开建模
 *    之后，来源与目标明确不重叠的计划必须**真的**走到 `registerOfficeExecution`，
 *    而登记失败、拒绝执行权、挂接索引落盘失败时原表必须一个单元格都不被写。
 *
 * 不 mock 的东西（这是本文件的全部意义）：
 * `executePlan` / `stageExcelNativeComparison` / `assertExcelNativeComparisonUnchanged`
 * / `captureWorkbookSourceFingerprint` / `verifyPlan` / `runPlan` / 工具桥
 * （`runBridgeUntilIdle` → `executeQueryTableTool` → `recordTrustedToolReads`）
 * / 撤销与核对（`reconcileOfficeRangeAction`）—— 全是产品代码，跑在替身宿主上。
 * 只 mock 网络（`./api`）与授权目标读取（避免弹窗）。
 *
 * 授权范围**不靠改 mock 造出来**，而是操作真实的范围选择弹层（`.scope-popover`）
 * 勾选并「确认」，与用户实际操作一致。
 *
 * 证据边界：本文件的全部结论都是**替身宿主自动化**证据。真实 Excel 实机证据见
 * `docs/SDK_WORKFLOW_DELIVERY_2026-10-09.md` §4.4（仍为未验证）。
 */
import { cleanup, fireEvent, render, waitFor } from "@testing-library/react";
import { act } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import { prepareExcelReviewLayout, arrangeExcelReviewLayout, executeWorksheetDeletion, openWorksheetCheckpoint, restoreWorksheetCheckpoint, confirmWorksheetCheckpointBackup } from "./api";
import { inspectOfficeAuthorizedTarget } from "./officeAuthorizedTarget";
import { checkIntent, streamAssistantResponse, understandMessage, continuePlanningTask, getPlanningTaskStatus } from "./api";
import {
  completeOfficeAction,
  finishOfficeExecution,
  getOfficeExecutionStatus,
  pollSdkBridgeTool,
  postSdkBridgeReceipt,
  reconcileOfficeActionReceipt,
  registerOfficeExecution,
  startOfficeAction,
  suspendOfficeExecution
} from "./api";
import {
  captureWorkbook,
  captureWorkbookCatalog,
  captureWorkbookStructure,
  currentOfficeWorkbookIdentity,
  invalidateWorkbookStructureCache
} from "./excel";
import { createOfficeTestHost, installOfficeTestHost, type TestOfficeHost } from "./officeTestHost";
import { loadPendingOfficeExecution, savePendingOfficeExecution } from "./officeOperationRecovery";
import { resetTrustedSourceReads } from "./planSourceDependency";

const SHEET = "综合票房";
/** 写入目标：与来源表不同的另一张**已授权**表。 */
const TARGET = "汇总表";
/** 工作簿里存在、但本次未勾选的工作表。 */
const UNAUTHORIZED_SHEET = "备用表";
const SHEET_IDS: Record<string, string> = {
  [SHEET]: "sheet-1",
  [TARGET]: "sheet-2",
  [UNAUTHORIZED_SHEET]: "sheet-3"
};
const ALL_SHEETS = [SHEET, TARGET, UNAUTHORIZED_SHEET];

/**
 * 授权目标读数的替身。
 *
 * `sheets` 是**文件里有哪些表**（目录），`selected` 才是本次授权的范围 ——
 * 与真实 `inspectOfficeAuthorizedTarget` 同语义：不能只返回被选中的表，否则
 * "授权在确认期间是否变化"就无从判定了。
 *
 * 目录**只有一个事实来源**：替身宿主的原表目录。删表用例调 `host.removeSheet()`
 * 让"原生删表已经发生"这件事真的落到宿主机型上，于是这里读到什么、
 * 产品 `scan()` 读到什么，永远是同一份。别在测试里再另抄一份"少了哪几张表"。
 */
function presentSheets(): string[] {
  return host ? host.sheetNames() : ALL_SHEETS;
}
/**
 * 目标读数的替身：`fileKey` / `fileRef` 一律取**宿主当下的真实读数**
 * （`currentOfficeWorkbookIdentity()` 与 `Office.context.document`），
 * 与真实 `inspectOfficeAuthorizedTarget` 同形 ——
 * `officeAnalysisTargetEvidence` 判的正是这两个值，替身若自己编一个身份串，
 * 证据那条判据就被替身替掉了（真实函数会直接返回 null，那是另一个故事）。
 */
function targetFor(names: readonly string[] | null) {
  const present = presentSheets();
  const sheets = present.map((name) => ({ id: SHEET_IDS[name], name }));
  // `null` = 跟随当前工作表：**活动表**是宿主的事实，不是测试写死的常量。
  const chosen = names === null ? [host ? host.activeSheetName() : SHEET] : [...names];
  return {
    fileKey: currentOfficeWorkbookIdentity() ?? null,
    fileRef: Office.context.document,
    name: "book.xlsx",
    sheets,
    selected: chosen
      .filter((name) => present.includes(name))
      .map((name) => SHEET_IDS[name])
      .filter(Boolean)
  };
}

vi.mock("./api", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./api")>()),
  checkHealth: vi.fn(async () => ({ status: "ok", modelConfigured: true })),
  listModels: vi.fn(async () => ({
    defaultModelId: "claude-agent-sdk",
    models: [
      {
        id: "claude-agent-sdk",
        label: "Claude Agent SDK",
        modelId: "claude-agent-sdk",
        baseUrl: "http://localhost",
        supportsVision: false,
        available: true
      }
    ]
  })),
  getModelSettings: vi.fn(async () => ({ connectionCount: 1 })),
  understandMessage: vi.fn(),
  checkIntent: vi.fn(),
  streamAssistantResponse: vi.fn(),
  getPlanningTaskStatus: vi.fn(),
  continuePlanningTask: vi.fn(),
  registerOfficeExecution: vi.fn(),
  startOfficeAction: vi.fn(),
  completeOfficeAction: vi.fn(),
  finishOfficeExecution: vi.fn(),
  suspendOfficeExecution: vi.fn(),
  getOfficeExecutionStatus: vi.fn(),
  claimOfficeExecutionResume: vi.fn(),
  reconcileOfficeActionReceipt: vi.fn(),
  // 工具桥：走真实状态机，只把"网络"换成替身（请求体与回执仍是产品代码拼的）。
  executeWorksheetDeletion: vi.fn(),
  confirmWorksheetCheckpointBackup: vi.fn(),
  openWorksheetCheckpoint: vi.fn(),
  restoreWorksheetCheckpoint: vi.fn(),
  pollSdkBridgeTool: vi.fn(),
  postSdkBridgeReceipt: vi.fn(),
  // 本机窗口排列服务：替身不模拟它，明确拒绝（产品会如实降级为"手动排列"）。
  arrangeExcelReviewLayout: vi.fn(),
  prepareExcelReviewLayout: vi.fn(async () => {
    throw new Error("test-host: 未提供本机窗口排列服务");
  })
}));

vi.mock("./officeAuthorizedTarget", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./officeAuthorizedTarget")>()),
  inspectOfficeAuthorizedTarget: vi.fn(async (names: readonly string[] | null) =>
    targetFor(names)
  ),
  inspectOfficeScopeAdmission: vi.fn(async (names: readonly string[]) => ({
    items: names.map((id) => ({ id, status: "ready", reason: "none" })),
    target: targetFor(names)
  })),
  /*
   * 证据函数**不替换**：`officeAnalysisTargetEvidence` 跑真实实现。
   * 它正是本轮缺陷所在（键的顺序敏感性），替身一旦给一个固定键，
   * 现场那个"两个键 ⇒ 证据冲突"就被替身悄悄抹掉了。
   */
  confirmOfficeAuthorizedScope: vi.fn(async (_names, confirm, expected) => confirm(expected))
}));

let BOOK_URL = "https://contoso.sharepoint.com/book.xlsx";
let host: TestOfficeHost;

function createHost(options: Partial<Parameters<typeof createOfficeTestHost>[0]> = {}): TestOfficeHost {
  return createOfficeTestHost({
    ...options,
    url: BOOK_URL,
    selectedRange: "A1:C3",
    activeSheet: SHEET,
    sheets: [
      {
        name: SHEET,
        values: [
          ["影片", "票房(万)", "场次"],
          ["甲", 1200, 30],
          ["乙", 980, 22]
        ]
      },
      // 来源之外的目标表（空表；写入目标不需要预先有数据）。
      { name: TARGET, values: [] },
      // 工作簿里确实存在、但本次没被勾选授权 —— 授权边界的判据要用到它。
      { name: UNAUTHORIZED_SHEET, values: [] }
    ]
  });
}

/** 恢复登记成功时的服务端状态（形状与 `./api` 契约一致）。 */
function officeLedgerState(overrides: Record<string, unknown> = {}) {
  return {
    operationId: "op-1",
    ownerToken: "owner-token",
    revision: 1,
    nextActionIndex: 0,
    status: "running",
    requiresReconciliation: false,
    checkpoint: null,
    ...overrides
  };
}

beforeEach(() => {
  BOOK_URL = "https://contoso.sharepoint.com/book.xlsx";
  localStorage.clear();
  vi.clearAllMocks();
  invalidateWorkbookStructureCache();
  resetTrustedSourceReads();
  host = createHost();
  installOfficeTestHost(host);
  // 账本默认是"登记成功"；需要验证失败路径的用例再逐条覆盖。
  vi.mocked(registerOfficeExecution).mockResolvedValue(officeLedgerState() as never);
  vi.mocked(startOfficeAction).mockResolvedValue(officeLedgerState() as never);
  vi.mocked(completeOfficeAction).mockResolvedValue(officeLedgerState() as never);
  vi.mocked(finishOfficeExecution).mockResolvedValue(
    officeLedgerState({ status: "finished" }) as never
  );
  vi.mocked(suspendOfficeExecution).mockResolvedValue(
    officeLedgerState({ status: "suspended" }) as never
  );
  // 桥默认没有待办：只有明确要读宿主的用例才投递调用。
  vi.mocked(pollSdkBridgeTool).mockResolvedValue({
    calls: [],
    generation: 0
  } as never);
  vi.mocked(postSdkBridgeReceipt).mockResolvedValue({
    status: "accepted",
    code: null
  } as never);
  vi.spyOn(console, "log").mockImplementation(() => {});
  Element.prototype.scrollIntoView = vi.fn();
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/** 原生审阅可用的计划：单张授权表内两个互不重叠的写值区域。 */
const reviewPlan = {
  id: "plan-review-1",
  title: "补齐上座率",
  summary: "在 E:F 补上座率，并在 G1 加合计标题",
  assumptions: [],
  warnings: [],
  actions: [
    {
      type: "writeValues" as const,
      sheet: SHEET,
      range: "E1:F3",
      values: [
        ["影片", "上座率"],
        ["甲", 0.42],
        ["乙", 0.31]
      ]
    },
    {
      type: "writeValues" as const,
      sheet: SHEET,
      range: "G1",
      values: [["合计"]]
    }
  ],
  /**
   * 显式验收条件：`verificationGaps` 只看**计划声明的**验收条件（inferredCriteria
   * 里的推断条件不算），没有它每步都会落进"缺少独立验收"。真实计划由编译器补上。
   */
  acceptanceCriteria: [
    {
      type: "rangeEquals" as const,
      sheet: SHEET,
      range: "E1:F3",
      expected: [
        ["影片", "上座率"],
        ["甲", 0.42],
        ["乙", 0.31]
      ]
    },
    { type: "rangeEquals" as const, sheet: SHEET, range: "G1", expected: [["合计"]] }
  ]
};

/**
 * 走「执行到 Excel」直接执行的计划。
 *
 * 为什么故意做成"两个重叠区域"：产品在计划可进原生审阅时会**只**渲染审阅卡
 * （不渲染「执行到 Excel」按钮）。要覆盖直接执行路径，就得让计划落在
 * `nativeComparisonPlanIssue` 的拒绝面上；重叠区域是其中最贴近真实的一种
 * （"表头再确认一次"），同时它仍是烘焙矩阵（`writeValues`），
 * 于是来源指纹判据照样生效。
 */
const directPlan = {
  id: "plan-direct-1",
  title: "改写场次",
  summary: "重写场次表头与数值",
  assumptions: [],
  warnings: [],
  actions: [
    {
      type: "writeValues" as const,
      sheet: SHEET,
      range: "C1:C3",
      values: [["场次"], [30], [22]]
    },
    {
      type: "writeValues" as const,
      sheet: SHEET,
      range: "C1",
      values: [["场次"]]
    }
  ]
};

/** 动作落在本次**未授权**（但工作簿里存在）的表上：执行期拒绝，一个单元格都不写。 */
const offSourcePlan = {
  id: "plan-off-source-1",
  title: "写到未授权表",
  summary: "把结果写到另一张表",
  assumptions: [],
  warnings: [],
  actions: [
    {
      type: "writeValues" as const,
      sheet: UNAUTHORIZED_SHEET,
      range: "A1",
      values: [["x"]]
    }
  ]
};

/**
 * 写**另一张表**（来源与目标明确不重叠）：恢复登记应当真的可达。
 *
 * 计划本身仍是烘焙矩阵（`writeValues`）—— 服务端编译出来的结果计划就是这个形状。
 */
const targetPlan = {
  id: "plan-target-1",
  title: "汇总到结果表",
  summary: "把「综合票房」的合计写到「汇总表」",
  assumptions: [],
  warnings: [],
  actions: [
    {
      type: "writeValues" as const,
      sheet: TARGET,
      range: "A1:B2",
      values: [
        ["影片", "合计"],
        ["甲", 1200]
      ]
    }
  ],
  acceptanceCriteria: [
    {
      type: "rangeEquals" as const,
      sheet: TARGET,
      range: "A1:B2",
      expected: [
        ["影片", "合计"],
        ["甲", 1200]
      ]
    }
  ]
};

/**
 * 两个动作的 **live** 计划（公式在执行那一刻求值）：
 * 用于"回执丢失后先核对、不重复执行已完成写入"。
 *
 * 为什么用公式计划而不是烘焙矩阵：烘焙矩阵的来源指纹覆盖本次授权范围（含写入
 * 目标），任务自己写过的目标表会让指纹变化，续跑前会要求重新确认固定值 —— 那是
 * 既有指纹口径的后果，本轮不改变它，细节见交付报告的限制清单。
 */
const liveTargetPlan = {
  id: "plan-target-live-1",
  title: "汇总合计（公式）",
  summary: "把合计公式写到「汇总表」",
  assumptions: [],
  warnings: [],
  actions: [
    {
      type: "writeFormulas" as const,
      sheet: TARGET,
      range: "A1:B1",
      formulas: [["=\"影片\"", "=1200"]]
    },
    {
      type: "writeFormulas" as const,
      sheet: TARGET,
      range: "A2:B2",
      formulas: [["=\"乙\"", "=980"]]
    }
  ]
};

/** 桥上的确定性读取：`scope: "active"` 只读当前工作表（= 来源依赖只含它）。 */
function bridgeQuery(scope: "active" | "selected") {
  return {
    mode: "rows" as const,
    scope,
    fields: ["影片"],
    limit: 10
  };
}

const BRIDGE_OP = `op_${"7".repeat(32)}`;
const BRIDGE_FP = "c".repeat(64);

interface PlanTurnOptions {
  pauseBeforePlan?: boolean;
  probe?: boolean;
  /** 通过真实范围选择弹层建立的授权范围。 */
  authorize?: string[];
  /** 让本轮真的经工具桥读一次宿主，从而产生可信来源记录。 */
  read?: { scope: "active" | "selected" };
}

/**
 * 送出一条需求并等计划卡出现。
 *
 * `options.read` 时走**真实**工具桥：规划流里投递 `bridgeTool` 指针 → 状态机
 * 轮询 → 复核当前授权 → 在替身宿主上跑既有 `query_table` → 回执，然后才返回
 * 计划。也就是说"计划读过谁"是**真的读过**，不是测试写进去的。
 */
async function sendPlan(
  plan: unknown,
  prompt = "把合计写到汇总表",
  options: PlanTurnOptions = {}
) {
  vi.mocked(understandMessage).mockResolvedValue({
    kind: "needs_workbook_context",
    summary: prompt,
    provider: "model"
  });
  vi.mocked(checkIntent).mockResolvedValue({
    kind: "proceed",
    summary: prompt,
    confirmedPrompt: prompt,
    provider: "model"
  });
  if (options.read) {
    vi.mocked(pollSdkBridgeTool).mockResolvedValue({
      calls: [
        {
          operationId: BRIDGE_OP,
          tool: "query_table",
          arguments: bridgeQuery(options.read.scope),
          requestFingerprint: BRIDGE_FP,
          expiresAt: 0
        }
      ],
      generation: 1
    } as never);
  }
  vi.mocked(streamAssistantResponse).mockImplementation(async (_request, hooks) => {
    if (options.read) {
      hooks?.onProgress?.({
        schemaVersion: 1,
        operationId: "turn-bridge-read",
        sequence: 1,
        timestamp: new Date().toISOString(),
        phase: "planning",
        stepId: "plan.tool-bridge",
        status: "progress",
        title: "正在读取授权范围",
        summary: "只在本机读取已授权工作表。",
        bridgeTool: { operationId: BRIDGE_OP, tool: "query_table" }
      });
      // 等这一次本地查询真的跑完（回执是执行之后发出的）。
      await waitFor(() =>
        expect(vi.mocked(postSdkBridgeReceipt).mock.calls.length).toBeGreaterThan(0)
      );
      const receipt = vi.mocked(postSdkBridgeReceipt).mock.calls[0]?.[0] as {
        workbookIdentity?: string;
        sourceFingerprint?: string | null;
        receipt?: { result?: { sourceSheets?: string[] }; errorCode?: string };
      };
      expect(
        receipt.receipt?.result?.sourceSheets,
        `工具执行记录必须真的读到了表（收到的是 ${JSON.stringify(receipt.receipt)}）`
      ).toBeDefined();
      /*
       * 三处必须是**同一个值**：规划请求里带的工作簿身份/来源指纹（服务端据此建立
       * 桥绑定与授权归属，并拿它与回执里的指纹比）、桥回执回带的那一份、以及执行前
       * 复核要重读的那一份。请求与回执只要差一点，服务端就会把回执判成
       * `source_changed`，模型永远拿不到结果值 —— 也就永远没有可信的"实际来源依赖"。
       */
      expect(_request.workbook.workbookIdentity).toBe(receipt.workbookIdentity);
      expect(_request.workbook.sourceFingerprint).toBe(receipt.sourceFingerprint);
      // 而且是**真身份 / 真读出来的指纹**，不是"同名兜底"或快照指纹。
      expect(receipt.workbookIdentity).toBe(`url:${BOOK_URL}`);
      expect(receipt.sourceFingerprint).toMatch(/^fnv1a32x2:/);
    }
    if (options.pauseBeforePlan) return {
      kind: "planning_pause", turnId: "resumed-write-turn", revision: 1, lastSequence: 0,
      reason: "规划已暂停", pauseType: "interrupted", extensionSeconds: 60,
      activeElapsedSeconds: 1, currentModelId: "claude-agent-sdk", modelSwitchCount: 0,
      canSwitchModel: false, availableModelConnections: []
    } as never;
    return { kind: "plan", provider: "model", plan: plan as never };
  });

  const view = render(<App />);
  await waitFor(() =>
    expect(view.container.querySelector(".composer textarea")).not.toBeNull()
  );
  await waitFor(() =>
    expect(
      view.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title
    ).toContain("当前模型")
  );
  if (options.authorize) await setAuthorizedScope(view, options.authorize);
  if (options.probe) {
    view.container.querySelector("details.recovery-probe-control")!.setAttribute("open", "");
    fireEvent.click(await view.findByRole("button", {name:"启用下一次第二步中断"}));
  }
  fireEvent.change(view.container.querySelector(".composer textarea")!, {
    target: { value: prompt }
  });
  fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
  if (options.pauseBeforePlan) await view.findByRole("button", { name: /继续分析/ });
  else await waitFor(() => expect(view.container.querySelector(".inline-plan")).not.toBeNull());
  return view;
}

/**
 * 操作**真实**的范围选择弹层建立授权：打开 → 选「选择多个」→ 勾选 → 「完成」。
 *
 * 不靠改 mock 造授权：`selectedNamesFor` 读的是弹层的草稿，`committedSheetNames`
 * 只在点「确认」并经目标守卫复核后才推进，两者都是产品实现。
 */
async function openScopePicker(view: ReturnType<typeof render>) {
  const trigger = view.container.querySelector<HTMLButtonElement>(".scope-trigger");
  if (!trigger) throw new Error("范围选择入口尚未出现");
  fireEvent.click(trigger);
  await waitFor(() =>
    expect(view.container.querySelector(".scope-popover")).not.toBeNull()
  );
}

async function setAuthorizedScope(view: ReturnType<typeof render>, names: readonly string[]) {
  await openScopePicker(view);
  const manual = [...view.container.querySelectorAll<HTMLButtonElement>(".scope-mode-list button")]
    .find((candidate) => candidate.textContent?.includes("选择多个"));
  if (!manual) throw new Error("「选择多个」入口尚未出现");
  fireEvent.click(manual);
  for (const name of ALL_SHEETS) {
    const button = sheetToggle(view, name);
    const selected = button.getAttribute("aria-pressed") === "true";
    if (selected !== names.includes(name)) fireEvent.click(button);
  }
  const commit = view.container.querySelector<HTMLButtonElement>(".scope-close");
  if (!commit) throw new Error("范围「确认」按钮尚未出现");
  await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
  fireEvent.click(commit);
  await waitFor(() =>
    expect(view.container.querySelector(".scope-popover")).toBeNull()
  );
}

/**
 * 按**指定顺序**确认多表。
 *
 * 目录顺序是 `ALL_SHEETS`（综合票房 → 汇总表 → 备用表），`names` 故意给成另一个
 * 顺序 —— 现场就是这么复现的：用户先点最后一张表。
 *
 * 先全部取消再依次勾上，这样草稿顺序**确定地**等于 `names`（`toggleSheet` 是追加）。
 */
async function setAuthorizedScopeInOrder(view: ReturnType<typeof render>, names: readonly string[]) {
  await openScopePicker(view);
  const manual = [...view.container.querySelectorAll<HTMLButtonElement>(".scope-mode-list button")]
    .find((candidate) => candidate.textContent?.includes("选择多个"));
  if (!manual) throw new Error("「选择多个」入口尚未出现");
  fireEvent.click(manual);
  for (const name of ALL_SHEETS) {
    const button = sheetToggle(view, name);
    if (button.getAttribute("aria-pressed") === "true") fireEvent.click(button);
  }
  for (const name of names) fireEvent.click(sheetToggle(view, name));
  const commit = view.container.querySelector<HTMLButtonElement>(".scope-close");
  if (!commit) throw new Error("范围「确认」按钮尚未出现");
  await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
  fireEvent.click(commit);
  await waitFor(() =>
    expect(view.container.querySelector(".scope-popover")).toBeNull()
  );
}

function sheetToggle(view: ReturnType<typeof render>, name: string) {
  const button = [...view.container.querySelectorAll<HTMLButtonElement>(".workbook-sheet-list button")]
    .find((candidate) => candidate.querySelector("strong")?.textContent === name);
  if (!button) throw new Error(`工作表「${name}」的勾选项尚未出现`);
  return button;
}

/** 只统计**单元格内容**写入，格式写（副本标色）单独看。 */
const contentWrites = (writes: ReturnType<TestOfficeHost["writes"]>) =>
  writes.filter((write) => write.kind !== "format");

async function stageReview(view: ReturnType<typeof render>) {
  await waitFor(() => expect(view.queryByRole("button", { name: "确认变更" })).not.toBeNull());
}
async function clickSave(view: ReturnType<typeof render>) {
  const save = view.getByRole("button", { name: "确认变更" });
  fireEvent.click(save);
  /*
   * 成功终态下固定条会**收起**（不再占着输入框上方），所以"确认成了"的观测点
   * 改成对话里的收尾消息 —— 那是状态真正的去处，也是这条判据想证的事。
   * 顺带钉住：收起的同时固定条不再提供任何审阅按钮。
   */
  await waitFor(() => expect(text(view)).toContain("已确认变更"));
  expect(view.container.querySelector(".result-review-bar")).toBeNull();
  expect(view.queryByRole("button", { name: "确认变更" })).toBeNull();
  return save;
}
const text = (view: ReturnType<typeof render>) => view.container.textContent ?? "";

describe("主表执行结果审阅（真实 App/执行器，宿主与网络替身）", () => {
  it.each(["confirm", "discard", "sourceChanged"])("acceptance entry interrupts after second write; remount %s", async outcome => {
    const sourceChanged = outcome === "sourceChanged";
    [111,222,333].forEach((v,i)=>host.setCell(SHEET, `B${i+2}`, v));
    [-101,-202,-303].forEach((v,i)=>host.setCell(TARGET, `B${i+2}`, v));
    let state: Awaited<ReturnType<typeof getOfficeExecutionStatus>>;
    vi.mocked(registerOfficeExecution).mockImplementation(async req => {
      state={...req,status:"running",revision:1,nextActionIndex:0,inFlightActionIndex:null,
        checkpoint:null,beforeSnapshot:null,canResume:false,requiresReconciliation:false,ownerToken:"owner"} as typeof state;
      return state;
    });
    vi.mocked(startOfficeAction).mockImplementation(async (_id,req)=>{
      state={...state,revision:state.revision+1,inFlightActionIndex:req.actionIndex,beforeSnapshot:req.beforeSnapshot,requiresReconciliation:true};return state;
    });
    vi.mocked(completeOfficeAction).mockImplementation(async(_id,req)=>{
      state={...state,revision:state.revision+1,inFlightActionIndex:null,nextActionIndex:req.actionIndex+1,checkpoint:req.checkpoint,requiresReconciliation:false};return state;
    });
    vi.mocked(suspendOfficeExecution).mockImplementation(async()=>{
      state={...state,status:"needs_reconciliation",revision:state.revision+1,ownerToken:undefined,canResume:true};return state;
    });
    vi.mocked(getOfficeExecutionStatus).mockImplementation(async()=>state);
    vi.mocked(reconcileOfficeActionReceipt).mockImplementation(async(_id,req)=>{
      expect(req.actionIndex).toBe(1);expect(req.outcome).toBe("applied");
      state={...state,status:"running",revision:state.revision+1,ownerToken:"new-owner",nextActionIndex:2,
        inFlightActionIndex:null,requiresReconciliation:false,checkpoint:req.checkpoint!};return state;
    });
    vi.mocked(finishOfficeExecution).mockImplementation(async()=>({...state,status:"completed"}));
    const plan={id:"probe-three",title:"三次值复制",summary:"复制到目标表",assumptions:[],warnings:[],
      actions:[2,3,4].map(r=>({type:"copyRange",sheet:TARGET,sourceSheet:SHEET,sourceRange:`B${r}`,
        targetRange:`B${r}`,copyType:"values",skipBlanks:false,transpose:false}))};
    const view=await sendPlan(plan,"依次执行三个值复制",{authorize:[SHEET,TARGET],probe:true});
    await waitFor(()=>expect(suspendOfficeExecution).toHaveBeenCalledTimes(1));
    expect([2,3,4].map(r=>host.cell(TARGET,`B${r}`))).toEqual([111,222,-303]);
    expect(vi.mocked(completeOfficeAction).mock.calls.map(c=>c[1].actionIndex)).toEqual([0]);
    expect(state!.nextActionIndex).toBe(1);expect(state!.inFlightActionIndex).toBe(1);
    expect(loadPendingOfficeExecution()?.operationId).toBe(state!.operationId);
    const before=contentWrites(host.originalWrites());expect(before).toHaveLength(2);
    view.unmount();
    if (sourceChanged) host.setCell(SHEET,"B3",999);
    const reopened=render(<App/>);
    const button=await reopened.findByRole("button",{name:"核对并继续"});
    await waitFor(()=>expect((button as HTMLButtonElement).disabled).toBe(false));fireEvent.click(button);
    if (sourceChanged) {
      await waitFor(()=>expect(text(reopened)).toContain("数据来源在预览后发生了变化"));
      expect(finishOfficeExecution).not.toHaveBeenCalled();
      expect(contentWrites(host.originalWrites())).toEqual(before);
      expect(loadPendingOfficeExecution()).not.toBeNull();return;
    }
    await waitFor(()=>expect(finishOfficeExecution, text(reopened)).toHaveBeenCalledTimes(1));
    expect(contentWrites(host.originalWrites()).slice(before.length).map(w=>w.address)).toEqual(["B4"]);
    expect([2,3,4].map(r=>host.cell(TARGET,`B${r}`))).toEqual([111,222,333]);
    expect(vi.mocked(startOfficeAction).mock.calls.map(c=>c[1].actionIndex)).toEqual([0,1,2]);
    expect(vi.mocked(completeOfficeAction).mock.calls.map(c=>c[1].actionIndex)).toEqual([0,2]);
    expect(registerOfficeExecution).toHaveBeenCalledTimes(1);expect(loadPendingOfficeExecution()).toBeNull();
    await stageReview(reopened);
    expect(reopened.container.querySelector('.result-review-bar')?.textContent).not.toContain('执行未完成');
    expect(text(reopened)).toContain('从第 3 步继续并完成');
    const writes=contentWrites(host.originalWrites());
    fireEvent.click(reopened.getByRole('button',{name:'双视图对比'}));
    await waitFor(()=>expect(host.createWorkbookCalls()).toHaveLength(1));
    const seed=host.seedEntries(0)['xl/worksheets/sheet1.xml'];
    expect(seed).toContain('-101');expect(seed).toContain('-202');expect(seed).toContain('-303');
    expect(contentWrites(host.originalWrites())).toEqual(writes);
    if(outcome==='confirm') {
      const confirm=reopened.getByRole('button',{name:'确认变更'});
      await waitFor(()=>expect((confirm as HTMLButtonElement).disabled).toBe(false));fireEvent.click(confirm);
      await waitFor(()=>expect(text(reopened)).toContain('已确认变更'));
      expect(reopened.queryByRole('button',{name:'确认变更'})).toBeNull();
      // This SharePoint host cannot close a native local copy; a settled cleanup notice is legitimate.
      expect(reopened.container.querySelector('.result-review-bar')?.textContent ?? '').not.toContain('执行未完成');
      expect(contentWrites(host.originalWrites())).toEqual(writes);
    } else {
      const discard=reopened.getByRole('button',{name:'放弃变更'});
      await waitFor(()=>expect((discard as HTMLButtonElement).disabled).toBe(false));fireEvent.click(discard);
      await waitFor(()=>expect(text(reopened)).toContain('已放弃并恢复'));
      expect([2,3,4].map(r=>host.cell(TARGET,`B${r}`))).toEqual([-101,-202,-303]);
    }
    expect(host.conditionalFormatCount(TARGET)).toBe(0);
  });

  it.each(['confirm', 'discard'])('value copy automatically executes three actions before %s', async outcome => {
    [111,222,333].forEach((v,i)=>host.setCell(SHEET, `B${i+2}`, v));
    [-101,-202,-303].forEach((v,i)=>host.setCell(TARGET, `B${i+2}`, v));
    const plan = {id:'auto-copy-three',title:'三次值复制',summary:'复制到目标表',assumptions:[],warnings:[],
      actions:[2,3,4].map(r=>({type:'copyRange',sheet:TARGET,sourceSheet:SHEET,
        sourceRange:`B${r}`,targetRange:`B${r}`,copyType:'values',skipBlanks:false,transpose:false}))};
    const view = await sendPlan(plan, '依次复制三格，直接执行后审阅', {authorize:[SHEET,TARGET]});
    await stageReview(view);
    expect([2,3,4].map(r=>host.cell(TARGET,`B${r}`))).toEqual([111,222,333]);
    expect(view.queryByRole('button',{name:'执行到 Excel'})).toBeNull();
    expect(host.conditionalFormatCount(TARGET)).toBe(3);
    const writes=contentWrites(host.originalWrites());
    expect(writes).toHaveLength(3);
    if(outcome==='confirm') {
      await clickSave(view);
      expect(contentWrites(host.originalWrites())).toEqual(writes);
    } else {
      fireEvent.click(view.getByRole('button',{name:'放弃变更'}));
      await waitFor(()=>expect(text(view)).toContain('已放弃并恢复'));
      expect([2,3,4].map(r=>host.cell(TARGET,`B${r}`))).toEqual([-101,-202,-303]);
    }
    expect(host.conditionalFormatCount(TARGET)).toBe(0);
    expect([2,3,4].map(r=>host.cell(SHEET,`B${r}`))).toEqual([111,222,333]);
    expect(view.container.querySelector('.scope-trigger')?.textContent).toContain('已固定 2 个工作表');
  });

  it('固定三表授权在执行、确认和重新打开范围选择器后保持不变', async () => {
    const view = await sendPlan(targetPlan, '汇总结果', { authorize: ALL_SHEETS });
    await stageReview(view);
    expect(view.container.querySelector('.scope-trigger')?.textContent).toContain('已固定 3 个工作表');
    await clickSave(view);
    expect(view.container.querySelector('.scope-trigger')?.textContent).toContain('已固定 3 个工作表');
    await openScopePicker(view);
    for (const name of ALL_SHEETS) expect(sheetToggle(view, name).getAttribute('aria-pressed')).toBe('true');
  });

  it("续跑交付计划后自动写入并回读，确认变更只清标记且不重复执行", async () => {
    vi.mocked(getPlanningTaskStatus).mockResolvedValue({
      kind: "planning_pause", turnId: "resumed-write-turn", status: "interrupted", revision: 1,
      lastSequence: 0, canContinue: true, pauseType: "interrupted", reason: "规划已暂停",
      extensionSeconds: 60, activeElapsedSeconds: 1, currentModelId: "claude-agent-sdk",
      canSwitchModel: false, availableModelConnections: [],
      continuation: { available: true, code: null, message: "", stage: "after_tool", sourceVerificationRequired: false }
    } as never);
    vi.mocked(continuePlanningTask).mockResolvedValue({ kind: "plan", provider: "model", plan: reviewPlan } as never);
    const view = await sendPlan(reviewPlan, "更新结果", { pauseBeforePlan: true });
    expect(host.originalWrites()).toHaveLength(0);
    fireEvent.click(await view.findByRole("button", { name: /继续分析/ }));
    await stageReview(view);
    expect(host.cell(SHEET, "G1")).toBe("合计");
    expect(contentWrites(host.originalWrites()).map(w => w.address)).toEqual(["E1:F3", "G1"]);
    expect(host.createWorkbookCalls()).toHaveLength(0);
    expect(view.queryByRole("button", { name: "确认执行" })).toBeNull();
    expect(view.queryByRole("button", { name: "生成副本并排审阅" })).toBeNull();
    expect(view.getByRole("button", { name: "放弃变更" })).toBeTruthy();
    expect(view.getByRole("button", { name: "双视图对比" })).toBeTruthy();
    const writes = contentWrites(host.originalWrites()).length;
    await clickSave(view);
    expect(contentWrites(host.originalWrites())).toHaveLength(writes);
    expect(continuePlanningTask).toHaveBeenCalledOnce();
  });
  it.each(["workbook-name", "cross-sheet-formula"])("含 %s 时双视图展示执行前目标值，不重写且放弃恢复", async kind => {
    host = createHost({ workbookNameCount: kind === "workbook-name" ? 1 : 0 });
    installOfficeTestHost(host);
    if (kind === "cross-sheet-formula") host.setCell(SHEET, "E1", "='汇总表'!A1", { asFormula: true });
    const before = host.matrix(SHEET, "A1:C3");
    const view = await sendPlan(reviewPlan);
    await stageReview(view);
    expect(host.cell(SHEET, "G1")).toBe("合计");
    const writes = contentWrites(host.originalWrites()).length;
    expect(writes).toBe(2);
    expect(host.createWorkbookCalls()).toHaveLength(0);
    fireEvent.click(view.getByRole("button", { name: "双视图对比" }));
    await waitFor(() => expect(host.createWorkbookCalls()).toHaveLength(1));
    const seed = host.seedEntries(0)["xl/worksheets/sheet1.xml"];
    expect(seed).not.toContain("<f>");
    expect(seed).not.toContain("票房");
    expect(text(view)).toContain("公式采用当时的值，不重新计算");
    expect(contentWrites(host.originalWrites())).toHaveLength(writes);
    expect(view.getByRole("button", { name: "确认变更" })).toBeTruthy();
    fireEvent.click(view.getByRole("button", { name: "放弃变更" }));
    await waitFor(() => expect(text(view)).toContain("已放弃并恢复"));
    expect(host.cell(SHEET, "G1")).toBe("");
    expect(host.matrix(SHEET, "A1:C3")).toEqual(before);
    expect(host.createWorkbookCalls()).toHaveLength(1);
  });
  it("含工作簿名称时确认变更仍只清标记，不重放写入", async () => {
    host = createHost({ workbookNameCount: 1 }); installOfficeTestHost(host);
    const view = await sendPlan(reviewPlan); await stageReview(view);
    const writes = contentWrites(host.originalWrites()).length;
    await clickSave(view);
    expect(contentWrites(host.originalWrites())).toHaveLength(writes);
    expect(host.cell(SHEET, "G1")).toBe("合计");
  });
  it("副本降级后执行恢复快照读取失败仍拒绝写入", async () => {
    host = createHost({ workbookNameCount: 1, corruptFormatRead: () => { throw new Error("恢复快照不可读"); } }); installOfficeTestHost(host);
    const view = await sendPlan({ ...reviewPlan, actions: [{ type: "setFont", sheet: SHEET, range: "E1:F3", bold: true }], acceptanceCriteria: [] });
    await waitFor(() => expect(text(view)).toContain("本次执行需要核对"));
    expect(host.originalWrites()).toHaveLength(0);
    expect(view.queryByRole("button", { name: "确认变更" })).toBeNull();
  });
  it("未知宿主读取错误仍阻止执行，不吞掉真实错误", async () => {
    host = createHost({ corruptFormatRead: () => { throw new Error("宿主读取中断"); } });
    installOfficeTestHost(host);
    const view = await sendPlan(reviewPlan);
    await waitFor(() => expect(text(view)).toContain("宿主读取中断"));
    expect(contentWrites(host.originalWrites())).toHaveLength(0);
  });
  it("主表立即显示结果，默认不创建对比工作簿", async () => {
    const view = await sendPlan(reviewPlan);
    await stageReview(view);
    expect(host.cell(SHEET, "G1")).toBe("合计");
    expect(contentWrites(host.originalWrites()).map(w => w.address)).toEqual(["E1:F3", "G1"]);
    expect(host.createWorkbookCalls()).toHaveLength(0);
    expect(view.container.querySelector(".composer-footer .result-review-bar") ?? view.container.querySelector(".result-review-bar")).not.toBeNull();
    expect(view.container.querySelector("button[title='保存为可重复使用的个人工具']")).toBeNull();
  });
  it.each([false, true])("双视图使用桥认可的副本身份并调用排列，排列失败=%s", async (fails) => {
    BOOK_URL = 'file:///C:/work/book.xlsx';
    host.setDocumentUrl(BOOK_URL);
    const ticket = { token: 'layout-ticket', layoutVersion: 3 as const };
    vi.mocked(prepareExcelReviewLayout).mockImplementationOnce(async request => {
      expect(request.copySheetName).toMatch(/^EB_[0-9a-f]{20}$/);
      return ticket;
    });
    if (fails) vi.mocked(arrangeExcelReviewLayout).mockRejectedValueOnce(new Error('排列超时'));
    else vi.mocked(arrangeExcelReviewLayout).mockResolvedValueOnce({ status: 'arranged', layoutVersion: 3,
      copySide: 'left', originalSide: 'right', synchronousScrolling: 'enabled' });
    const view = await sendPlan(reviewPlan); await stageReview(view);
    const writes = contentWrites(host.originalWrites()).length;
    fireEvent.click(view.getByRole('button', { name: '双视图对比' }));
    await waitFor(() => expect(arrangeExcelReviewLayout).toHaveBeenCalledWith(ticket));
    expect(host.createWorkbookCalls()).toHaveLength(1);
    const request = vi.mocked(prepareExcelReviewLayout).mock.calls.at(-1)![0];
    expect(host.seedEntries(0)['xl/workbook.xml']).toContain(request.copySheetName);
    if (fails) await waitFor(() => expect(text(view)).toContain('未能自动并排：排列超时'));
    else expect(text(view)).not.toContain('未能自动并排');
    expect(contentWrites(host.originalWrites())).toHaveLength(writes);
  });
  it("双视图的文件种子来自执行前，不重新执行计划", async () => {
    host.setCell(SHEET, "E1", "执行前旧值");
    host.setCell(SHEET, "G2", "不应复制的区域外数据");
    const view = await sendPlan(reviewPlan); await stageReview(view);
    const writes = contentWrites(host.originalWrites()).length;
    fireEvent.click(view.getByRole("button", { name: "双视图对比" }));
    await waitFor(() => expect(host.createWorkbookCalls()).toHaveLength(1));
    const xml = host.seedEntries(0)["xl/worksheets/sheet1.xml"];
    expect(xml).not.toContain("票房");
    expect(xml).toContain("执行前旧值");
    expect(xml).not.toContain("不应复制的区域外数据");
    expect(xml).toContain('r="E1"');
    expect(xml).not.toContain("<f>");
    expect(xml).not.toContain("上座率");
    expect(xml).not.toContain("合计");
    await waitFor(() => expect(text(view)).toContain("未能自动并排"));
    expect(contentWrites(host.originalWrites())).toHaveLength(writes);
  });
  /*
   * 非任务型 busy 不得给出"可暂停"的按钮（2026-10-23）。
   *
   * 结果已生成、等用户确认/放弃时 `busy` 为真，但并没有可中止的执行。
   * 显示 ■ 会让人以为能停，点下去却什么都不发生 —— 这正是实机上
   * "输入框暂停按钮偶发失效"的一条独立路径。
   */
  it("does not offer a pause button while nothing is actually executing", async () => {
    const view = await sendPlan(reviewPlan);
    await stageReview(view);
    const composerButton = view.container.querySelector<HTMLButtonElement>(".composer .send-button");
    expect(composerButton).not.toBeNull();
    expect(composerButton!.textContent?.trim()).not.toBe("■");
    expect(composerButton!.getAttribute("aria-label")).not.toBe("暂停任务");
    expect(view.container.querySelector('[aria-label="暂停任务"]')).toBeNull();
    expect(text(view)).toContain("等待审阅结果");
  });

  /*
   * 悬浮「回到底部」浮标的运行状态误判（与 WPS 同一条用户口径）。
   *
   * 等待用户确认不是"运行中"。
   *
   * 根因：`running` 接了 `busy`，而 `busy` 含 `phase === 'reviewing'`
   * （结果已就绪、等用户确认/放弃）⇒ 执行早结束了，浮标还在跳三点。
   *
   * 这里**不**只测"共享按钮收到 false 会画锚"——那证明不了接线对。验的是
   * App 自己算出来的 `running`：
   *   ① 结果已生成、reviewing、pending=false ⇒ 锚；
   *   ② 恢复这类实际操作在途（pending=true）⇒ 三点；
   *   ③ 操作完成 ⇒ 回锚。
   * `busy` 的防重复执行/互斥边界必须仍在（恢复入口在非 pending 时可点）。
   */
  it("waiting for the user is not running: the scroll hint shows the anchor while reviewing", async () => {
    const view = await sendPlan(reviewPlan);
    await stageReview(view);
    // 滚离底部：浮标只在离开底部时渲染（判据是位置，不要求有新内容）。
    const stream = view.container.querySelector<HTMLElement>(".message-stream")!;
    Object.defineProperties(stream, {
      scrollHeight: { configurable: true, value: 1000 },
      clientHeight: { configurable: true, value: 300 }
    });
    fireEvent.scroll(stream, { target: { scrollTop: 200 } });

    /*
     * ① 结果已就绪、等人：固定条 `aria-busy=false` 证明审阅通道没有操作在途
     *    （不是"按钮替身画了个锚"），此时浮标必须是**锚**。
     */
    const bar = view.container.querySelector<HTMLElement>(".result-review-bar")!;
    expect(bar.getAttribute("aria-busy")).toBe("false");
    expect(text(view)).toContain("请查看主表中的执行结果");
    await waitFor(() =>
      expect(view.container.querySelector<HTMLButtonElement>(".new-progress-button")?.dataset.mode).toBe("idle"));
    const hint = view.container.querySelector<HTMLButtonElement>(".new-progress-button")!;
    expect(hint.querySelector(".progress-anchor-icon")).not.toBeNull();
    expect(hint.querySelector(".progress-dots")).toBeNull();

    /*
     * ② 实际操作在途：挂住恢复流程里的 `targetGuard.check()`
     *    （→ `inspectOfficeAuthorizedTarget`）。这是**真的在跑**：
     *    `resultReviewState.pending` 会一直是 true，而 `phase` 仍是 `reviewing`
     *    —— 正好证明"在途"与"等人"是两件事。
     */
    let releaseGuard: (() => void) | undefined;
    vi.mocked(inspectOfficeAuthorizedTarget).mockImplementationOnce(
      () => new Promise((resolve) => { releaseGuard = () => resolve(targetFor(null) as never); })
    );
    fireEvent.click(view.getByRole("button", { name: "放弃变更" }));
    await waitFor(() =>
      expect(view.container.querySelector<HTMLButtonElement>(".new-progress-button")?.dataset.mode).toBe("running"));
    expect(view.container.querySelector<HTMLButtonElement>(".new-progress-button")!
      .querySelectorAll(".progress-dots > span")).toHaveLength(3);
    // 同一时刻审阅通道确实在途：按钮被禁用、`aria-busy` 为真。
    expect(view.container.querySelector<HTMLElement>(".result-review-bar")!.getAttribute("aria-busy")).toBe("true");
    expect((view.getByRole("button", { name: "放弃变更" }) as HTMLButtonElement).disabled).toBe(true);

    /*
     * ③ 操作完成：回到锚。`busy` 的互斥边界没被破坏 —— 恢复成功、固定条收起，
     *    而恢复本身真的发生了（互斥一直都在，只是不再冒充"在跑"）。
     */
    await act(async () => { releaseGuard?.(); });
    await waitFor(() => expect(text(view)).toContain("已放弃并恢复"));
    await waitFor(() =>
      expect(view.container.querySelector<HTMLButtonElement>(".new-progress-button")?.dataset.mode).toBe("idle"));
    expect(view.container.querySelector<HTMLButtonElement>(".new-progress-button")).toBe(hint);
    const settledBar = view.container.querySelector<HTMLElement>(".result-review-bar");
    if (settledBar) expect(settledBar.getAttribute("aria-busy")).toBe("false");
  });
  it("确认仅去除本次标记，不重复写入，并开放保存工具", async () => {
    const view = await sendPlan(reviewPlan); await stageReview(view);
    const writes = contentWrites(host.originalWrites());
    expect(host.conditionalFormatCount(SHEET)).toBe(2);
    await clickSave(view);
    expect(host.conditionalFormatCount(SHEET)).toBe(0);
    expect(contentWrites(host.originalWrites())).toEqual(writes);
    expect(host.cell(SHEET, "G1")).toBe("合计");
    expect(text(view)).toContain("并通过独立验证");
    expect(view.container.querySelector("button[title='保存为可重复使用的个人工具']")).not.toBeNull();
  });
  it("放弃调用真正撤销，恢复目标且保留区域外内容", async () => {
    const view = await sendPlan(reviewPlan); await stageReview(view);
    expect(host.cell(SHEET, "G1")).toBe("合计");
    fireEvent.click(view.getByRole("button", { name: "放弃变更" }));
    await waitFor(() => expect(text(view)).toContain("已放弃并恢复"));
    expect(host.cell(SHEET, "G1")).toBe("");
    expect(host.cell(SHEET, "E1")).toBe("");
    expect(host.cell(SHEET, "B2")).toBe(1200);
    expect(host.windowNames()).toEqual(["book.xlsx"]);
  });
  it("执行后目标被人工修改，确认与放弃都不覆盖该修改", async () => {
    const view = await sendPlan(reviewPlan); await stageReview(view);
    host.setCell(SHEET, "G1", "人工修改");
    const count = contentWrites(host.originalWrites()).length;
    fireEvent.click(view.getByRole("button", { name: "确认变更" }));
    await waitFor(() => expect(text(view)).toContain("又发生了变化"));
    fireEvent.click(view.getByRole("button", { name: "放弃变更" }));
    await waitFor(() => expect(text(view)).toContain("未能放弃"));
    expect(host.cell(SHEET, "G1")).toBe("人工修改");
    expect(contentWrites(host.originalWrites())).toHaveLength(count);
    expect(view.container.querySelector(".result-review-bar")?.textContent).not.toContain("已确认变更");
  });
  it("执行后切换工作簿，旧确认不产生任何额外写入", async () => {
    const view = await sendPlan(reviewPlan); await stageReview(view);
    const count = contentWrites(host.originalWrites()).length;
    host.setDocumentUrl("https://contoso.sharepoint.com/other.xlsx");
    fireEvent.click(view.getByRole("button", { name: "确认变更" }));
    await waitFor(() => expect(text(view)).toContain("工作簿、授权或计划已变化"));
    expect(contentWrites(host.originalWrites())).toHaveLength(count);
  });
  it("回读不符不开放确认或保存工具，不自动重放", async () => {
    host.tamperNextWrite(event => event.address === "E1:F3" && event.store === "original",
      matrix => [matrix[0], [matrix[1][0], 0.99], matrix[2]]);
    const view = await sendPlan(reviewPlan);
    await waitFor(() => expect(text(view)).toContain("本次执行需要核对"));
    expect(host.cell(SHEET, "F2")).toBe(0.99);
    expect(view.queryByRole("button", { name: "确认变更" })).toBeNull();
    expect(contentWrites(host.originalWrites()).map(w => w.address)).toEqual(["E1:F3", "G1"]);
    expect(view.container.querySelector("button[title='保存为可重复使用的个人工具']")).toBeNull();
  });
  it("部分失败保留成功步骤的恢复记录，不重放失败步骤", async () => {
    host.failWrites(event => event.store === "original" && event.address === "G1", "宿主拒绝写入 G1");
    const view = await sendPlan(reviewPlan);
    await waitFor(() => expect(text(view)).toContain("已有 1 步成功写入当前工作簿"));
    expect(host.cell(SHEET, "E1")).toBe("影片");
    expect(host.cell(SHEET, "G1")).toBe("");
    expect(contentWrites(host.originalWrites()).map(w => w.address)).toEqual(["E1:F3"]);
    expect(view.queryByRole("button", { name: "确认变更" })).toBeNull();
  });
});

async function clickRunOnExcel(view: ReturnType<typeof render>) {
  const button = await waitFor(() => {
    const found = [...view.container.querySelectorAll("button")].find(
      (candidate) => candidate.textContent === "执行到 Excel"
    ) as HTMLButtonElement | undefined;
    if (!found) throw new Error("「执行到 Excel」按钮尚未出现");
    return found;
  });
  fireEvent.click(button);
  return button;
}

describe("Excel 写侧：直接执行路径的写前判据", () => {
  it("计划含未授权工作表的动作：执行期拒绝，一个单元格都不写", async () => {
    const view = await sendPlan(offSourcePlan);
    await clickRunOnExcel(view);

    await waitFor(() => expect(text(view)).toContain("执行前检查未通过"));
    expect(host.writes()).toEqual([]);
    expect(host.cell(SHEET, "C2")).toBe(30);
  });

  it("等待期间源值被改（烘焙矩阵计划）：给出二次确认，不自动写入", async () => {
    const view = await sendPlan(directPlan);
    host.setCell(SHEET, "B2", 4321);
    await clickRunOnExcel(view);

    await waitFor(() =>
      expect(view.container.querySelector(".stale-source-message")).not.toBeNull()
    );
    expect(host.originalWrites()).toEqual([]);
    expect(host.cell(SHEET, "C2")).toBe(30);
  });

  it("源表**公式文本**变化也会改变来源指纹（不是只算值）", async () => {
    const view = await sendPlan(directPlan);
    // 显示结果不变，变的是单元格里的公式文本。
    host.setCell(SHEET, "C2", "=30+0", { asFormula: true });
    await clickRunOnExcel(view);

    await waitFor(() =>
      expect(view.container.querySelector(".stale-source-message")).not.toBeNull()
    );
    expect(host.originalWrites()).toEqual([]);
  });

  it("源值、身份都没变时照常写入，并如实给出恢复限制说明", async () => {
    const view = await sendPlan(directPlan);
    await clickRunOnExcel(view);

    await waitFor(() => expect(host.cell(SHEET, "C2")).toBe(30));
    expect(host.matrix(SHEET, "C1:C3")).toEqual([["场次"], [30], [22]]);
    // 没有可信来源记账 → 不登记断点，如实说明"本次不承诺跨重启自动续跑"。
    expect(text(view)).toContain("本次不承诺跨重启自动续跑");
    expect(registerOfficeExecution).not.toHaveBeenCalled();
  });
});

/**
 * 断点恢复登记的**真实可达性**。
 *
 * 本轮修复的缺陷：App 曾把 `sourceFingerprintSheets` 写成"本次授权表集合"，而
 * 动作又必须落在授权表上 —— 于是 `officeRecoverySupport` 的"会修改来源表"必然
 * 命中，`registerOfficeExecution` 成了死代码。现在授权范围／实际来源依赖／写入
 * 目标三者分开建模，来源依赖只来自本机可信记录，来源与目标明确不重叠的计划必须
 * 真的登记成功。
 */
describe("Excel 写侧：断点恢复登记的真实可达性", () => {
  const SCOPE = [SHEET, TARGET];

  it("合法计划真的登记，并且登记成功后才开始原表写入", async () => {
    const view = await sendPlan(targetPlan, "把合计写到汇总表", {
      authorize: SCOPE,
      read: { scope: "active" }
    });
    // 已授权的新流程在计划抵达后自动执行；下面核对实际登记/写入结果。

    // 登记真的发生了，而且带上的是**本机可信记账**出来的来源依赖（只读过的表）。
    await waitFor(() => expect(registerOfficeExecution).toHaveBeenCalledTimes(1));
    const registered = vi.mocked(registerOfficeExecution).mock.calls[0][0] as {
      plan: { sourceDependencySheets?: string[]; sourceFingerprintSheets?: string[] };
      authorizedSheetNames: string[];
    };
    expect(registered.plan.sourceDependencySheets).toEqual([SHEET]);
    // 指纹覆盖范围仍然是本次授权范围（既有口径未改）；两者刻意不同。
    expect(registered.plan.sourceFingerprintSheets).toEqual(SCOPE);
    expect(registered.authorizedSheetNames).toEqual(SCOPE);
    expect(officeRecoveryAttestedOnce()).toBe(true);

    // 写入落到目标表，来源表一个单元格都没动。
    await waitFor(() => expect(host.cell(TARGET, "A1")).toBe("影片"));
    expect(host.matrix(TARGET, "A1:B2")).toEqual([
      ["影片", "合计"],
      ["甲", 1200]
    ]);
    expect(host.matrix(SHEET, "A1:C3")).toEqual([
      ["影片", "票房(万)", "场次"],
      ["甲", 1200, 30],
      ["乙", 980, 22]
    ]);
    expect(contentWrites(host.originalWrites())).toEqual([
      { store: "original", sheet: TARGET, address: "A1:B2", kind: "values", cells: 4 }
    ]);
    // 界面对"这一级断点"的范围如实说明，而不是笼统宣称"支持恢复"。
    expect(text(view)).toContain("已登记动作级断点");
    expect(text(view)).toContain("只续跑本次登记的动作");
    expect(text(view)).not.toContain("本次不承诺跨重启自动续跑");
  });

  it("网络失败：登记不到就不再往下走，原表零写入", async () => {
    vi.mocked(registerOfficeExecution).mockRejectedValue(
      new Error("无法连接本地服务") as never
    );
    const view = await sendPlan(targetPlan, "把合计写到汇总表", {
      authorize: SCOPE,
      read: { scope: "active" }
    });
    // 已授权的新流程在计划抵达后自动执行；下面核对实际登记/写入结果。

    await waitFor(() =>
      expect(registerOfficeExecution).toHaveBeenCalledTimes(1)
    );
    await waitFor(() => expect(text(view)).toContain("无法连接本地服务"));
    expect(host.originalWrites()).toEqual([]);
    expect(agentWrites()).toBe(0);
  });

  it("服务端没授予写入权：抛错并停止，原表零写入", async () => {
    vi.mocked(registerOfficeExecution).mockResolvedValue(
      officeLedgerState({ ownerToken: "" }) as never
    );
    const view = await sendPlan(targetPlan, "把合计写到汇总表", {
      authorize: SCOPE,
      read: { scope: "active" }
    });
    // 已授权的新流程在计划抵达后自动执行；下面核对实际登记/写入结果。

    await waitFor(() =>
      expect(text(view)).toContain("服务端没有授予本次 Office 写入权")
    );
    expect(registerOfficeExecution).toHaveBeenCalledTimes(1);
    expect(host.writes()).toEqual([]);
  });

  it("挂接索引无法落盘：抛错并停止，原表零写入", async () => {
    let restoreStorage: (() => void) | undefined;
    vi.mocked(registerOfficeExecution).mockImplementationOnce(async () => {
      // Inject only after registration, before the real attachment-index write.
      const spy = vi.spyOn(Storage.prototype, "setItem").mockImplementation(() => {
        throw new Error("quota exceeded");
      });
      restoreStorage = () => spy.mockRestore();
      return officeLedgerState() as never;
    });
    const view = await sendPlan(targetPlan, "把合计写到汇总表", {
      authorize: SCOPE,
      read: { scope: "active" }
    });
    // 挂载期间不要拦 localStorage（工具库自己也在读写），只在本步写盘前拦。
    try {
      // 已授权的新流程在计划抵达后自动执行；下面核对实际登记/写入结果。

      await waitFor(() =>
        expect(text(view)).toContain("无法保存 Office 任务挂接索引")
      );
      expect(host.writes()).toEqual([]);
    } finally {
      restoreStorage?.();
    }
  });

  it("动作级写入权被拒：该步零写入", async () => {
    vi.mocked(startOfficeAction).mockRejectedValue(
      new Error("Office 写入权已经失效") as never
    );
    const view = await sendPlan(targetPlan, "把合计写到汇总表", {
      authorize: SCOPE,
      read: { scope: "active" }
    });
    // 已授权的新流程在计划抵达后自动执行；下面核对实际登记/写入结果。

    await waitFor(() => expect(startOfficeAction).toHaveBeenCalledTimes(1));
    await waitFor(() => expect(text(view)).toContain("Office 写入权已经失效"));
    expect(host.originalWrites()).toEqual([]);
  });

  it("服务端自报的来源依赖一律被剥离：不能靠声明拿到恢复", async () => {
    /*
     * 计划里带着一个"很好"的来源集合（把目标表摘掉了）。这是**声明**，不是记账：
     * 本轮没有本机读取记录，所以它必须被剥掉，恢复登记必须不发生。
     */
    const forged = {
      ...targetPlan,
      id: "plan-forged-1",
      sourceFingerprintSheets: SCOPE,
      sourceDependencySheets: [SHEET]
    };
    const view = await sendPlan(forged, "把合计写到汇总表", { authorize: SCOPE });
    // 已授权的新流程在计划抵达后自动执行；下面核对实际登记/写入结果。

    await waitFor(() => expect(host.cell(TARGET, "A1")).toBe("影片"));
    expect(registerOfficeExecution).not.toHaveBeenCalled();
    expect(text(view)).toContain("本次不承诺跨重启自动续跑");
  });

  it("工具真的读过目标表：依赖含目标 → 不登记（不靠摘掉目标造来源集合）", async () => {
    /*
     * `scope: "selected"` 会读**全部授权表**，其中就有写入目标「汇总表」。
     * 于是来源依赖 = [综合票房, 汇总表] 与目标重叠 —— 即便"把目标从来源里摘掉"
     * 就能换来登记成功，也不摘：记账说的是"读过谁"，它就是读过目标表。
     */
    const view = await sendPlan(targetPlan, "把合计写到汇总表", {
      authorize: SCOPE,
      read: { scope: "selected" }
    });
    // 已授权的新流程在计划抵达后自动执行；下面核对实际登记/写入结果。

    await waitFor(() => expect(host.cell(TARGET, "A1")).toBe("影片"));
    expect(registerOfficeExecution).not.toHaveBeenCalled();
    expect(text(view)).toContain("会修改来源工作表");
  });

  /*
   * 稳定身份不是"顺手带的元数据"：工具桥绑定、服务端桥归属、断点恢复账本全都读它。
   * 三个捕获点各自漏写，都会让同一个工作簿在"绑定"与"执行前复核"两边拿到不同的
   * 身份（一边真身份、一边同名兜底），自比不自等 —— 桥必然判换簿，恢复登记随之
   * 不可达。所以逐个捕获点各断言一次，而不是只看其中一个。
   */
  it("工作簿快照当场记下稳定身份：三个捕获点一致，未保存时才为空", async () => {
    const expected = `url:${BOOK_URL}`;
    expect((await captureWorkbookCatalog()).workbookIdentity).toBe(expected);
    expect((await captureWorkbookStructure([SHEET])).workbookIdentity).toBe(expected);
    expect((await captureWorkbook([SHEET])).workbookIdentity).toBe(expected);

    // 未保存的工作簿没有稳定身份：如实留空，不拿文件名冒充一个"看起来能用"的身份。
    installOfficeTestHost(
      createOfficeTestHost({
        selectedRange: "A1:C3",
        activeSheet: SHEET,
        sheets: [{ name: SHEET, values: [["影片"]] }]
      })
    );
    expect((await captureWorkbookCatalog()).workbookIdentity).toBeNull();
  });
});

/** 原表上由 Agent 写下的内容（副本写入不计）。 */
function agentWrites(): number {
  return host.originalWrites().length;
}

/** 断言登记只发生过一次（把"真登记"与"重试"分开看）。 */
function officeRecoveryAttestedOnce(): boolean {
  return vi.mocked(registerOfficeExecution).mock.calls.length === 1;
}

/**
 * 回执丢失／进程中断之后：**先核对，再续跑**，绝不重复执行已完成的写入。
 */
describe("Excel 写侧：中断后先核对再续跑", () => {
  const SCOPE = [SHEET, TARGET];
  const OPERATION_ID = "office-resume-0001";

  function pendingIndex() {
    savePendingOfficeExecution({
      schemaVersion: 1,
      operationId: OPERATION_ID,
      conversationId: "conversation-1",
      activityMessageId: "activity-1",
      workbookIdentity: `url:${BOOK_URL}`,
      createdAt: new Date().toISOString()
    });
  }

  function resumeState(overrides: Record<string, unknown> = {}) {
    return {
      operationId: OPERATION_ID,
      status: "running",
      revision: 2,
      workbookIdentity: `url:${BOOK_URL}`,
      plan: {
        ...liveTargetPlan,
        sourceFingerprintSheets: SCOPE,
        sourceDependencySheets: [SHEET]
      },
      authorizedSheetNames: SCOPE,
      allowedExternalFiles: [],
      nextActionIndex: 0,
      inFlightActionIndex: 0,
      beforeSnapshot: {
        actionIndex: 0,
        sheet: TARGET,
        range: "A1:B1",
        formulas: [[null, null]],
        numberFormat: [["General", "General"]],
        fillColor: "#FFFFFF",
        fontBold: false,
        fontColor: "#000000"
      },
      checkpoint: null,
      result: null,
      error: null,
      canResume: true,
      requiresReconciliation: true,
      ownerToken: "owner-token",
      ...overrides
    };
  }

  it("重载后按挂接索引核对回执：只汇报，不重复执行已完成的动作", async () => {
    pendingIndex();
    vi.mocked(getOfficeExecutionStatus).mockResolvedValue({
      ...resumeState({ status: "completed", nextActionIndex: 2, inFlightActionIndex: null }),
      result: { actionResults: [], verification: { passed: true, status: "verified", unverifiedActions: [] } }
    } as never);

    const view = render(<App />);
    await waitFor(() =>
      expect(text(view)).toContain("不会重复执行已经完成的动作")
    );

    expect(getOfficeExecutionStatus).toHaveBeenCalledTimes(1);
    expect(host.writes()).toEqual([]);
    // 挂接索引在核对完成后被清掉：下次启动不会再提示同一件事。
    expect(loadPendingOfficeExecution()).toBeNull();
  });

  it.each(["applied", "not_applied", "unknown"] as const)("A1 三动作中途断开：替身宿主现场核对为 %s 后不重写已完成动作", async (outcome) => {
    host.setCell(TARGET, "A1", '="影片"', { asFormula: true });
    host.setCell(TARGET, "B1", "=1200", { asFormula: true });
    if (outcome === "applied") {
      host.setCell(TARGET, "A2", '="乙"', { asFormula: true });
      host.setCell(TARGET, "B2", "=980", { asFormula: true });
    } else if (outcome === "unknown") {
      host.setCell(TARGET, "A2", "人工改值");
    }
    const beforeWrites = contentWrites(host.originalWrites()).length;
    const originalFirstRow = host.matrix(TARGET, "A1:B1");
    const plan = { ...liveTargetPlan, sourceFingerprintSheets: SCOPE, sourceDependencySheets: [SHEET],
      actions: [...liveTargetPlan.actions, { type: "writeFormulas" as const, sheet: TARGET, range: "A3:B3", formulas: [['="丙"', "=600"]] }] };
    pendingIndex();
    vi.mocked(getOfficeExecutionStatus).mockResolvedValue(resumeState({ plan, nextActionIndex: 1, inFlightActionIndex: 1,
      beforeSnapshot: { ...resumeState().beforeSnapshot, actionIndex: 1, range: "A2:B2", formulas: [["", ""]], fillColor: "" } }) as never);
    vi.mocked(reconcileOfficeActionReceipt).mockResolvedValue(resumeState({ plan,
      nextActionIndex: outcome === "applied" ? 2 : 1, inFlightActionIndex: null, requiresReconciliation: false, revision: 3 }) as never);
    const view = render(<App />);
    const button = await view.findByRole("button", { name: "核对并继续" });
    fireEvent.click(button);
    if (outcome === "unknown") {
      await waitFor(() => expect(text(view)).toContain("目标范围既不匹配执行前状态"));
      expect(reconcileOfficeActionReceipt).not.toHaveBeenCalled();
      expect(startOfficeAction).not.toHaveBeenCalled();
      expect(contentWrites(host.originalWrites()).length).toBe(beforeWrites);
      return;
    }
    await waitFor(() => expect(finishOfficeExecution).toHaveBeenCalledTimes(1));
    expect(contentWrites(host.originalWrites()).slice(beforeWrites).map((write) => write.address))
      .toEqual(outcome === "applied" ? ["A3:B3"] : ["A2:B2", "A3:B3"]);
    expect(host.matrix(TARGET, "A1:B1")).toEqual(originalFirstRow);
    expect(host.cell(TARGET, "A3")).toBe('="丙"');
    expect(vi.mocked(startOfficeAction).mock.calls.map((call) => call[1].actionIndex))
      .toEqual(outcome === "applied" ? [2] : [1, 2]);
    expect(reconcileOfficeActionReceipt).toHaveBeenCalledWith(OPERATION_ID, expect.objectContaining({ actionIndex: 1, outcome }));
  });

  it("在途动作先核对为「已写入」：跳过该步，只写剩余步骤，不重复写", async () => {
    // 头一步已经真的写进宿主了（回执丢失），第二部还没写。
    host.setCell(TARGET, "A1", "=\"影片\"", { asFormula: true });
    host.setCell(TARGET, "B1", "=1200", { asFormula: true });
    const firstWrites = host.originalWrites().length;

    pendingIndex();
    vi.mocked(getOfficeExecutionStatus).mockResolvedValue(resumeState() as never);
    vi.mocked(reconcileOfficeActionReceipt).mockResolvedValue(
      resumeState({ nextActionIndex: 1, inFlightActionIndex: null, requiresReconciliation: false, revision: 3 }) as never
    );
    vi.mocked(finishOfficeExecution).mockResolvedValue(
      officeLedgerState({ status: "finished", revision: 4 }) as never
    );

    const view = render(<App />);
    // 挂接索引在挂载时先核对一次：在途动作 → 恢复卡（要求先核对）。
    const resumeButton = await waitFor(() => {
      const card = view.container.querySelector(".office-resume-card");
      if (!card) throw new Error("恢复卡尚未出现");
      const button = [...card.querySelectorAll("button")].find(
        (candidate) => candidate.textContent === "核对并继续"
      ) as HTMLButtonElement | undefined;
      if (!button) throw new Error("「核对并继续」按钮尚未出现");
      return button;
    });
    fireEvent.click(resumeButton);

    // 续跑：只写第二步（A2:B2），第一步的 A1:B1 不再被写一次。
    await waitFor(() => expect(host.cell(TARGET, "A2")).toBe("=\"乙\""));
    const writes = contentWrites(host.originalWrites()).slice(firstWrites);
    expect(writes.map((write) => write.address)).toEqual(["A2:B2"]);
    expect(reconcileOfficeActionReceipt).toHaveBeenCalledTimes(1);
    expect(
      (vi.mocked(reconcileOfficeActionReceipt).mock.calls[0][1] as { outcome: string }).outcome
    ).toBe("applied");
  });
});

/**
 * 完整点击链的**收尾**：确认写入之后必须还能**真的撤销回去**。
 *
 * 前两段（预览零写入、确认写精确目标 + 独立回读）已由上面的用例覆盖；这里补
 * 第 4 条要求里明确点名的两件事：
 *
 * 1. **撤销范围必须显式，且恢复结果要真的核对过** —— 撤销按钮只在
 *    `postStateCoverage === "complete"` 时可用；点了以后原表必须逐单元格回到
 *    执行前状态，并且"回到执行前"这件事是由 `assertUndoPostState` 在**写之前**
 *    读宿主现场核出来的，不是界面上的一句宣称。
 * 2. **同表不重叠**（来源与写入目标在同一张表、但区域不重叠）的支持程度要
 *    **如实说明**：可以直接写入（来源指纹覆盖它），但按 `officeRecoverySupport`
 *    它**不在**跨重启断点恢复的白名单里 —— 不伪造来源依赖、也不放松恢复闸。
 */
describe("Excel 写侧：撤销与同表不重叠的真实支持程度", () => {
  it("确认写入 → 界面给出显式撤销范围 → 撤销后逐单元格回到执行前状态", async () => {
    const view = await sendPlan(reviewPlan);
    await stageReview(view);
    await clickSave(view);

    // 前置：写入确实发生，且界面只在**执行后状态已完整核对**时给出可用的撤销。
    await waitFor(() => expect(host.cell(SHEET, "G1")).toBe("合计"));
    await waitFor(() => expect(text(view)).toContain("并已核对执行后状态。"));

    const undoButton = [...view.container.querySelectorAll("button")].find(
      (candidate) => candidate.textContent === "↶ 撤销本次执行"
    ) as HTMLButtonElement | undefined;
    expect(
      undoButton,
      "执行后状态已完整核对，撤销按钮必须真的出现（不是笼统占位）"
    ).toBeDefined();
    expect(undoButton!.disabled).toBe(false);
    // The recovery control stays with this execution's verified result, not a global footer.
    expect(undoButton!.closest(".execution-result-card")?.textContent).toContain("项符合预期");
    // 撤销范围显式：撤销数据项数 = 本次实际写入的区域数（E:F 与 G1）。
    expect(text(view)).toContain("已记录 2 项本次执行撤销数据");

    fireEvent.click(undoButton!);

    // 恢复结果**真的核对过**：两个写入区域逐单元格回到执行前（原表种子）。
    await waitFor(() => expect(text(view)).toContain("已撤销上一次执行"));
    expect(host.matrix(SHEET, "E1:F3")).toEqual([
      ["", ""],
      ["", ""],
      ["", ""]
    ]);
    expect(host.cell(SHEET, "G1")).toBe("");
    // ★ 撤销成功后，原结果卡不再把旧验证结论当当前状态：标题换成「已撤销并恢复」，
    //   且不再出现「验证通过」/「项符合预期」这类旧摘要（否则就是"撤销了却仍显示验证通过"）。
    //   撤销按钮也会随卡降级而消失，因此这里**重新查询**卡片（不沿用已摘除按钮的旧引用）。
    await waitFor(() => expect(text(view)).toContain("已撤销并恢复"));
    const resultCard = view.container.querySelector(".execution-result-card");
    expect(resultCard).not.toBeNull();
    expect(resultCard!.textContent).toContain("已撤销并恢复");
    expect(resultCard!.textContent).toContain("本次执行已撤销");
    expect(resultCard!.textContent).not.toContain("验证通过");
    expect(resultCard!.textContent).not.toContain("项符合预期");
    // 撤销按钮已随卡降级消失：撤销完成后不再保留可点按钮。
    expect(
      [...view.container.querySelectorAll("button")].some(
        (candidate) => candidate.textContent === "↶ 撤销本次执行"
      )
    ).toBe(false);
    // 只撤销本次改动的区域：没被本次碰过的列原样保留（撤过头同样是错的）。
    expect(host.matrix(SHEET, "A1:C3")).toEqual([
      ["影片", "票房(万)", "场次"],
      ["甲", 1200, 30],
      ["乙", 980, 22]
    ]);
    // 撤销写日志里的**目标范围**必须与执行时写入的范围一一对应，而不是整表重刷。
    const undoWrites = contentWrites(host.originalWrites()).slice(2);
    expect(undoWrites.map((write) => write.address).sort()).toEqual(["E1:F3", "G1"]);
  });

  it("执行后人工又改了目标区域：撤销被拒，原表不再被自动覆盖", async () => {
    const view = await sendPlan(reviewPlan);
    await stageReview(view);
    await clickSave(view);
    await waitFor(() => expect(host.cell(SHEET, "G1")).toBe("合计"));
    const afterExecution = host.originalWrites().length;

    // 执行之后用户自己动了同一区域 —— 撤销必须停下，不能把用户的编辑盖掉。
    host.setCell(SHEET, "G1", "我手动改的");
    const undoButton = [...view.container.querySelectorAll("button")].find(
      (candidate) => candidate.textContent === "↶ 撤销本次执行"
    ) as HTMLButtonElement;
    fireEvent.click(undoButton);

    await waitFor(() =>
      // 报冲突时**指名"将要写回的字段"**（本步是值/公式），而不是笼统说"区域变了" ——
      // 判据覆盖范围已收窄到实际恢复范围，文案也必须对齐这一点。
      expect(text(view)).toContain(
        "在本次撤销将要写回的字段（值/公式）上又发生了变化，已停止撤销"
      )
    );
    // 负向断言的正向锚点：撤销尝试没有对原表产生任何新写入。
    expect(host.originalWrites().length).toBe(afterExecution);
    expect(host.cell(SHEET, "G1")).toBe("我手动改的");
    // ★ 撤销**冲突**绝不能把卡标成"已撤销"：标题仍是验证结论，不出现"已撤销并恢复"。
    const cardAfterConflict = undoButton.closest(".execution-result-card");
    expect(cardAfterConflict).not.toBeNull();
    expect(cardAfterConflict!.textContent).not.toContain("已撤销并恢复");
    expect(cardAfterConflict!.textContent).not.toContain("本次执行已撤销");
    expect(cardAfterConflict!.textContent).toContain("验证通过");
  });

  it("来源与写入目标同表、区域不重叠：可以写入，但如实说明不支持跨重启断点恢复", async () => {
    /*
     * 同一张表（`SHEET`）既是来源也是写入目标，写入区域 E1:F3 与来源区域
     * A1:C3 **不重叠**。计划刻意声明了"来源依赖只含本表"（服务端编译结果计划
     * 的形状），但本轮**没有**任何本机读取记账 —— 声明必须被剥掉，于是
     * `officeRecoverySupport` 只能退回指纹范围，落在"不确定重叠"的保守拒绝面上。
     */
    const sameSheetPlan = {
      ...reviewPlan,
      id: "plan-same-sheet-1",
      sourceFingerprintSheets: [SHEET],
      sourceDependencySheets: [SHEET]
    };
    const view = await sendPlan(sameSheetPlan);
    await stageReview(view);
    await clickSave(view);

    // 写入照常成功（同表不重叠是可写的），且回读对得上。
    await waitFor(() => expect(host.cell(SHEET, "G1")).toBe("合计"));
    expect(host.matrix(SHEET, "E1:F3")).toEqual([
      ["影片", "上座率"],
      ["甲", 0.42],
      ["乙", 0.31]
    ]);
    expect(text(view)).toContain("并通过独立验证");
    // 撤销可用（执行后状态完整核对）—— 与"断点恢复"是两件事，分开如实说明。
    expect(text(view)).toContain("并已核对执行后状态。");

    // 但**不支持跨重启断点恢复**：界面必须照实说，而不是因为它是"同表"就放宽闸。
    expect(text(view)).toContain("本次不承诺跨重启自动续跑");
    expect(text(view)).not.toContain("已登记动作级断点");
    expect(registerOfficeExecution).not.toHaveBeenCalled();
  });
});


describe("Excel native worksheet checkpoint routing", () => {
  it.each([false, true])("a fresh pure delete plan uses the checkpoint batch, backup consent=%s", async needsConsent => {
    BOOK_URL = "file:///C:/work/book.xlsx";
    host = createHost(); installOfficeTestHost(host);
    const executedOperationIds: string[] = [];
    vi.mocked(executeWorksheetDeletion).mockImplementation(async req => {
      expect(req.authorizedSheets).toEqual(expect.arrayContaining(ALL_SHEETS));
      expect(req.host).toBe("excel");
      executedOperationIds.push(req.operationId);
      // operationId 由执行时生成（32 位 hex），不是计划 id。
      expect(req.operationId).toMatch(/^[a-f0-9]{32}$/);
      if (needsConsent) return { operationId: req.operationId, status: "needs_backup_consent", backupSheets: ["#EB_RULES"], deletedSheets: [], remainingSheets: [], recoveryAvailable: false, message: "请确认额外备份许可" };
      return { operationId: req.operationId, status: "completed", deletedSheets: [TARGET], remainingSheets: [SHEET, UNAUTHORIZED_SHEET], recoveryAvailable: true, message: "已核对剩余工作表。" };
    });
    vi.mocked(confirmWorksheetCheckpointBackup).mockImplementation(async req => ({ operationId: req.operationId, status: "completed", deletedSheets: [TARGET], remainingSheets: [SHEET, UNAUTHORIZED_SHEET, "#EB_RULES"], recoveryAvailable: true, message: "已核对剩余工作表。" }));
    vi.mocked(restoreWorksheetCheckpoint).mockImplementation(async (operationId, deleted) => {
      // 原地撤销真的把表放回原簿（替身宿主上同一份目录），后续断言才不是空转。
      for (const name of deleted) host.addSheet(name);
      return {
        operationId, status: "completed", restoredSheets: [...deleted],
        failedSheets: [], unverifiedSheets: [], repairedFormulas: 0, repairedNames: 0,
        limitations: [], message: `已把 ${deleted.length} 张工作表放回原工作簿。`
      } as never;
    });
    const plan = { id: "delete-checkpoint", title: "删除汇总表", summary: "只删除汇总表", assumptions: [], warnings: [], acceptanceCriteria: [], actions: [{ type: "deleteWorksheet", sheet: TARGET }] };
    const view = await sendPlan(plan, "删除汇总表", { authorize: ALL_SHEETS });
    if (needsConsent) {
      const button = await view.findByRole("button", { name: "允许仅本地备份并继续" });
      expect(confirmWorksheetCheckpointBackup).not.toHaveBeenCalled();
      fireEvent.click(button);
    }
    await waitFor(() => expect(view.container.textContent).toContain("工作表删除已完成"));
    expect(view.getByRole("button", { name: "↶ 撤销本次执行" })).toBeTruthy();
    expect(executeWorksheetDeletion).toHaveBeenCalledOnce();
    expect(view.queryByRole("button", { name: "执行到 Excel" })).toBeNull();
    /*
     * 结果卡必须住在**发起这次删除的那条请求消息**里，而不是聊天底部：
     *  - 卡片所在的 `.message-row` 同时也是这条请求的计划卡；
     *  - 消息流最后一行不是删除卡（旧实现把整份检查点列表挂在末尾）。
     */
    const undoButton = view.getByRole("button", { name: "↶ 撤销本次执行" });
    const card = undoButton.closest(".worksheet-checkpoint-card") as HTMLElement;
    expect(card).not.toBeNull();
    expect(card.closest(".message-row")?.querySelector(".inline-plan")).not.toBeNull();
    const stream = view.container.querySelector(".message-stream") as HTMLElement;
    expect(stream.lastElementChild?.querySelector(".worksheet-checkpoint-card")).toBeNull();

    const rowsBeforeOpen = view.container.querySelectorAll(".message-row").length;
    fireEvent.click(undoButton);
    await waitFor(() => expect(restoreWorksheetCheckpoint).toHaveBeenCalledOnce());
    // 撤销是**原地**做的：对象是这次执行（同一个 operationId），且不再打开任何副本文件。
    expect(vi.mocked(restoreWorksheetCheckpoint).mock.calls[0][0]).toBe(executedOperationIds.at(-1));
    expect(vi.mocked(restoreWorksheetCheckpoint).mock.calls[0][1]).toEqual([TARGET]);
    expect(openWorksheetCheckpoint).not.toHaveBeenCalled();
    await waitFor(() => expect(view.container.textContent).toContain("已撤销本次删除"));
    expect(card.textContent).toContain("已恢复 1 张工作表");
    // 只更新这一张卡：不新增执行过程卡、不新增消息行。
    expect(view.container.querySelectorAll(".message-row").length).toBe(rowsBeforeOpen);
    expect(view.container.querySelectorAll(".worksheet-checkpoint-card").length).toBe(1);
    // 撤销不等于"打开副本"，也不宣称原表被整体还原成某个副本。
    expect(view.container.textContent).not.toContain("已打开恢复副本");
    expect(view.container.textContent).not.toContain("已还原");
    // 成功之后按钮收起（终态），不给第二次点击的机会。
    expect(view.queryByRole("button", { name: "↶ 撤销本次执行" })).toBeNull();
    expect(registerOfficeExecution).not.toHaveBeenCalled();
    expect(executeWorksheetDeletion).toHaveBeenCalledOnce();

    // 新会话不再继承上一段的删除卡（旧实现会把它们继续挂在消息流末尾）。
    fireEvent.click(view.getByRole("button", { name: "新对话" }));
    await waitFor(() => expect(view.container.querySelector(".worksheet-checkpoint-card")).toBeNull());
    expect(executeWorksheetDeletion).toHaveBeenCalledOnce();
  });

  it("每条删除结果留在各自请求消息内，不堆到后续消息下面", async () => {
    BOOK_URL = "file:///C:/work/book.xlsx";
    host = createHost(); installOfficeTestHost(host);
    vi.mocked(executeWorksheetDeletion).mockImplementation(async req => ({
      operationId: req.operationId, status: "completed", deletedSheets: [req.plan.actions[0]?.sheet ?? TARGET],
      remainingSheets: [SHEET, UNAUTHORIZED_SHEET], recoveryAvailable: true, message: "已核对剩余工作表。"
    }));
    vi.mocked(restoreWorksheetCheckpoint).mockImplementation(async (operationId, deleted) => {
      // 原地撤销真的把表放回原簿（替身宿主上同一份目录），后续断言才不是空转。
      for (const name of deleted) host.addSheet(name);
      return {
        operationId, status: "completed", restoredSheets: [...deleted],
        failedSheets: [], unverifiedSheets: [], repairedFormulas: 0, repairedNames: 0,
        limitations: [], message: `已把 ${deleted.length} 张工作表放回原工作簿。`
      } as never;
    });
    const first = { id: "delete-checkpoint-a", title: "删除汇总表", summary: "只删除汇总表", assumptions: [], warnings: [], acceptanceCriteria: [], actions: [{ type: "deleteWorksheet", sheet: TARGET }] };
    const view = await sendPlan(first, "删除汇总表", { authorize: ALL_SHEETS });
    await waitFor(() => expect(view.container.querySelectorAll(".worksheet-checkpoint-card").length).toBe(1));

    // 第二次删除：另一条请求（不同消息、不同计划），结果卡应落在它自己的位置上。
    const second = { ...first, id: "delete-checkpoint-b", title: "再删除一张", actions: [{ type: "deleteWorksheet", sheet: UNAUTHORIZED_SHEET }] };
    vi.mocked(streamAssistantResponse).mockResolvedValue({ kind: "plan", provider: "model", plan: second } as never);
    await waitFor(() => expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.disabled).toBe(false));
    fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "再删一张" } });
    fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
    await waitFor(() => expect(view.container.querySelectorAll(".worksheet-checkpoint-card").length).toBe(2));

    // 之后问一个普通问题：旧结果留在各自的消息里，不会漂到新消息下面。
    vi.mocked(streamAssistantResponse).mockResolvedValue({ kind: "answer", provider: "model", message: "这次只是查询。" } as never);
    await waitFor(() => expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.disabled).toBe(false));
    fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "不用改，看看就行" } });
    fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
    await waitFor(() => expect(view.container.textContent).toContain("这次只是查询。"));

    const cards = [...view.container.querySelectorAll<HTMLElement>(".worksheet-checkpoint-card")];
    expect(cards).toHaveLength(2);
    const rows = cards.map(card => card.closest(".message-row") as HTMLElement);
    expect(new Set(rows).size).toBe(2);
    for (const row of rows) expect(row.querySelector(".inline-plan")).not.toBeNull();
    const stream = view.container.querySelector(".message-stream") as HTMLElement;
    expect(stream.lastElementChild?.querySelector(".worksheet-checkpoint-card")).toBeNull();
    expect(executeWorksheetDeletion).toHaveBeenCalledTimes(2);
  });

  /*
   * 删表成功后，工作簿真的少了一张表。这**不是**用户改了授权范围。
   *
   * 旧行为：下一轮提问时目标守卫判 `scope_changed`，弹「工作表范围发生变化 /
   * 尚未写入」，用户被迫再确认一次范围；确认又会退役所有旧计划（结果卡随之消失）。
   *
   * 现在：删除成功 → 重读目录 → 用回执核对"剩下的表就是回执里的那一份" →
   * 只把授权范围与目标基线**收窄**到剩下的表。下一轮提问正常进行，
   * 既不弹错误提示，也不动已经完成的结果卡与撤销入口。
   */
  it("删表后的范围收缩按预期结果同步：不弹范围变化，旧待执行状态收尾，回执与撤销入口仍在", async () => {
    BOOK_URL = "file:///C:/work/book.xlsx";
    host = createHost(); installOfficeTestHost(host);
    vi.mocked(executeWorksheetDeletion).mockImplementation(async req => {
      // 删除真的发生了：此后宿主读到的表少一张，授权范围随之收缩。
      host.removeSheet(TARGET);
      return {
        operationId: req.operationId, status: "completed", deletedSheets: [TARGET],
        remainingSheets: [SHEET, UNAUTHORIZED_SHEET], recoveryAvailable: true,
        message: "已删除 1 张工作表，并核对剩余工作表。"
      };
    });
    vi.mocked(restoreWorksheetCheckpoint).mockImplementation(async (operationId, deleted) => {
      // 原地撤销真的把表放回原簿（替身宿主上同一份目录），后续断言才不是空转。
      for (const name of deleted) host.addSheet(name);
      return {
        operationId, status: "completed", restoredSheets: [...deleted],
        failedSheets: [], unverifiedSheets: [], repairedFormulas: 0, repairedNames: 0,
        limitations: [], message: `已把 ${deleted.length} 张工作表放回原工作簿。`
      } as never;
    });
    const plan = { id: "delete-then-shrink", title: "删除汇总表", summary: "只删除汇总表", assumptions: [], warnings: [], acceptanceCriteria: [], actions: [{ type: "deleteWorksheet", sheet: TARGET }] };
    const view = await sendPlan(plan, "删除汇总表", { authorize: ALL_SHEETS });
    await waitFor(() => expect(view.container.textContent).toContain("工作表删除已完成"));

    const card = view.container.querySelector<HTMLElement>(".worksheet-checkpoint-card");
    expect(card).not.toBeNull();
    const requestRow = card!.closest(".message-row") as HTMLElement;
    expect(requestRow).not.toBeNull();
    expect(card!.textContent).toContain("已删除 1 张工作表");

    /*
     * 完成状态收尾：同一条计划不再显示「准备就绪」「验证后保存」。
     * 计划内容仍留在折叠明细里（不是删掉计划）。
     */
    expect(requestRow.textContent).not.toContain("准备就绪");
    expect(requestRow.textContent).not.toContain("验证后保存");
    expect(requestRow.querySelector(".plan-details")).not.toBeNull();

    /*
     * 顺序：计划/过程在结果之前，撤销按钮属于结果卡内部。
     * 详情折叠时首屏只有一句结论。
     */
    const planBlock = requestRow.querySelector(".inline-plan") as HTMLElement;
    expect(planBlock).not.toBeNull();
    expect(planBlock.compareDocumentPosition(card!) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(card!.querySelector(".execution-recovery-footer")?.querySelector("button")?.textContent?.trim())
      .toBe("↶ 撤销本次执行");
    expect(card!.querySelector("details")?.open).toBe(false);

    // 下一轮提问：范围收缩已经同步，不再要求重新确认，也不出现"尚未写入"。
    vi.mocked(streamAssistantResponse).mockResolvedValue({ kind: "answer", provider: "model", message: "现在还剩两张表。" } as never);
    await waitFor(() => expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.disabled).toBe(false));
    fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "现在还剩几张表" } });
    fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
    await waitFor(() => expect(view.container.textContent).toContain("现在还剩两张表。"));
    expect(view.container.querySelector(".target-scope-notice")).toBeNull();
    expect(view.container.textContent).not.toContain("工作表范围发生变化");
    expect(view.container.textContent).not.toContain("尚未写入");
    expect(view.container.querySelectorAll(".worksheet-checkpoint-card")).toHaveLength(1);
    expect(executeWorksheetDeletion).toHaveBeenCalledOnce();

    // 撤销是**原地**的：不重跑删除、不新增消息行、不打开副本，成功只更新这张卡。
    const undoButton = view.getByRole("button", { name: "↶ 撤销本次执行" });
    const rowsBeforeOpen = view.container.querySelectorAll(".message-row").length;
    fireEvent.click(undoButton);
    await waitFor(() => expect(restoreWorksheetCheckpoint).toHaveBeenCalledOnce());
    expect(openWorksheetCheckpoint).not.toHaveBeenCalled();
    await waitFor(() => expect(view.container.textContent).toContain("已撤销本次删除"));
    expect(card!.textContent).toContain("已恢复 1 张工作表");
    expect(view.container.querySelectorAll(".message-row").length).toBe(rowsBeforeOpen);
    expect(view.container.querySelectorAll(".worksheet-checkpoint-card").length).toBe(1);
    expect(executeWorksheetDeletion).toHaveBeenCalledOnce();
    expect(view.container.textContent).not.toContain("已打开恢复副本");
    expect(view.container.textContent).not.toContain("已还原");

    /*
     * 撤销把「汇总表」放回了工作簿，但**授权范围没变**（恢复出来的表不会被自动授权）。
     * 下一轮提问不能因为这次预期内的目录增加而弹「目标发生变化」——这就是"撤销后
     * 误报范围变化"的反面判据。
     */
    vi.mocked(streamAssistantResponse).mockResolvedValue({ kind: "answer", provider: "model", message: "撤销后仍只授权两张表。" } as never);
    await waitFor(() => expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.disabled).toBe(false));
    fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "撤销后再问一次" } });
    fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
    await waitFor(() => expect(view.container.textContent).toContain("撤销后仍只授权两张表。"));
    expect(view.container.querySelector(".target-scope-notice")).toBeNull();
    expect(view.container.textContent).not.toContain("工作表范围发生变化");
    expect(view.container.textContent).not.toContain("尚未写入");
    expect(executeWorksheetDeletion).toHaveBeenCalledOnce();
    expect(restoreWorksheetCheckpoint).toHaveBeenCalledOnce();
  });

  /*
   * 反面：回执与实际剩余表**对不上**时，绝不能把范围收缩当成预期结果。
   *
   * 这里模拟"外部又少了一张已授权表"：授权 A/B/C，回执只说删了 B，
   * 实际 A/C 里的 C 也不见了。核对必须拒绝，下一轮提问仍被拦下。
   *
   * 后台目录刷新保留三表勾选，由目标守卫明确提示范围变化；不得静默代选。
   */
  it("实际表集合与回执不一致时不静默收窄，仍按范围变化拦截", async () => {
    BOOK_URL = "file:///C:/work/book.xlsx";
    host = createHost(); installOfficeTestHost(host);
    vi.mocked(executeWorksheetDeletion).mockImplementation(async req => {
      // 宿主少了两张表，回执却只承认删了一张 —— 有一处变化不是这次执行造成的。
      host.removeSheet(TARGET);
      host.removeSheet(UNAUTHORIZED_SHEET);
      return {
        operationId: req.operationId, status: "completed", deletedSheets: [TARGET],
        remainingSheets: [SHEET], recoveryAvailable: true,
        message: "已删除 1 张工作表，并核对剩余工作表。"
      };
    });
    const plan = { id: "delete-receipt-mismatch", title: "删除汇总表", summary: "只删除汇总表", assumptions: [], warnings: [], acceptanceCriteria: [], actions: [{ type: "deleteWorksheet", sheet: TARGET }] };
    const view = await sendPlan(plan, "删除汇总表", { authorize: ALL_SHEETS });
    await waitFor(() => expect(view.container.textContent).toContain("工作表删除已完成"));

    vi.mocked(streamAssistantResponse).mockResolvedValue({ kind: "answer", provider: "model", message: "不该跑到这一步。" } as never);
    const planTurns = vi.mocked(streamAssistantResponse).mock.calls.length;
    await waitFor(() => expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.disabled).toBe(false));
    fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "继续处理" } });
    fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
    await waitFor(() =>
      expect(view.container.querySelector(".authorized-target-error")?.textContent
        ?? view.container.querySelector(".target-scope-notice")?.textContent
        ?? "").toContain("工作表范围发生变化")
    );
    expect(view.container.querySelector(".scope-trigger")?.textContent).toContain("已固定 3 个工作表");
    // 没有执行、没有读字段。
    expect(view.container.textContent).not.toContain("不该跑到这一步。");
    expect(vi.mocked(streamAssistantResponse).mock.calls.length).toBe(planTurns);
    // 打开、关闭选择器也不收窄固定选择。
    if (!view.container.querySelector(".scope-popover")) {
      fireEvent.click(view.container.querySelector(".scope-trigger")!);
    }
    const discard = await waitFor(() => {
      const node = view.container.querySelector<HTMLButtonElement>(".scope-close");
      if (!node) throw new Error("范围选择器没有关闭入口");
      return node;
    });
    fireEvent.click(discard);
    await waitFor(() => expect(view.container.querySelector(".scope-popover")).toBeNull());
    expect(view.container.querySelector(".scope-trigger")?.textContent).toContain("已固定 3 个工作表");
    // 回执本身仍留在原请求位置：拦截针对的是下一轮，不是抹掉已发生的结果。
    expect(view.container.querySelectorAll(".worksheet-checkpoint-card")).toHaveLength(1);
    expect(executeWorksheetDeletion).toHaveBeenCalledOnce();
  });
});

/* ────────────────────────────────────────────────────────────────────────────
 * 数据范围授权同步：**确认过的范围不许被要求再确认一次**
 *
 * 现场反馈：勾选工作表并「确认」之后，仍反复弹「工作表范围发生变化」。
 *
 * 这一组全部挂在**真实链路**上：真实 `<App/>`、真实 `useAuthorizedTarget`
 * （含 2s 监视器与 focus 重查）、真实 `useScopeSelection`（草稿 / 已提交两态）、
 * 真实 `.scope-popover` 弹层（点「选择多个」、勾表、点「确认」/「取消」）。
 * 只有网络与目标读数（`inspectOfficeAuthorizedTarget`）是替身。
 *
 * 写法上的两条纪律：
 *   · 先断言**请求真的跑通**（回答出现），再断言没有提示 —— 只断言"没有提示"
 *     在被拦下时同样成立，而那正是缺陷的样子；
 *   · 每条用例都带上"这一轮检查真的发生了"的锚点（读数次数增加），
 *     否则"没有提示"可能只是守卫根本没跑。
 * ──────────────────────────────────────────────────────────────────────────── */

function scopeNotice(view: ReturnType<typeof render>): string | null {
  const node = view.container.querySelector(".target-scope-notice");
  return node ? (node.textContent ?? "") : null;
}

function scopeModeButton(view: ReturnType<typeof render>, label: "跟随当前" | "选择多个") {
  const button = [...view.container.querySelectorAll<HTMLButtonElement>(".scope-mode-list > button")]
    .find((candidate) => (candidate.textContent ?? "").includes(label));
  if (!button) throw new Error(`范围面板缺少「${label}」`);
  return button;
}

async function bootApp() {
  const view = render(<App />);
  await waitFor(() => expect(view.container.querySelector(".composer textarea")).not.toBeNull());
  await waitFor(() =>
    expect(view.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title).toContain("当前模型")
  );
  return view;
}

/** 一轮后台检查：与 2s 轮询、窗口重新获得焦点是同一条出口。 */
async function backgroundCheck(view: ReturnType<typeof render>) {
  const before = vi.mocked(inspectOfficeAuthorizedTarget).mock.calls.length;
  await act(async () => {
    window.dispatchEvent(new Event("focus"));
    await Promise.resolve();
  });
  await waitFor(() =>
    expect(
      vi.mocked(inspectOfficeAuthorizedTarget).mock.calls.length,
      "后台检查必须真的跑了一轮，否则后面那句「没有提示」是空的"
    ).toBeGreaterThan(before)
  );
  await act(async () => { await Promise.resolve(); });
}

async function sendAndExpectAnswer(
  view: ReturnType<typeof render>,
  prompt: string,
  answer: string
) {
  vi.mocked(understandMessage).mockResolvedValue({
    kind: "needs_workbook_context",
    summary: prompt,
    provider: "model"
  } as never);
  vi.mocked(checkIntent).mockResolvedValue({
    kind: "proceed",
    summary: prompt,
    confirmedPrompt: prompt,
    provider: "model"
  } as never);
  vi.mocked(streamAssistantResponse).mockResolvedValue({
    kind: "answer",
    provider: "model",
    message: answer
  } as never);
  await waitFor(() =>
    expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.disabled).toBe(false)
  );
  fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: prompt } });
  fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
  await waitFor(() => {
    const notice = scopeNotice(view);
    expect(notice, `请求被范围提示拦下了：${notice ?? ""}`).toBeNull();
    expect(text(view)).toContain(answer);
  });
}

describe("数据范围授权同步（真实 App / 守卫 / 范围 Hook）", () => {
  /*
   * ① 最常走的那条路：勾两张表 → 「确认」→ 等后台检查跑一轮 → 再发一条需求。
   *
   * 判据的顺序是有意的：先证"下一句需求真的跑通了"，再证"没有范围提示"。
   */
  it("①多表确认后：后台检查与下一条需求都不再要求确认", async () => {
    const view = await bootApp();
    await setAuthorizedScope(view, [SHEET, TARGET]);
    expect(view.container.querySelector(".scope-trigger")?.textContent).toContain("已固定 2 个工作表");

    await backgroundCheck(view);
    expect(scopeNotice(view), "确认过的范围被后台检查判成了变化").toBeNull();

    await sendAndExpectAnswer(view, "按这两张表汇总", "已按两张表汇总。");
    expect(scopeNotice(view)).toBeNull();
  });

  /*
   * ③ 固定多表授权后切换活动工作表：守卫读的是**已提交授权**，不是活动表。
   * 活动表变化只影响"跟随当前"口径，不该动固定范围。
   */
  it("③固定多表授权后切换活动工作表：不误报范围变化", async () => {
    const view = await bootApp();
    await setAuthorizedScope(view, [SHEET, TARGET]);
    // 用户在 Excel 里点了另一张表的标签 —— 宿主的活动表真的变了。
    host.activateSheet(TARGET);
    expect(host.activeSheetName()).toBe(TARGET);

    await backgroundCheck(view);
    expect(scopeNotice(view), "切了活动表就把固定范围判成了变化").toBeNull();

    await sendAndExpectAnswer(view, "换到汇总表看总数", "汇总表总数已读取。");
    expect(scopeNotice(view)).toBeNull();
  });

  /*
   * ④ 「只是看看」不是授权变更：在弹层里切到「跟随当前」再点「取消」，
   * 已提交授权与守卫基线必须一字未改，下一条需求照常跑。
   *
   * 这条同时钉住两个面：触发器仍报"已固定 2 个工作表"（草稿退回已提交的**口径**，
   * 不只是名字），以及守卫读的是已提交口径而不是弹层里的草稿口径。
   */
  it("④切换模式即授权，收起不会回退也不会重复提示", async () => {
    const view = await bootApp();
    await setAuthorizedScope(view, [SHEET, TARGET]);
    await openScopePicker(view);
    fireEvent.click(scopeModeButton(view, "跟随当前"));
    await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
    fireEvent.click(view.container.querySelector(".scope-close")!);
    expect(view.container.querySelector(".scope-trigger")?.textContent).toContain("当前表 综合票房");
    await backgroundCheck(view);
    expect(scopeNotice(view)).toBeNull();
    await sendAndExpectAnswer(view, "按当前表继续", "已按当前表处理。");
    expect(scopeNotice(view)).toBeNull();
  });

  it("⑥确认期间在飞的旧读数：返回后既不弹提示，也不放行旧请求", async () => {
    const view = await bootApp();
    await setAuthorizedScope(view, [SHEET]);

    // 从这里开始读数受测试控制：**第一份**挂起（模拟慢宿主），其余立刻返回。
    let holdOnce = true;
    let release: (() => void) | null = null;
    vi.mocked(inspectOfficeAuthorizedTarget).mockImplementation(((names: readonly string[] | null) => {
      const snapshot = targetFor(names);
      if (holdOnce) {
        holdOnce = false;
        return new Promise((resolve) => { release = () => resolve(snapshot); }) as never;
      }
      return Promise.resolve(snapshot) as never;
    }) as never);
    const releaseHeld = () => {
      const resume = release;
      release = null;
      resume?.();
    };

    act(() => { window.dispatchEvent(new Event("focus")); });
    await waitFor(() =>
      expect(vi.mocked(inspectOfficeAuthorizedTarget).mock.calls.length).toBeGreaterThan(1)
    );

    // 旧读数还挂着的时候，用户确认了一个更大的范围。
    await openScopePicker(view);
    fireEvent.click(sheetToggle(view, TARGET));
    const commit = view.container.querySelector<HTMLButtonElement>(".scope-close")!;

    // 确认路径按规则先等这份在飞读数 —— 放行它（它描述的是旧范围）。
    await act(async () => {
      releaseHeld();
      await Promise.resolve();
    });
    await waitFor(() => expect(commit.getAttribute("data-scope-pending")).toBe("false"));
    fireEvent.click(commit);
    await waitFor(() => expect(view.container.querySelector(".scope-popover")).toBeNull());
    // 旧读数返回后的那一次比对在这里发生。
    await act(async () => { await Promise.resolve(); });
    expect(scopeNotice(view), "确认期间的旧读数把新基线判成了变化").toBeNull();

    await sendAndExpectAnswer(view, "用两张表继续", "已按两张表继续。");
    expect(scopeNotice(view)).toBeNull();
  });

  /*
   * ⑧ 「跟随当前」口径也要按自己的口径判：确认时授权的是"这个文件的活动表"。
   * 切到另一张表再说同一件事，读取目标就变了 —— 必须先确认，不能静默改读。
   */
  it("⑧跟随当前口径：确认后切到另一张表要先确认，不会静默改读", async () => {
    const view = await bootApp();
    // 默认就是「跟随当前」，直接点「确认」。
    await openScopePicker(view);
    fireEvent.click(scopeModeButton(view, "跟随当前"));
    const commit = view.container.querySelector<HTMLButtonElement>(".scope-close")!;
    await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
    fireEvent.click(commit);
    await waitFor(() => expect(view.container.querySelector(".scope-popover")).toBeNull());
    expect(view.container.querySelector(".scope-trigger")?.textContent).toContain("当前表 综合票房");

    host.activateSheet(TARGET);
    await backgroundCheck(view);
    expect(scopeNotice(view), "活动表变了却仍按旧的活动表放行").toContain("工作表范围发生变化");

    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "汇总", provider: "model" } as never);
    vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", summary: "汇总", confirmedPrompt: "汇总", provider: "model" } as never);
    vi.mocked(streamAssistantResponse).mockResolvedValue({ kind: "answer", provider: "model", message: "不该读到汇总表。" } as never);
    fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "汇总一下" } });
    fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
    await waitFor(() => expect(view.container.textContent).toContain("工作表范围发生变化"));
    expect(text(view)).not.toContain("不该读到汇总表。");
    expect(streamAssistantResponse).not.toHaveBeenCalled();
  });

  /*
   * ⑨ 草稿口径不是授权口径。
   *
   * 用户点开面板、切到「跟随当前」、然后**直接把弹层关掉**（不点「取消」也不点
   * 「确认」），之后又在 Excel 里翻到另一张表 —— 这是最常见的"我只是看看"。
   * 三件事必须同时成立：
   *   · 后台检查不能因此弹「工作表范围发生变化」（守卫读**已提交**口径）；
   *   · 发送也不能拿草稿口径去读表：草稿口径与已提交不一致时，请求停在
   *     「正在核对勾选的工作表」，一个字段都不读；
   *   · 尤其不能因为"名字看起来一样"就把活动表当授权表读出去。
   *
   * 这一条刻意让**名字完全相同**（草稿与已提交都是「综合票房」），只有口径不同：
   * 否则"名字脏"会抢在前面把请求拦下，口径这条判据就没被单独验证。
   */
  it("⑨关闭选择器不撤回新口径，随后真切表仍被拦截", async () => {
    const view = await bootApp();
    await setAuthorizedScope(view, [SHEET]);
    await openScopePicker(view);
    fireEvent.click(scopeModeButton(view, "跟随当前"));
    await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
    fireEvent.click(view.container.querySelector(".scope-trigger")!);
    host.activateSheet(UNAUTHORIZED_SHEET);
    await backgroundCheck(view);
    expect(scopeNotice(view)).toContain("工作表范围发生变化");
    expect(streamAssistantResponse).not.toHaveBeenCalled();
  });

  it("A.按非目录顺序确认三张表后，首条需求不再要求确认目标", async () => {
    const view = await bootApp();
    host.activateSheet(UNAUTHORIZED_SHEET);
    await setAuthorizedScopeInOrder(view, [UNAUTHORIZED_SHEET, TARGET, SHEET]);
    expect(view.container.querySelector(".scope-trigger")?.textContent).toContain("已固定 3 个工作表");

    await sendAndExpectAnswer(view, "把这三张表合起来看", "已按三张表汇总。");
    expect(scopeNotice(view)).toBeNull();
  });

  /*
   * ⑦ 真正的变化仍然拦得住：外部删掉一张已授权表，不能因为"刚确认过"就放行。
   */
  it("⑦外部删掉已授权表：仍然按范围变化拦截，且不执行需求", async () => {
    const view = await bootApp();
    await setAuthorizedScope(view, [SHEET, TARGET]);
    // 外部把「汇总表」删了（不是这次执行造成的，没有回执）。
    host.removeSheet(TARGET);
    await backgroundCheck(view);

    vi.mocked(understandMessage).mockResolvedValue({
      kind: "needs_workbook_context",
      summary: "继续",
      provider: "model"
    } as never);
    vi.mocked(checkIntent).mockResolvedValue({
      kind: "proceed",
      summary: "继续",
      confirmedPrompt: "继续",
      provider: "model"
    } as never);
    vi.mocked(streamAssistantResponse).mockResolvedValue({
      kind: "answer",
      provider: "model",
      message: "不该跑到这一步。"
    } as never);
    fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "继续处理" } });
    fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
    await waitFor(() => expect(scopeNotice(view)).not.toBeNull());
    expect(scopeNotice(view)).toContain("工作表范围发生变化");
    expect(text(view)).not.toContain("不该跑到这一步。");
    expect(streamAssistantResponse).not.toHaveBeenCalled();
  });
});
