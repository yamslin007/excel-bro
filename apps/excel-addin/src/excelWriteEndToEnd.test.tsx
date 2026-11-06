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

/**
 * 恢复操作区那颗主按钮的**名字族**（2026-11-02）。
 *
 * 主操作按证据分档：「检查已写入结果」/「检查写入结果」/「检查当前结果」，
 * 续跑那档带剩余步数「继续剩余 N 步」。测试若写死某一个名字，就等于默认"分档不会发生"——
 * 而分档恰恰是"状态未知时不许宣称已写入"的落地点。所以这里按名字族找：
 * 负向用例（`queryByRole(...).toBeNull()`）也因此仍然在测真东西，不是自动通过。
 */
const RECOVERY_CHECK_BUTTON = /^检查/;
const RECOVERY_RESUME_BUTTON = /^继续剩余 \d+ 步$/;
import { act } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import { closeExcelReviewCopy, selectExcelReviewCopySheet, prepareExcelReviewLayout, arrangeExcelReviewLayout, executeWorksheetDeletion, openWorksheetCheckpoint, restoreWorksheetCheckpoint, confirmWorksheetCheckpointBackup } from "./api";
import { inspectOfficeAuthorizedTarget } from "./officeAuthorizedTarget";
import { checkIntent, streamAssistantResponse, understandMessage, continuePlanningTask, getPlanningTaskStatus } from "./api";
import {
  reportOfficeHostWrites,
  listPendingOfficeExecutions,
  claimOfficeExecutionResume,
  completeOfficeAction,
  finishOfficeExecution,
  getOfficeExecutionStatus,
  pollSdkBridgeTool,
  postSdkBridgeReceipt,
  reconcileOfficeActionReceipt,
  recordExecutionDisposition,
  registerOfficeExecution,
  sendExecutionReview,
  startOfficeAction,
  suspendOfficeExecution
} from "./api";
import {
  executePlan,
  reconcileOfficeRangeAction,
  captureWorkbook,
  captureWorkbookCatalog,
  captureWorkbookStructure,
  currentOfficeWorkbookIdentity,
  invalidateWorkbookStructureCache
} from "./excel";
import { createOfficeTestHost, installOfficeTestHost, type TestOfficeHost } from "./officeTestHost";
import {
  clearPendingOfficeExecution,
  loadPendingOfficeExecution,
  savePendingOfficeExecution
} from "./officeOperationRecovery";
import { resetTrustedSourceReads } from "./planSourceDependency";

/*
 * 附件接缝的**替身**：真机里 `prepareImageFile` 要 `new Image()` + canvas 解码压缩，
 * jsdom 没有这套能力（`Image` 永不触发 onload）。这里只替掉"解码与压缩"这一步，
 * 尺寸限额、状态机、主按钮判据全都是真的 —— 于是"只贴图片、没打字"这条分支可以被
 * 断言，而不是靠"同源结构"推断。
 */
vi.mock("./imageAttachments", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./imageAttachments")>();
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

const SHEET = "综合票房";
/** 写入目标：与来源表不同的另一张**已授权**表。 */
const TARGET = "汇总表";
/** 工作簿里存在、但本次未勾选的工作表。 */
const UNAUTHORIZED_SHEET = "备用表";
const SHEET_IDS: Record<string, string> = {
  [SHEET]: "sheet-1",
  [TARGET]: "sheet-2",
  [UNAUTHORIZED_SHEET]: "sheet-3",
  // 三步公式计划用的三张结果表（`Excel 写侧：三步公式计划自动执行`）。
  "结果一": "sheet-4",
  "结果二": "sheet-5",
  "结果三": "sheet-6",
  // 「重开后撤销未完成执行」用的夹具：与真机验收同一个形状（来源 + 三张结果表）。
  "来源": "sheet-7",
  "结果1": "sheet-8",
  "结果2": "sheet-9",
  "结果3": "sheet-10"
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
  reportOfficeHostWrites: vi.fn(async () => ({})),
  registerOfficeExecution: vi.fn(),
  startOfficeAction: vi.fn(),
  completeOfficeAction: vi.fn(),
  finishOfficeExecution: vi.fn(),
  suspendOfficeExecution: vi.fn(),
  getOfficeExecutionStatus: vi.fn(),
  /* 发现未完成执行：默认"这个工作簿没有待续写的执行"。
     未替身时产品会真发 GET，测试的网络闸门会直接让用例失败 —— 也就是说，
     每条挂载 App 的用例都必须显式表态"这个簿有没有未完成执行"。 */
  listPendingOfficeExecutions: vi.fn(async (workbookIdentity: string) => ({
    schemaVersion: 1 as const,
    namespace: "office" as const,
    workbookIdentity,
    limit: 20,
    truncated: false,
    candidates: []
  })),
  claimOfficeExecutionResume: vi.fn(),
  reconcileOfficeActionReceipt: vi.fn(),
  /* 用户处置留痕（放弃/撤销）：真实网络出口，必须替身，否则测试会往用户服务端发请求。 */
  recordExecutionDisposition: vi.fn(async () => undefined),
  /* 审阅决定落账（审阅生命周期）：也是真实网络出口（POST /api/{ns}/executions/{id}/review）。
     不替身就会让结果审阅的确认/放弃真的发到用户服务端。 */
  sendExecutionReview: vi.fn(async () => undefined),
  // 工具桥：走真实状态机，只把"网络"换成替身（请求体与回执仍是产品代码拼的）。
  executeWorksheetDeletion: vi.fn(),
  confirmWorksheetCheckpointBackup: vi.fn(),
  openWorksheetCheckpoint: vi.fn(),
  restoreWorksheetCheckpoint: vi.fn(),
  pollSdkBridgeTool: vi.fn(),
  postSdkBridgeReceipt: vi.fn(),
  // 本机窗口排列服务：替身不模拟它，明确拒绝（产品会如实降级为"手动排列"）。
  arrangeExcelReviewLayout: vi.fn(),
  closeExcelReviewCopy: vi.fn(),
  selectExcelReviewCopySheet: vi.fn(async () => ({status:'selected'})),
  setExcelReviewCopyViewing: vi.fn(async (_request, viewing) => viewing
    ? {status:'arranged',layoutVersion:3,copySide:'left',originalSide:'right',synchronousScrolling:'enabled'}
    : {status:'suspended',copyRetained:true}),
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
  ],
  // 真实产品里编译器会给 writeFormulas 计划补 formulasEqual 验收条件；
  // `verificationGaps` 只认计划声明的验收条件（inferredCriteria 里的推断条件不算），
  // 没有它每步都会落进"缺少独立验收"。
  acceptanceCriteria: [
    {
      type: "formulasEqual" as const,
      sheet: TARGET,
      range: "A1:B1",
      expected: [["=\"影片\"", "=1200"]]
    },
    {
      type: "formulasEqual" as const,
      sheet: TARGET,
      range: "A2:B2",
      expected: [["=\"乙\"", "=980"]]
    }
  ]
};

/** 桥上的确定性读取：`scope: "active"` 只读当前工作表（= 来源依赖只含它）。 */
function bridgeQuery(scope: "active" | "selected", sheets?: string[]) {
  /*
   * `scope: "selected"` 不再等于"把已授权表全读一遍"：多张表时必须显式绑定
   * （`sheets` 或 range 带表名），否则执行器直接拒绝。夹具要"真的读整片授权区"
   * 就得把表名列出来 —— 这也正是新语义要的显式性。
   */
  return {
    mode: "rows" as const,
    scope,
    ...(sheets && sheets.length > 0 ? { sheets } : {}),
    fields: ["影片"],
    limit: 10
  };
}

const BRIDGE_OP = `op_${"7".repeat(32)}`;
const BRIDGE_FP = "c".repeat(64);

interface PlanTurnOptions {
  planOrigin?: "literal_range_edit";
  pauseBeforePlan?: boolean;
  probe?: boolean;
  /** 通过真实范围选择弹层建立的授权范围。 */
  authorize?: string[];
  /** 弹层里要遍历的表清单（默认 `ALL_SHEETS`）；自定义夹具时传它。 */
  authorizePool?: string[];
  /** 让本轮真的经工具桥读一次宿主，从而产生可信来源记录。 */
  read?: { scope: "active" | "selected"; sheets?: string[] };
  /**
   * 计划**返回之前**的一下改动（改源值、改公式文本…）。
   *
   * 自动执行上线后，计划一到就开始跑，"计划到达之后再改源"变成了竞态；
   * 真正要验的场景是"本轮读过源（锚点已成立）→ 源变了 → 才返回计划"，
   * 这个钩子就是那段顺序。
   */
  beforePlan?: () => void;
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
          arguments: bridgeQuery(options.read.scope, options.read.sheets),
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
        expect(vi.mocked(postSdkBridgeReceipt).mock.calls.length).toBeGreaterThan(0),
        { timeout: 5000 }
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
    options.beforePlan?.();
    return { kind: "plan", provider: "model", plan: plan as never, planOrigin: options.planOrigin };
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
  if (options.authorize) await setAuthorizedScope(view, options.authorize, options.authorizePool);
  if (options.probe) {
    view.container.querySelector("details.recovery-probe-control")!.setAttribute("open", "");
    fireEvent.click(await view.findByRole("button", {name:"启用下一次第二步中断"}));
  }
  fireEvent.change(view.container.querySelector(".composer textarea")!, {
    target: { value: prompt }
  });
  fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
  if (options.pauseBeforePlan) await view.findByRole("button", { name: /继续分析/ });
  else await waitFor(() => expect(view.container.querySelector(".inline-plan")).not.toBeNull(), { timeout: 4000 });
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

async function setAuthorizedScope(view: ReturnType<typeof render>, names: readonly string[],
  pool: readonly string[] = ALL_SHEETS) {
  await openScopePicker(view);
  const manual = [...view.container.querySelectorAll<HTMLButtonElement>(".scope-mode-list button")]
    .find((candidate) => candidate.textContent?.includes("选择多个"));
  if (!manual) throw new Error("「选择多个」入口尚未出现");
  fireEvent.click(manual);
  for (const name of pool) {
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
  await waitFor(() => expect(view.queryByRole("button", { name: "确认全部变更" })).not.toBeNull());
}
async function clickSave(view: ReturnType<typeof render>) {
  const save = view.getByRole("button", { name: "确认全部变更" });
  fireEvent.click(save);
  /*
   * 成功终态下固定条会**收起**（不再占着输入框上方），所以"确认成了"的观测点
   * 改成对话里的收尾消息 —— 那是状态真正的去处，也是这条判据想证的事。
   * 顺带钉住：收起的同时固定条不再提供任何审阅按钮。
   */
  await waitFor(() => expect(text(view)).toContain("已确认变更"));
  expect(view.container.querySelector(".result-review-bar")).toBeNull();
  expect(view.queryByRole("button", { name: "确认全部变更" })).toBeNull();
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
    vi.mocked(claimOfficeExecutionResume).mockImplementation(async()=>{
      state={...state,status:"running",revision:state.revision+1,ownerToken:"resume-owner"};return state;
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
    /*
     * 恢复卡是**两步**，不是一个「核对并继续」：
     *   ①「核对在途动作」只补记回执（零单元格写入）；
     *   ②「继续剩余动作」只跑剩下的动作（已完成的第 2 步不重放）。
     */
    const reconcileButton=await reopened.findByRole("button",{name: RECOVERY_CHECK_BUTTON});
    await waitFor(()=>expect((reconcileButton as HTMLButtonElement).disabled).toBe(false));
    expect(reopened.queryByRole("button",{name:RECOVERY_RESUME_BUTTON})).toBeNull();
    act(()=>{fireEvent.click(reconcileButton);fireEvent.click(reconcileButton);});
    await waitFor(()=>expect(reconcileOfficeActionReceipt).toHaveBeenCalledTimes(1));
    // 核对这一下：一个单元格都没写。
    expect(contentWrites(host.originalWrites())).toEqual(before);
    const resumeButton=await reopened.findByRole("button",{name:RECOVERY_RESUME_BUTTON});
    await waitFor(()=>expect((resumeButton as HTMLButtonElement).disabled).toBe(false));
    fireEvent.click(resumeButton);
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
    await waitFor(() => expect(vi.mocked(reportOfficeHostWrites).mock.calls.flatMap(call => call[1].entries).map(entry => entry.actionIndex)).toEqual([0, 1, 2]));
    const resumedAudit = vi.mocked(reportOfficeHostWrites).mock.calls.flatMap(call => call[1].entries).find(entry => entry.actionIndex === 2)!;
    expect(resumedAudit).toMatchObject({outcome: "written", readback: "verified"});
    expect(resumedAudit.hostWriteCalls).toBeGreaterThan(0);
    await stageReview(reopened);
    expect(reopened.container.querySelector('.result-review-bar')?.textContent).not.toContain('执行未完成');
    expect(text(reopened)).toContain('从第 3 步继续并完成');
    const writes=contentWrites(host.originalWrites());
    fireEvent.click(reopened.getByRole('button',{name:'双视图对比'}));
    await waitFor(()=>expect(text(reopened)).toContain('没有保留完整的执行前表快照'));
    expect(host.createWorkbookCalls()).toHaveLength(0); // Never manufacture a blank "whole sheet" from partial undo data.
    expect(contentWrites(host.originalWrites())).toEqual(writes);
    if(outcome==='confirm') {
      const confirm=reopened.getByRole('button',{name:'确认全部变更'});
      await waitFor(()=>expect((confirm as HTMLButtonElement).disabled).toBe(false));fireEvent.click(confirm);
      await waitFor(()=>expect(text(reopened)).toContain('已确认变更'));
      expect(reopened.queryByRole('button',{name:'确认全部变更'})).toBeNull();
      // This SharePoint host cannot close a native local copy; a settled cleanup notice is legitimate.
      expect(reopened.container.querySelector('.result-review-bar')?.textContent ?? '').not.toContain('执行未完成');
      // ③确认只清标记：续跑 + 确认的整个链路没有产生第二次写入。
      expect(contentWrites(host.originalWrites())).toEqual(writes);
    } else {
      const discard=reopened.getByRole('button',{name:'放弃全部变更'});
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
      fireEvent.click(view.getByRole('button',{name:'放弃全部变更'}));
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

  it("续跑交付计划后自动写入并回读，确认全部变更只清标记且不重复执行", async () => {
    vi.mocked(getPlanningTaskStatus).mockResolvedValue({
      kind: "planning_pause", turnId: "resumed-write-turn", status: "interrupted", revision: 1,
      lastSequence: 0, canContinue: true, pauseType: "interrupted", reason: "规划已暂停",
      extensionSeconds: 60, activeElapsedSeconds: 1, currentModelId: "claude-agent-sdk",
      canSwitchModel: false, availableModelConnections: [],
      continuation: { available: true, code: null, message: "", stage: "after_tool", sourceVerificationRequired: true }
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
    expect(view.getByRole("button", { name: "放弃全部变更" })).toBeTruthy();
    expect(view.getByRole("button", { name: "双视图对比" })).toBeTruthy();
    const writes = contentWrites(host.originalWrites()).length;
    await clickSave(view);
    expect(contentWrites(host.originalWrites())).toHaveLength(writes);
    expect(continuePlanningTask).toHaveBeenCalledOnce();
  });
  /**
   * **规划继续**这一条入口也要核对锚点（不是只有"正常收计划"那条路）。
   *
   * 服务端明确说"这次断点不需要核实来源" ⇒ 本轮**没有**现场读数，也就没有可信锚点；
   * 而计划里写的是固定值 ⇒ 必须拒绝，不能拿"续跑那一刻再读一次"凑一个基线出来。
   */
  it("续跑没有核实来源时：写固定值的计划写入 0 次，并说明缺的是记录而不是数据变化", async () => {
    vi.mocked(getPlanningTaskStatus).mockResolvedValue({
      kind: "planning_pause", turnId: "resumed-no-verify-turn", status: "interrupted", revision: 1,
      lastSequence: 0, canContinue: true, pauseType: "interrupted", reason: "规划已暂停",
      extensionSeconds: 60, activeElapsedSeconds: 1, currentModelId: "claude-agent-sdk",
      canSwitchModel: false, availableModelConnections: [],
      continuation: { available: true, code: null, message: "", stage: "after_tool", sourceVerificationRequired: false }
    } as never);
    vi.mocked(continuePlanningTask).mockResolvedValue({ kind: "plan", provider: "model", plan: reviewPlan } as never);
    const view = await sendPlan(reviewPlan, "更新结果", { pauseBeforePlan: true });
    const before = host.originalWrites().length;
    expect(before).toBe(0);
    fireEvent.click(await view.findByRole("button", { name: /继续分析/ }));
    await waitFor(() => {
      if (!(view.container.textContent ?? "").includes("没有记录它是按哪一版数据算出来的")) {
        throw new Error(
          `续跑的计划没有被"缺锚点"拦下；当前写入 ${host.originalWrites().length} 次`
        );
      }
    });
    // 目标区一格没动；文案说的是"没有记录"，不是"数据已经变化"。
    expect(host.originalWrites()).toHaveLength(0);
    expect(host.cell(SHEET, "G1")).not.toBe("合计");
    expect(view.container.textContent).not.toContain("数据来源在预览后发生了变化");
  }, 30_000);

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
    expect(seed).toContain("票房"); // Unchanged context on the authorized target is part of the before image.
    expect(text(view)).toContain("公式采用执行前的值，不重新计算");
    expect(contentWrites(host.originalWrites())).toHaveLength(writes);
    expect(view.getByRole("button", { name: "确认全部变更" })).toBeTruthy();
    fireEvent.click(view.getByRole("button", { name: "放弃全部变更" }));
    await waitFor(() => expect(text(view)).toContain("已放弃并恢复"));
    expect(host.cell(SHEET, "G1")).toBe("");
    expect(host.matrix(SHEET, "A1:C3")).toEqual(before);
    expect(host.createWorkbookCalls()).toHaveLength(1);
  });
  it("含工作簿名称时确认全部变更仍只清标记，不重放写入", async () => {
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
    await waitFor(() => expect(text(view)).toContain("未执行"));
    expect(text(view)).toContain("恢复快照不可读");
    expect(host.originalWrites()).toHaveLength(0);
    expect(view.queryByRole("button", { name: "确认全部变更" })).toBeNull();
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
    expect(view.container.querySelector(".composer-box .result-review-composer")).not.toBeNull();
    expect(view.container.querySelectorAll(".result-review-bar")).toHaveLength(1);
    expect(view.container.querySelector(".composer-box textarea")).toBeNull();
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
    expect(xml).toContain("票房");
    expect(xml).toContain("执行前旧值");
    expect(xml).toContain("不应复制的区域外数据"); // Same authorized sheet; not an unauthorized source sheet.
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
    expect(composerButton).toBeNull();
    expect(view.container.querySelector(".composer-box .result-review-composer")).not.toBeNull();
    expect(view.container.querySelector('[aria-label="暂停任务"]')).toBeNull();
    expect(text(view)).not.toContain("等待审阅结果");
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
    expect(text(view)).toContain("待审阅");
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
    fireEvent.click(view.getByRole("button", { name: "放弃全部变更" }));
    await waitFor(() =>
      expect(view.container.querySelector<HTMLButtonElement>(".new-progress-button")?.dataset.mode).toBe("running"));
    expect(view.container.querySelector<HTMLButtonElement>(".new-progress-button")!
      .querySelectorAll(".progress-dots > span")).toHaveLength(3);
    // 同一时刻审阅通道确实在途：按钮被禁用、`aria-busy` 为真。
    expect(view.container.querySelector<HTMLElement>(".result-review-bar")!.getAttribute("aria-busy")).toBe("true");
    expect((view.getByRole("button", { name: "放弃全部变更" }) as HTMLButtonElement).disabled).toBe(true);

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
    expect(view.container.querySelector(".composer-box textarea")).not.toBeNull();
    expect(view.container.querySelector(".result-review-composer")).toBeNull();
    expect(contentWrites(host.originalWrites())).toEqual(writes);
    expect(host.cell(SHEET, "G1")).toBe("合计");
    expect(text(view)).toContain("并通过独立验证");
    expect(view.container.querySelector("button[title='保存为可重复使用的个人工具']")).not.toBeNull();
  });
  it("放弃调用真正撤销，恢复目标且保留区域外内容", async () => {
    const view = await sendPlan(reviewPlan); await stageReview(view);
    expect(host.cell(SHEET, "G1")).toBe("合计");
    expect(host.conditionalFormatCount(SHEET)).toBeGreaterThan(0);
    fireEvent.click(view.getByRole("button", { name: "放弃全部变更" }));
    await waitFor(() => expect(text(view)).toContain("已放弃并恢复"));
    expect(host.conditionalFormatCount(SHEET)).toBe(0);
    expect(host.cell(SHEET, "G1")).toBe("");
    expect(host.cell(SHEET, "E1")).toBe("");
    expect(host.cell(SHEET, "B2")).toBe(1200);
    expect(host.windowNames()).toEqual(["book.xlsx"]);
  });
  it("放弃全部变更 → 服务端留痕记「放弃」，不是「撤销」", async () => {
    const view = await sendPlan(reviewPlan); await stageReview(view);
    fireEvent.click(view.getByRole("button", { name: "放弃全部变更" }));
    await waitFor(() => expect(text(view)).toContain("已放弃并恢复"));
    /*
     * 留痕是 best-effort（`void` 掉的），所以等的是"端点被调用过"，不是某条消息。
     * 它证的正是账本里补出来的那一行：用户**主动撤回**了这次执行。
     * 这一格以前完全不存在 —— 2026-10-29 真机确认全部变更后，账本里连"用户点过确认"
     * 都查不到（见 `outputs/a1-acceptance-20261029-2/EVIDENCE.md` §6）。
     */
    await waitFor(() => expect(vi.mocked(recordExecutionDisposition)).toHaveBeenCalledTimes(1));
    const [namespace, executionId, disposition] = vi.mocked(recordExecutionDisposition).mock.calls[0];
    expect(namespace).toBe("office");
    expect(typeof executionId).toBe("string");
    expect((executionId as string).length).toBeGreaterThan(0);
    expect(disposition).toBe("abandoned"); // 不是 "undone"：两种取值不许互相冒充
    // 留痕不改行为：单元格确实回到执行前，且只留了一次。
    expect(host.cell(SHEET, "G1")).toBe("");
    expect(host.cell(SHEET, "B2")).toBe(1200);
  });
  it("留痕上报失败：放弃照样成功、不重试、不再撤一次", async () => {
    /*
     * 留痕是**观测**，不是恢复的前置条件。这里钉三件事：
     * 1. 上报抛错不影响业务结论（单元格确实回到执行前、文案仍是「已放弃并恢复」）；
     * 2. **不自动重试**（端点只被叫一次）——重试若发生在错误的位置，就会变成"再撤一次"；
     * 3. 上报前后宿主写入次数不再增加（证明没有第二次恢复）。
     */
    vi.mocked(recordExecutionDisposition).mockRejectedValueOnce(new Error("网络不可用"));
    const view = await sendPlan(reviewPlan); await stageReview(view);
    fireEvent.click(view.getByRole("button", { name: "放弃全部变更" }));
    await waitFor(() => expect(text(view)).toContain("已放弃并恢复"));
    await waitFor(() => expect(vi.mocked(recordExecutionDisposition)).toHaveBeenCalledTimes(1));
    const writes = host.originalWrites().length;
    await waitFor(() => expect(text(view)).toContain("已放弃并恢复"));
    expect(host.originalWrites().length).toBe(writes);
    expect(host.cell(SHEET, "G1")).toBe("");
    expect(text(view)).not.toContain("无法记录这次执行的处置");
  });
  it("确认全部变更不是处置：不写「放弃/撤销」留痕", async () => {
    const view = await sendPlan(reviewPlan); await stageReview(view);
    fireEvent.click(view.getByRole("button", { name: "确认全部变更" }));
    await waitFor(() => expect(text(view)).toContain("已确认变更"));
    expect(vi.mocked(recordExecutionDisposition)).not.toHaveBeenCalled();
    expect(host.cell(SHEET, "G1")).toBe("合计");
  });
  it("执行后目标被人工修改，确认与放弃都不覆盖该修改", async () => {
    const view = await sendPlan(reviewPlan); await stageReview(view);
    host.setCell(SHEET, "G1", "人工修改");
    const count = contentWrites(host.originalWrites()).length;
    fireEvent.click(view.getByRole("button", { name: "确认全部变更" }));
    await waitFor(() => expect(text(view)).toContain("又发生了变化"));
    fireEvent.click(view.getByRole("button", { name: "放弃全部变更" }));
    await waitFor(() => expect(text(view)).toContain("未能放弃"));
    expect(host.cell(SHEET, "G1")).toBe("人工修改");
    expect(contentWrites(host.originalWrites())).toHaveLength(count);
    expect(view.container.querySelector(".result-review-bar")?.textContent).not.toContain("已确认变更");
    /*
     * ★ 恢复没发生 ⇒ 两种留痕一个都不写（`abandoned` 也不行）。
     * 这条同时钉住 `revertedByDiscard` 的清理：失败分支必须把"这次是放弃"的标记撤掉，
     * 否则同一次执行之后的**真正撤销**会被贴成 `abandoned`。
     */
    expect(vi.mocked(recordExecutionDisposition)).not.toHaveBeenCalled();
  });
  it("执行后切换工作簿，旧确认不产生任何额外写入", async () => {
    const view = await sendPlan(reviewPlan); await stageReview(view);
    const count = contentWrites(host.originalWrites()).length;
    host.setDocumentUrl("https://contoso.sharepoint.com/other.xlsx");
    fireEvent.click(view.getByRole("button", { name: "确认全部变更" }));
    await waitFor(() => expect(text(view)).toContain("工作簿、授权或计划已变化"));
    expect(contentWrites(host.originalWrites())).toHaveLength(count);
  });
  it("回读不符不开放确认或保存工具，不自动重放", async () => {
    host.tamperNextWrite(event => event.address === "E1:F3" && event.store === "original",
      matrix => [matrix[0], [matrix[1][0], 0.99], matrix[2]]);
    const view = await sendPlan(reviewPlan);
    await waitFor(() => expect(text(view)).toContain("本次执行需要核对"));
    expect(host.cell(SHEET, "F2")).toBe(0.99);
    expect(view.queryByRole("button", { name: "确认全部变更" })).toBeNull();
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
    expect(view.queryByRole("button", { name: "确认全部变更" })).toBeNull();
  });

});

/**
 * 旧入口已删除（2026-10-30）。
 *
 * 这里曾经有个 `clickRunOnExcel`：它在"副本比对做不了"的计划上找「执行到 Excel」
 * 并点下去。真机反馈很直接 —— 用户已经确认授权，界面却还要他再点一下，而那一下
 * 是**绕过审阅**的执行。现在计划到达即自动执行一次：
 *
 * - 能生成副本 ⇒ 审阅带双视图对比；
 * - 不能生成副本 ⇒ 审阅降级（条子上说明原因，不给"对比"按钮），执行照旧；
 * - 越界（目标/来源不在授权内）⇒ **不执行**，卡片直接给原因。
 *
 * 所以下面的用例都不再点击：判据变成"自动执行有没有被正确拦住/放行"。
 */
describe("Excel 写侧：非副本计划同样自动执行，写前判据不变", () => {
  it("计划含未授权工作表的动作：不自动执行，卡片给出原因，一个单元格都不写", async () => {
    const view = await sendPlan(offSourcePlan);

    await waitFor(() => expect(text(view)).toContain("计划写入目标超出本次授权，未执行。"));
    // 没有手动入口可以绕过授权，也没有任何写入。
    expect(view.queryByRole("button", { name: "执行到 Excel" })).toBeNull();
    expect(host.writes()).toEqual([]);
    expect(host.cell(SHEET, "C2")).toBe(30);
    expect(registerOfficeExecution).not.toHaveBeenCalled();
  });

  it("本轮读过源、计划到达前源值又变了：给出二次确认，不自动写入", async () => {
    /*
     * 顺序就是真机顺序：本轮经工具桥读了一次源（锚点因此在场）→ 用户/别人改了源
     * → 计划才送达 → 自动执行被来源复核拦下。
     *
     * 旧写法是"送达之后再改、再点按钮"：自动执行上线后那段窗口是竞态，
     * 而它真正想验的只是"锚点成立之后源变化必须拦住"。
     */
    const view = await sendPlan(directPlan, "把合计写到汇总表", {
      read: { scope: "active" },
      beforePlan: () => host.setCell(SHEET, "B2", 4321)
    });

    await waitFor(() =>
      expect(view.container.querySelector(".stale-source-message")).not.toBeNull()
    );
    expect(host.originalWrites()).toEqual([]);
    expect(host.cell(SHEET, "C2")).toBe(30);
  });

  it("源表**公式文本**变化也会改变来源指纹（不是只算值）", async () => {
    const view = await sendPlan(directPlan, "把合计写到汇总表", {
      read: { scope: "active" },
      // 显示结果不变（30 = 30），变的是单元格里的公式文本。
      beforePlan: () => host.setCell(SHEET, "C2", "=30+0", { asFormula: true })
    });

    await waitFor(() =>
      expect(view.container.querySelector(".stale-source-message")).not.toBeNull()
    );
    expect(host.originalWrites()).toEqual([]);
  });

  it("源值、身份都没变时照常自动写入，并如实给出恢复限制说明", async () => {
    const view = await sendPlan(directPlan);

    await waitFor(() => expect(host.cell(SHEET, "C2")).toBe(30));
    expect(host.matrix(SHEET, "C1:C3")).toEqual([["场次"], [30], [22]]);
    // 没有可信来源记账 → 不登记断点，如实说明"本次不承诺跨重启自动续跑"。
    await waitFor(() => expect(text(view)).toContain("本次不承诺跨重启自动续跑"));
    expect(registerOfficeExecution).not.toHaveBeenCalled();
  });
});

/**
 * 2026-10-30 真机阻断条件的回归。
 *
 * 现场：用户已经确认授权，三步公式计划（三张结果表各写一条 `=来源!B*`）却只摆出
 * 「执行到 Excel」—— 因为那时"要不要自动执行"由 `resultReviewPlanIssue` 决定，
 * 而"多张写入目标表"与"跨表公式"都落在"副本比对做不了"这一档上。
 *
 * 这条用例钉三件事：
 *  1. 计划到达即自动执行，**没有**手动入口（不再需要用户点一下绕过审阅）；
 *  2. 确定性验收入口断在第二步回读之后，第三步一个单元格都不写；
 *  3. 原授权、来源复核与恢复登记都还在：登记一次、授权集合原样、公式动作原样。
 */
describe("Excel 写侧：三步公式计划自动执行（阻断条件回归）", () => {
  const FORMULA_TARGETS = ["结果一", "结果二", "结果三"];
  const FORMULA_SCOPE = [SHEET, ...FORMULA_TARGETS];
  /** 来源表 + 三张空的公式结果表（`createHost()` 的来源值原样保留，工具桥按「影片」读）。 */
  function installFormulaHost() {
    host = createOfficeTestHost({
      url: BOOK_URL,
      selectedRange: "A1:C3",
      activeSheet: SHEET,
      sheets: [
        {
          name: SHEET,
          values: [["影片", "票房(万)", "场次"], ["甲", 1200, 30], ["乙", 980, 22]]
        },
        ...FORMULA_TARGETS.map((name) => ({ name, values: [] }))
      ]
    });
    installOfficeTestHost(host);
  }
  /** 三张结果表各一条**跨表**公式：任何一条都做不了副本比对（跨表 + 多张目标表）。 */
  function threeFormulaPlan(id: string) {
    return {
      id,
      title: "三条公式",
      summary: "三张结果表各写一条公式",
      assumptions: [],
      warnings: [],
      actions: FORMULA_TARGETS.map((name, index) => ({
        type: "writeFormulas" as const,
        sheet: name,
        range: "B2",
        formulas: [[`=${SHEET}!${["B2", "C2", "B3"][index]}*${index + 1}`]]
      })),
      sourceDependencySheets: [SHEET]
    };
  }

  it.each([[true, true], [false, true], [false, false]])("三张结果表的公式计划自动执行并在第二步中断（模型读取=%s，同一执行=%s）", async (modelRead, sameExecution) => {
    installFormulaHost();
    const plan = threeFormulaPlan("plan-formula-three");
    /*
     * 中断后的现场：有在途动作 ⇒ 恢复卡按「待核对」摆出来（不是一张空白失败卡）。
     *
     * 这里必须给**完整**的账本读数（`plan` / `authorizedSheetNames` / `workbookIdentity`
     * / 写前快照），因为卡片正文里的"第几步 / 还剩几步"来自 `plan.actions.length`，
     * 够不够格摆成"可继续"来自共享判据 `pendingResumeReadiness`（身份 → 状态 →
     * 授权 → 在途材料）。只给 `status` 与两个索引，卡片构造会在读 `plan.actions` 时抛错，
     * 表现为"恢复入口根本没出现"——那是替身缺料，不是产品少闸门。
     */
    vi.mocked(suspendOfficeExecution).mockImplementation(async operationId => ({
      operationId: sameExecution ? operationId : "op-formula-three",
      status: "suspended",
      revision: 2,
      workbookIdentity: `url:${BOOK_URL}`,
      plan,
      authorizedSheetNames: FORMULA_SCOPE,
      allowedExternalFiles: [],
      nextActionIndex: 1,
      inFlightActionIndex: 1,
      beforeSnapshot: {
        actionIndex: 1,
        sheet: "结果二",
        range: "B2",
        formulas: [[""]],
        numberFormat: [["General"]],
        fillColor: "#FFFFFF",
        fontBold: false,
        fontColor: "#000000"
      },
      checkpoint: null,
      result: null,
      error: null,
      canResume: true,
      requiresReconciliation: true,
      ownerToken: "owner-token"
    } as never));
    const view = await sendPlan(plan, "把三条公式分别写到三张结果表", {
      authorize: FORMULA_SCOPE,
      authorizePool: FORMULA_SCOPE,
      ...(modelRead ? { read: { scope: "active" as const } } : {}),
      probe: true
    });

    // ① 没有手动入口：计划到达即执行。
    expect(view.queryByRole("button", { name: "执行到 Excel" })).toBeNull();
    /*
     * 而且**真的进了执行**——只断言"没有按钮"是不够的：旧缺陷下按钮不出现
     * （工作簿模式下它本来就不画），计划却停在"等人点"，一样没有按钮。
     * 判别点是"登记是否发生"：计划没被送进执行器就永远不会有这一步。
     */
    await waitFor(() => expect(registerOfficeExecution).toHaveBeenCalledTimes(1));

    // ② 第二步回读后中断：前两步真的写进去了，第三步没跑。
    await waitFor(() => expect(suspendOfficeExecution).toHaveBeenCalledTimes(1));
    expect(host.cell("结果一", "B2")).toBe(`=${SHEET}!B2*1`);
    expect(host.cell("结果二", "B2")).toBe(`=${SHEET}!C2*2`);
    expect(host.cell("结果三", "B2")).toBe("");
    expect(contentWrites(host.originalWrites()).map((write) => `${write.sheet}!${write.address}`))
      .toEqual(["结果一!B2", "结果二!B2"]);
    // 第 2 步的完成回执没有提交（中断就在回读之后、回执之前）。
    expect(vi.mocked(completeOfficeAction).mock.calls.map((call) => call[1].actionIndex)).toEqual([0]);

    // ③ 授权、来源复核与恢复登记都还在。
    expect(registerOfficeExecution).toHaveBeenCalledTimes(1);
    const registered = vi.mocked(registerOfficeExecution).mock.calls[0][0] as {
      plan: { actions: { type: string; sheet: string }[]; sourceDependencySheets?: string[] };
      authorizedSheetNames: string[];
    };
    expect(registered.authorizedSheetNames).toEqual(FORMULA_SCOPE);
    expect(registered.plan.actions.map((action) => action.type))
      .toEqual(["writeFormulas", "writeFormulas", "writeFormulas"]);
    expect(registered.plan.sourceDependencySheets).toEqual([SHEET]);
    if (!modelRead) expect(postSdkBridgeReceipt).not.toHaveBeenCalled();
    // 来源表一个单元格都没动。
    expect(host.matrix(SHEET, "A1:C2")).toEqual([["影片", "票房(万)", "场次"], ["甲", 1200, 30]]);
    // 恢复入口按结论给出（有在途动作 ⇒ 待核对），可点、且只有一条。
    expect(await view.findByRole("button", { name: RECOVERY_CHECK_BUTTON })).toBeTruthy();
    expect(view.container.querySelector(".recovery-composer")).not.toBeNull();
    if (sameExecution) expect(view.container.querySelector(".result-review-bar")).toBeNull();
    else expect(view.container.querySelector(".result-review-bar")).not.toBeNull();
  });

  it("探针取消显示具体同表依赖原因和未执行，不显示需要核对", async () => {
    installFormulaHost();
    const plan = threeFormulaPlan('formula-local-probe');
    plan.actions[0].formulas = [['=A1*2']];
    const view = await sendPlan(plan, '写三条公式', {
      authorize: FORMULA_SCOPE, authorizePool: FORMULA_SCOPE, probe: true, read: { scope: 'active' }
    });
    await waitFor(() => expect(view.container.querySelector('.result-review-bar')?.textContent).toContain('未执行'));
    expect(view.container.querySelector('.result-review-bar')?.textContent).toContain('写入目标表自身');
    expect(view.container.querySelector('.result-review-bar')?.textContent).not.toContain('需要核对');
    expect(registerOfficeExecution).not.toHaveBeenCalled();
    expect(contentWrites(host.originalWrites())).toEqual([]);
  });

  /*
   * 修复的另一半：副本比对做不了时**降级**，不是"不许执行"。
   *
   * 旧代码把 `resultReviewPlanIssue` 当执行许可，于是这类计划连执行都进不去；
   * 现在它照常执行一次，审阅条如实说"本次没有副本"，**不画**点下去只会报错的
   * 「双视图对比」，而「确认全部变更 / 放弃全部变更」照旧 —— 比对是审阅的辅助，不是许可。
   *
   * 这里用**值**计划而不是公式计划：公式计划的 `verified` 依赖服务端补
   * `formulasEqual`（替身服务端没有），会先卡在"结果未完整验证"上，测不到降级这段。
   */
  it("多张目标表按各自执行前快照切换对比，确认不重写", async () => {
    BOOK_URL = 'file:///c:/work/multi-review.xlsx';
    vi.mocked(registerOfficeExecution).mockResolvedValue(officeLedgerState() as never);
    vi.mocked(startOfficeAction).mockResolvedValue(officeLedgerState() as never);
    vi.mocked(completeOfficeAction).mockResolvedValue(officeLedgerState() as never);
    vi.mocked(finishOfficeExecution).mockResolvedValue(officeLedgerState({status:'finished'}) as never);

    vi.mocked(closeExcelReviewCopy).mockResolvedValue({status:'closed',copyClosed:true,recoveryCopyPath:null} as never);
    installFormulaHost();
    FORMULA_TARGETS.forEach((name,index)=>host.setCell(name,'A1',`执行前${index+1}`));
    const rows = [["甲", 1200], ["乙", 980], ["甲", 30]];
    const plan = {
      id: "plan-three-values",
      title: "三条结果",
      summary: "三张结果表各写一行",
      assumptions: [],
      warnings: [],
      actions: FORMULA_TARGETS.map((name, index) => ({
        type: "writeValues" as const,
        sheet: name,
        range: "A1:B1",
        values: [rows[index]]
      })),
      acceptanceCriteria: FORMULA_TARGETS.map((name, index) => ({
        type: "rangeEquals" as const,
        sheet: name,
        range: "A1:B1",
        expected: [rows[index]]
      }))
    };
    const view = await sendPlan(plan, "把三行分别写到三张结果表", {
      authorize: FORMULA_SCOPE,
      authorizePool: FORMULA_SCOPE,
      read: { scope: "active" }
    });

    expect(view.queryByRole("button", { name: "执行到 Excel" })).toBeNull();
    /*
     * 判别点是"计划是否真的被送进执行器"，而不是"有没有按钮"：旧缺陷下
     * 工作簿模式本来就不画按钮，计划却停在"等人点"。没有这一步，下面"等写入"
     * 只能以超时收场，红得说不清原因。
     */
    await waitFor(() => expect(registerOfficeExecution).toHaveBeenCalledTimes(1));
    // 三张表都写进去了：没有副本也没收走自动执行。
    await waitFor(() => expect(host.cell("结果三", "A1")).toBe("甲"), { timeout: 5000 });
    expect(FORMULA_TARGETS.map((name) => host.matrix(name, "A1:B1")))
      .toEqual([[["甲", 1200]], [["乙", 980]], [["甲", 30]]]);

    await view.findByRole("region", { name: "本次结果审阅" }, { timeout: 5000 });
    const written=host.originalWrites().length;
    expect(view.queryByRole('button', {name:'结果一'})).toBeNull();
    fireEvent.click(view.getByRole('button', {name:'双视图对比'}));
    await waitFor(()=>expect(host.createWorkbookCalls()).toHaveLength(1));
    const picker = await view.findByRole('button', {name:'结果一'});
    fireEvent.click(view.getByRole('button', {name:'结果二'}));
    await waitFor(()=>expect(selectExcelReviewCopySheet).toHaveBeenCalledTimes(1));
    expect(host.createWorkbookCalls()).toHaveLength(1);
    expect(host.seedEntries(0)['xl/worksheets/sheet1.xml']).toContain('执行前1');
    expect(host.seedEntries(0)['xl/worksheets/sheet2.xml']).toContain('执行前2');
    expect(host.seedEntries(0)['xl/worksheets/sheet3.xml']).toContain('执行前3');
    expect(host.seedEntries(0)['xl/workbook.xml']).toContain('<sheet name="结果一"');
    expect(host.seedEntries(0)['xl/workbook.xml']).toContain('<sheet name="结果二"');
    expect(host.seedEntries(0)['xl/workbook.xml']).toContain('<sheet name="结果三"');
    expect(host.seedEntries(0)['xl/workbook.xml']).not.toContain('<sheet name="EB_');
    expect(host.seedEntries(0)['xl/worksheets/sheet2.xml']).not.toContain('执行前1');
    expect(host.seedEntries(0)['xl/worksheets/sheet2.xml']).not.toContain('980');
    expect(host.originalWrites()).toHaveLength(written);
    await waitFor(()=>expect((view.getByRole('button', {name:'结果二'}) as HTMLButtonElement).disabled).toBe(false));
    fireEvent.click(view.getByRole('button', {name:'退出双视图'}));
    await view.findByRole('button', {name:'双视图对比'});
    fireEvent.click(view.getByRole('button', {name:'双视图对比'}));
    await waitFor(()=>expect((view.getByRole('button', {name:'退出双视图'}) as HTMLButtonElement).disabled).toBe(false));
    fireEvent.click(view.getByRole('button', {name:'结果一'}));
    await waitFor(()=>expect((view.getByRole('button', {name:'结果一'}) as HTMLButtonElement).disabled).toBe(false));
    expect(host.createWorkbookCalls()).toHaveLength(1);
    expect(host.originalWrites()).toHaveLength(written);
    fireEvent.click(view.getByRole('button',{name:'确认全部变更'}));
    await waitFor(()=>expect(view.container.querySelector('.result-review-composer')).toBeNull());
    expect(host.originalWrites()).toHaveLength(written);

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
     * 显式把整片授权区都列进 `sheets`，让本轮**真的**读过「汇总表」——
     * 于是来源依赖 = [综合票房, 汇总表] 与目标重叠。即便"把目标从来源里摘掉"
     * 就能换来登记成功，也不摘：记账说的是"读过谁"，它就是读过目标表。
     * （旧夹具靠 `scope: "selected"` 隐式读全部授权表；那条隐式路径已经取消，
     * 多张表必须显式绑定，所以这里把表名列出来。）
     */
    const view = await sendPlan(targetPlan, "把合计写到汇总表", {
      authorize: SCOPE,
      read: { scope: "selected", sheets: SCOPE }
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

  it("已完成前缀但恢复快照丢失：拒绝续写，不能生成仅含后缀的撤销点", async () => {
    await expect(executePlan(liveTargetPlan, {
      authorizedSheetNames:SCOPE, allowedExternalFiles:new Set<string>(), operationId:"office-material-test",
      recovery:{startActionIndex:1,undoSnapshot:null,
        onBeforeAction:async()=>undefined,onActionCompleted:async()=>undefined}
    })).rejects.toThrow("缺少已完成动作的恢复快照");
    expect(host.writes()).toEqual([]);
  });

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
    view.unmount();
    const reopened = render(<App />);
    await waitFor(() => expect(reopened.container.querySelector(".composer textarea")).not.toBeNull());
    expect(getOfficeExecutionStatus).toHaveBeenCalledTimes(1);
    expect(host.writes()).toEqual([]);
    for (const mutation of [registerOfficeExecution, startOfficeAction, completeOfficeAction,
      reconcileOfficeActionReceipt, finishOfficeExecution, suspendOfficeExecution]) {
      expect(mutation).not.toHaveBeenCalled();
    }
    expect(loadPendingOfficeExecution()).toBeNull();
  });

  it.each(["applied", "not_applied", "unknown"] as const)("A1 三动作中途断开：核对只补记回执（零写入），继续才只跑剩余动作（%s）", async (outcome) => {
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
    const prior=await reconcileOfficeRangeAction({plan,operationId:OPERATION_ID,
      workbookIdentity:resumeState().workbookIdentity,actionIndex:0,
      before:resumeState().beforeSnapshot!,previous:null});
    expect(prior.outcome).toBe("applied");
    /*
     * 替身服务端按真实时序推进：核对回执提交之后，在途动作才被清掉、进度才前移。
     * 用 `reconciled` 开关而不是"第几次调用"，是为了让"核对前/核对后"两个现场
     * 在任意次数的只读查询下都稳定。
     */
    let reconciled = false;
    let reconciledCheckpoint: unknown = null;
    const inFlightState = () => resumeState({ plan, nextActionIndex: 1, inFlightActionIndex: 1,
      checkpoint: prior.checkpoint,
      beforeSnapshot: { ...resumeState().beforeSnapshot, actionIndex: 1, range: "A2:B2", formulas: [["", ""]], fillColor: "" } });
    const settledState = () => resumeState({ plan,
      checkpoint: reconciledCheckpoint ?? prior.checkpoint,
      nextActionIndex: outcome === "applied" ? 2 : 1, inFlightActionIndex: null,
      requiresReconciliation: false, revision: 3 });
    vi.mocked(getOfficeExecutionStatus).mockImplementation(async () => (reconciled ? settledState() : inFlightState()) as never);
    vi.mocked(reconcileOfficeActionReceipt).mockImplementation(async (_id, request) => {
      expect(request.actionIndex).toBe(1);
      expect(request.outcome).toBe(outcome);
      reconciled = true;
      // 真实服务端会把核对累积的 checkpoint（含第 1、2 步的写后指纹）存进账本并返回；
      // 替身也必须传播它，否则"已核对"之后的账本缺了第 1 步的写后指纹。
      reconciledCheckpoint = request.checkpoint;
      return settledState() as never;
    });
    vi.mocked(claimOfficeExecutionResume).mockImplementation(async () => settledState() as never);
    vi.mocked(finishOfficeExecution).mockResolvedValue(
      resumeState({ plan, nextActionIndex: 3, inFlightActionIndex: null, revision: 4 }) as never
    );
    const view = render(<App />);
    const reconcileButton = await view.findByRole("button", { name: RECOVERY_CHECK_BUTTON });
    // 待核对时**不许**同时给续跑按钮：那正是"一个按钮里先补记再续跑"的老行为。
    expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    /*
     * 信息层级（2026-11-02）：输入框里那块操作区一句话说完"写了哪些 / 还没写哪些 /
     * 点下去不会写表"，上头不再另摆一张说同一件事的大卡。
     */
    const card = text(view);
    expect(view.container.querySelector(".office-recovery-panel"), "接管时不许再有大卡").toBeNull();
    expect(card).toContain("任务中断 · 第 2 步结果待核对");
    expect(card).toContain("已写入：");
    expect(card).toContain("是否已写入还没确认");
    expect(card).toContain("尚未写入：");
    expect(card).toContain("不会修改表格");
    // 状态未知时一个字都不许往"已写入"上靠：进在途那格不进 `已写入` 清单。
    expect(card).not.toContain("已写入：汇总表!A1:B1、汇总表!A2:B2");
    fireEvent.click(reconcileButton);
    if (outcome === "unknown") {
      await waitFor(() => expect(text(view)).toContain("目标范围既不匹配执行前状态"));
      expect(reconcileOfficeActionReceipt).not.toHaveBeenCalled();
      expect(startOfficeAction).not.toHaveBeenCalled();
      expect(contentWrites(host.originalWrites()).length).toBe(beforeWrites);
      return;
    }
    /*
     * ①核对：只有回执，没有单元格写入、没有动作执行权、没有收尾。
     * 这一条是"核对只补记"的判据 —— 以前"核对并继续"会在这一下把剩下的也写掉。
     */
    await waitFor(() => expect(reconcileOfficeActionReceipt).toHaveBeenCalledTimes(1));
    // 逐格断言"这一下写了哪几格"而不是只比次数：失败时原文要点出被多写的那一格。
    expect(contentWrites(host.originalWrites()).slice(beforeWrites).map((write) => write.address))
      .toEqual([]);
    expect(host.matrix(TARGET, "A1:B1")).toEqual(originalFirstRow);
    expect(startOfficeAction).not.toHaveBeenCalled();
    expect(finishOfficeExecution).not.toHaveBeenCalled();
    // ②继续：只剩没跑过的动作，已完成的第 2 步不重放。
    fireEvent.click(await view.findByRole("button", { name: RECOVERY_RESUME_BUTTON }));
    await waitFor(() => expect(finishOfficeExecution).toHaveBeenCalledTimes(1));
    expect(contentWrites(host.originalWrites()).slice(beforeWrites).map((write) => write.address))
      .toEqual(outcome === "applied" ? ["A3:B3"] : ["A2:B2", "A3:B3"]);
    expect(host.matrix(TARGET, "A1:B1")).toEqual(originalFirstRow);
    expect(host.cell(TARGET, "A3")).toBe('="丙"');
    expect(vi.mocked(startOfficeAction).mock.calls.map((call) => call[1].actionIndex))
      .toEqual(outcome === "applied" ? [2] : [1, 2]);
  });


  it("连点「检查」的第二下不会派发第二个请求，也不会跨到续跑把表写了", async () => {
    /*
     * 现场风险：核对与续跑住在**同一个位置**（输入框里的操作区）。双击时浏览器会派发
     * `detail=1` 与 `detail=2` 两下；若两下都派发，第二下可能落在"视图已变成
     * 继续剩余 N 步"的那一帧上 —— 用户以为自己在检查，结果把表写了。
     *
     * 判据两层：点击层的 `detail > 1` 不派发 + 核对→继续**换节点**（`key=view.action`）。
     * 这里断言端到端的可观察结果：只提交一次回执、一格都没写、没有任何执行权请求。
     */
    // 在途的就是第 2 步（A2:B2），现场已经是**写后**状态：回执丢了，但值在。
    host.setCell(TARGET, "A1", '="影片"', { asFormula: true });
    host.setCell(TARGET, "B1", "=1200", { asFormula: true });
    host.setCell(TARGET, "A2", '="乙"', { asFormula: true });
    host.setCell(TARGET, "B2", "=980", { asFormula: true });
    const beforeWrites = contentWrites(host.originalWrites()).length;
    const originalFirstRow = host.matrix(TARGET, "A1:B1");
    const plan = { ...liveTargetPlan, sourceFingerprintSheets: SCOPE, sourceDependencySheets: [SHEET],
      actions: [...liveTargetPlan.actions, { type: "writeFormulas" as const, sheet: TARGET, range: "A3:B3", formulas: [['="丙"', "=600"]] }] };
    pendingIndex();
    const prior = await reconcileOfficeRangeAction({ plan, operationId: OPERATION_ID,
      workbookIdentity: resumeState().workbookIdentity, actionIndex: 0,
      before: resumeState().beforeSnapshot!, previous: null });
    let reconciled = false;
    let reconciledCheckpoint: unknown = null;
    const inFlightState = () => resumeState({ plan, nextActionIndex: 1, inFlightActionIndex: 1,
      checkpoint: prior.checkpoint,
      beforeSnapshot: { ...resumeState().beforeSnapshot, actionIndex: 1, range: "A2:B2" } });
    const settledState = () => resumeState({ plan,
      checkpoint: reconciledCheckpoint ?? prior.checkpoint,
      nextActionIndex: 2, inFlightActionIndex: null, requiresReconciliation: false, revision: 3 });
    vi.mocked(getOfficeExecutionStatus).mockImplementation(async () => (reconciled ? settledState() : inFlightState()) as never);
    vi.mocked(reconcileOfficeActionReceipt).mockImplementation(async (_id, request) => {
      reconciled = true;
      // 真实服务端把核对累积的 checkpoint（含两步写后指纹）落库并返回，替身也必须传播它。
      reconciledCheckpoint = request.checkpoint;
      return settledState() as never;
    });
    vi.mocked(claimOfficeExecutionResume).mockImplementation(async () => settledState() as never);
    vi.mocked(finishOfficeExecution).mockResolvedValue(
      resumeState({ plan, nextActionIndex: 3, inFlightActionIndex: null, revision: 4 }) as never
    );
    const view = render(<App />);
    const reconcileButton = await view.findByRole("button", { name: RECOVERY_CHECK_BUTTON });

    // 双击：第一下 detail=1，第二下 detail=2（浏览器对同一次双击的真实派发形状）。
    fireEvent.click(reconcileButton, { detail: 1 });
    fireEvent.click(reconcileButton, { detail: 2 });

    await waitFor(() => expect(reconcileOfficeActionReceipt).toHaveBeenCalledTimes(1));
    // 给第二下"落空"之后可能出现的异步续跑留出机会，再断言它确实没发生。
    await act(async () => { await Promise.resolve(); });
    expect(reconcileOfficeActionReceipt).toHaveBeenCalledTimes(1);
    expect(contentWrites(host.originalWrites()).slice(beforeWrites)).toEqual([]);
    expect(host.matrix(TARGET, "A1:B1")).toEqual(originalFirstRow);
    // 只补记、不写表：回执说的就是"现场与写后一致"，第 3 步 A3:B3 一格都没动。
    expect((vi.mocked(reconcileOfficeActionReceipt).mock.calls[0][1] as { outcome: string }).outcome)
      .toBe("applied");
    expect(claimOfficeExecutionResume).not.toHaveBeenCalled();
    expect(startOfficeAction).not.toHaveBeenCalled();
    expect(finishOfficeExecution).not.toHaveBeenCalled();
    /*
     * 结算之后停在"继续剩余 N 步"，**没有**自动续跑：按钮出现是等来的（异步状态推进），
     * 出现之后立刻再断一次"一格都没写"。
     */
    const resumeButton = await view.findByRole("button", { name: RECOVERY_RESUME_BUTTON });
    expect(resumeButton).toBeTruthy();
    // 只剩第 3 步可跑：文案里的步数来自账本游标，不按动作总数硬算。
    expect(resumeButton.textContent ?? "").toContain("1 步");
    expect(contentWrites(host.originalWrites()).slice(beforeWrites)).toEqual([]);
    expect(claimOfficeExecutionResume).not.toHaveBeenCalled();
  });

  it("在途动作先核对为「已写入」：跳过该步，只写剩余步骤，不重复写", async () => {
    // 头一步已经真的写进宿主了（回执丢失），第二部还没写。
    host.setCell(TARGET, "A1", "=\"影片\"", { asFormula: true });
    host.setCell(TARGET, "B1", "=1200", { asFormula: true });
    const firstWrites = host.originalWrites().length;

    pendingIndex();
    // 替身服务端在核对回执提交之后才清掉在途动作、把进度推到第 2 步，并记下新的检查点。
    let reconciled = false;
    let checkpoint: unknown = null;
    const settledState = (revision: number) => resumeState({
      checkpoint: checkpoint as never,
      nextActionIndex: 1, inFlightActionIndex: null, requiresReconciliation: false, revision
    });
    vi.mocked(getOfficeExecutionStatus).mockImplementation(async () =>
      (reconciled ? settledState(3) : resumeState()) as never
    );
    vi.mocked(reconcileOfficeActionReceipt).mockImplementation(async (_id,request)=>{
      reconciled = true;
      checkpoint = request.checkpoint;
      return settledState(3) as never;
    });
    vi.mocked(claimOfficeExecutionResume).mockImplementation(async () => settledState(4) as never);
    vi.mocked(finishOfficeExecution).mockResolvedValue(
      resumeState({ nextActionIndex: 2, inFlightActionIndex: null, revision: 5 }) as never
    );

    const view = render(<App />);
    // 挂接索引在挂载时先核对一次：在途动作 → 恢复卡只给「核对在途动作」。
    const reconcileButton = await view.findByRole("button", { name: RECOVERY_CHECK_BUTTON });
    expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    fireEvent.click(reconcileButton);

    // ①核对只补记：一个单元格都没写，也没有领取任何动作的执行权。
    await waitFor(() => expect(reconcileOfficeActionReceipt).toHaveBeenCalledTimes(1));
    expect(contentWrites(host.originalWrites()).slice(firstWrites)).toEqual([]);
    expect(startOfficeAction).not.toHaveBeenCalled();

    // ②继续：只写第二步（A2:B2），第一步的 A1:B1 不再被写一次。
    fireEvent.click(await view.findByRole("button", { name: RECOVERY_RESUME_BUTTON }));
    await waitFor(() => expect(host.cell(TARGET, "A2")).toBe("=\"乙\""));
    const writes = contentWrites(host.originalWrites()).slice(firstWrites);
    expect(writes.map((write) => write.address)).toEqual(["A2:B2"]);
    expect(reconcileOfficeActionReceipt).toHaveBeenCalledTimes(1);
    expect(
      (vi.mocked(reconcileOfficeActionReceipt).mock.calls[0][1] as { outcome: string }).outcome
    ).toBe("applied");
  });

  it("全部动作已写入、只缺最终回执：只给「补记完成（不重写原表）」，点它不写单元格也不执行动作", async () => {
    // 现场就是"两个动作都已经写进原表"：进度 2/2，没有在途动作，只差最终回执。
    host.setCell(TARGET, "A1", '="影片"', { asFormula: true });
    host.setCell(TARGET, "B1", "=1200", { asFormula: true });
    host.setCell(TARGET, "A2", '="乙"', { asFormula: true });
    host.setCell(TARGET, "B2", "=980", { asFormula: true });
    const firstWrites = contentWrites(host.originalWrites()).length;
    pendingIndex();
    const ledgerOnly = (revision: number) => resumeState({
      nextActionIndex: 2, inFlightActionIndex: null, requiresReconciliation: false, revision
    });
    vi.mocked(getOfficeExecutionStatus).mockImplementation(async () => ledgerOnly(2) as never);
    vi.mocked(claimOfficeExecutionResume).mockImplementation(async () => ledgerOnly(3) as never);
    vi.mocked(finishOfficeExecution).mockResolvedValue(
      resumeState({ status: "completed", nextActionIndex: 2, inFlightActionIndex: null, revision: 4 }) as never
    );

    const view = render(<App />);
    const backfill = await view.findByRole("button", { name: "补记完成（不重写原表）" });
    // 只剩回执时不画「继续剩余动作」：没有动作可执行，画了就是假按钮。
    expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    fireEvent.click(backfill);

    await waitFor(() => expect(finishOfficeExecution).toHaveBeenCalledTimes(1));
    expect((vi.mocked(finishOfficeExecution).mock.calls[0][1] as { result: unknown }).result)
      .toEqual({ backfilled: true, actionCount: 2 });
    /*
     * 补记只改账本：一个单元格都没写，也没有领取任何动作的执行权。
     * 以前这条路会先 `executePlan(startActionIndex = 动作总数)` 再收尾 —— 一个动作都不该跑。
     */
    expect(contentWrites(host.originalWrites()).length).toBe(firstWrites);
    expect(startOfficeAction).not.toHaveBeenCalled();
    expect(completeOfficeAction).not.toHaveBeenCalled();
    expect(reconcileOfficeActionReceipt).not.toHaveBeenCalled();
    expect(loadPendingOfficeExecution()).toBeNull();
    await waitFor(() => expect(text(view)).toContain("补记只改账本状态，原表未再做任何写入"));
  });

  it.each(["confirm", "discard"] as const)("公式续跑完成后进入审阅收尾（%s）：确认不重放、放弃恢复原公式与格式", async (outcome) => {
    // 执行前现场：A1:B1 原本是空（先记住，再模拟"第一步已写进原表、回执丢失"）。
    const beforeFirstRow = host.matrix(TARGET, "A1:B1");
    host.setCell(TARGET, "A1", '="影片"', { asFormula: true });
    host.setCell(TARGET, "B1", "=1200", { asFormula: true });
    host.setCellStyle(TARGET, "A1", { fontBold: true });
    host.setCellDisplay(TARGET, "A1", { format: "0.00" });
    const firstWrites = contentWrites(host.originalWrites()).length;

    pendingIndex();
    // 替身服务端：核对第一步已写入 → 推进到第 2 步；续跑完成 → 记账本完成。
    let reconciled = false;
    let checkpoint: unknown = null;
    const settledState = (revision: number) => resumeState({
      checkpoint: checkpoint as never,
      nextActionIndex: 1, inFlightActionIndex: null, requiresReconciliation: false, revision
    });
    vi.mocked(getOfficeExecutionStatus).mockImplementation(async () =>
      (reconciled ? settledState(3) : resumeState()) as never
    );
    vi.mocked(reconcileOfficeActionReceipt).mockImplementation(async (_id, request) => {
      reconciled = true;
      checkpoint = request.checkpoint;
      return settledState(3) as never;
    });
    vi.mocked(claimOfficeExecutionResume).mockImplementation(async () => settledState(4) as never);
    vi.mocked(finishOfficeExecution).mockResolvedValue(
      resumeState({ nextActionIndex: 2, inFlightActionIndex: null, status: "completed", revision: 5 }) as never
    );

    const view = render(<App />);
    fireEvent.click(await view.findByRole("button", { name: RECOVERY_CHECK_BUTTON }));
    await waitFor(() => expect(reconcileOfficeActionReceipt).toHaveBeenCalledTimes(1));
    // ①核对只补记：一个单元格都没写。
    expect(contentWrites(host.originalWrites()).slice(firstWrites)).toEqual([]);

    // ②继续：只写第二步（A2:B2），第一步的 A1:B1 不重写。
    fireEvent.click(await view.findByRole("button", { name: RECOVERY_RESUME_BUTTON }));
    await waitFor(() => expect(host.cell(TARGET, "A2")).toBe('="乙"'));
    expect(contentWrites(host.originalWrites()).slice(firstWrites).map((w) => w.address))
      .toEqual(["A2:B2"]);
    await waitFor(() => expect(finishOfficeExecution).toHaveBeenCalledTimes(1));

    // ③审阅收尾：续跑完成后必须能进入审阅（有确认/放弃），而不是因为
    //    "前几步判据在上一次运行里"被判成 executed_unverified 后整卡 failed。
    //    这是本用例要钉住的关键：续跑完成 ≠ 不能审阅。
    await waitFor(() => expect(view.queryByRole("button", { name: "确认全部变更" })).not.toBeNull());
    expect(view.getByRole("button", { name: "放弃全部变更" })).toBeTruthy();

    if (outcome === "confirm") {
      // ④确认不重放：点确认后 A1:B1 与 A2:B2 都不再被写一次。
      const writesBeforeConfirm = contentWrites(host.originalWrites()).length;
      fireEvent.click(view.getByRole("button", { name: "确认全部变更" }));
      await waitFor(() => expect(text(view)).toContain("已确认变更"));
      expect(contentWrites(host.originalWrites())).toHaveLength(writesBeforeConfirm);
      // 公式仍在（确认只清标记，不回写）。
      expect(host.cell(TARGET, "A1")).toBe('="影片"');
      expect(host.cell(TARGET, "A2")).toBe('="乙"');
    } else {
      // ⑤放弃恢复：原公式与格式一起回到执行前，不留本次写入的公式。
      fireEvent.click(view.getByRole("button", { name: "放弃全部变更" }));
      await waitFor(() => expect(text(view)).toContain("已放弃并恢复"));
      expect(host.matrix(TARGET, "A1:B1")).toEqual(beforeFirstRow);
      expect(host.matrix(TARGET, "A2:B2")).toEqual([["", ""]]);
    }
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
    /*
     * ★「确认之后撤销」在 Excel 上是**真的存在**的路径（2026-10-29 核实结论）。
     *
     * 它走的是结果卡里的「↶ 撤销本次执行」→ `undoExecutionTo(executionId)` →
     * `useUndoSnapshot` 逐条 `await undoExecution(snapshot)` 成功后才回调
     * `onRestored` —— 留痕因此只在**宿主真的恢复完**之后才写，且这里必须是
     * `undone`（"确认之后撤销"），**不是** `abandoned`（"审阅阶段放弃"）。
     * 两个取值互不冒充：把这条断言改成 `abandoned` 就会红。
     */
    await waitFor(() => expect(vi.mocked(recordExecutionDisposition)).toHaveBeenCalledTimes(1));
    const [undoNamespace, undoneId, undoKind] = vi.mocked(recordExecutionDisposition).mock.calls[0];
    expect(undoNamespace).toBe("office");
    expect(typeof undoneId).toBe("string");
    expect((undoneId as string).length).toBeGreaterThan(0);
    expect(undoKind).toBe("undone"); // 不是 "abandoned"
    // 留痕用的是**这次执行的身份**：同一条结果卡撤销的就是它自己，不借用别的记录。
    expect(text(view)).toContain("已撤销上一次执行");
    expect(host.originalWrites().length).toBeGreaterThan(0);
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
    /*
     * ★ 恢复失败 ⇒ **不留痕**。留痕的语义是"用户确实把这次执行撤回了"；撤销被冲突
     * 挡下时宿主一个单元格都没回退，写 `undone_at` 就是造假。判据的位置在
     * `useUndoSnapshot`：`onRestored` 只在 `await undoExecution(snapshot)` 成功之后
     * 才被调用，抛错分支只发一条消息。
     */
    expect(vi.mocked(recordExecutionDisposition)).not.toHaveBeenCalled();
    // ★ 撤销**冲突**绝不能把卡标成"已撤销"：标题仍是验证结论，不出现"已撤销并恢复"。
    const cardAfterConflict = undoButton.closest(".execution-result-card");
    expect(cardAfterConflict).not.toBeNull();
    expect(cardAfterConflict!.textContent).not.toContain("已撤销并恢复");
    expect(cardAfterConflict!.textContent).not.toContain("本次执行已撤销");
    expect(cardAfterConflict!.textContent).toContain("验证通过");
  });

  it("确认后撤销的留痕上报失败：不报错、不重试，撤销入口随之消失（没有第二次撤销）", async () => {
    /*
     * 「重试」这件事在本产品里只有一种合法形态：**重发那一条 `POST …/disposition`**。
     * 服务端首次写赢（`WHERE undone_at IS NULL`），所以重发是幂等的；而"再撤一次"
     * 是另一件事 —— 宿主恢复是不可重复的副作用，绝不能挂在重试路径上。
     *
     * 这条用例钉住调用方那一侧：上报失败 ⇒ 端点只被叫 **1** 次（不自动重试）、
     * 不弹错误、单元格照样回到执行前；撤销入口随栈消费一起消失 ⇒ 界面上根本没有
     * 第二次撤销的入口。
     */
    vi.mocked(recordExecutionDisposition).mockRejectedValueOnce(new Error("网络不可用"));
    const view = await sendPlan(reviewPlan);
    await stageReview(view);
    await clickSave(view);
    await waitFor(() => expect(host.cell(SHEET, "G1")).toBe("合计"));

    const undoButton = [...view.container.querySelectorAll("button")].find(
      (candidate) => candidate.textContent === "↶ 撤销本次执行"
    ) as HTMLButtonElement;
    fireEvent.click(undoButton);
    await waitFor(() => expect(text(view)).toContain("已撤销上一次执行"));
    await waitFor(() => expect(vi.mocked(recordExecutionDisposition)).toHaveBeenCalledTimes(1));

    // 留痕失败不改判：单元格确实回退了，界面也不出现任何"无法记录"的失败提示。
    expect(host.cell(SHEET, "G1")).toBe("");
    expect(text(view)).not.toContain("无法记录这次执行的处置");
    expect(vi.mocked(recordExecutionDisposition).mock.calls[0][2]).toBe("undone");

    // 入口已消费：结果卡进终态，`↶ 撤销本次执行` 不再存在 ⇒ 不存在第二次撤销的路径。
    const resultCard = view.container.querySelector(".execution-result-card");
    expect(resultCard).not.toBeNull();
    expect(resultCard!.textContent).toContain("本次执行已撤销");
    expect(view.queryByRole("button", { name: "↶ 撤销本次执行" })).toBeNull();
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
    const discard = view.getByRole("button", { name: "关闭数据范围选择" });
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
  await waitFor(() => {
    // 恢复待办会替换普通输入区，模型选择器此时不渲染。
    const modelReady = view.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型");
    const recoveryReady = view.container.querySelector(".composer .recovery-composer");
    expect(Boolean(modelReady || recoveryReady)).toBe(true);
  });
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
  it("scope confirmation commits the changed active sheet without toggling the mode", async () => {
    const view = await bootApp();
    await backgroundCheck(view);
    host.activateSheet(TARGET);
    await backgroundCheck(view);
    await waitFor(() => expect(scopeNotice(view)).toContain("工作表范围发生变化"));
    await openScopePicker(view);
    await waitFor(() => expect(view.container.querySelector(".scope-current-file")?.textContent ?? view.container.querySelector(".scope-popover")?.textContent).toContain(TARGET));
    // Only the footer confirmation is clicked; no checkbox or mode gesture.
    fireEvent.click(view.container.querySelector<HTMLButtonElement>(".scope-close")!);
    await waitFor(() => expect(view.container.querySelector(".scope-popover")).toBeNull());
    await waitFor(() => expect(scopeNotice(view)).toBeNull());
    await sendAndExpectAnswer(view, "检查当前目标表", "已核验新目标表");
    expect(registerOfficeExecution).not.toHaveBeenCalled();
    host.activateSheet(SHEET);
    await backgroundCheck(view);
    await waitFor(() => expect(scopeNotice(view)).toContain("工作表范围发生变化"));
  });

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


it("executes a trusted literal without a source read and confirmation does not rewrite", async () => {
  host.setCell(SHEET, "A1", "旧值");
  const plan = {id: "literal-native-s3", title: "写入完成", summary: "写入完成", assumptions: [], warnings: [],
    actions: [{type: "writeValues", sheet: SHEET, range: "A1", values: [["完成"]]}],
    acceptanceCriteria: [{type: "rangeEquals", sheet: SHEET, range: "A1", expected: [["完成"]]}]};
  const view = await sendPlan(plan, "仅在 A1 写入完成", {planOrigin: "literal_range_edit"});
  await waitFor(() => expect(view.queryByRole("button", {name: "确认全部变更"}), text(view)).not.toBeNull());
  expect(host.cell(SHEET, "A1")).toBe("完成");
  const writes = contentWrites(host.originalWrites()); expect(writes).toHaveLength(1);
  fireEvent.click(await view.findByRole("button", {name: "确认全部变更"}));
  await waitFor(() => expect(text(view)).toContain("已确认变更"));
  expect(contentWrites(host.originalWrites())).toEqual(writes);
});

it.each(["confirm", "discard"])("executes three trusted literal actions in order without a source read before %s", async outcome => {
  [-101,-202,-303].forEach((value,index)=>host.setCell(SHEET,`B${index+2}`,value));
  const plan = {id:"literal-ordered-three",title:"依次写入常量",summary:"三步独立写入",assumptions:[],warnings:[],
    actions:[111,222,333].map((value,index)=>({type:"writeValues",sheet:SHEET,range:`B${index+2}`,values:[[value]]})),
    acceptanceCriteria:[111,222,333].map((value,index)=>({type:"rangeEquals",sheet:SHEET,range:`B${index+2}`,expected:[[value]]}))};
  const view = await sendPlan(plan,"B2/B3/B4依次写入111/222/333，不合并动作",{planOrigin:"literal_range_edit"});
  await waitFor(()=>expect(view.queryByRole("button",{name:"确认全部变更"}),text(view)).not.toBeNull());
  expect([2,3,4].map(row=>host.cell(SHEET,`B${row}`))).toEqual([111,222,333]);
  const writes = contentWrites(host.originalWrites());
  expect(writes).toHaveLength(3);
  expect(writes.map(write=>write.address)).toEqual(["B2","B3","B4"]);
  fireEvent.click(await view.findByRole("button",{name:outcome==="confirm"?"确认全部变更":"放弃全部变更"}));
  await waitFor(()=>expect(text(view)).toContain(outcome==="confirm"?"已确认变更":"已放弃并恢复"));
  if(outcome==="confirm") expect(contentWrites(host.originalWrites())).toEqual(writes);
  else expect([2,3,4].map(row=>host.cell(SHEET,`B${row}`))).toEqual([-101,-202,-303]);
});

/**
 * CAP-03（Excel）：完整退出 / 换机器之后**重新发现**未完成执行。
 *
 * 此前 Excel 端只有一条 `localStorage` 索引把用户带回断点，索引不在就再也看不到那次执行。
 * 这一组用例钉三件事：
 *
 * 1. **多候选不自动选**：只给选择入口，一个状态查询都不发；
 * 2. **发现 ≠ 可以续跑**：身份不符只给阻断 + 只读重核，绝不摆出续跑按钮；
 * 3. **同一条执行会被重新核验**：用户确认收窄授权后，卡面按新授权变红，而不是停在旧结论上。
 *
 * 全部断言都带"零写入"：发现这条路只读，claim / 补记 / 执行一个都不许发生。
 */
describe("Excel 未完成执行的发现入口（真实 App / 宿主与网络替身）", () => {
  const OPERATION = "office-discovery-0001";
  const identity = () => `url:${BOOK_URL}`;
  const mutationSpies = () => [
    registerOfficeExecution,
    startOfficeAction,
    completeOfficeAction,
    reconcileOfficeActionReceipt,
    finishOfficeExecution,
    suspendOfficeExecution
  ];

  function candidate(overrides: Record<string, unknown> = {}) {
    return {
      operationId: OPERATION,
      taskId: "task-discovery-0001",
      // 发现候选现在带"性质 + 审阅状态"（2026-11-03）：替身必须与服务端同形，
      // 否则准入闸会以"缺少性质标记"整条拒绝 —— 那会掩盖真正的发现判据。
      kind: "resume",
      status: "running",
      reviewState: null,
      createdAt: 1_790_000_000_000,
      updatedAt: 1_790_000_600_000,
      nextActionIndex: 0,
      totalActions: 2,
      inFlightActionIndex: null,
      ownerActive: false,
      requiresReconciliation: false,
      canResume: true,
      ...overrides
    };
  }

  function discovery(candidates: unknown[], overrides: Record<string, unknown> = {}) {
    vi.mocked(listPendingOfficeExecutions).mockResolvedValue({
      schemaVersion: 1,
      namespace: "office",
      workbookIdentity: identity(),
      limit: 20,
      truncated: false,
      candidates,
      ...overrides
    } as never);
  }

  /** 一条两动作的计划：都写 `TARGET`（汇总表）。 */
  function executionState(overrides: Record<string, unknown> = {}) {
    return {
      operationId: OPERATION,
      status: "running",
      revision: 3,
      workbookIdentity: identity(),
      plan: {
        id: "plan-discovery",
        title: "补两格",
        summary: "写入汇总表",
        assumptions: [],
        warnings: [],
        actions: [
          { type: "writeValues", sheet: TARGET, range: "B2", values: [[111]] },
          { type: "writeValues", sheet: TARGET, range: "B3", values: [[222]] }
        ]
      },
      authorizedSheetNames: [SHEET, TARGET],
      allowedExternalFiles: [],
      nextActionIndex: 0,
      inFlightActionIndex: null,
      beforeSnapshot: null,
      checkpoint: null,
      result: null,
      error: null,
      canResume: true,
      requiresReconciliation: false,
      ownerToken: null,
      ...overrides
    };
  }

  it("多条候选只给选择入口：不自动选、不查状态、不写任何单元格", async () => {
    discovery([
      candidate({ operationId: "op-newer", status: "failed", updatedAt: 1_790_000_900_000 }),
      candidate()
    ]);
    const view = await bootApp();
    await waitFor(() => expect(view.container.querySelector(".office-recovery-choices")).not.toBeNull());
    // 两条候选各一个按钮：界面不替用户挑"最新的一条"。
    expect(view.getAllByRole("button", { name: "核对这一次" })).toHaveLength(2);
    expect(getOfficeExecutionStatus).not.toHaveBeenCalled();
    expect(text(view)).not.toContain("继续剩余动作");
    /*
     * 没有可选的**那一条**时，恢复操作区不该顶掉输入框（没东西可操作），
     * 但这条待办也不许"因为没接管输入框就消失"：它仍然在固定区域里的恢复视图里。
     */
    expect(view.container.querySelector("[data-recovery-composer]")).toBeNull();
    expect(view.container.querySelector(".composer-box textarea")).not.toBeNull();
    const dockRecovery = view.container.querySelector<HTMLElement>('[data-dock-view="recovery"]');
    expect(dockRecovery, "恢复待办必须在固定区域里看得见").not.toBeNull();
    expect(dockRecovery!.contains(view.container.querySelector(".office-recovery-choices"))).toBe(true);
    expect(host.writes()).toEqual([]);
    for (const spy of mutationSpies()) expect(spy).not.toHaveBeenCalled();
  });

  it("审阅在等决定时：恢复待办退成 chip 但不消失，输入框不被顶掉，也不谎称入口在里面", async () => {
    /*
     * 两个待办同时在场：审阅（优先级 1）与未收尾的写入（优先级 2）。
     * 输入区只能归一件事 —— 审阅期间用户可能正在打字决定要不要确认，所以恢复操作区
     * **不接管**。这条钉四件事：①恢复没被藏起来（chip 在，点得开）；
     * ②输入框还在（不是被置换掉）；③卡面不说"在输入框里"；④全程零写入。
     */
    discovery([candidate()]);
    vi.mocked(getOfficeExecutionStatus).mockResolvedValue(executionState() as never);
    const view = await bootApp();
    await waitFor(() => expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).not.toBeNull());
    // 基线：没人跟它抢的时候，输入区确实归恢复操作区（没有 textarea）。
    expect(view.container.querySelector("[data-recovery-composer]")).not.toBeNull();
    expect(view.container.querySelector(".composer-box textarea")).toBeNull();

    /*
     * 用户先做别的事：「另提需求」把输入框要回来，发一条会写单元格的任务并走到审阅。
     * 这一下才让**两件待办同时在场**（审阅在等决定 + 上一次写入没收尾）。
     * 必须用**同一个 App 实例**发（`sendPlan` 会另挂一份 App，等于换了个现场）。
     */
    await setAuthorizedScope(view, [SHEET, TARGET]);
    fireEvent.click(view.getByRole("button", { name: "另提需求" }));
    await waitFor(() => expect(view.container.querySelector(".composer-box textarea")).not.toBeNull());
    vi.mocked(understandMessage).mockResolvedValue({
      kind: "needs_workbook_context",
      summary: "补上座率",
      provider: "model"
    } as never);
    vi.mocked(checkIntent).mockResolvedValue({
      kind: "proceed",
      summary: "补上座率",
      confirmedPrompt: "补上座率",
      provider: "model"
    } as never);
    vi.mocked(streamAssistantResponse).mockResolvedValue({
      kind: "plan",
      provider: "model",
      plan: reviewPlan
    } as never);
    fireEvent.change(view.container.querySelector(".composer textarea")!, {
      target: { value: "补上座率" }
    });
    fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
    await stageReview(view);

    const dock = view.container.querySelector<HTMLElement>(".composer-action-dock")!;
    expect(dock.getAttribute("data-dock-primary")).toBe("review");
    const chip = view.container.querySelector<HTMLElement>('[data-dock-switch="recovery"]');
    expect(chip, "恢复待办不许因为输入框被审阅占着就消失").not.toBeNull();
    expect(chip!.textContent).toContain("恢复核对");
    // 审阅置换输入框，恢复待办仍可从 chip 查看。
    expect(view.container.querySelector(".composer-box textarea")).toBeNull();
    expect(view.container.querySelector(".composer-box .result-review-composer")).not.toBeNull();
    expect(view.container.querySelector("[data-recovery-composer]")).toBeNull();
    expect(view.container.querySelector(".recovery-pending-entry")).toBeNull();
    // 卡面此刻没展开（展开的是审阅那一块）：恢复视图连渲染都没渲染，自然不该有指路句。
    expect(view.container.querySelector(".office-recovery-panel")).toBeNull();

    // 显式切过去：卡与入口都还在，用户随时能看清这条待办是什么。
    const writesBefore = host.writes().length;
    fireEvent.click(chip!);
    await waitFor(() =>
      expect(view.container.querySelector('[data-dock-view="recovery"]')).not.toBeNull()
    );
    expect(view.container.querySelector('[data-dock-view="review"]')).toBeNull();
    expect(view.container.querySelector('[data-dock-switch="review"]')).not.toBeNull();
    expect(view.container.querySelector(".office-recovery-panel")).not.toBeNull();
    // 判别句只属于"输入区归别人"这一档：不写"在输入框里"（此刻那里是普通 textarea）。
    const hint = view.container.querySelector<HTMLElement>("[data-recovery-primary-hint]")!;
    expect(hint.textContent).toContain("等当前事项处理完后");
    // 切过来也没有主操作可点 —— 输入区仍然归审阅，不许在这里另画一颗。
    expect(view.container.querySelector(".composer-box [data-recovery-primary]")).toBeNull();

    // 发现这条链路依然只读：一次状态查询（发现时那一次），零补记、零写入。
    expect(
      vi.mocked(getOfficeExecutionStatus).mock.calls.every(([operationId]) => operationId === OPERATION)
    ).toBe(true);
    expect(reconcileOfficeActionReceipt).not.toHaveBeenCalled();
    expect(host.writes()).toHaveLength(writesBefore);
  });

  it("点选的一条身份不符时：只给阻断与只读重核，不给续跑按钮", async () => {
    /*
     * 让这条**排在最前**：共享排序把"有在途动作"放第一档（现场可能已被改过）。
     * 不靠"最新的一条"来选 —— 第 1 条用例已经钉住"界面不替用户挑"。
     */
    discovery([
      candidate({
        operationId: "op-other-book",
        status: "needs_reconciliation",
        inFlightActionIndex: 1,
        requiresReconciliation: true
      }),
      candidate()
    ]);
    vi.mocked(getOfficeExecutionStatus).mockResolvedValue(
      executionState({ workbookIdentity: "url:https://contoso.sharepoint.com/another.xlsx" }) as never
    );
    const view = await bootApp();
    await waitFor(() => expect(view.container.querySelector(".office-recovery-choices")).not.toBeNull());
    fireEvent.click(view.getAllByRole("button", { name: "核对这一次" })[0]);
    await waitFor(() => expect(text(view)).toContain("当前不是产生该断点的工作簿"));
    expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(view.queryByRole("button", { name: "核对并继续" })).toBeNull();
    expect(view.getByRole("button", { name: "重新核对（只读）" })).toBeTruthy();
    // 点选本身只读：一个写入类接口都不许被调用。
    expect(getOfficeExecutionStatus).toHaveBeenCalledTimes(1);
    expect(vi.mocked(getOfficeExecutionStatus).mock.calls[0][0]).toBe("op-other-book");
    expect(host.writes()).toEqual([]);
    for (const spy of mutationSpies()) expect(spy).not.toHaveBeenCalled();
  });

  it("单条候选自动摆卡；用户确认收窄授权后同一条执行被重新核验并如实阻断", async () => {
    discovery([candidate()]);
    vi.mocked(getOfficeExecutionStatus).mockResolvedValue(executionState() as never);
    const view = await bootApp();
    await waitFor(() => expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).not.toBeNull());
    // 只看"读的是哪一条"，不钉次数：工作簿目录晚到会再核对一轮（只读，不写入）。
    expect(
      vi.mocked(getOfficeExecutionStatus).mock.calls.every(([operationId]) => operationId === OPERATION)
    ).toBe(true);

    /*
     * 用户在范围面板里只留「综合票房」并确认：`TARGET` 被收窄掉。
     * 卡面必须跟着变红 —— 停在"可继续"就是这次要修的那种"用户已经改正、界面还在说旧结论"。
     */
    await setAuthorizedScope(view, [SHEET]);
    await waitFor(() => expect(text(view)).toContain("当前授权里没有工作表「汇总表」"));
    expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(view.getByRole("button", { name: "重新核对（只读）" })).toBeTruthy();

    // 重新核验的是**同一条**执行：每一次状态查询的 operationId 都相同，没有被换成别的候选。
    const calls = vi.mocked(getOfficeExecutionStatus).mock.calls;
    expect(calls.length).toBeGreaterThanOrEqual(2);
    expect(calls.every(([operationId]) => operationId === OPERATION)).toBe(true);
    expect(host.writes()).toEqual([]);
    for (const spy of mutationSpies()) expect(spy).not.toHaveBeenCalled();
  });

  it("迟到的发现结果不许落盘：授权变过之后旧响应既不摆卡也不换对象", async () => {
    /*
     * 现场：第一次发现请求发出后用户就改了授权（请求在飞、界面已经换了现场）。
     * 迟到的响应属于**旧代数**，既不许把候选摆出来，也不许把用户的入口换掉。
     *
     * 只有这条能钉住代数闸：既有三条判据（同一条 / 次数 / 只读）在"没有闸"时
     * 照样成立 —— 旧响应会自己走完 adopt 流程，把上一个现场摆到当前界面上。
     */
    let releaseStale: ((value: unknown) => void) | null = null;
    const stale = new Promise((resolve) => {
      releaseStale = resolve;
    });
    let issued = 0;
    vi.mocked(listPendingOfficeExecutions).mockImplementation(async () => {
      issued += 1;
      if (issued === 1) return stale as never;
      return {
        schemaVersion: 1,
        namespace: "office",
        workbookIdentity: identity(),
        limit: 20,
        truncated: false,
        candidates: []
      } as never;
    });
    const view = await bootApp();
    await waitFor(() => expect(issued).toBeGreaterThanOrEqual(1));

    // 授权变一次：代数推进，第一份响应从此作废。
    await setAuthorizedScope(view, [SHEET, TARGET]);
    await waitFor(() => expect(issued).toBeGreaterThanOrEqual(2));

    // 现在才让**旧**那条响应落地：它带着一条看起来完全可以续跑的候选。
    await act(async () => {
      releaseStale?.({
        schemaVersion: 1,
        namespace: "office",
        workbookIdentity: identity(),
        limit: 20,
        truncated: false,
        candidates: [candidate({ operationId: "op-stale-generation" })]
      });
      await Promise.resolve();
    });

    expect(getOfficeExecutionStatus).not.toHaveBeenCalled();
    expect(view.container.querySelector(".office-recovery-choices")).toBeNull();
    expect(text(view)).not.toContain("继续剩余动作");
    expect(host.writes()).toEqual([]);
    for (const spy of mutationSpies()) expect(spy).not.toHaveBeenCalled();
  });

  /*
   * 下面五条钉的是"发现与续跑**不扩大授权**"，以及四个真机上最常见的异常分支。
   *
   * 全部用**确定性**手段造现场：探针在第 2 步回读之后、完成回执之前中断（不靠手动暂停碰时机）、
   * `clearPendingOfficeExecution` 模拟"换机器 / 清缓存"（不要求用户手删 localStorage）、
   * `host.setDocumentUrl` 模拟换簿（不要求用户真去另存为）、把状态替身改成"在途但没快照"
   * （不要求用户停真实服务）。
   */

  it("窗格重开且本机索引不在：靠后端发现摆卡，只续剩余动作，再走审阅收尾", async () => {
    // 与「acceptance entry」同一套服务端状态机：第 2 步写入回读后、完成回执前中断。
    [111, 222, 333].forEach((value, index) => host.setCell(SHEET, `B${index + 2}`, value));
    [-101, -202, -303].forEach((value, index) => host.setCell(TARGET, `B${index + 2}`, value));
    let state: Awaited<ReturnType<typeof getOfficeExecutionStatus>>;
    vi.mocked(registerOfficeExecution).mockImplementation(async (request) => {
      state = { ...request, status: "running", revision: 1, nextActionIndex: 0, inFlightActionIndex: null,
        checkpoint: null, beforeSnapshot: null, canResume: false, requiresReconciliation: false,
        ownerToken: "owner" } as typeof state;
      return state;
    });
    vi.mocked(startOfficeAction).mockImplementation(async (_id, request) => {
      state = { ...state, revision: state.revision + 1, inFlightActionIndex: request.actionIndex,
        beforeSnapshot: request.beforeSnapshot, requiresReconciliation: true };
      return state;
    });
    vi.mocked(completeOfficeAction).mockImplementation(async (_id, request) => {
      state = { ...state, revision: state.revision + 1, inFlightActionIndex: null,
        nextActionIndex: request.actionIndex + 1, checkpoint: request.checkpoint, requiresReconciliation: false };
      return state;
    });
    vi.mocked(suspendOfficeExecution).mockImplementation(async () => {
      state = { ...state, status: "needs_reconciliation", revision: state.revision + 1,
        ownerToken: undefined, canResume: true };
      return state;
    });
    vi.mocked(getOfficeExecutionStatus).mockImplementation(async () => state);
    vi.mocked(reconcileOfficeActionReceipt).mockImplementation(async (_id, request) => {
      expect(request.actionIndex).toBe(1);
      expect(request.outcome).toBe("applied");
      state = { ...state, status: "running", revision: state.revision + 1, ownerToken: "new-owner",
        nextActionIndex: 2, inFlightActionIndex: null, requiresReconciliation: false,
        checkpoint: request.checkpoint! };
      return state;
    });
    vi.mocked(claimOfficeExecutionResume).mockImplementation(async () => {
      state = { ...state, status: "running", revision: state.revision + 1, ownerToken: "resume-owner" };
      return state;
    });
    vi.mocked(finishOfficeExecution).mockImplementation(async () => ({ ...state, status: "completed" } as never));

    const plan = { id: "remount-three", title: "三次值复制", summary: "复制到目标表", assumptions: [], warnings: [],
      actions: [2, 3, 4].map((row) => ({ type: "copyRange", sheet: TARGET, sourceSheet: SHEET,
        sourceRange: `B${row}`, targetRange: `B${row}`, copyType: "values", skipBlanks: false, transpose: false })) };
    const view = await sendPlan(plan, "依次执行三个值复制", { authorize: [SHEET, TARGET], probe: true });
    await waitFor(() => expect(suspendOfficeExecution).toHaveBeenCalledTimes(1));
    expect([2, 3, 4].map((row) => host.cell(TARGET, `B${row}`))).toEqual([111, 222, -303]);
    expect(state!.nextActionIndex).toBe(1);
    expect(state!.inFlightActionIndex).toBe(1);
    const operationId = state!.operationId;
    const before = contentWrites(host.originalWrites());
    expect(before).toHaveLength(2);

    /*
     * 窗格重开**且本机索引不在**：换机器、清缓存、换端口都是这个样子。
     * 不要求用户手删 localStorage —— 测试直接清掉，真机上是 DEV 面板里那个"模拟本机索引丢失"。
     */
    view.unmount();
    clearPendingOfficeExecution(operationId);
    expect(loadPendingOfficeExecution()).toBeNull();
    vi.mocked(listPendingOfficeExecutions).mockImplementation(async (workbookIdentity) => ({
      schemaVersion: 1,
      namespace: "office",
      workbookIdentity,
      limit: 20,
      truncated: false,
      candidates: [candidate({ operationId, nextActionIndex: 1, inFlightActionIndex: 1,
        requiresReconciliation: true, status: "needs_reconciliation" })]
    }) as never);

    const reopened = render(<App />);
    const reconcileButton = await reopened.findByRole("button", { name: RECOVERY_CHECK_BUTTON });
    await waitFor(() => expect((reconcileButton as HTMLButtonElement).disabled).toBe(false));
    /*
     * 发现只补"挂接记录"，不改授权：续跑读的就是这份索引。
     *
     * 授权范围本身**不跨窗格重开持久化**（已提交范围是内存态，重开后回到首读播种的
     * "跟随当前表"），这是既有产品语义，不是发现改的 —— 所以这里不拿"已固定 2 张表"
     * 当判据，而是看续跑有没有越过授权边界写出去（下面那两条）。
     */
    expect(loadPendingOfficeExecution()?.operationId).toBe(operationId);

    fireEvent.click(reconcileButton);
    await waitFor(() => expect(reconcileOfficeActionReceipt).toHaveBeenCalledTimes(1));
    // 核对只补记回执：第 3 步还没跑，这一下不写任何单元格。
    expect(contentWrites(host.originalWrites()).slice(before.length)).toEqual([]);
    expect(finishOfficeExecution).not.toHaveBeenCalled();
    fireEvent.click(await reopened.findByRole("button", { name: RECOVERY_RESUME_BUTTON }));
    await waitFor(() => expect(finishOfficeExecution, text(reopened)).toHaveBeenCalledTimes(1));
    // 只补第 3 步：第 2 步已经在中断前写完（核对判定 applied），不许重放。
    expect(contentWrites(host.originalWrites()).slice(before.length).map((write) => write.address)).toEqual(["B4"]);
    expect([2, 3, 4].map((row) => host.cell(TARGET, `B${row}`))).toEqual([111, 222, 333]);
    expect(vi.mocked(completeOfficeAction).mock.calls.map((call) => call[1].actionIndex)).toEqual([0, 2]);
    expect(registerOfficeExecution).toHaveBeenCalledTimes(1);
    expect(loadPendingOfficeExecution()).toBeNull();

    // 审阅收尾：断点续跑的结论与回执都在，且不残留"执行未完成"。
    await stageReview(reopened);
    expect(text(reopened)).toContain("从第 3 步继续并完成");
    expect(reopened.container.querySelector(".result-review-bar")?.textContent ?? "").not.toContain("执行未完成");
  });

  it("从没确认过范围（confirmedScope=null）时不误判成撤销：卡可继续，且不动授权", async () => {
    discovery([candidate()]);
    vi.mocked(getOfficeExecutionStatus).mockResolvedValue(executionState() as never);
    const view = await bootApp();
    const scopeBefore = view.container.querySelector(".scope-trigger")?.textContent;
    const button = await view.findByRole("button", { name: RECOVERY_RESUME_BUTTON });
    expect((button as HTMLButtonElement).disabled).toBe(false);
    /*
     * 首读播种把已提交范围写成**当前活动表**（综合票房），而计划要写汇总表。
     * 那不是用户的手势 —— 拿它比对就会凭空说"你把汇总表取消了"。
     */
    expect(text(view)).not.toContain("当前授权里没有工作表");
    expect(view.container.querySelector(".scope-trigger")?.textContent).toBe(scopeBefore);
    expect(loadPendingOfficeExecution()?.operationId).toBe(OPERATION);
    expect(host.writes()).toEqual([]);
    for (const spy of mutationSpies()) expect(spy).not.toHaveBeenCalled();
  });

  it("授权被撤销（清空并确认）后：卡面立刻阻断，且发现不会把表加回来", async () => {
    discovery([candidate()]);
    vi.mocked(getOfficeExecutionStatus).mockResolvedValue(executionState() as never);
    const view = await bootApp();
    await waitFor(() => expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).not.toBeNull());

    await openScopePicker(view);
    fireEvent.click(view.getByRole("button", { name: "清空" }));
    await waitFor(() =>
      expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false")
    );
    fireEvent.click(view.container.querySelector<HTMLButtonElement>(".scope-close")!);
    await waitFor(() => expect(view.container.querySelector(".scope-popover")).toBeNull());

    await waitFor(() => expect(text(view)).toContain("当前授权里没有工作表「汇总表」"));
    expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(view.getByRole("button", { name: "重新核对（只读）" })).toBeTruthy();
    // 撤销之后仍然只有只读入口：重新核对不许借机把授权恢复。
    fireEvent.click(view.getByRole("button", { name: "重新核对（只读）" }));
    await waitFor(() => expect(text(view)).toContain("当前授权里没有工作表「汇总表」"));

    /*
     * 刷新一轮目录（不是勾选手势）：草稿会被重新播种成"跟随当前表"，但已提交的授权
     * 仍是空的 —— 卡面不许因为这一轮刷新又变成"可继续"。
     */
    await openScopePicker(view);
    fireEvent.click(view.getByRole("button", { name: /刷新/ }));
    await waitFor(() =>
      expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false")
    );
    await waitFor(() => expect(text(view)).toContain("当前授权里没有工作表「汇总表」"));
    expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(host.writes()).toEqual([]);
    for (const spy of mutationSpies()) expect(spy).not.toHaveBeenCalled();
  });

  it("换到另一个工作簿再切回：按身份改判又回来，全程零写入", async () => {
    const otherUrl = "https://contoso.sharepoint.com/sites/finance/other-workbook.xlsx";
    vi.mocked(listPendingOfficeExecutions).mockImplementation(async (workbookIdentity) => ({
      schemaVersion: 1,
      namespace: "office",
      workbookIdentity,
      limit: 20,
      truncated: false,
      candidates: workbookIdentity === `url:${BOOK_URL}` ? [candidate()] : []
    }) as never);
    vi.mocked(getOfficeExecutionStatus).mockResolvedValue(executionState() as never);
    const view = await bootApp();
    await waitFor(() => expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).not.toBeNull());

    await act(async () => {
      host.setDocumentUrl(otherUrl);
      document.dispatchEvent(new Event("visibilitychange"));
      await Promise.resolve();
    });
    await waitFor(() => expect(text(view)).toContain("当前不是产生该断点的工作簿"));
    expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(view.getByRole("button", { name: "重新核对（只读）" })).toBeTruthy();
    // 换走时**不清**本机索引：切回来还要用它把同一条接上。
    expect(loadPendingOfficeExecution()?.operationId).toBe(OPERATION);

    await act(async () => {
      host.setDocumentUrl(BOOK_URL);
      document.dispatchEvent(new Event("visibilitychange"));
      await Promise.resolve();
    });
    await waitFor(() => expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).not.toBeNull());
    expect(
      vi.mocked(getOfficeExecutionStatus).mock.calls.every(([operationId]) => operationId === OPERATION)
    ).toBe(true);
    expect(host.writes()).toEqual([]);
    for (const spy of mutationSpies()) expect(spy).not.toHaveBeenCalled();
  });

  /** 输入区上方那张**唯一**可点击的恢复卡（不是消息里的记录）。 */
  function recoveryPanelOf(view: ReturnType<typeof render>): HTMLElement | null {
    return view.container.querySelector<HTMLElement>(".office-recovery-panel");
  }

  /**
   * 输入框里的恢复操作区（接管态）。
   *
   * 2026-11-02 起它与上面那张大卡**互斥**：结论是待核对/可继续时由操作区接管，
   * 卡被收走；结论给不出"一个主操作"（阻断/证据不足/重新规划）时操作区退出，
   * 卡回来。所以"入口"永远只有一块，测试要按状态取对的那一块。
   */
  function recoveryComposerOf(view: ReturnType<typeof render>): HTMLElement | null {
    return view.container.querySelector<HTMLElement>(".composer-box .recovery-composer");
  }

  /** 当前那一块入口（两种节点取其一，不会同时有两个）。 */
  function recoveryEntryOf(view: ReturnType<typeof render>): HTMLElement | null {
    return recoveryComposerOf(view) ?? recoveryPanelOf(view);
  }

  it("恢复入口在输入框里：不在消息滚动容器里，历史消息只留没有按钮的记录", async () => {
    discovery([candidate()]);
    vi.mocked(getOfficeExecutionStatus).mockResolvedValue(executionState() as never);
    const view = await bootApp();
    await waitFor(() => expect(recoveryComposerOf(view)).not.toBeNull());
    const panel = recoveryComposerOf(view)!;
    /*
     * ⓪**只有一处**：操作区接管输入框时，固定区域那张大卡不再渲染 —— 同一个结论
     * 不画两遍（标题/正文/进度逐句重复，用户要先读两遍才知道点哪）。
     */
    expect(recoveryPanelOf(view), "接管之后上方不许再有那张大卡").toBeNull();
    // ①唯一性：全页只有一个可点的续跑入口，消息里那份记录一个按钮都没有。
    // 先断这一条再断位置：两者是**两个**属性（`mutate_office_recovery_placement.mjs`
    // 的 P1 把卡搬回消息里时，唯一性仍成立、只有位置那条红）。
    expect(view.getAllByRole("button", { name: RECOVERY_RESUME_BUTTON })).toHaveLength(1);
    const record = view.container.querySelector<HTMLElement>(".office-recovery-record");
    expect(record, "历史消息里仍要留下执行记录").not.toBeNull();
    expect(record!.querySelectorAll("button")).toHaveLength(0);
    expect(record!.closest(".message-stream"), "记录留在历史里，入口不在").not.toBeNull();
    expect(record!.textContent).toContain("操作入口在输入框里");
    /*
     * ②位置（jsdom 能钉死的那一半）：可点击的那一份**不在** `.message-stream`
     * 这个滚动容器里，而在 `.composer` 里。核对后追加的日志再怎么滚，也改不了它的位置。
     * 视口内的证据在 `scripts/check_recovery_entry_viewport.mjs`（真实样式 + 真实布局），
     * 这里不拿"元素存在"冒充"看得见"。
     */
    expect(panel.closest(".message-stream")).toBeNull();
    expect(panel.closest(".composer")).not.toBeNull();
    expect(panel.getAttribute("data-recovery-composer")).toBe("resume");
    /*
     * ③**唯一主入口**：那颗按钮就在这块操作区里，不是卡面上另画的一颗、也不是
     * 右下角那颗图标按钮。全页与这块里都只有一颗。
     */
    const primaryEntry = panel.querySelector<HTMLElement>("[data-recovery-primary]");
    expect(primaryEntry, "主入口必须在输入框里").not.toBeNull();
    expect(primaryEntry!.getAttribute("aria-label")).toMatch(RECOVERY_RESUME_BUTTON);
    expect(view.getByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBe(primaryEntry);
    expect(panel.querySelectorAll("[data-recovery-primary]")).toHaveLength(1);
    /*
     * ④**置换**而不是叠加：恢复操作区占着输入框时，里面没有 textarea，也没有
     * 右下角那颗主按钮 —— 一屏只有一个"现在该点哪"的答案。
     */
    expect(view.container.querySelector(".composer-box textarea")).toBeNull();
    expect(view.container.querySelector(".composer .composer-primary")).toBeNull();
    /*
     * ⑤「另提需求」把输入框**原样还回来**：普通 textarea 回来、主按钮回到「发送」、
     * 旧任务退成一行待办入口（不是撤销、不是放弃，任务与材料都没动）。
     */
    fireEvent.click(view.getByRole("button", { name: "另提需求" }));
    await waitFor(() => expect(view.container.querySelector(".composer-box textarea")).not.toBeNull());
    expect(view.container.querySelector(".recovery-pending-entry")).not.toBeNull();
    expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(
      view.container.querySelector(".composer .composer-primary")!.getAttribute("aria-label")
    ).toBe("发送");
    /*
     * ⑥**草稿不丢、也不被误提交**：写一句话之后从待办入口点回操作区，再按
     * 「另提需求」回到输入框，那句话还在；全过程一次核对/续跑/发送都没发生。
     */
    const composerTextarea = view.container.querySelector<HTMLTextAreaElement>(
      ".composer-box textarea"
    )!;
    fireEvent.change(composerTextarea, { target: { value: "再补一列金额" } });
    fireEvent.click(view.getByRole("button", { name: /恢复待办/ }));
    await waitFor(() => expect(view.getByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeTruthy());
    fireEvent.click(view.getByRole("button", { name: "另提需求" }));
    await waitFor(() =>
      expect(
        view.container.querySelector<HTMLTextAreaElement>(".composer-box textarea")!.value
      ).toBe("再补一列金额")
    );
    // 进度是共享层那一句，不是另拼的。
    expect(record!.textContent).toContain("已完成 0 / 2 步 · 剩余 2 步");
    expect(host.writes()).toEqual([]);
    expect(reconcileOfficeActionReceipt).not.toHaveBeenCalled();
    for (const spy of mutationSpies()) expect(spy).not.toHaveBeenCalled();
  });

  it("输入法还在组合时不算空输入框：恢复待办退成入口行，不卸载正在打字的 textarea", async () => {
    /*
     * 现场：用户正在用拼音打一句话（组合未上屏），这时一次**晚到**的发现结果摆出了恢复待办。
     * 组合期间浏览器可能还没把这几个字提交进 `value` —— 只看"文字/附件非空"就会判定
     * "输入框是空的"，于是把 textarea 整块**卸载**，用户眼睁睁看着正在打的字没了。
     * 判据因此多看一件事：`compositionstart` 到 `compositionend` 之间也算有草稿。
     */
    /*
     * 这本簿用自己的身份：上一条用例如果在收尾时往浏览器索引里写了"还有一条待办"，
     * 它带的是另一本簿的身份。发现按身份精确匹配，换一个身份就不会把它继承过来 ——
     * 否则开场就已经是恢复态，`.composer-box` 里根本没有 textarea 可以组合输入。
     */
    BOOK_URL = "https://contoso.sharepoint.com/ime-compose.xlsx";
    host.setDocumentUrl(BOOK_URL);
    /*
     * 把**每一次**发现都按住，直到测试自己 `release`：
     * 身份监视的 3s 定时器、窗口 focus/visibilitychange 都可能在同一时刻再触发一轮
     * 发现，若只按住第一轮、后面的立刻返回候选，恢复操作区就会在 `compositionStart`
     * 之前接管输入框（textarea 被卸载）。按住全部几轮，测试对"发现到底触发几次"就不敏感。
     */
    let issued = 0;
    const pending: ((value: unknown) => void)[] = [];
    vi.mocked(listPendingOfficeExecutions).mockImplementation(() => {
      issued += 1;
      return new Promise((resolve) => {
        pending.push(resolve);
      }) as never;
    });
    const releaseAll = () => {
      for (const resolve of pending.splice(0)) {
        resolve({
          schemaVersion: 1,
          namespace: "office",
          workbookIdentity: identity(),
          limit: 20,
          truncated: false,
          candidates: [candidate()]
        });
      }
    };
    vi.mocked(getOfficeExecutionStatus).mockResolvedValue(executionState() as never);
    const view = await bootApp();
    const textarea = view.container.querySelector<HTMLTextAreaElement>(".composer-box textarea")!;
    expect(textarea).not.toBeNull();

    fireEvent.compositionStart(textarea);
    await act(async () => {
      releaseAll();
      await Promise.resolve();
    });
    await waitFor(() => expect(view.container.querySelector(".office-recovery-panel")).not.toBeNull());
    // 组合中：**同一颗** textarea 原地不动，恢复待办退成一行看得见、点得开的入口。
    expect(view.container.querySelector(".composer-box textarea")).toBe(textarea);
    expect(view.container.querySelector("[data-recovery-composer]")).toBeNull();
    expect(view.container.querySelector(".recovery-pending-entry")).not.toBeNull();

    // 组合结束、而且没上屏（用户取消了这次输入）：恢复操作区立刻回来，不用再点一次。
    fireEvent.compositionEnd(textarea);
    await waitFor(() =>
      expect(view.container.querySelector("[data-recovery-composer]")).not.toBeNull()
    );
    expect(view.container.querySelector(".recovery-pending-entry")).toBeNull();
    expect(host.writes()).toEqual([]);
  });

  it("换簿再切回：结论就地改判，入口始终只有一块（操作区 ⇄ 卡）、按钮不重复", async () => {
    const otherUrl = "https://contoso.sharepoint.com/sites/finance/other-workbook.xlsx";
    vi.mocked(listPendingOfficeExecutions).mockImplementation(async (workbookIdentity) => ({
      schemaVersion: 1,
      namespace: "office",
      workbookIdentity,
      limit: 20,
      truncated: false,
      candidates: workbookIdentity === identity() ? [candidate()] : []
    }) as never);
    vi.mocked(getOfficeExecutionStatus).mockResolvedValue(executionState() as never);
    const view = await bootApp();
    await waitFor(() => expect(recoveryComposerOf(view)).not.toBeNull());
    expect(recoveryComposerOf(view)!.getAttribute("data-recovery-composer")).toBe("resume");
    expect(recoveryPanelOf(view), "接管时上方不许有那张大卡").toBeNull();

    await act(async () => {
      host.setDocumentUrl(otherUrl);
      document.dispatchEvent(new Event("visibilitychange"));
      await Promise.resolve();
    });
    await waitFor(() => expect(text(view)).toContain("当前不是产生该断点的工作簿"));
    /*
     * 结论变成"不是产生断点的工作簿"：这种结论交不出**一个**主操作，操作区退出，
     * 卡回来（阻断原因与"重新核对（只读）"只有卡上能放）。
     * 判据不是"节点没变"，而是**同一时刻只有一块入口** —— 操作区与大卡互斥，
     * 任何时刻都不许两处都说这件事。
     */
    await waitFor(() => expect(recoveryPanelOf(view)).not.toBeNull());
    const blockedPanel = recoveryPanelOf(view)!;
    expect(recoveryComposerOf(view)).toBeNull();
    expect(blockedPanel.getAttribute("data-recovery-kind")).toBe("blocked");
    expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(view.getByRole("button", { name: "重新核对（只读）" })).toBeTruthy();
    expect(view.container.querySelectorAll(".office-recovery-panel")).toHaveLength(1);
    expect(view.container.querySelectorAll(".office-recovery-record button")).toHaveLength(0);

    await act(async () => {
      host.setDocumentUrl(BOOK_URL);
      document.dispatchEvent(new Event("visibilitychange"));
      await Promise.resolve();
    });
    await waitFor(() => expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).not.toBeNull());
    // 又变回可继续：操作区重新接管，大卡再被收走 —— 仍然只有一块。
    expect(recoveryComposerOf(view), "可继续时操作区重新接管输入框").not.toBeNull();
    expect(recoveryPanelOf(view)).toBeNull();
    expect(view.getAllByRole("button", { name: RECOVERY_RESUME_BUTTON })).toHaveLength(1);
    expect(host.writes()).toEqual([]);
    for (const spy of mutationSpies()) expect(spy).not.toHaveBeenCalled();
  });

  it("核对在途：状态就地换成「正在核对…」，日志追加之后入口仍在滚动容器之外", async () => {
    discovery([candidate({ inFlightActionIndex: 1, requiresReconciliation: true })]);
    let hang = false;
    vi.mocked(getOfficeExecutionStatus).mockImplementation(async () => {
      if (hang) return new Promise<never>(() => undefined) as never;
      return executionState({
        inFlightActionIndex: 1,
        requiresReconciliation: true,
        nextActionIndex: 1,
        beforeSnapshot: {
          actionIndex: 1, sheet: TARGET, range: "B3", formulas: [[null]],
          numberFormat: [["General"]], fillColor: "#FFFFFF", fontBold: false, fontColor: "#000000"
        }
      }) as never;
    });
    const view = await bootApp();
    const reconcileButton = await view.findByRole("button", { name: RECOVERY_CHECK_BUTTON });
    const panel = recoveryComposerOf(view)!;
    expect(panel.getAttribute("data-recovery-composer")).toBe("reconcile");
    expect(recoveryPanelOf(view)).toBeNull();
    const rowsBefore = view.container.querySelectorAll(".message-stream .message-row").length;

    hang = true;
    fireEvent.click(reconcileButton);
    /*
     * 进行中：**同一块**操作区改的是自己的状态（`data-recovery-busy`），不另插一块；
     * 主按钮仍是同一颗，只是禁用 + 文案改成「正在检查…」。防双击就靠这一条：
     * 按钮在飞的时候是 disabled，第二次点击落不下去。
     */
    await waitFor(() =>
      expect(recoveryComposerOf(view)!.getAttribute("data-recovery-busy")).toBe("true")
    );
    expect(recoveryComposerOf(view)).toBe(panel);
    const primary = panel.querySelector<HTMLButtonElement>("[data-recovery-primary]");
    expect(primary).not.toBeNull();
    expect(primary!.disabled).toBe(true);
    expect(primary!.textContent).toBe("正在检查…");
    expect(panel.querySelectorAll("[data-recovery-primary]")).toHaveLength(1);
    // 核对真的追加了日志（消息流变长），入口却**不在**那个滚动容器里。
    await waitFor(() =>
      expect(view.container.querySelectorAll(".message-stream .message-row").length)
        .toBeGreaterThan(rowsBefore)
    );
    expect(recoveryComposerOf(view)).toBe(panel);
    expect(panel.closest(".message-stream")).toBeNull();
    expect(panel.closest(".composer")).not.toBeNull();
    expect(view.container.querySelectorAll(".office-recovery-record button")).toHaveLength(0);
    // 只读这一下：一个写入类接口都没被调用。
    expect(host.writes()).toEqual([]);
    for (const spy of mutationSpies()) expect(spy).not.toHaveBeenCalled();
  });

  it("续跑不信任卡面：状态里出现没有写前快照的在途动作时立刻停下，零写入", async () => {
    discovery([candidate()]);
    vi.mocked(getOfficeExecutionStatus).mockResolvedValue(executionState() as never);
    const view = await bootApp();
    const button = await view.findByRole("button", { name: RECOVERY_RESUME_BUTTON });
    /*
     * 卡片摆好**之后**服务端才出现"在途但没快照"（另一个现场/另一次中断留下的）。
     * 续跑必须重新读状态并按同一条纪律停下 —— 卡面说"能继续"不等于能继续。
     */
    vi.mocked(getOfficeExecutionStatus).mockResolvedValue(
      executionState({ inFlightActionIndex: 1, beforeSnapshot: null }) as never
    );
    await waitFor(() => expect((button as HTMLButtonElement).disabled).toBe(false));
    fireEvent.click(button);
    await waitFor(() => expect(text(view)).toContain("在途动作缺少写前快照"));
    expect(startOfficeAction).not.toHaveBeenCalled();
    expect(registerOfficeExecution).not.toHaveBeenCalled();
    expect(host.writes()).toEqual([]);
  });

  it("DEV 面板「模拟本机索引丢失」只清本机挂接，不动账本与其它本机键", async () => {
    discovery([]);
    const view = await bootApp();
    /*
     * 真机验收不该要求人去 DevTools 里删 localStorage：这里就是那条出口的判据 ——
     * 清掉挂接记录（模拟换机器/清缓存），**只**清这一个键。
     */
    savePendingOfficeExecution({
      schemaVersion: 1,
      operationId: "op-index-only",
      conversationId: "conversation-1",
      activityMessageId: "message-1",
      workbookIdentity: identity(),
      createdAt: new Date().toISOString()
    });
    const probeKey = "excel-bro.recovery-probe.v1";
    localStorage.setItem(probeKey, "keep-me");
    view.container.querySelector("details.recovery-probe-control")!.setAttribute("open", "");

    fireEvent.click(view.getByRole("button", { name: "模拟本机索引丢失" }));

    await waitFor(() => expect(loadPendingOfficeExecution()).toBeNull());
    expect(localStorage.getItem(probeKey)).toBe("keep-me");
    await waitFor(() => expect(text(view)).toContain("已清空本机挂接记录"));
    expect(host.writes()).toEqual([]);
    expect(registerOfficeExecution).not.toHaveBeenCalled();
  });
});

/**
 * 重开之后**还能撤销**这次未完成的写入。
 *
 * 真机缺陷（2026-10-31，`C:/Users/jia/Desktop/测试/公式来源证据验收-Excel-B.xlsx`）：
 * 三步公式计划在第二步回读后中断 → 用户核对在途动作 → 把 `来源!B5` 从 5 改成 500 →
 * 点继续得到"来源已变化，不能沿用旧结果续跑。"（**这条拒绝是对的**）→ 重开 Excel，
 * 恢复卡仍在、两个公式仍在表里，却**没有**「撤销本次执行」，只能手工一格一格擦。
 *
 * 根因不是材料丢了：`checkpoint`（每个已完成动作的写前值/公式 + 写后指纹）一直由
 * `complete_office_action` 落库、也随状态返回。是撤销栈**只在 `completed` 分支**用
 * `result.undoSnapshot` 重建，未完成/被阻断的分支一条都没挂 —— 而"能不能续跑"与
 * "能不能撤销"本来就是两种资格：前者依赖来源一致，后者依赖**写前恢复点 + 目标现场
 * 仍是它留下的样子**。所以本文件这一组的判据全部围绕一件事：
 * **来源变了不许续跑，但撤销照样可用，并且只写回已写入的目标。**
 *
 * 证据边界：真实 `<App/>`、真实 `executePlan` / `reconcileOfficeRangeAction` /
 * `undoExecution`，宿主是替身（`officeTestHost`），网络是替身。真机仍待按
 * `docs/FORMULA_SOURCE_RESUME_ACCEPTANCE_2026-10-30.md` 的收尾步骤点一次。
 */
describe("Excel 未完成执行：重开后从账本材料撤销已写入的内容", () => {
  const SOURCE = "来源";
  const RESULTS = ["结果1", "结果2", "结果3"];
  const SCOPE = [SOURCE, ...RESULTS];
  const OPERATION = "op-formula-undo-0001";
  /** 与真机夹具同一组公式：`来源!B2*1` / `来源!B4*2` / `来源!B5*3`。 */
  const FORMULAS = [`=${SOURCE}!B2*1`, `=${SOURCE}!B4*2`, `=${SOURCE}!B5*3`];
  const UNDO_LABEL = "↶ 撤销本次执行";
  const SOURCE_CHANGED = "来源数据已改变，不能继续原任务。";

  type LedgerState = Awaited<ReturnType<typeof getOfficeExecutionStatus>>;

  /** 来源表 + 三张空的结果表（结果表 B2 执行前都是空）。 */
  function installSourceHost() {
    host = createOfficeTestHost({
      url: BOOK_URL,
      selectedRange: "A1:B5",
      activeSheet: SOURCE,
      sheets: [
        {
          /*
           * 列名与来源表真机夹具同形（`影片` / `票房(万)`）：本轮要走**真实**工具桥
           * 读一次来源表（`query_table` 按字段名取列），列名对不上就读不出可信来源记录，
           * 计划也就登记不了公式来源证据 —— 那是替身缺料，不是产品少闸门。
           * 公式引用的三个格：B2=1、B4=3、B5=5。
           */
          name: SOURCE,
          values: [["影片", "票房(万)"], ["甲", 1], ["乙", 2], ["丙", 3], ["丁", 5]]
        },
        ...RESULTS.map((name) => ({ name, values: [] }))
      ]
    });
    installOfficeTestHost(host);
  }

  function formulaPlan() {
    return {
      id: "plan-formula-undo",
      title: "三条公式",
      summary: "三张结果表各写一条公式",
      assumptions: [],
      warnings: [],
      actions: RESULTS.map((name, index) => ({
        type: "writeFormulas" as const,
        sheet: name,
        range: "B2",
        formulas: [[FORMULAS[index]]]
      })),
      sourceDependencySheets: [SOURCE]
    };
  }

  /**
   * 账本替身按**真实时序**推进：登记 → 每步开始/回执 → 挂起 → 核对 → 续跑 → 收尾。
   *
   * 关键是 `checkpoint` 不手写：它就是产品代码交给 `complete_office_action` /
   * `reconcileOfficeActionReceipt` 的那一份（真实的写前值 + 写后指纹）。手搓一份
   * "看起来像"的快照会让本组的判据全变成替身的自说自话。
   */
  function installLedger() {
    let state: LedgerState | null = null;
    vi.mocked(registerOfficeExecution).mockImplementation(async (request) => {
      state = {
        ...request,
        status: "running",
        revision: 1,
        nextActionIndex: 0,
        inFlightActionIndex: null,
        checkpoint: null,
        beforeSnapshot: null,
        canResume: false,
        requiresReconciliation: false,
        ownerToken: "owner",
        result: null,
        error: null,
        disposition: null
      } as never;
      return state;
    });
    vi.mocked(startOfficeAction).mockImplementation(async (_id, request) => {
      state = {
        ...state!,
        revision: state!.revision + 1,
        inFlightActionIndex: request.actionIndex,
        beforeSnapshot: request.beforeSnapshot,
        requiresReconciliation: true
      } as never;
      return state;
    });
    vi.mocked(completeOfficeAction).mockImplementation(async (_id, request) => {
      state = {
        ...state!,
        revision: state!.revision + 1,
        inFlightActionIndex: null,
        nextActionIndex: request.actionIndex + 1,
        checkpoint: request.checkpoint ?? null,
        requiresReconciliation: false
      } as never;
      return state;
    });
    vi.mocked(suspendOfficeExecution).mockImplementation(async () => {
      state = {
        ...state!,
        status: "needs_reconciliation",
        revision: state!.revision + 1,
        ownerToken: null,
        canResume: true
      } as never;
      return state;
    });
    vi.mocked(getOfficeExecutionStatus).mockImplementation(async () => state!);
    vi.mocked(reconcileOfficeActionReceipt).mockImplementation(async (_id, request) => {
      state = {
        ...state!,
        status: "running",
        revision: state!.revision + 1,
        ownerToken: "reconcile-owner",
        nextActionIndex: request.actionIndex + 1,
        inFlightActionIndex: null,
        requiresReconciliation: false,
        checkpoint: request.checkpoint ?? null
      } as never;
      return state;
    });
    vi.mocked(claimOfficeExecutionResume).mockImplementation(async () => {
      state = { ...state!, revision: state!.revision + 1, ownerToken: "resume-owner" } as never;
      return state;
    });
    vi.mocked(finishOfficeExecution).mockImplementation(async () => {
      state = { ...state!, status: "completed" } as never;
      return state;
    });
    return { current: () => state! };
  }

  /** 走到真机那个现场：第 2 步写入并回读之后、完成回执之前中断。 */
  async function runToInterrupted() {
    installSourceHost();
    const ledger = installLedger();
    const view = await sendPlan(formulaPlan(), "把三条公式分别写到三张结果表", {
      authorize: SCOPE,
      authorizePool: SCOPE,
      read: { scope: "active" },
      probe: true
    });
    await waitFor(() => expect(suspendOfficeExecution).toHaveBeenCalledTimes(1));
    expect(RESULTS.map((name) => host.cell(name, "B2")))
      .toEqual([FORMULAS[0], FORMULAS[1], ""]);
    expect(vi.mocked(completeOfficeAction).mock.calls.map((call) => call[1].actionIndex))
      .toEqual([0]);
    return { view, ledger };
  }

  /** 核对在途动作（只补记回执）→ 账本进度 2/3，写前快照补齐第 2 步。 */
  async function reconcileInFlight(view: ReturnType<typeof render>) {
    const before = contentWrites(host.originalWrites()).length;
    fireEvent.click(await view.findByRole("button", { name: RECOVERY_CHECK_BUTTON }));
    await waitFor(() => expect(reconcileOfficeActionReceipt).toHaveBeenCalledTimes(1));
    expect(contentWrites(host.originalWrites()).slice(before)).toEqual([]);
    return before;
  }

  it("中断 → 核对 → 改来源（续跑被拒）→ 重开且本机索引不在 → 撤销只回退已写入的目标，来源保持用户的修改", async () => {
    const { view, ledger } = await runToInterrupted();
    const operationId = ledger.current().operationId;
    const before = await reconcileInFlight(view);
    expect(ledger.current().nextActionIndex).toBe(2);
    // 材料真的在**账本**里：第 1、2 步的写前快照都登记了，且写后状态已核对过。
    expect(ledger.current().checkpoint?.completedActionCount).toBe(2);
    expect(ledger.current().checkpoint?.postStateCoverage).toBe("complete");

    /*
     * 可继续，也可撤销 —— 两件事同时成立。撤销那条先看**说明**：这句话只在
     * "材料够 + 身份对 + 未处置"时渲染，是判据的正文，比"按钮在不在"更能定位原因。
     */
    expect(await view.findByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeTruthy();
    await waitFor(() =>
      expect(text(view)).toContain(`来源工作表「${SOURCE}」保持你现在的修改`)
    );
    expect(view.getByRole("button", { name: UNDO_LABEL })).toBeTruthy();

    /*
     * 用户把 `来源!B5` 从 5 改成 500 再点继续：**必须**拒绝，而且一格都不写。
     * 拒绝不是本组要修的东西，这里钉住它是为了下面那条 —— 撤销**不因此**消失。
     */
    host.setCell(SOURCE, "B5", 500);
    fireEvent.click(view.getByRole("button", { name: RECOVERY_RESUME_BUTTON }));
    await waitFor(() => expect(text(view)).toContain(SOURCE_CHANGED));
    expect(contentWrites(host.originalWrites()).slice(before)).toEqual([]);
    expect(host.cell(RESULTS[2], "B2")).toBe("");
    // 续跑被拒之后，撤销入口仍在（来源变化只挡住续跑那一条资格）。
    expect(view.queryByRole("button", { name: UNDO_LABEL })).not.toBeNull();

    /*
     * 重开窗格 + **本机索引不在**：`localStorage.clear()` 连挂接记录和对话一起清掉，
     * 内存里的撤销栈自然也没了。此刻界面关于这次执行的全部信息只能来自账本。
     */
    view.unmount();
    localStorage.clear();
    expect(loadPendingOfficeExecution()).toBeNull();
    vi.mocked(listPendingOfficeExecutions).mockResolvedValue({
      schemaVersion: 1,
      namespace: "office",
      workbookIdentity: `url:${BOOK_URL}`,
      limit: 20,
      truncated: false,
      candidates: [
        {
          operationId,
          taskId: "task-formula-undo",
          kind: "resume",
          reviewState: null,
          status: "needs_reconciliation",
          createdAt: 1_790_000_000_000,
          updatedAt: 1_790_000_600_000,
          nextActionIndex: 2,
          totalActions: 3,
          inFlightActionIndex: null,
          ownerActive: false,
          requiresReconciliation: false,
          canResume: true
        }
      ]
    } as never);

    const reopened = render(<App />);
    /*
     * 先看**说明**再看按钮：说明就是判据的正文（"只回退这次写入的目标；来源保持你现在
     * 的修改；不重新执行计划"）。先断言文本，是因为"按钮没出现"有多种原因（判据阻断、
     * 材料没挂进栈），而这句话只有"材料与结论都成立"时才会渲染。
     */
    await waitFor(() =>
      expect(text(reopened)).toContain(`来源工作表「${SOURCE}」保持你现在的修改`)
    );
    expect(text(reopened)).toContain("撤销不会重新执行计划");
    /*
     * 错误 #1 的回归：重开时 `来源!B5` 已经是 500（来源变了），卡面**不得**再宣称
     * "来源已核对、可继续"，也不该给「继续剩余动作」按钮 —— 发现阶段读了来源指纹，
     * 结论是"来源已变化"（阻断），而不是上一轮那种"显示可继续、点下去才被拒"。
     */
    expect(reopened.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(text(reopened)).not.toContain("来源和目标已核对");
    expect(text(reopened)).toContain(SOURCE_CHANGED);
    const undoButton = await reopened.findByRole("button", { name: UNDO_LABEL });
    // 重开只读了一次状态（不额外查询），也没有任何写入。
    expect(host.cell(SOURCE, "B5")).toBe(500);
    const writesBeforeUndo = contentWrites(host.originalWrites()).length;
    const registrations = vi.mocked(registerOfficeExecution).mock.calls.length;

    fireEvent.click(undoButton);
    await waitFor(() => expect(text(reopened)).toContain("已撤销本次执行"));
    // ①只回退**已写入**的目标：三张结果表 B2 都回到空（第 3 步从来没写过）。
    expect(RESULTS.map((name) => host.cell(name, "B2"))).toEqual(["", "", ""]);
    expect(
      contentWrites(host.originalWrites()).slice(writesBeforeUndo)
        .map((write) => `${write.sheet}!${write.address}`).sort()
    ).toEqual([`${RESULTS[0]}!B2`, `${RESULTS[1]}!B2`]);
    // ②来源保持用户现在的修改：撤销不碰来源表，也不悄悄改回 5。
    expect(host.cell(SOURCE, "B5")).toBe(500);
    expect(host.matrix(SOURCE, "A1:B5")).toEqual([
      ["影片", "票房(万)"], ["甲", 1], ["乙", 2], ["丙", 3], ["丁", 500]
    ]);
    // ③没有重新执行计划：一次登记都没有新增，也没有新的动作执行权/完成回执。
    expect(vi.mocked(registerOfficeExecution).mock.calls.length).toBe(registrations);
    expect(vi.mocked(startOfficeAction).mock.calls.length).toBe(2);
    expect(vi.mocked(completeOfficeAction).mock.calls.length).toBe(1);
    expect(finishOfficeExecution).not.toHaveBeenCalled();
    // ④留痕一次，且是「撤销」不是「放弃」。
    await waitFor(() => expect(vi.mocked(recordExecutionDisposition)).toHaveBeenCalledTimes(1));
    const [namespace, recordedId, disposition] = vi.mocked(recordExecutionDisposition).mock.calls[0];
    expect(namespace).toBe("office");
    expect(recordedId).toBe(operationId);
    expect(disposition).toBe("undone");
    // ⑤入口被消费：撤销完不再给按钮，只剩结论（终态摘要，在消息里可展开）。
    expect(reopened.queryByRole("button", { name: UNDO_LABEL })).toBeNull();
    expect(reopened.container.querySelector('[data-recovery-undo="done"]')).toBeNull();
    /*
     * 错误 #2 的回归：撤销成功是**终态**。同一张卡不得同时出现"已撤销本次执行"
     * （撤销行）与"可继续剩余动作"（结论 + 按钮）—— 上一轮真机正是两句话并存、
     * 按钮还能点。终态下「继续剩余动作」和「补记完成」都不得再出现。
     */
    expect(reopened.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(reopened.queryByRole("button", { name: "补记完成（不重写原表）" })).toBeNull();
    expect(reopened.queryByRole("button", { name: RECOVERY_CHECK_BUTTON })).toBeNull();
    expect(text(reopened)).toContain("已撤销本次执行");
    expect(text(reopened)).not.toContain("可继续剩余动作");
    /*
     * 终态不再占固定区域：可点的恢复卡（`data-recovery-kind="undone"` 那张面板）已经
     * 退出，结论改放进**所属任务的消息**里，一条可展开的终态摘要（`data-recovery-terminal`）。
     * 记录一条不删，只是不再伪装成"还有一件待办"。
     */
    expect(reopened.container.querySelector('[data-recovery-kind="undone"]')).toBeNull();
    expect(reopened.container.querySelector('[data-recovery-terminal="undone"]')).not.toBeNull();
    // 再触发一次渲染也不会撤第二遍。
    await act(async () => { await Promise.resolve(); });
    expect(host.cell(RESULTS[0], "B2")).toBe("");
    expect(vi.mocked(recordExecutionDisposition)).toHaveBeenCalledTimes(1);
  });

  it("旧任务已撤销 → 发送普通聊天 → 固定区域不再显示旧恢复卡，只剩消息里的可展开终态摘要", async () => {
    const { view, ledger } = await runToInterrupted();
    const operationId = ledger.current().operationId;
    await reconcileInFlight(view);
    // 走到终态：撤销（真实处理器），来源保持原样（不引入"来源变化"这条额外变量）。
    fireEvent.click(view.getByRole("button", { name: UNDO_LABEL }));
    await waitFor(() => expect(text(view)).toContain("已撤销本次执行"));

    // 终态退出固定区域：可点的恢复卡面板不再渲染，结论变成消息里一条可展开摘要。
    await waitFor(() =>
      expect(view.container.querySelector('[data-recovery-terminal="undone"]')).not.toBeNull()
    );
    expect(view.container.querySelector(".office-recovery-panel")).toBeNull();
    expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(view.queryByRole("button", { name: RECOVERY_CHECK_BUTTON })).toBeNull();
    const summary = view.container.querySelector('[data-recovery-terminal="undone"]')!;
    expect(summary.textContent).toContain("已撤销");
    expect(summary.textContent).toContain("展开");
    expect(summary.querySelector("summary")).not.toBeNull();

    // 记录一条不删：账本仍在、写前材料仍在（终态只是"退出固定区域"，不是清理）。
    expect(ledger.current().operationId).toBe(operationId);
    expect(contentWrites(host.originalWrites()).length).toBeGreaterThan(0);

    // 发一句普通聊天：固定区域**仍然**不显示旧恢复卡，摘要还留在它自己的那条消息里。
    const textarea = view.container.querySelector<HTMLTextAreaElement>(".composer textarea")!;
    fireEvent.change(textarea, { target: { value: "随便聊聊" } });
    await waitFor(() =>
      expect(view.container.querySelector(".composer .composer-primary")!.getAttribute("aria-label"))
        .toBe("发送")
    );
    expect(view.container.querySelector(".office-recovery-panel")).toBeNull();
    expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    // 终态摘要没有消失，也没有被重新摆回固定区域。
    expect(view.container.querySelector('[data-recovery-terminal="undone"]')).not.toBeNull();
    expect(view.container.querySelector(".office-recovery-panel")).toBeNull();
  });

  it("重开后目标被人工改过：撤销在任何写入之前停下，不留成功留痕，也不清空目标", async () => {
    const { view, ledger } = await runToInterrupted();
    const operationId = ledger.current().operationId;
    const before = await reconcileInFlight(view);
    view.unmount();
    localStorage.clear();
    vi.mocked(listPendingOfficeExecutions).mockResolvedValue({
      schemaVersion: 1,
      namespace: "office",
      workbookIdentity: `url:${BOOK_URL}`,
      limit: 20,
      truncated: false,
      candidates: [
        {
          operationId,
          taskId: "task-formula-undo",
          kind: "resume",
          reviewState: null,
          status: "needs_reconciliation",
          createdAt: 1_790_000_000_000,
          updatedAt: 1_790_000_600_000,
          nextActionIndex: 2,
          totalActions: 3,
          inFlightActionIndex: null,
          ownerActive: false,
          requiresReconciliation: false,
          canResume: true
        }
      ]
    } as never);

    const reopened = render(<App />);
    const undoButton = await reopened.findByRole("button", { name: UNDO_LABEL });
    /*
     * 执行之后、撤销之前，用户手改了第 1 步写进去的那一格。
     * 撤销的现场判据（`assertUndoPostState`）必须在**任何写入之前**发现它，而不是
     * 先按旧快照覆盖、再报"有点不对"。
     */
    host.setCell(RESULTS[0], "B2", "我手动改的");
    const writesBefore = contentWrites(host.originalWrites()).length;
    fireEvent.click(undoButton);
    await waitFor(() => expect(text(reopened)).toContain("又发生了变化"));

    // 零恢复写入：被改的那一格仍是用户的值，另一格仍是这次执行写下的公式。
    expect(contentWrites(host.originalWrites()).length).toBe(writesBefore);
    expect(host.cell(RESULTS[0], "B2")).toBe("我手动改的");
    expect(host.cell(RESULTS[1], "B2")).toBe(FORMULAS[1]);
    expect(host.cell(SOURCE, "B5")).toBe(5);
    // 没恢复成功 ⇒ 不写"撤销"留痕，也不宣称"已撤销"。
    expect(vi.mocked(recordExecutionDisposition)).not.toHaveBeenCalled();
    expect(text(reopened)).not.toContain("已撤销本次执行");
    /*
     * 入口**仍在**（不是"失败即消失"）：目标被改过时用户可能只是改错了一格，
     * 撤销栈里那一份留着，改回去再点就能重试。这里不把"按钮还在"当成失败证据，
     * 失败证据是"一格都没被覆盖"＋"没有成功留痕"。
     */
    expect(reopened.queryByRole("button", { name: UNDO_LABEL })).not.toBeNull();
    expect(contentWrites(host.originalWrites()).slice(before)).toEqual([]);
  });

  it("重开后授权被收窄：不给撤销按钮，只给原因，零写入", async () => {
    const { view, ledger } = await runToInterrupted();
    const operationId = ledger.current().operationId;
    await reconcileInFlight(view);
    view.unmount();
    localStorage.clear();
    vi.mocked(listPendingOfficeExecutions).mockResolvedValue({
      schemaVersion: 1,
      namespace: "office",
      workbookIdentity: `url:${BOOK_URL}`,
      limit: 20,
      truncated: false,
      candidates: [
        {
          operationId,
          taskId: "task-formula-undo",
          kind: "resume",
          reviewState: null,
          status: "needs_reconciliation",
          createdAt: 1_790_000_000_000,
          updatedAt: 1_790_000_600_000,
          nextActionIndex: 2,
          totalActions: 3,
          inFlightActionIndex: null,
          ownerActive: false,
          requiresReconciliation: false,
          canResume: true
        }
      ]
    } as never);

    const writesBeforeOpen = host.writes().length;
    const reopened = render(<App />);
    await reopened.findByRole("button", { name: RECOVERY_RESUME_BUTTON });
    // 用户重开后只勾了来源表并确认：撤销要写回的两张结果表已被收窄掉。
    await setAuthorizedScope(reopened, [SOURCE], SCOPE);
    await waitFor(() => expect(text(reopened)).toContain("当前授权里没有工作表"));
    expect(reopened.queryByRole("button", { name: UNDO_LABEL })).toBeNull();
    expect(reopened.container.querySelector('[data-recovery-undo="blocked"]')).not.toBeNull();
    expect(host.writes().length).toBe(writesBeforeOpen);
    expect(vi.mocked(recordExecutionDisposition)).not.toHaveBeenCalled();
  });

  it("重开后就地撤销又被处置过（账本已有 undone 留痕）：不再重建入口", async () => {
    const { view, ledger } = await runToInterrupted();
    const operationId = ledger.current().operationId;
    await reconcileInFlight(view);
    view.unmount();
    localStorage.clear();
    // 账本说这次执行已经撤销过（另一台机器上撤的）：入口不许再出现。
    ledger.current().disposition = { undoneAt: 1_790_000_999_000, abandonedAt: null };
    vi.mocked(listPendingOfficeExecutions).mockResolvedValue({
      schemaVersion: 1,
      namespace: "office",
      workbookIdentity: `url:${BOOK_URL}`,
      limit: 20,
      truncated: false,
      candidates: [
        {
          operationId,
          taskId: "task-formula-undo",
          kind: "resume",
          reviewState: null,
          status: "needs_reconciliation",
          createdAt: 1_790_000_000_000,
          updatedAt: 1_790_000_600_000,
          nextActionIndex: 2,
          totalActions: 3,
          inFlightActionIndex: null,
          ownerActive: false,
          requiresReconciliation: false,
          canResume: true
        }
      ]
    } as never);

    const writesBeforeOpen = host.writes().length;
    const reopened = render(<App />);
    await waitFor(() => expect(text(reopened)).toContain("这次执行已经撤销过"));
    expect(reopened.queryByRole("button", { name: UNDO_LABEL })).toBeNull();
    // 已经撤销过的执行也不该再被摆成"可继续"：两件事读的是同一条账本事实。
    expect(reopened.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(host.writes().length).toBe(writesBeforeOpen);
  });

  it("重开后账本没有写前快照：不给撤销，也不用当前单元格反推一个", async () => {
    const { view, ledger } = await runToInterrupted();
    const operationId = ledger.current().operationId;
    await reconcileInFlight(view);
    view.unmount();
    localStorage.clear();
    // 材料缺失：进度照旧是 2/3，但写前快照没落库。
    ledger.current().checkpoint = null;
    vi.mocked(listPendingOfficeExecutions).mockResolvedValue({
      schemaVersion: 1,
      namespace: "office",
      workbookIdentity: `url:${BOOK_URL}`,
      limit: 20,
      truncated: false,
      candidates: [
        {
          operationId,
          taskId: "task-formula-undo",
          kind: "resume",
          reviewState: null,
          status: "needs_reconciliation",
          createdAt: 1_790_000_000_000,
          updatedAt: 1_790_000_600_000,
          nextActionIndex: 2,
          totalActions: 3,
          inFlightActionIndex: null,
          ownerActive: false,
          requiresReconciliation: false,
          canResume: true
        }
      ]
    } as never);

    const writesBeforeOpen = host.writes().length;
    const reopened = render(<App />);
    await waitFor(() => expect(text(reopened)).toContain("不会用当前单元格反推"));
    expect(reopened.queryByRole("button", { name: UNDO_LABEL })).toBeNull();
    expect(host.writes().length).toBe(writesBeforeOpen);
    expect(host.cell(RESULTS[0], "B2")).toBe(FORMULAS[0]);
    expect(host.cell(RESULTS[1], "B2")).toBe(FORMULAS[1]);
  });

  it("来源恢复登记基线后：点「重新核对（只读）」按当前现场重新判断，不永久锁死", async () => {
    const { view, ledger } = await runToInterrupted();
    const operationId = ledger.current().operationId;
    await reconcileInFlight(view);
    // 把来源改掉再重开：重开后的发现必须**不**宣称"来源已核对、可继续"。
    host.setCell(SOURCE, "B5", 500);
    view.unmount();
    localStorage.clear();
    vi.mocked(listPendingOfficeExecutions).mockResolvedValue({
      schemaVersion: 1,
      namespace: "office",
      workbookIdentity: `url:${BOOK_URL}`,
      limit: 20,
      truncated: false,
      candidates: [
        {
          operationId,
          taskId: "task-formula-undo",
          kind: "resume",
          reviewState: null,
          status: "needs_reconciliation",
          createdAt: 1_790_000_000_000,
          updatedAt: 1_790_000_600_000,
          nextActionIndex: 2,
          totalActions: 3,
          inFlightActionIndex: null,
          ownerActive: false,
          requiresReconciliation: false,
          canResume: true
        }
      ]
    } as never);

    const reopened = render(<App />);
    // 阻断：没有"继续"，也没有"来源已核对"，只给只读重核。
    await waitFor(() => expect(text(reopened)).toContain(SOURCE_CHANGED));
    expect(reopened.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(text(reopened)).not.toContain("来源和目标已核对");
    const reverify = await reopened.findByRole("button", { name: "重新核对（只读）" });

    // 用户把来源改回登记基线，点「重新核对」：结论按**现在的现场**重新算，恢复可继续。
    host.setCell(SOURCE, "B5", 5);
    fireEvent.click(reverify);
    await waitFor(() =>
      expect(reopened.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).not.toBeNull()
    );
    // 复核全程零写入：来源改回去是用户自己做的，重核没有写任何单元格。
    expect(host.cell(SOURCE, "B5")).toBe(5);
  });

  it("撤销成功后迟到的发现/重核不把终态覆盖回可继续", async () => {
    const { view, ledger } = await runToInterrupted();
    const operationId = ledger.current().operationId;
    await reconcileInFlight(view);
    view.unmount();
    localStorage.clear();
    vi.mocked(listPendingOfficeExecutions).mockResolvedValue({
      schemaVersion: 1,
      namespace: "office",
      workbookIdentity: `url:${BOOK_URL}`,
      limit: 20,
      truncated: false,
      candidates: [
        {
          operationId,
          taskId: "task-formula-undo",
          kind: "resume",
          reviewState: null,
          status: "needs_reconciliation",
          createdAt: 1_790_000_000_000,
          updatedAt: 1_790_000_600_000,
          nextActionIndex: 2,
          totalActions: 3,
          inFlightActionIndex: null,
          ownerActive: false,
          requiresReconciliation: false,
          canResume: true
        }
      ]
    } as never);

    const reopened = render(<App />);
    const undoButton = await reopened.findByRole("button", { name: UNDO_LABEL });
    fireEvent.click(undoButton);
    await waitFor(() => expect(text(reopened)).toContain("已撤销本次执行"));
    expect(reopened.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();

    /*
     * 终态之后又发生一次"重新发现"（例如身份观察触发、或用户又点了一次只读重核）：
     * 迟到的响应**不得**把已经撤销的执行重新算回"可继续"。这里直接再跑一次发现，
     * 卡面必须仍是"已撤销"，且没有任何续跑/补记/再撤按钮。
     */
    await act(async () => {
      // 触发一次身份观察驱动的只读重核：换簿后身份观察会调 reverifySelectedOfficeRecovery。
      await Promise.resolve();
    });
    expect(reopened.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(reopened.queryByRole("button", { name: UNDO_LABEL })).toBeNull();
    expect(text(reopened)).toContain("已撤销本次执行");
    expect(text(reopened)).not.toContain("可继续剩余动作");
    // 终态后任何渲染都不该再产生写入。
    const writesAfter = contentWrites(host.originalWrites()).length;
    await act(async () => { await Promise.resolve(); });
    expect(contentWrites(host.originalWrites()).length).toBe(writesAfter);
  });

  it("回归⑥：撤销成功但留痕上报失败 → 本会话保持已撤销且不重复宿主操作；重开后无留痕也不重新开放续跑", async () => {
    const { view, ledger } = await runToInterrupted();
    const operationId = ledger.current().operationId;
    await reconcileInFlight(view);
    // 来源没改（B5=5），核对后卡面是「继续剩余动作」—— 这才是"撤销后会不会被重新开放"的对照起点。
    expect(await view.findByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeTruthy();

    // 撤销的留痕上报失败（网络不可用）：best-effort，不重试、不弹错、不改判。
    vi.mocked(recordExecutionDisposition).mockRejectedValueOnce(new Error("网络不可用"));
    const writesBeforeUndo = contentWrites(host.originalWrites()).length;
    const registrations = vi.mocked(registerOfficeExecution).mock.calls.length;
    const claims = vi.mocked(claimOfficeExecutionResume).mock.calls.length;

    fireEvent.click(view.getByRole("button", { name: UNDO_LABEL }));
    await waitFor(() => expect(text(view)).toContain("已撤销本次执行"));

    // ① 留痕确实只上报一次（best-effort 不重试），且是 undone。
    await waitFor(() => expect(vi.mocked(recordExecutionDisposition)).toHaveBeenCalledTimes(1));
    expect(vi.mocked(recordExecutionDisposition).mock.calls[0][2]).toBe("undone");
    // ② 宿主恢复只发生一次：撤销写回两个目标格，之后没有新的登记 / 领取执行权 / 执行。
    expect(contentWrites(host.originalWrites()).slice(writesBeforeUndo)
      .map((write) => `${write.sheet}!${write.address}`).sort())
      .toEqual([`${RESULTS[0]}!B2`, `${RESULTS[1]}!B2`]);
    expect(vi.mocked(registerOfficeExecution).mock.calls.length).toBe(registrations);
    expect(vi.mocked(claimOfficeExecutionResume).mock.calls.length).toBe(claims);
    // ③ 本会话保持已撤销：终态退出固定区域，没有继续/补记/再撤，只剩消息里的终态摘要。
    expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(view.container.querySelector('[data-recovery-terminal="undone"]')).not.toBeNull();
    expect(view.container.querySelector('[data-recovery-kind="undone"]')).toBeNull();

    // 重开窗格：本机终态 ref/state 随组件销毁，账本又因为上报失败没有 undoneAt 留痕。
    view.unmount();
    localStorage.clear();
    vi.mocked(listPendingOfficeExecutions).mockResolvedValue({
      schemaVersion: 1,
      namespace: "office",
      workbookIdentity: `url:${BOOK_URL}`,
      limit: 20,
      truncated: false,
      candidates: [
        {
          operationId,
          taskId: "task-formula-undo",
          kind: "resume",
          reviewState: null,
          status: "needs_reconciliation",
          createdAt: 1_790_000_000_000,
          updatedAt: 1_790_000_600_000,
          nextActionIndex: 2,
          totalActions: 3,
          inFlightActionIndex: null,
          ownerActive: false,
          requiresReconciliation: false,
          canResume: true
        }
      ]
    } as never);

    const reopened = render(<App />);
    /*
     * 账本没有处置留痕，但**现场已经被撤销清空**：重开后靠写后现场核对（`verifyResumeScene`）
     * 发现第 1、2 步的目标不再是执行后的样子 ⇒ 阻断续跑，而不是重新开放「继续剩余动作」。
     */
    await waitFor(() =>
      expect(text(reopened)).toContain("已经不是这次执行写入后的样子")
    );
    expect(reopened.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(reopened.queryByRole("button", { name: "补记完成（不重写原表）" })).toBeNull();
    expect(reopened.queryByRole("button", { name: UNDO_LABEL })).toBeNull();
    // 重开后没有任何新写入。
    expect(contentWrites(host.originalWrites()).length).toBe(writesBeforeUndo + 2);
  });

  it("回归⑧：卡片生成后账本处置变成「已撤销」，点原「继续剩余动作」入口 → 处理器按最新处置拒绝，零写入", async () => {
    const { view, ledger } = await runToInterrupted();
    const operationId = ledger.current().operationId;
    await reconcileInFlight(view);
    // 卡片此刻是「继续剩余动作」（来源 matched）。保存这个入口元素本身（= 旧回调的持有者）。
    const resumeButton = await view.findByRole("button", { name: RECOVERY_RESUME_BUTTON });

    /*
     * 在点击之前，账本处置被推进成「已撤销」（例如另一台机器 / 另一窗格已经撤销并留痕）。
     * 按钮还是那一颗（旧回调还在），但 `resumeOfficeRemainingActions` 必须先按最新账本
     * 读数复核处置，拒绝续跑 —— 而不是拿旧卡面的结论直接去 claim + 执行。
     */
    ledger.current().disposition = { undoneAt: 1_790_000_700_000, abandonedAt: null };
    const claims = vi.mocked(claimOfficeExecutionResume).mock.calls.length;
    const writesBefore = contentWrites(host.originalWrites()).length;

    fireEvent.click(resumeButton);
    await waitFor(() => expect(text(view)).toContain("已经撤销过（账本已留处置痕）"));
    // 处理器拒绝：没有领取执行权，也没有任何写入。
    expect(vi.mocked(claimOfficeExecutionResume).mock.calls.length).toBe(claims);
    expect(contentWrites(host.originalWrites()).length).toBe(writesBefore);
    // 卡面收敛到终态，继续/补记/再撤都不给（终态退出固定区域，只剩消息里的摘要）。
    expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(view.container.querySelector('[data-recovery-terminal="undone"]')).not.toBeNull();
    expect(view.container.querySelector('[data-recovery-kind="undone"]')).toBeNull();
  });

  it("已撤销终态不出现续跑引导句，也不给任何写入按钮（说明按状态派生）", async () => {
    /*
     * 2026-10-31 真机：已撤销的卡上，"已撤销本次执行"下面仍跟着写死的
     * `EXECUTION_RECOVERY_NOTE`（"续跑通过复核后执行剩余动作，再审阅结果"）——与终态
     * 矛盾。这里钉住渲染层：`undone` 结论下卡面**不再**出现这句续跑引导，也不给写入按钮。
     */
    const { view, ledger } = await runToInterrupted();
    const operationId = ledger.current().operationId;
    await reconcileInFlight(view);
    // 撤销前：操作区是「继续剩余 N 步」，续跑引导句收在「查看详情」里（技术口径）。
    expect(await view.findByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeTruthy();
    fireEvent.click(view.getByRole("button", { name: "查看详情" }));
    expect(text(view)).toContain("续跑通过复核后执行剩余动作");

    // 撤销：账本处置推进成已撤销（等价于撤销成功后留痕），卡面收敛到 undone 终态。
    ledger.current().disposition = { undoneAt: 1_790_000_700_000, abandonedAt: null };
    fireEvent.click(view.getByRole("button", { name: RECOVERY_RESUME_BUTTON }));
    await waitFor(() =>
      expect(view.container.querySelector('[data-recovery-terminal="undone"]')).not.toBeNull()
    );

    // 终态说明替换续跑引导：不再出现"续跑"引导句或审阅措辞。
    expect(text(view)).not.toContain("续跑通过复核后执行剩余动作");
    expect(text(view)).toContain("已撤销本次执行");
    // 不给任何写入/补记/继续按钮。
    expect(view.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    expect(view.queryByRole("button", { name: "补记完成（不重写原表）" })).toBeNull();
    expect(view.queryByRole("button", { name: UNDO_LABEL })).toBeNull();
  });

  /**
   * 2026-11-03 的主线缺陷：写入完成、回读通过，用户还没点「确认全部变更」就刷新窗格 ——
   * 待审阅界面消失、输入框回到普通状态，甚至写「本次没有需要你处理的后续操作」，
   * 而原表停在一个没人收尾的写入现场上。
   *
   * 修法在账本里（`review_state`）与发现里（`kind: "review"`）：重开时按**显式审阅状态**
   * 找回审阅，走 `adoptCompleted`（那条路上没有 `prepare`/`execute`，所以重开一次不会
   * 多写一次）。这条用例钉住三件事：入口回来了、内容没被重写、确认只清标记。
   */
  it.each(['confirm', 'discard', 'changed'] as const)("执行完成但没处理审阅 → 重开窗格后审阅入口回来，且一个单元格都不再写（%s）", async outcome => {
    const { view, ledger } = await runToInterrupted();
    const operationId = ledger.current().operationId;
    await reconcileInFlight(view);
    const writesAfterResume = contentWrites(host.originalWrites()).length;
    const snapshot = ledger.current();

    /*
     * 账本此刻的事实：三个动作都写完了、回读通过、`review_state = pending`
     * （服务端把审阅初始状态与完成回执写在**同一条 UPDATE** 里）。
     *
     * 这里只补这份**账本事实**，不补任何单元格内容、也不触碰宿主：被测的是
     * "重开时能不能凭账本把审阅找回来"，不是"能不能把表再写一遍"。
     */
    Object.assign(snapshot, {
      status: "completed",
      nextActionIndex: 3,
      inFlightActionIndex: null,
      requiresReconciliation: false,
      canResume: false,
      review: { state: "pending", decidedAt: 1_790_000_000_000 },
      result: { verification: { status: "verified" }, review: [], undoSnapshot: snapshot.checkpoint }
    });
    const actionCalls = {
      start: vi.mocked(startOfficeAction).mock.calls.length,
      complete: vi.mocked(completeOfficeAction).mock.calls.length,
      finish: vi.mocked(finishOfficeExecution).mock.calls.length
    };

    view.unmount();
    localStorage.clear();
    expect(loadPendingOfficeExecution()).toBeNull();
    /*
     * 再放一份**上一次审阅遗留的标记归属记录**：挂载时那段"认领并清理"会照着它
     * 去删条件格式。这一步只允许清展示标记 —— 它既不是"用户确认过"，也不许顺手
     * 把审阅状态写掉（见下面 `sendExecutionReview` 在点确认之前没被调用）。
     */
    localStorage.setItem("excel-bro.reviewMarks.v1", JSON.stringify({
      v: 1,
      identity: `url:${BOOK_URL}`,
      formula: '=N("EB_REVIEW_0000000000000000000")=0',
      entries: [{ sheet: RESULTS[0], range: "B2", id: "rule-legacy-review-mark" }]
    }));
    vi.mocked(listPendingOfficeExecutions).mockResolvedValue({
      schemaVersion: 1,
      namespace: "office",
      workbookIdentity: `url:${BOOK_URL}`,
      limit: 20,
      truncated: false,
      candidates: [
        {
          operationId,
          taskId: "task-review-restore",
          kind: "review",
          status: "completed",
          reviewState: "pending",
          createdAt: 1_790_000_000_000,
          updatedAt: 1_790_000_600_000,
          nextActionIndex: 3,
          totalActions: 3,
          inFlightActionIndex: null,
          ownerActive: false,
          requiresReconciliation: false,
          // 待审阅**没有**续跑资格：一步都不能再写。
          canResume: false
        }
      ]
    } as never);

    const reopened = render(<App />);
    // 入口回来了（不是"本次没有需要你处理的后续操作"）。
    await stageReview(reopened);
    expect(text(reopened)).not.toContain("没有需要你处理的后续操作");
    // 恢复审阅**没有**执行任何动作、也没有再提交一次完成回执。
    expect(contentWrites(host.originalWrites())).toHaveLength(writesAfterResume);
    expect(vi.mocked(startOfficeAction).mock.calls.length).toBe(actionCalls.start);
    expect(vi.mocked(completeOfficeAction).mock.calls.length).toBe(actionCalls.complete);
    expect(vi.mocked(finishOfficeExecution).mock.calls.length).toBe(actionCalls.finish);
    // 也不给续跑入口（那是"待续写"，不是"待审阅"）。
    expect(reopened.queryByRole("button", { name: RECOVERY_RESUME_BUTTON })).toBeNull();
    /*
     * 启动时那次"清理遗留标记"**没有**顺手确认：审阅还在，决定还没落账。
     * 这是"清理标记 ≠ 用户确认"的可观测证据（标记清理本身是展示层收尾）。
     */
    expect(vi.mocked(sendExecutionReview)).not.toHaveBeenCalled();

    if (outcome !== 'confirm') {
      if (outcome === 'changed') host.setCell(RESULTS[0], 'B2', '用户后续修改');
      const beforeDiscard = contentWrites(host.originalWrites()).length;
      fireEvent.click(reopened.getByRole('button', {name:'放弃全部变更'}));
      if (outcome === 'discard') {
        await waitFor(() => expect(sendExecutionReview).toHaveBeenCalledWith('office', operationId, 'discarded'));
        expect(RESULTS.map(name => host.cell(name, 'B2'))).toEqual(['','','']);
        expect(contentWrites(host.originalWrites()).length).toBeGreaterThan(beforeDiscard);
        expect(reopened.queryByRole('button', {name:'放弃全部变更'})).toBeNull();
      } else {
        await waitFor(() => expect(text(reopened)).toContain('未能放弃本次变更'));
        expect(host.cell(RESULTS[0], 'B2')).toBe('用户后续修改');
        expect(host.cell(RESULTS[1], 'B2')).toBe(FORMULAS[1]);
        expect(contentWrites(host.originalWrites())).toHaveLength(beforeDiscard);
        expect(sendExecutionReview).not.toHaveBeenCalled();
      }
      expect(vi.mocked(startOfficeAction).mock.calls.length).toBe(actionCalls.start);
      expect(vi.mocked(finishOfficeExecution).mock.calls.length).toBe(actionCalls.finish);
      return;
    }

    fireEvent.click(reopened.getByRole("button", { name: "确认全部变更" }));
    await waitFor(() =>
      expect(vi.mocked(sendExecutionReview)).toHaveBeenCalledWith("office", operationId, "confirmed")
    );
    // 确认只接受已有结果：没有新写入，也没有重跑计划。
    expect(contentWrites(host.originalWrites())).toHaveLength(writesAfterResume);
    expect(vi.mocked(startOfficeAction).mock.calls.length).toBe(actionCalls.start);
    expect(vi.mocked(finishOfficeExecution).mock.calls.length).toBe(actionCalls.finish);
    // 单元格里还是第一次写进去的那条公式（没被重写成值，也没写第二遍）。
    expect(host.matrix(RESULTS[0], "B2")).toEqual([[FORMULAS[0]]]);
  });
});
