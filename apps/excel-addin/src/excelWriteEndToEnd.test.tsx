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
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import { checkIntent, streamAssistantResponse, understandMessage } from "./api";
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
 */
function targetFor(names: readonly string[] | null) {
  const sheets = ALL_SHEETS.map((name) => ({ id: SHEET_IDS[name], name }));
  const chosen = names === null ? [SHEET] : [...names];
  return {
    fileKey: "office:book-1",
    name: "book.xlsx",
    sheets,
    selected: chosen.map((name) => SHEET_IDS[name]).filter(Boolean)
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
  pollSdkBridgeTool: vi.fn(),
  postSdkBridgeReceipt: vi.fn(),
  // 本机窗口排列服务：替身不模拟它，明确拒绝（产品会如实降级为"手动排列"）。
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
  confirmOfficeAuthorizedScope: vi.fn(async (_names, confirm, expected) => confirm(expected)),
  officeAnalysisTargetEvidence: vi.fn(async () => ({ key: "target-1", conclusive: true }))
}));

const BOOK_URL = "https://contoso.sharepoint.com/book.xlsx";
let host: TestOfficeHost;

function createHost(): TestOfficeHost {
  return createOfficeTestHost({
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
  fireEvent.change(view.container.querySelector(".composer textarea")!, {
    target: { value: prompt }
  });
  fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
  await waitFor(() =>
    expect(view.container.querySelector(".inline-plan")).not.toBeNull()
  );
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
  for (const name of [SHEET, TARGET]) {
    const button = sheetToggle(view, name);
    const selected = button.getAttribute("aria-pressed") === "true";
    if (selected !== names.includes(name)) fireEvent.click(button);
  }
  const commit = view.container.querySelector<HTMLButtonElement>(".scope-confirm");
  if (!commit) throw new Error("范围「确认」按钮尚未出现");
  await waitFor(() => expect(commit.disabled).toBe(false));
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
  await waitFor(() => expect(view.container.querySelector(".result-review-bar")?.textContent).toContain("已确认变更"));
  return save;
}
const text = (view: ReturnType<typeof render>) => view.container.textContent ?? "";

describe("主表执行结果审阅（真实 App/执行器，宿主与网络替身）", () => {
  it("主表立即显示结果，默认不创建对比工作簿", async () => {
    const view = await sendPlan(reviewPlan);
    await stageReview(view);
    expect(host.cell(SHEET, "G1")).toBe("合计");
    expect(contentWrites(host.originalWrites()).map(w => w.address)).toEqual(["E1:F3", "G1"]);
    expect(host.createWorkbookCalls()).toHaveLength(0);
    expect(view.container.querySelector(".composer-footer .result-review-bar") ?? view.container.querySelector(".result-review-bar")).not.toBeNull();
    expect(view.container.querySelector("button[title='保存为可重复使用的个人工具']")).toBeNull();
  });
  it("双视图的文件种子来自执行前，不重新执行计划", async () => {
    const view = await sendPlan(reviewPlan); await stageReview(view);
    const writes = contentWrites(host.originalWrites()).length;
    fireEvent.click(view.getByRole("button", { name: "双视图对比" }));
    await waitFor(() => expect(host.createWorkbookCalls()).toHaveLength(1));
    const xml = host.seedEntries(0)["xl/worksheets/sheet1.xml"];
    expect(xml).toContain("票房");
    expect(xml).not.toContain("上座率");
    expect(xml).not.toContain("合计");
    expect(contentWrites(host.originalWrites())).toHaveLength(writes);
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
    await clickRunOnExcel(view);

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
    await clickRunOnExcel(view);

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
    await clickRunOnExcel(view);

    await waitFor(() =>
      expect(text(view)).toContain("服务端没有授予本次 Office 写入权")
    );
    expect(registerOfficeExecution).toHaveBeenCalledTimes(1);
    expect(host.writes()).toEqual([]);
  });

  it("挂接索引无法落盘：抛错并停止，原表零写入", async () => {
    const view = await sendPlan(targetPlan, "把合计写到汇总表", {
      authorize: SCOPE,
      read: { scope: "active" }
    });
    // 挂载期间不要拦 localStorage（工具库自己也在读写），只在本步写盘前拦。
    const setItem = vi
      .spyOn(Storage.prototype, "setItem")
      .mockImplementation(() => {
        throw new Error("quota exceeded");
      });
    try {
      await clickRunOnExcel(view);

      await waitFor(() =>
        expect(text(view)).toContain("无法保存 Office 任务挂接索引")
      );
      expect(host.writes()).toEqual([]);
    } finally {
      setItem.mockRestore();
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
    await clickRunOnExcel(view);

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
    await clickRunOnExcel(view);

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
    await clickRunOnExcel(view);

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
      nextActionIndex: 1,
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

  it("在途动作先核对为「已写入」：跳过该步，只写剩余步骤，不重复写", async () => {
    // 头一步已经真的写进宿主了（回执丢失），第二部还没写。
    host.setCell(TARGET, "A1", "=\"影片\"", { asFormula: true });
    host.setCell(TARGET, "B1", "=1200", { asFormula: true });
    const firstWrites = host.originalWrites().length;

    pendingIndex();
    vi.mocked(getOfficeExecutionStatus).mockResolvedValue(resumeState() as never);
    vi.mocked(reconcileOfficeActionReceipt).mockResolvedValue(
      resumeState({ inFlightActionIndex: null, requiresReconciliation: false, revision: 3 }) as never
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
