import * as historyApi from "./api";
// @vitest-environment jsdom
import { act, cleanup, fireEvent, render, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import { createOfficeTestHost } from "./officeTestHost";
import { ApiRequestError, checkIntent, understandMessage, streamAssistantResponse, getPlanningTaskStatus, getPlanningTaskEvents, continuePlanningTask, pausePlanningTask } from "./api";
import { captureWorkbookCatalog, captureWorkbookStructure } from "./excel";
import { inspectOfficeAuthorizedTarget } from "./officeAuthorizedTarget";

/**
 * 回复段快照：每段过程卡的**文本**与**步骤行文本**。
 *
 * 「继续之后旧段不被改写」只能靠逐字比较来证明：只看"有几张卡"无法区分
 * "历史段被冻结"和"历史段被就地改写后又复制了一份"。
 */
function segmentSnapshots(container: HTMLElement) {
  return [...container.querySelectorAll<HTMLElement>(".task-progress-timeline")].map(
    (card) => ({
      text: card.textContent ?? "",
      steps: [...card.querySelectorAll<HTMLElement>(".activity-steps > li")].map(
        (row) => row.textContent ?? ""
      )
    })
  );
}

/** 当前活动段的数量：同一任务同一时刻最多一个。 */
function liveSegmentCount(container: HTMLElement): number {
  return container.querySelectorAll(".activity-card.is-live").length;
}

/** 第 index 段过程卡所在的 `.message-row` 在消息流里的位置。 */
function segmentRowIndex(container: HTMLElement, index: number): number {
  const rows = [...container.querySelectorAll<HTMLElement>(".message-row")];
  const card = container.querySelectorAll(".task-progress-timeline")[index];
  return rows.findIndex((row) => card && row.contains(card));
}

/** 最后一条用户消息所在的位置。 */
function lastUserRowIndex(container: HTMLElement): number {
  const rows = [...container.querySelectorAll<HTMLElement>(".message-row.user")];
  const all = [...container.querySelectorAll<HTMLElement>(".message-row")];
  const last = rows[rows.length - 1];
  return last ? all.indexOf(last) : -1;
}

const fixtures = vi.hoisted(() => ({  order: [] as string[],
  snapshot: { name: "book.xlsx", capturedAt: "2026-10-08T00:00:00Z", activeWorksheet: "Sheet1", worksheets: [
    { name: "Sheet1", usedRange: "A1:B3", rowCount: 3, columnCount: 2, headers: ["Category", "Level"], dataRows: [], truncated: false }
  ] },
  target: { fileKey: "office:book-1", name: "book.xlsx", sheets: [{ id: "sheet-1", name: "Sheet1" }], selected: ["sheet-1"] }
}));

vi.mock("./api", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./api")>();
  const { guardNetworkSeam } = await import("@excel-bro/core/testNetworkGate");
  return {
  ...actual,
  checkHealth: vi.fn(async () => ({ status: "ok", modelConfigured: true })),
  listModels: vi.fn(async () => ({ defaultModelId: "model-1", models: [{ id: "model-1", label: "模型", modelId: "test-model", baseUrl: "http://localhost", supportsVision: true, available: true }] })),
  getModelSettings: vi.fn(async () => ({ connectionCount: 1 })),
  understandMessage: vi.fn(), checkIntent: vi.fn(), streamAssistantResponse: vi.fn(), getPlanningTaskStatus: vi.fn(), continuePlanningTask: vi.fn(),
  /*
   * 暂停接口默认**没有服务端**（拒绝），与原先"真 fetch 在 jsdom 里必然失败"同效。
   * 需要断言三态（acknowledged / local-stopped / unconfirmed）的用例自己接管它。
   */
  pausePlanningTask: vi.fn(),
  /*
   * 只读补发端点（「原任务仍在运行」时接回进度）。
   *
   * 默认**读不到**：绝大多数用例走到"仍在运行"就该停下，不偷偷发真请求。
   * 需要这条路径的用例自己 `mockResolvedValueOnce` 提供分页。
   */
  getPlanningTaskEvents: vi.fn(async () => { throw new TypeError("Failed to fetch"); }),
  /*
   * 只读轮询桥（`POST /api/sdk/tool-bridge/poll`）：接回原任务进度时会问一次有没有
   * 待执行的读表指针。以前没有替身 —— 这条请求真打到 `127.0.0.1:8765`，用例却断言
   * "宿主写入 0"。这条接缝有几条用例要**故意**跑真实实现（自己 stub fetch 驱动轮询），
   * 所以用 `guardNetworkSeam`：本用例装了 fetch 替身才真的调用，否则按"连不上"失败。
   */
  pollSdkBridgeTool: vi.fn(guardNetworkSeam(actual.pollSdkBridgeTool))
  };
});
vi.mock("./excel", async (importOriginal) => ({
  ...await importOriginal<typeof import("./excel")>(),
  isRunningInExcel: () => true,
  captureWorkbookCatalog: vi.fn(async () => ({ ...fixtures.snapshot, catalogOnly: true, worksheets: fixtures.snapshot.worksheets.map((sheet) => ({ ...sheet, headers: [], dataRows: [] })) })),
  captureWorkbookStructure: vi.fn(async () => { fixtures.order.push("schema"); return fixtures.snapshot; }),
  captureWorkbook: vi.fn(async () => fixtures.snapshot),
  watchWorkbookStructureChanges: vi.fn(async () => () => {}),
  currentOfficeWorkbookIdentity: () => "office:book-1"
}));
vi.mock("./officeAuthorizedTarget", async (importOriginal) => ({
  ...await importOriginal<typeof import("./officeAuthorizedTarget")>(),
  inspectOfficeAuthorizedTarget: vi.fn(async () => { fixtures.order.push("authorization"); return fixtures.target; }),
  inspectOfficeScopeAdmission: vi.fn(async (names: readonly string[]) => ({
    items: names.map((id) => ({ id, status: "ready", reason: "none" })),
    target: await inspectOfficeAuthorizedTarget(names)
  })),
  confirmOfficeAuthorizedScope: vi.fn(async (_names, confirm, expected) => confirm(expected)),
  officeAnalysisTargetEvidence: vi.fn(async () => ({ key: "target-1", conclusive: true }))
}));

/**
 * 读表桥外面套一层**可控返回值**（2026-10-20）。
 *
 * 桥自己的行为由 `hooks/useSdkToolBridge.test.tsx` 的 23 条用例与四条变异证明，
 * 这里要证的是**调用层**：面板拿到 `false` 之后到底还发不发请求。
 *
 * 为什么必须可控：`begin`/`ensureLive` 返回 false 只发生在"初始化（读宿主）期间
 * 被取消/收尾/被另一轮取代"这一瞬间，靠真实时序在 jsdom 里碰上它等于靠运气 ——
 * 而这是"继续分析之后一动不动"最贵的一条路径，不能用"碰不到就没问题"来收口。
 *
 * 替身只在**开关打开时**替换返回值，其余一律转发给真实 hook（同一个实例、同一份
 * 状态），所以其他用例跑的仍是完整实现。
 */
const bridgeStub = vi.hoisted(() => ({ begin: false, ensureLive: false }));

vi.mock("./hooks/useSdkToolBridge", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./hooks/useSdkToolBridge")>();
  const React = await import("react");
  type Controller = ReturnType<typeof actual.useSdkToolBridge>;
  return {
    ...actual,
    useSdkToolBridge: (dependencies?: Parameters<typeof actual.useSdkToolBridge>[0]) => {
      const real = actual.useSdkToolBridge(dependencies);
      const latest = React.useRef<Controller>(real);
      latest.current = real;
      return React.useMemo(
        () =>
          new Proxy({} as Controller, {
            get(_target, key: string) {
              if (key === "begin" && bridgeStub.begin) return async () => false;
              if (key === "ensureLive" && bridgeStub.ensureLive) return async () => false;
              const value = (latest.current as unknown as Record<string, unknown>)[key];
              return typeof value === "function" ? value.bind(latest.current) : value;
            }
          }),
        []
      );
    }
  };
});

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
  bridgeStub.begin = false;
  bridgeStub.ensureLive = false;
  vi.spyOn(console, "log").mockImplementation(() => {});
  vi.mocked(getPlanningTaskEvents).mockReset();
  vi.mocked(getPlanningTaskEvents).mockRejectedValue(new TypeError("Failed to fetch"));
  vi.mocked(pausePlanningTask).mockReset();
  vi.mocked(pausePlanningTask).mockRejectedValue(new TypeError("Failed to fetch"));
  vi.mocked(inspectOfficeAuthorizedTarget).mockImplementation(async () => { fixtures.order.push("authorization"); return fixtures.target; });
  fixtures.order = [];
  vi.stubGlobal("Office", { context: { document: { url: "file:///C:/Data/book.xlsx" } }, onReady: (cb: () => void) => cb() });
  Element.prototype.scrollIntoView = vi.fn();
});
afterEach(() => { cleanup(); vi.unstubAllGlobals(); vi.restoreAllMocks(); });

async function send(text: string) {
  const view = render(<App />);
  await waitFor(() => expect(vi.mocked(captureWorkbookCatalog)).toHaveBeenCalled());
  await waitFor(() => expect(view.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title).toContain("当前模型"));
  fixtures.order = [];
  vi.mocked(inspectOfficeAuthorizedTarget).mockClear();
  fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: text } });
  fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
  return view;
}

describe("Excel actual composer: message first", () => {
  it("opens EB utilities from the composer logo and places recovery and history beside new conversation without a header", async () => {
    const view = render(<App />);
    expect(view.container.querySelector('.chat-header')).toBeNull();
    expect(view.container.querySelector('.scope-entry-row .folder-backup-entry')).not.toBeNull();
    expect(view.container.querySelector('.chat-header .brand-mark')).toBeNull();
    expect(view.container.querySelector('.attach-image-button')).toBeNull();
    const logo = view.getByRole('button', { name: 'EB 菜单' });
    fireEvent.click(logo);
    expect([...view.container.querySelectorAll('.composer-app-menu .view-menu > button > strong')].slice(0, 5).map(item => item.textContent)).toEqual(['工具', '宠物', '壁纸', 'EB 函数说明', '专注窗口']);
    fireEvent.click(view.getByRole('menuitemcheckbox'));
    fireEvent.click(logo);
    expect(view.getByRole('menuitemcheckbox').getAttribute('aria-checked')).toBe('false');
    fireEvent.click(view.getByRole('menuitem', { name: /^壁纸/ }));
    expect(document.querySelector('.appearance-dialog')).not.toBeNull();
    expect(document.querySelector('.appearance-preview')).not.toBeNull();
    fireEvent.click(view.getByRole('button', { name: '关闭壁纸设置' }));
    expect(document.querySelector('.appearance-dialog')).toBeNull();
    expect(vi.mocked(understandMessage)).not.toHaveBeenCalled();
  });

  it("keeps the model picker in the toolbar and displays the configured model ID when switching", async () => {
    const models = [
      { id: "connection:one", modelId: "model-one", label: "连接一", provider: "model" as const, available: true, supportsVision: true },
      { id: "connection:two", modelId: "model-two", label: "连接二", provider: "model" as const, available: true, supportsVision: false }
    ];
    vi.mocked(historyApi.listModels).mockResolvedValueOnce({ defaultModelId: models[0].id, models });
    const view = render(<App />);
    const picker = () => view.container.querySelector<HTMLButtonElement>(".composer-toolbar .model-picker-trigger");
    await waitFor(() => expect(picker()?.textContent).toBe("model-one"));
    expect(view.container.querySelector(".chat-header .model-picker")).toBeNull();
    fireEvent.click(picker()!);
    fireEvent.click(view.getByRole("menuitemradio", { name: "连接二" }));
    expect(picker()?.textContent).toBe("model-two");
    expect(localStorage.getItem("excel-bro.model.v2")).toBe("connection:two");
    expect(vi.mocked(understandMessage)).not.toHaveBeenCalled();
  });

  it.each(["openai", "anthropic"])("routes %s connections despite legacy Agent capability metadata", async (protocol) => {
    const model = { id: `connection:generic-${protocol}`, modelId: "custom-model", label: "自定义接口", provider: "model" as const,
      available: true, supportsVision: false, agentAvailable: false, formulaAvailable: false };
    vi.mocked(historyApi.listModels).mockResolvedValueOnce({ defaultModelId: model.id, models: [model] });
    vi.mocked(understandMessage).mockResolvedValue({ kind: "conversation", message: "连接已响应", provider: "model" });
    const view = render(<App />);
    await waitFor(() => expect(view.container.querySelector(".model-picker-trigger")?.textContent).toBe("custom-model"));
    fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "你好" } });
    fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
    await waitFor(() => expect(view.container.textContent).toContain("连接已响应"));
    expect(vi.mocked(understandMessage).mock.calls[0][0].modelId).toBe(model.id);
    expect(streamAssistantResponse).not.toHaveBeenCalled();
  });

  it("history alignment: cancel/confirm deletion preserves other guards and navigation never writes", async () => {
    const host = createOfficeTestHost({ url: "file:///C:/Data/book.xlsx", sheets: [{ name: "Sheet1", values: [["合成数据"]] }] }); vi.stubGlobal("Excel", host.Excel);
    const mutations = [
      vi.spyOn(historyApi, "registerOfficeExecution"), vi.spyOn(historyApi, "startOfficeAction"),
      vi.spyOn(historyApi, "completeOfficeAction"), vi.spyOn(historyApi, "finishOfficeExecution"),
      vi.spyOn(historyApi, "claimOfficeExecutionResume"), vi.spyOn(historyApi, "reconcileOfficeActionReceipt"),
      vi.spyOn(historyApi, "recordExecutionDisposition"), vi.spyOn(historyApi, "deleteExecutionRecoveryPackage")
    ];
    for (const mutation of mutations) mutation.mockRejectedValue(new Error("History must never mutate execution"));
    const key = "excel-bro.chat.v4";
    localStorage.setItem(key, JSON.stringify({ activeConversationId: "alpha", legacyCredentialConversationIds: ["guarded"],
      conversations: ["alpha", "beta", "guarded"].map((id) => ({ id, title: id,
        createdAt: "2026-10-01T00:00:00Z", updatedAt: "2026-10-01T00:00:00Z",
        messages: [{ id: `m-${id}`, role: "user", text: `普通聊天 ${id}`, createdAt: "2026-10-01T00:00:00Z" }] })) }));
    const view = render(<App />);
    await waitFor(() => expect(view.container.textContent).toContain("普通聊天 alpha"));
    fireEvent.click(view.getByRole("button", { name: "历史对话" }));
    const saved = () => JSON.parse(localStorage.getItem(key)!);
    fireEvent.click(view.getByRole("button", { name: "删除「beta」" }));
    fireEvent.click(view.getByRole("button", { name: /^取消$/ }));
    expect(saved().conversations.map((c: {id: string}) => c.id)).toEqual(["alpha", "beta", "guarded"]);
    fireEvent.click(view.getByRole("button", { name: "删除「beta」" }));
    fireEvent.click(view.getByRole("button", { name: /^确认删除$/ }));
    expect(saved().conversations.map((c: {id: string}) => c.id)).toEqual(["alpha", "guarded"]);
    expect(saved().legacyCredentialConversationIds).toEqual(["guarded"]);
    fireEvent.click([...view.container.querySelectorAll<HTMLButtonElement>(".history-open-button")].find((b) => b.textContent?.includes("guarded"))!);
    expect(view.container.querySelector(".message-stream")?.textContent).toContain("普通聊天 guarded");
    expect(saved().legacyCredentialConversationIds).toEqual(["guarded"]);
    expect(host.writes()).toHaveLength(0);
    for (const mutation of mutations) expect(mutation).not.toHaveBeenCalled();
    expect(understandMessage).not.toHaveBeenCalled(); expect(streamAssistantResponse).not.toHaveBeenCalled();
  });

  it.each([
    ["MODEL_QUOTA_EXHAUSTED", "模型 API 余额不足或调用额度已用尽", "模型设置"],
    ["MODEL_CONTEXT_LIMIT", "上下文长度上限", "拆分任务"],
    ["MODEL_OUTPUT_LIMIT", "未获得完整结果", "输出上限"]
  ])("shows actionable model limit %s and retains the request", async (code, title, action) => {
    vi.mocked(understandMessage).mockRejectedValueOnce(new ApiRequestError(502, code, "PRIVATE_PROVIDER_BODY", true));
    const view = await send("检查当前表");
    await waitFor(() => expect(view.container.textContent).toContain(title));
    expect(view.container.textContent).toContain(action);
    expect(view.container.textContent).toContain("请以执行结果卡为准");
    expect(view.container.textContent).not.toContain("PRIVATE_PROVIDER_BODY");
    expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value).toBe("检查当前表");
    expect(view.getByRole("button", { name: /^发送$/ })).toBeTruthy();
    expect(captureWorkbookStructure).not.toHaveBeenCalled();
    expect(streamAssistantResponse).not.toHaveBeenCalled();
  });
  it("a submitted answer survives an intent network error; changing that answer replaces it on retry", async () => {
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "填写等级", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValueOnce({ kind: "clarification", provider: "model", clarification: {
      id: "retry-mode", summary: "填写口径", question: "如何填写？", reason: "需要口径", scopeLabel: "Sheet1", options: []
    } }).mockRejectedValueOnce(new TypeError("确认连接中断"))
      .mockResolvedValue({ kind: "proceed", confirmedPrompt: "根据类别填写等级", summary: "填写等级", provider: "model" });
    vi.mocked(streamAssistantResponse).mockResolvedValue({ kind: "answer", provider: "model", message: "收到新口径" });
    const view = await send("根据类别填写等级");
    await waitFor(() => expect(view.container.querySelector(".composer textarea")).not.toBeNull());
    /*
     * 自由输入就在输入框本身：不再有第二个光标（`.clarification-custom-input`），
     * 也不再有自己的「提交补充」按钮 —— 提交统一发生在右下角那一颗
     * （此时它的可访问名是「发送回答」）。
     */
    expect(view.container.querySelector(".clarification-custom-input")).toBeNull();
    expect(view.queryByRole("button", { name: "提交补充" })).toBeNull();
    fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "逐行填写" } });
    fireEvent.click(view.getByRole("button", { name: "发送回答" }));
    await waitFor(() => expect(view.container.textContent).toContain("确认连接中断"));
    const history = JSON.parse(localStorage.getItem("excel-bro.chat.v4")!);
    const card = history.conversations.flatMap((item: { messages: unknown[] }) => item.messages)
      .find((item: { clarification?: { id: string } }) => item.clarification?.id === "retry-mode");
    expect(card.clarification.taskContext.confirmations).toEqual([
      { questionId: "retry-mode", label: "填写口径", value: "逐行填写" }
    ]);
    fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "分组汇总" } });
    fireEvent.click(view.getByRole("button", { name: "发送回答" }));
    await waitFor(() => expect(view.container.textContent).toContain("收到新口径"));
    const request = vi.mocked(checkIntent).mock.calls[2][0];
    expect(request.taskContext?.confirmations).toEqual([
      { questionId: "retry-mode", label: "填写口径", value: "分组汇总" }
    ]);
    expect(request.prompt).toBe("根据类别填写等级\n补充确认：分组汇总");
  });
  it("continuous clarification carries both decisions into planning; a new task starts clean", async () => {
    const clarification = (id: string, summary: string, yes: string, no: string) => ({
      kind: "clarification" as const, provider: "model" as const,
      clarification: { id, summary, question: summary, reason: "需要决定", scopeLabel: "Sheet1",
        options: [{ id: "yes", label: yes, description: yes, resolution: yes },
          { id: "no", label: no, description: no, resolution: no }] }
    });
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "填写等级", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValueOnce(clarification("mode", "填写口径", "逐行填写", "分组汇总"))
      .mockResolvedValue({ kind: "proceed", confirmedPrompt: "调整列宽", summary: "调整列宽", provider: "model" });
    vi.mocked(streamAssistantResponse).mockResolvedValueOnce(clarification("blank", "空值处理", "空值不写入", "清空目标"))
      .mockResolvedValue({ kind: "answer", provider: "model", message: "完整规则已收到" });
    const view = await send("根据类别填写等级");
    /*
     * 快捷回答**点只选中、不提交**：以前点一下就把这一轮答案发出去了，
     * 想改口径没机会回头。现在选中只是本地草稿，发送才提交。
     */
    const yes = await view.findByRole("button", { name: /逐行填写/ });
    fireEvent.click(yes);
    expect(yes.getAttribute("aria-pressed")).toBe("true");
    expect(streamAssistantResponse).not.toHaveBeenCalled();
    fireEvent.click(view.getByRole("button", { name: "发送回答" }));
    await waitFor(() => expect(view.getByRole("button", { name: /空值不写入/ })).toBeTruthy());
    fireEvent.click(view.getByRole("button", { name: /空值不写入/ }));
    fireEvent.click(view.getByRole("button", { name: "发送回答" }));
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledTimes(2));
    const context = vi.mocked(streamAssistantResponse).mock.calls[1][0].taskContext;
    expect(context?.originalPrompt).toBe("根据类别填写等级");
    expect(context?.confirmations.map((item) => item.value)).toEqual(["逐行填写", "空值不写入"]);
    await waitFor(() => expect(view.container.textContent).toContain("完整规则已收到"));
    fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "调整列宽" } });
    fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledTimes(3));
    expect(vi.mocked(streamAssistantResponse).mock.calls[2][0].taskContext)
      .toEqual({ schemaVersion: 1, originalPrompt: "调整列宽", confirmations: [] });
  });

  /*
   * 回车在**澄清回答**这件事上的三种真实输入。
   *
   * 中文输入法的组合期按 Enter 是"选定候选词"，不是"发送"：以前这里只判
   * `key === "Enter" && !shiftKey`，一句话还没选完词就整句发出去了 —— 在问题待回答
   * 的时候它还会被当成"答案"，把这道题用半截话结掉。
   *
   * **口径**：组合期那一拍**允许候选字上屏**（我们一个字节都不拦：不 `preventDefault`），
   * 要求的是"不发送消息、不触发暂停／继续"。别把它写成"不上屏" —— 那是把输入法
   * 本来该做的事也当成了要禁止的行为。
   */
  it("组合期的 Enter 只确认候选词：候选字照常上屏，但不发送、不触发暂停或继续", async () => {
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "填写等级", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "clarification", provider: "model", clarification: {
      id: "ime-mode", summary: "填写口径", question: "如何填写？", reason: "需要口径",
      scopeLabel: "Sheet1", options: []
    } });
    const view = await send("根据类别填写等级");
    await waitFor(() => expect(view.getByRole("button", { name: "发送回答" })).toBeTruthy());
    const textarea = view.container.querySelector<HTMLTextAreaElement>(".composer textarea")!;
    fireEvent.change(textarea, { target: { value: "汇总一下" } });
    const before = vi.mocked(checkIntent).mock.calls.length;

    /*
     * ① 组合期（`isComposing`）：这一拍是"确认候选词"。我们**不拦默认行为**
     *    （`fireEvent` 返回 true = 没有 `preventDefault`），所以候选字照常进输入框；
     *    但一个字节都不发、也不暂停。
     */
    const composingSunk = fireEvent.keyDown(textarea, { key: "Enter", isComposing: true });
    await act(async () => {});
    expect(composingSunk).toBe(true);
    expect(vi.mocked(checkIntent).mock.calls.length).toBe(before);
    // ② 有的输入法在组词结束那一拍只留 `keyCode === 229`；同样不许发送。
    const composingEnd = fireEvent.keyDown(textarea, { key: "Enter", keyCode: 229 });
    await act(async () => {});
    expect(composingEnd).toBe(true);
    expect(vi.mocked(checkIntent).mock.calls.length).toBe(before);
    // 输入框里的字**还在**（既没清空，也没被当成新需求发走）。
    expect(textarea.value).toBe("汇总一下");
    expect(view.container.querySelector(".message-row.user")).not.toBeNull();

    // ③ Shift + Enter 是换行，同样不提交。
    fireEvent.keyDown(textarea, { key: "Enter", shiftKey: true });
    await act(async () => {});
    expect(vi.mocked(checkIntent).mock.calls.length).toBe(before);

    // ④ 组合结束后的那一次 Enter 才是提交。
    fireEvent.keyDown(textarea, { key: "Enter" });
    await waitFor(() => expect(vi.mocked(checkIntent).mock.calls.length).toBe(before + 1));
    expect(vi.mocked(checkIntent).mock.calls.at(-1)![0].prompt).toContain("汇总一下");
  });

  it("不读表的追问替换输入框：其他与框选互斥，提交仍不读取或写入", async () => {
    vi.mocked(understandMessage).mockResolvedValueOnce({ kind: "clarification", provider: "model", clarification: {
      id: "before-read", summary: "规则", question: "如何分级？", reason: "先确认", scopeLabel: "仅讨论",
      options: [
        { id: "range", label: "区间分级", description: "按区间", resolution: "按区间分级" },
        { id: "map", label: "逐项对应", description: "按映射", resolution: "按映射分级" }
      ]
    }}).mockResolvedValue({ kind: "needs_workbook_context", summary: "需要读表", provider: "model" });
    const view = await send("先询问分级规则，不要读取或修改工作表");
    const option = await view.findByRole("button", { name: /区间分级/ });
    const other = view.getByPlaceholderText("其他");
    expect(view.container.querySelector(".message-stream .clarification-card.pending")).toBeNull();
    expect(view.getAllByText("如何分级？")).toHaveLength(1);
    expect(option.closest(".composer-box.is-answering")).not.toBeNull();
    fireEvent.click(option);
    expect(option.getAttribute("aria-pressed")).toBe("true");
    fireEvent.change(other, { target: { value: "我的规则" } });
    expect(option.getAttribute("aria-pressed")).toBe("false");
    fireEvent.click(option);
    expect((other as HTMLTextAreaElement).value).toBe("");
    expect(understandMessage).toHaveBeenCalledTimes(1);
    fireEvent.click(view.getByRole("button", { name: "发送回答" }));
    await waitFor(() => expect(view.container.textContent).toContain("回答已记录，尚未读取或修改工作表"));
    expect(understandMessage).toHaveBeenCalledTimes(2);
    expect(vi.mocked(understandMessage).mock.calls[1][0].prompt).toContain("不要读取或修改工作表");
    expect(captureWorkbookStructure).not.toHaveBeenCalled();
    expect(checkIntent).not.toHaveBeenCalled();
    expect(streamAssistantResponse).not.toHaveBeenCalled();
    expect(view.container.querySelector(".composer-box.is-answering")).toBeNull();
    expect(view.container.querySelector(".message-stream .clarification-card.resolved")?.textContent).toContain("如何分级？");
  });

  it("选中选项后回车即可提交，且点选项本身不提交", async () => {
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "填写等级", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValueOnce({ kind: "clarification", provider: "model", clarification: {
      id: "enter-mode", summary: "填写口径", question: "如何填写？", reason: "需要口径", scopeLabel: "Sheet1",
      options: [{ id: "row", label: "逐行填写", description: "逐行填写", resolution: "逐行填写" }]
    } }).mockResolvedValue({ kind: "proceed", confirmedPrompt: "根据类别填写等级", summary: "填写等级", provider: "model" });
    vi.mocked(streamAssistantResponse).mockResolvedValue({ kind: "answer", provider: "model", message: "已收到口径" });
    const view = await send("根据类别填写等级");
    const option = await view.findByRole("button", { name: /逐行填写/ });
    expect(option.closest(".composer-box")).not.toBeNull();
    expect(view.container.querySelector(".clarification-bypass-toggle")?.closest(".composer-box")).not.toBeNull();
    expect(view.container.querySelectorAll(".clarification-composer")).toHaveLength(1);
    fireEvent.click(option);
    expect(option.getAttribute("aria-pressed")).toBe("true");
    expect(vi.mocked(checkIntent).mock.calls).toHaveLength(1);
    fireEvent.keyDown(view.container.querySelector(".composer textarea")!, { key: "Enter" });
    await waitFor(() => expect(view.container.textContent).toContain("已收到口径"));
    expect(view.container.querySelector(".composer-box.is-answering")).toBeNull();
    expect(vi.mocked(checkIntent).mock.calls).toHaveLength(2);
  });

  /*
   * 「旧问题不得重复提交」+「对话切换时答案失效」。
   *
   * 绑定身份是 `conversationId + messageId + clarificationId`。换到新对话之后，
   * 那道题属于**上一段**对话，输入框必须回到"提新需求"，而不是继续给旧题发答案。
   */
  it("换到新对话后旧题的选中与答案一起失效，发送恢复成新需求", async () => {
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "填写等级", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValueOnce({ kind: "clarification", provider: "model", clarification: {
      id: "switch-mode", summary: "填写口径", question: "如何填写？", reason: "需要口径", scopeLabel: "Sheet1",
      options: [{ id: "row", label: "逐行填写", description: "逐行填写", resolution: "逐行填写" }]
    } }).mockResolvedValue({ kind: "proceed", confirmedPrompt: "新的需求", summary: "新的需求", provider: "model" });
    vi.mocked(streamAssistantResponse).mockResolvedValue({ kind: "answer", provider: "model", message: "新需求已收到" });
    const view = await send("根据类别填写等级");
    fireEvent.click(await view.findByRole("button", { name: /逐行填写/ }));
    expect(view.getByRole("button", { name: "发送回答" })).toBeTruthy();

    fireEvent.click(view.getByRole("button", { name: "新对话" }));
    await waitFor(() => expect(view.queryByRole("button", { name: "发送回答" })).toBeNull());
    expect(view.queryByLabelText("回答待确认的问题")).toBeNull();
    expect(view.container.querySelector(".clarification-answer-region")).toBeNull();

    fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "新的需求" } });
    fireEvent.keyDown(view.container.querySelector(".composer textarea")!, { key: "Enter" });
    await waitFor(() => expect(view.container.textContent).toContain("新需求已收到"));
    // 走的是新需求链路（understandMessage 又跑了一次），不是给旧题补答案。
    expect(vi.mocked(understandMessage).mock.calls).toHaveLength(2);
  });
  it.each(["new", "history"])("clears the previous conversation's progress hint on %s chat navigation, showing the idle anchor for a finished turn", async (navigation) => {
    vi.mocked(understandMessage).mockResolvedValue({ kind: "conversation", message: "历史回答", provider: "model" });
    const view = await send("历史问题");
    await waitFor(() => expect(view.container.textContent).toContain("历史回答"));
    const sendMore = async (question: string, answer: string) => {
      vi.mocked(understandMessage).mockResolvedValue({ kind: "conversation", message: answer, provider: "model" });
      fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: question } });
      fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
      await waitFor(() => expect(view.container.textContent).toContain(answer));
    };
    if (navigation === "history") {
      fireEvent.click(view.getByRole("button", { name: "新对话" }));
      await sendMore("另一段对话", "另一段回答");
    }
    const stream = view.container.querySelector<HTMLElement>(".message-stream")!;
    Object.defineProperties(stream, {
      scrollHeight: { configurable: true, value: 1000 },
      clientHeight: { configurable: true, value: 300 }
    });
    stream.scrollTop = 200;
    fireEvent.scroll(stream);
    await sendMore("补充问题", "补充回答");
    // 这一轮**已经跑完**了，所以是"停下"那一态：锚，而不是三点。
    const hint = view.getByRole("button", { name: "回到底部查看最新内容" });
    expect(hint.dataset.mode).toBe("idle");
    expect(hint.querySelector(".progress-anchor-icon")).not.toBeNull();
    expect(hint.querySelector(".progress-dots")).toBeNull();
    expect(stream.scrollTop).toBe(200);

    if (navigation === "new") fireEvent.click(view.getByRole("button", { name: "新对话" }));
    else {
      fireEvent.click(view.getByRole("button", { name: "历史对话" }));
      const history = Array.from(view.container.querySelectorAll<HTMLButtonElement>(".history-open-button"))
        .find((button) => button.textContent?.includes("历史问题"))!;
      fireEvent.click(history);
    }
    expect(view.queryByRole("button", { name: "回到底部查看最新内容" })).toBeNull();
    expect(stream.scrollTop).toBe(1000);
    expect(captureWorkbookStructure).not.toHaveBeenCalled();
    expect(checkIntent).not.toHaveBeenCalled();
  });

  it("keeps the three-dot indicator while the turn is in flight and swaps to the anchor once it lands", async () => {
    // 把消息理解挂住：`running` 只有在**真的还跑着**的时候才为真，不挂住就只能看到终态，
    // 于是"运行中显示三点"这条接线没有任何用例约束得住。
    let finishUnderstanding:
      | ((value: Awaited<ReturnType<typeof understandMessage>>) => void)
      | undefined;
    vi.mocked(understandMessage).mockImplementation(
      () => new Promise((resolve) => { finishUnderstanding = resolve; })
    );

    const view = render(<App />);
    await waitFor(() => expect(vi.mocked(captureWorkbookCatalog)).toHaveBeenCalled());
    await waitFor(() =>
      expect(view.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title).toContain("当前模型")
    );

    // 先把用户"滚走"：浮标只在滚离底部时渲染（判据是位置，不要求有新内容）。
    const stream = view.container.querySelector<HTMLElement>(".message-stream")!;
    Object.defineProperties(stream, {
      scrollHeight: { configurable: true, value: 1000 },
      clientHeight: { configurable: true, value: 300 }
    });
    stream.scrollTop = 200;
    fireEvent.scroll(stream);

    fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "统计当前表" } });
    fireEvent.click(view.getByRole("button", { name: /^发送$/ }));

    // 运行时：三点动效。等到的是 running 那一态本身 —— 如果接线写成 `running={false}`，
    // 这里永远等不到，用例会在这里超时失败，而不是悄悄放行。
    await waitFor(() =>
      expect(
        view.container.querySelector<HTMLButtonElement>(".new-progress-button")?.dataset.mode
      ).toBe("running")
    );
    const hint = view.getByRole("button", { name: "查看最新执行进度" });
    expect(hint.querySelectorAll(".progress-dots > span")).toHaveLength(3);
    expect(hint.querySelector(".progress-anchor-icon")).toBeNull();

    await act(async () => {
      finishUnderstanding?.({ kind: "conversation", message: "统计完成", provider: "model" });
    });

    // 跑完：同一个按钮换成锚（不是消失再出现）。
    await waitFor(() =>
      expect(
        view.container.querySelector<HTMLButtonElement>(".new-progress-button")?.dataset.mode
      ).toBe("idle")
    );
    expect(view.getByRole("button", { name: "回到底部查看最新内容" })).toBe(hint);
    expect(hint.querySelector(".progress-anchor-icon")).not.toBeNull();
    expect(hint.querySelector(".progress-dots")).toBeNull();
  });

  it.each(["running", "unavailable", "no_checkpoint", "completed"])("manual pause automatically inspects %s without starting another plan", async (kind) => {
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析当前表", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "分析当前表", summary: "分析", provider: "model", turnId: "excel-paused-turn" });
    vi.mocked(streamAssistantResponse).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
      hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    if (kind === "unavailable") vi.mocked(getPlanningTaskStatus).mockRejectedValue(new TypeError("Failed to fetch"));
    else vi.mocked(getPlanningTaskStatus).mockResolvedValue({ turnId: "excel-paused-turn", status: kind === "no_checkpoint" ? "interrupted" : kind,
      revision: 7, lastSequence: 12, canContinue: false, reason: "没有检查点",
      ...(kind === "completed" ? { result: { kind: "answer", provider: "model", message: "原任务结果已取回" } } : {}) } as never);
    const view = await send("分析当前表");
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    fireEvent.click(await view.findByRole("button", { name: /^暂停$/ }));
    await waitFor(() => expect(getPlanningTaskStatus).toHaveBeenCalledOnce());
    await waitFor(() => expect(view.container.querySelector('[aria-label="暂停任务"]')).toBeNull());
    expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value).toBe("");
    expect(continuePlanningTask).not.toHaveBeenCalled();
    expect(checkIntent).toHaveBeenCalledOnce();
    expect(understandMessage).toHaveBeenCalledOnce();
    if (kind === "running" || kind === "unavailable") expect(view.getByRole("button", { name: "核对任务状态" })).toBeTruthy();
    // 产品文案已统一为「继续分析」；这里仍要证明"不能继续"这件事被如实说出来、
    // 而且**没有**自动继续（下一行的 continuePlanningTask 断言钉住这一条）。
    if (kind === "no_checkpoint") expect(view.container.textContent).toContain("不能继续分析");
    if (kind === "completed") {
      expect(view.container.textContent).toContain("结果已保留");
      expect(view.container.textContent).not.toContain("原任务结果已取回");
      const before = segmentSnapshots(view.container);
      fireEvent.click(await view.findByRole('button', { name: '核对任务状态' }));
      await waitFor(() => expect(view.container.textContent).toContain('原任务结果已取回'));
      const after = segmentSnapshots(view.container);
      expect(after).toHaveLength(before.length + 1);
      expect(after.slice(0, before.length).map((segment) => segment.steps)).toEqual(before.map((segment) => segment.steps));
      expect(continuePlanningTask).not.toHaveBeenCalled();
      expect(streamAssistantResponse).toHaveBeenCalledOnce();
    }
    // 暂停之后那次自动核对是**只读查询**，不去轮询补发进度：用户刚按下暂停，
    // 不该被立刻接回一条"其实还在进行中"。接回只属于用户明确要求继续的那一次。
    expect(getPlanningTaskEvents).not.toHaveBeenCalled();
  });

  /**
   * 「原任务仍在服务端运行」这条恢复路径（Excel 端）。
   *
   * 服务端还在跑时继续请求会被 409 挡住（再发一次就是重复规划、重复读表），但
   * "点了没反应"同样是错的：任务确实在跑，进度也确实写进了账本。正确做法是
   * **只读补发**已落盘的事件（`GET /api/turn/{id}/events`），把后续进度落到
   * **新的回复段**（同任务、多段回复），等它结束后再走同一条结果处理链取回结果。
   */
  it.each([false, true])("running 接回处理读表回执，包括暂停前已收到的指针 %s", async (alreadyReceived) => {
    const host = createOfficeTestHost({ url: "file:///C:/Data/book.xlsx", sheets: [
      { name: "Sheet1", values: [["Category", "Level"], ["A", 1], ["B", 2]] }
    ] });
    vi.stubGlobal("Excel", host.Excel);
    vi.stubGlobal("Office", host.Office);
    const OP = "op_" + "b".repeat(32);
    let published = alreadyReceived;
    let receipted = false;
    const receipts: any[] = [];
    vi.stubGlobal("fetch", vi.fn(async (url: string, init?: RequestInit) => {
      const body = JSON.parse(String(init?.body ?? "{}"));
      if (String(url).includes("/poll")) return new Response(JSON.stringify({ generation: 1,
        calls: published && !receipted ? [{ operationId: OP, tool: "query_table",
          arguments: { mode: "aggregate", scope: "selected", metrics: [{ operation: "count", outputName: "rows" }] }, requestFingerprint: "d".repeat(64), expiresAt: 0 }] : [] }), { status: 200 });
      if (String(url).includes("/receipt")) {
        receipts.push(body); receipted = true;
        return new Response(JSON.stringify({ status: "accepted", code: null }), { status: 200 });
      }
      return new Response("{}", { status: 200 });
    }));
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "分析", summary: "分析", provider: "model", turnId: "excel-paused-turn" });
    vi.mocked(streamAssistantResponse).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
      hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    vi.mocked(getPlanningTaskStatus).mockImplementation(async () => ({ turnId: "excel-paused-turn",
      status: receipted ? "completed" : "running", canContinue: false, revision: 7, lastSequence: 31,
      ...(receipted ? { result: { kind: "answer", provider: "model", message: "收到真实读表回执后完成" } } : {}) } as never));
    vi.mocked(getPlanningTaskEvents).mockImplementation(async () => {
      const first = !published; published = true;
      return { turnId: "excel-paused-turn", active: !receipted, status: receipted ? "completed" : "running", lastSequence: 31,
        events: first ? [{ schemaVersion: 1, operationId: "excel-paused-turn", sequence: 31, phase: "planning",
          stepId: "bridge.query_table:bbbb", status: "started", title: "读取授权表", timestamp: "2026-10-22T00:00:00Z",
          bridgeTool: { operationId: OP, tool: "query_table" } }] : [] } as never;
    });
    const view = await send("分析当前表");
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    fireEvent.click(await view.findByRole("button", { name: "暂停任务" }));
    await waitFor(() => expect((view.getByRole("button", { name: "核对任务状态" }) as HTMLButtonElement).disabled).toBe(false));
    const cards = view.container.querySelectorAll(".task-progress-timeline").length;
    const before = segmentSnapshots(view.container);
    fireEvent.click(await view.findByRole("button", { name: "核对任务状态" }));
    await waitFor(() => expect(view.container.textContent).toContain("收到真实读表回执后完成"), { timeout: 6000 });
    const after = segmentSnapshots(view.container);
    // 判别点 1：继续产生**新的回复段**；历史段的步骤行逐字保留它停止时的内容。
    expect(after).toHaveLength(cards + 1);
    expect(after.slice(0, before.length).map((segment) => segment.steps))
      .toEqual(before.map((segment) => segment.steps));
    // 旧段只多了一个轻量定位入口，没有被改写成"最新进展"。
    expect(after[0].text).toContain("查看后续进展");
    expect(after[0].text).not.toContain("收到真实读表回执后完成");
    // 判别点 2：同一时刻只有一个活动段（同一任务仍只有一个执行实例）。
    expect(liveSegmentCount(view.container)).toBeLessThanOrEqual(1);
    expect(continuePlanningTask).not.toHaveBeenCalled();
    expect(streamAssistantResponse).toHaveBeenCalledOnce();
    expect(receipts).toHaveLength(1);
    expect(receipts[0]).toMatchObject({ taskId: "excel-paused-turn", operationId: OP });
    expect(receipts[0].receipt.errorCode).toBeUndefined();
    expect(receipts[0].receipt.result).toBeDefined();
    expect(host.writes()).toEqual([]);
  });

  it.each([false, true])("继续只补收未见事件，轮询暂停忽略迟到页面 %s", async (pausePoll) => {
    const host = createOfficeTestHost({ url: "file:///C:/Data/book.xlsx", sheets: [
      { name: "Sheet1", values: [["Category", "Level"], ["A", 1], ["B", 2]] }
    ] });
    vi.stubGlobal("Excel", host.Excel);
    vi.stubGlobal("Office", host.Office);
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析当前表", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "分析当前表", summary: "分析", provider: "model", turnId: "excel-paused-turn" });
    vi.mocked(streamAssistantResponse).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
      hooks?.onProgress?.({ operationId: 'excel-paused-turn', sequence: 12, stepId: 'query-done',
        title: '查询已经完成', phase: 'planning', status: 'completed', schemaVersion: 1 } as never);
      hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    vi.mocked(getPlanningTaskStatus)
      // ① 暂停后的自动核对：任务确实还在服务端跑。
      .mockResolvedValueOnce({ turnId: "excel-paused-turn", status: "running", revision: 7, lastSequence: 12, canContinue: false } as never)
      // ② 点「核对任务状态」时的第 1 次读：仍在跑 → 去接回已落盘进度。
      .mockResolvedValueOnce({ turnId: "excel-paused-turn", status: "running", revision: 7, lastSequence: 12, canContinue: false } as never)
      // ③ 接回之后任务已结束：同一条处理链取回结果。
      .mockResolvedValueOnce({ turnId: "excel-paused-turn", status: "completed", revision: 8, lastSequence: 14,
        canContinue: false, result: { kind: "answer", provider: "model", message: "原任务结果已取回" } } as never);
    vi.mocked(getPlanningTaskEvents)
      // 第 1 页：任务还在跑，但第 7 步的进度已经落盘。
      .mockResolvedValueOnce({ turnId: "excel-paused-turn", status: "running", active: true, lastSequence: 13, events: [
        { operationId: 'excel-paused-turn', sequence: 12, stepId: 'query-done', phase: 'planning', status: 'completed', title: '重复查询事件不应出现' },
        { schemaVersion: 1, operationId: "excel-paused-turn", sequence: 13, phase: "planning",
          stepId: "step-7", status: "progress", title: "步骤7继续", timestamp: "2026-10-21T00:00:00Z" }
      ] } as never)
      // 第 2 页：任务结束，最后一次补发带出第 8 步。
      .mockResolvedValueOnce({ turnId: "excel-paused-turn", status: "completed", active: false, lastSequence: 14, events: [
        { schemaVersion: 1, operationId: "excel-paused-turn", sequence: 14, phase: "planning",
          stepId: "step-8", status: "completed", title: "步骤8", timestamp: "2026-10-21T00:00:00Z" }
      ] } as never);

    const view = await send("分析当前表");
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    fireEvent.click(await view.findByRole("button", { name: /^暂停$/ }));
    await waitFor(() => expect(getPlanningTaskStatus).toHaveBeenCalledOnce());
    await waitFor(() => expect((view.getByRole("button", { name: "核对任务状态" }) as HTMLButtonElement).disabled).toBe(false));
    const before = segmentSnapshots(view.container);
    const timelineCountBefore = before.length;
    expect(before[0].steps.length).toBeGreaterThan(0);

    let finish!: (page: any) => void;
    if (pausePoll) vi.mocked(getPlanningTaskEvents).mockReset().mockImplementation(() => new Promise((resolve) => { finish = resolve; }));
    fireEvent.click(await view.findByRole("button", { name: "核对任务状态" }));
    if (pausePoll) {
      await waitFor(() => expect(finish).toBeTypeOf('function'));
      fireEvent.click(await view.findByRole('button', { name: '暂停任务' }));
      await waitFor(() => expect(view.queryByRole('button', { name: '暂停任务' })).toBeNull());
      const paused = segmentSnapshots(view.container);
      await act(async () => finish({ turnId: 'excel-paused-turn', active: false, status: 'completed', lastSequence: 13,
        events: [{ operationId: 'excel-paused-turn', sequence: 13, stepId: 'late', title: '迟到页面不应出现', phase: 'planning', status: 'completed' }] }));
      expect(segmentSnapshots(view.container)).toEqual(paused);
      expect(view.container.textContent).not.toContain('迟到页面不应出现');
      expect(continuePlanningTask).not.toHaveBeenCalled();
      expect(host.writes()).toEqual([]);
      return;
    }

    // 补发按 1200ms 退避，等的是真实墙钟时间，不是 settle 次数。
    await waitFor(() => expect(view.container.textContent).toContain("原任务结果已取回"), { timeout: 6000 });
    expect(vi.mocked(getPlanningTaskEvents).mock.calls[0][1]).toBe(12);
    expect(view.container.textContent).not.toContain('重复查询事件不应出现');
    const after = segmentSnapshots(view.container);
    // 判别点 1：继续开了一个**新的回复段**，旧段的步骤行逐字不变（不是被就地改写）。
    expect(after).toHaveLength(timelineCountBefore + 1);
    expect(after.slice(0, timelineCountBefore).map((segment) => segment.steps))
      .toEqual(before.map((segment) => segment.steps));
    expect(after[0].text).toContain("查看后续进展");
    expect(after[0].text).not.toContain("步骤8");
    // 判别点 2：补发出来的两步落在**最新的那一段**里，且是接在原步骤之后
    // （不重置编号，也不把"接回进度"另开一行 —— 接回的是原任务那两步本身）。
    const lastSegment = view.container.querySelectorAll<HTMLElement>(".task-progress-timeline");
    const rows = lastSegment[lastSegment.length - 1].querySelectorAll(".activity-steps > li");
    // 新段 = 一条真实的"接回原任务进度"步骤 + 补发出来的两步。
    expect(rows).toHaveLength(3);
    expect(rows[0].textContent).toContain("已接回原任务进度");
    expect(rows[0].getAttribute("data-status")).toBe("completed");
    expect(rows[rows.length - 2].textContent).toContain("步骤7继续");
    expect(rows[rows.length - 1].textContent).toContain("步骤8");
    // 判别点 3：只有一个活动段 —— 同一任务仍只有一个有效执行实例。
    expect(liveSegmentCount(view.container)).toBeLessThanOrEqual(1);
    // 判别点 4：只读补发 —— 没有新的继续请求，也没有新的规划轮次与意图确认。
    expect(continuePlanningTask).not.toHaveBeenCalled();
    expect(streamAssistantResponse).toHaveBeenCalledOnce();
    expect(checkIntent).toHaveBeenCalledOnce();
    expect(understandMessage).toHaveBeenCalledOnce();
  });

  it("a model-routed continue queries the paused task instead of starting a new workbook task", async () => {
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析当前表", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "分析当前表", summary: "分析", provider: "model", turnId: "excel-paused-turn" });
    vi.mocked(streamAssistantResponse).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
      hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    vi.mocked(getPlanningTaskStatus).mockResolvedValue({ turnId: "excel-paused-turn", status: "running",
      revision: 7, lastSequence: 12, canContinue: false } as never);
    const view = await send("分析当前表");
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    fireEvent.click(await view.findByRole("button", { name: /^暂停$/ }));
    await waitFor(() => expect(getPlanningTaskStatus).toHaveBeenCalledOnce());
    await waitFor(() => expect((view.getByRole("button", { name: "核对任务状态" }) as HTMLButtonElement).disabled).toBe(false));
    vi.mocked(understandMessage).mockResolvedValueOnce({ kind: "conversation", message: "稍后可以核对。", provider: "model" });
    fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "解释一下这个状态" } });
    fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
    await waitFor(() => expect(view.container.textContent).toContain("稍后可以核对"));
    vi.mocked(understandMessage).mockResolvedValueOnce({ kind: "task_recovery", turnId: "excel-paused-turn", provider: "model" });
    fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "继续" } });
    fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
    await waitFor(() => expect(getPlanningTaskStatus).toHaveBeenCalledTimes(2));
    expect(vi.mocked(understandMessage).mock.calls.at(-1)?.[0].recoveryTasks).toEqual(["excel-paused-turn"]);
    expect(vi.mocked(getPlanningTaskStatus).mock.calls.at(-1)?.[0]).toBe("excel-paused-turn");
    expect(checkIntent).toHaveBeenCalledOnce();
    expect(streamAssistantResponse).toHaveBeenCalledOnce();
    expect(continuePlanningTask).not.toHaveBeenCalled();
    await waitFor(() => expect((view.getByRole("button", { name: "核对任务状态" }) as HTMLButtonElement).disabled).toBe(false));
  });

  it("an unoffered model recovery ID cannot query a different task or enter workbook planning", async () => {
    vi.mocked(understandMessage).mockResolvedValue({ kind: "task_recovery", turnId: "invented", provider: "model" });
    const view = await send("继续");
    await waitFor(() => expect(view.container.textContent).toContain("无法确定要恢复的原任务"));
    expect(getPlanningTaskStatus).not.toHaveBeenCalled();
    expect(checkIntent).not.toHaveBeenCalled();
    expect(continuePlanningTask).not.toHaveBeenCalled();
  });

  it("automatic inspection preserves an unsent new draft, even when that inspection fails", async () => {
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析当前表", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "分析当前表", summary: "分析", provider: "model", turnId: "excel-paused-turn" });
    vi.mocked(streamAssistantResponse).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
      hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    let rejectStatus!: (error: Error) => void;
    vi.mocked(getPlanningTaskStatus).mockImplementation(() => new Promise((_resolve, reject) => { rejectStatus = reject; }));
    const view = await send("分析当前表");
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    const input = view.container.querySelector<HTMLTextAreaElement>(".composer textarea")!;
    fireEvent.change(input, { target: { value: "尚未发送的新问题" } });
    fireEvent.click(await view.findByRole("button", { name: /^暂停$/ }));
    await waitFor(() => expect(getPlanningTaskStatus).toHaveBeenCalledOnce());
    await act(async () => { rejectStatus(new TypeError("Failed to fetch")); });
    await waitFor(() => expect(view.getByRole("button", { name: "核对任务状态" })).toBeTruthy());
    expect(input.value).toBe("尚未发送的新问题");
    expect(continuePlanningTask).not.toHaveBeenCalled();
    expect(understandMessage).toHaveBeenCalledOnce();
  });

  it("pausing message understanding has no checkpoint and keeps the request in history, not in the send box", async () => {
    vi.mocked(understandMessage).mockImplementation((_request, signal) => new Promise((_resolve, reject) => {
      signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    const view = await send("分析当前表");
    await waitFor(() => expect(understandMessage).toHaveBeenCalledOnce());
    fireEvent.click(await view.findByRole("button", { name: /^暂停$/ }));
    await waitFor(() => expect(view.getByRole("button", { name: "继续分析" })).toBeTruthy());
    expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value).toBe("");
    expect(view.container.querySelector(".message-row.user")?.textContent).toContain("分析当前表");
    expect(getPlanningTaskStatus).not.toHaveBeenCalled();
    expect(captureWorkbookStructure).not.toHaveBeenCalled();
    expect(checkIntent).not.toHaveBeenCalled();
  });

  it.each((["message", "reading", "fields"] as const).flatMap((stage) =>
    (["text", "button"] as const).map((via) => ({ stage, via }))
  ))("early $stage pause restarts via $via without recovering the consumed task again", async ({ stage, via }) => {
    const blocked = (_request: unknown, signal?: AbortSignal) => new Promise<never>((_resolve, reject) => {
      signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    });
    let releaseReading!: (value: typeof fixtures.snapshot) => void;
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析", provider: "model" });
    if (stage === "message") vi.mocked(understandMessage).mockImplementationOnce(blocked);
    else if (stage === "reading") vi.mocked(captureWorkbookStructure).mockReturnValueOnce(
      new Promise((resolve) => { releaseReading = resolve; }));
    else vi.mocked(checkIntent).mockImplementationOnce(blocked);
    const view = await send("按会员等级汇总订单并列出异常");
    await waitFor(() => expect(stage === "message" ? understandMessage
      : stage === "reading" ? captureWorkbookStructure : checkIntent).toHaveBeenCalledOnce());
    fireEvent.click(await view.findByRole("button", { name: /^暂停$/ }));
    await waitFor(() => expect(view.getByRole("button", { name: "继续分析" })).toBeTruthy());
    if (stage === "reading") await act(async () => releaseReading(fixtures.snapshot));
    // 暂停时就地取快照：后面要逐字比较"历史段有没有被这一轮改写"。
    const before = segmentSnapshots(view.container);
    expect(before).toHaveLength(1);
    // A second semantic call may select any offered paused task, even when its
    // prompt is the original request. Consumed IDs must not be offered again.
    vi.mocked(understandMessage).mockImplementation(async (request) => request.recoveryTasks?.length
      ? { kind: "task_recovery", turnId: request.recoveryTasks![0], action: "continue", provider: "model" }
      : { kind: "needs_workbook_context", summary: request.prompt, provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "conversation", message: "原需求重新分析已接通", provider: "model" });
    const semanticCallsBeforeRestart = vi.mocked(understandMessage).mock.calls.length;
    if (via === "text") {
      fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "继续" } });
      fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
    } else {
      expect(view.container.querySelector(".analysis-pause-inline")).not.toBeNull();
      expect(view.container.textContent).not.toContain("尚未生成规划检查点");
      fireEvent.click(await view.findByRole("button", { name: "继续任务" }));
    }
    await waitFor(() => expect(view.container.textContent).toContain("原需求重新分析已接通"));
    expect(understandMessage).toHaveBeenCalledTimes(semanticCallsBeforeRestart + (via === "text" ? 1 : 0) + (stage === "message" ? 1 : 0));
    if (stage === "message") {
      expect(vi.mocked(understandMessage).mock.calls.at(-1)?.[0].prompt).toBe("按会员等级汇总订单并列出异常");
      expect(vi.mocked(understandMessage).mock.calls.at(-1)?.[0].recoveryTasks).toBeUndefined();
    }
    expect(view.container.querySelectorAll(".task-progress-timeline")).toHaveLength(2);
    expect(view.container.textContent).not.toContain("已接回原需求，重新分析");
    expect(vi.mocked(checkIntent).mock.calls.at(-1)?.[0].prompt).toBe("按会员等级汇总订单并列出异常");
    expect(getPlanningTaskStatus).not.toHaveBeenCalled();
    expect(continuePlanningTask).not.toHaveBeenCalled();
    expect(streamAssistantResponse).not.toHaveBeenCalled();
    expect([...view.container.querySelectorAll(".message-row.user")]
      .filter((item) => item.textContent?.includes("按会员等级汇总订单并列出异常"))).toHaveLength(1);

    // —— 回复段（缺陷 1）——
    const after = segmentSnapshots(view.container);
    // 判别点 1：继续开了一个**新的回复段**，不是把进度接回旧卡。
    expect(after).toHaveLength(before.length + 1);
    // 判别点 2：历史段逐字保留它暂停时的步骤行与耗时（不是被这一轮就地改写）。
    expect(after.slice(0, before.length).map((segment) => segment.steps))
      .toEqual(before.map((segment) => segment.steps));
    // 历史段只多了一个轻量定位入口，并保留"已暂停"的结论。
    expect(after[before.length - 1].text).toContain("查看后续进展");
    // 判别点 3：新段排在**用户那句话之后**（那句话换来了下面这段进展）。
    expect(segmentRowIndex(view.container, after.length - 1))
      .toBeGreaterThan(lastUserRowIndex(view.container));
    // 判别点 4：同一时刻只有一个活动段；历史段不再提供"继续"。
    expect(liveSegmentCount(view.container)).toBeLessThanOrEqual(1);
    expect([...view.container.querySelectorAll("button")]
      .filter((button) => button.textContent?.trim() === "继续分析")).toHaveLength(0);
    // 续跑段不从"步骤 1"重开：它是一条"继续分析"的承接行。
    /*
     * 进行中的卡片默认**收起成一行**（`processPanelExpanded` 对 `live` 直接返回收起：
     * 那一行显示的是"当前正在做的事"）。收起是呈现，不是内容被删 —— 所以先展开再读，
     * 与本文件另一处（"收起之后仍要能证明带来了新动作"）用同一手法。
     */
    const newSegmentCard = view.container
      .querySelectorAll<HTMLElement>(".task-progress-timeline")[after.length - 1];
    expect(newSegmentCard?.getAttribute("data-expanded")).toBe("false");
    fireEvent.click(newSegmentCard!.querySelector<HTMLButtonElement>(".task-progress-toggle")!);
    expect(segmentSnapshots(view.container)[after.length - 1].steps[0]).toMatch(/继续分析|已理解消息/);

    // 判别点 5：已经完成的模型阶段不重做 —— 只有暂停在**那一阶段本身**才会重跑它。
    // · 消息理解阶段没跑完 ⇒ 重跑一次；跑完了 ⇒ 一次都不再调（复用原始理解结果）。
    // · 表结构读取是本地读宿主，续跑一律按现场重读（不拿半截快照顶替）。
    expect(understandMessage).toHaveBeenCalledTimes(
      semanticCallsBeforeRestart + (via === "text" ? 1 : 0) + (stage === "message" ? 1 : 0)
    );
    expect(captureWorkbookStructure).toHaveBeenCalledTimes(stage === "message" ? 1 : 2);
    expect(checkIntent).toHaveBeenCalledTimes(stage === "fields" ? 2 : 1);

    // 判别点 6：早期恢复**没有**服务端任务身份 —— 段上不许出现伪造的 taskId，
    // 但段链本身要落盘（旧段 → 新段），下次继续才能接在最新那段后面。
    await waitFor(() => {
      type StoredMessage = { id: string; replySegment?: Record<string, unknown> };
      const stored = JSON.parse(localStorage.getItem("excel-bro.chat.v4") ?? "{}");
      const conversations: Array<{ messages: StoredMessage[] }> =
        Array.isArray(stored.conversations) ? stored.conversations : [];
      const persisted = conversations.flatMap((conversation) => conversation.messages);
      const segments = persisted.filter((message) => message.replySegment);
      const newest = segments.at(-1)!;
      expect(newest.replySegment).toMatchObject({ status: "live", continuesTask: true });
      expect(newest.replySegment).not.toHaveProperty("taskId");
      const superseded = segments.filter(
        (message) => message.replySegment?.status === "superseded"
      );
      expect(superseded).toHaveLength(1);
      expect(superseded[0].replySegment).toMatchObject({ nextSegmentId: newest.id });
    });
  });

  it.each(["composer", "card"])("field intent pause via %s settles before the response, including after continue", async (via) => {
    const releases: Array<(value: Awaited<ReturnType<typeof checkIntent>>) => void> = [];
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析", provider: "model" });
    vi.mocked(checkIntent).mockImplementation(() => new Promise((resolve) => releases.push(resolve)));
    const view = await send("分析当前表");
    for (let attempt = 1; attempt <= 2; attempt++) {
      await waitFor(() => expect(checkIntent).toHaveBeenCalledTimes(attempt));
      /*
       * 「请求已在飞」不等于「界面已提交可暂停状态」—— 这是本条用例曾偶发抓不到
       * 按钮的真实原因（取证：outputs/pause-fix-2026-10-23/evidence/race-probe.log）。
       *
       * 探针在 `checkIntent` **派发那一刻**同步读 DOM，两轮结果不同：
       *   · 第 1 轮：`■ / aria-label=暂停任务 / is-stop / disabled=false` 已在，
       *     活动卡 1 张（点击到派发之间隔了几次 await，React 已提交）；
       *   · 第 2 轮（点「继续任务」之后）：DOM 里**只有** `▶ / aria-label=继续任务`，
       *     活动卡 0 张，一个暂停按钮也没有；再等一个宏任务才变成 ■ + 1 张活动卡。
       * 所以失败瞬间产品并没有"进不了可暂停状态"，只是那一帧还没画出来。
       *
       * 等待条件用按钮的**可访问名**，不用时间：这个按钮叫「暂停任务」本身
       * 就等于"当前确实存在可中止的执行"，也正是用户能按下去的那个条件。
       */
      const pauseName = via === "composer" ? "暂停任务" : /^暂停$/;
      const pauseButton = await view.findByRole("button", { name: pauseName });
      expect(pauseButton.hasAttribute("disabled")).toBe(false);
      fireEvent.click(pauseButton);
      // Intentionally keep the transport pending: pausing must not wait for its response.
      await waitFor(() => expect(view.queryByRole("button", { name: "暂停任务" })).toBeNull());
      expect(view.getByRole("button", { name: "继续任务" })).toBeTruthy();
      const frozen = segmentSnapshots(view.container);
      await act(async () => releases[attempt - 1]({ kind: "proceed", provider: "model", turnId: `late-${attempt}`,
        summary: "迟到结果", confirmedPrompt: "分析当前表" }));
      expect(segmentSnapshots(view.container)).toEqual(frozen);
      expect(streamAssistantResponse).not.toHaveBeenCalled();
      if (attempt === 1) fireEvent.click(await view.findByRole("button", { name: "继续任务" }));
    }
  });

  it("late field intent after pause cannot start planning and leaves a usable early analysis card", async () => {
    let release!: (value: Awaited<ReturnType<typeof checkIntent>>) => void;
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析", provider: "model" });
    vi.mocked(checkIntent).mockReturnValueOnce(new Promise((resolve) => { release = resolve; }));
    const view = await send("分析当前表");
    await waitFor(() => expect(checkIntent).toHaveBeenCalledOnce());
    fireEvent.click(await view.findByRole("button", { name: /^暂停$/ }));
    await act(async () => release({ kind: "proceed", provider: "model", turnId: "late-intent-id",
      summary: "分析", confirmedPrompt: "分析当前表" }));
    await waitFor(() => expect(view.getByRole("button", { name: "继续分析" })).toBeTruthy());
    expect(streamAssistantResponse).not.toHaveBeenCalled();
    expect(getPlanningTaskStatus).not.toHaveBeenCalled();
  });

  it("automatic pause inspection stops at changed authorization and never leaves a permanently loading card", async () => {
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析当前表", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "分析当前表", summary: "分析", provider: "model", turnId: "excel-paused-turn" });
    vi.mocked(streamAssistantResponse).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
      hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    const view = await send("分析当前表");
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    vi.mocked(inspectOfficeAuthorizedTarget).mockResolvedValue({ ...fixtures.target, fileKey: "office:another-file" });
    fireEvent.click(await view.findByRole("button", { name: "暂停任务" }));
    await waitFor(() => expect(view.container.querySelector(".target-scope-notice")).not.toBeNull());
    await waitFor(() => expect((view.getByRole("button", { name: "核对任务状态" }) as HTMLButtonElement).disabled).toBe(false));
    expect(getPlanningTaskStatus).not.toHaveBeenCalled();
    expect(continuePlanningTask).not.toHaveBeenCalled();
    expect(understandMessage).toHaveBeenCalledOnce();
    expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value).toBe("");
  });

  it("running 查询期间换文件，接回前重新核验且不留下恢复中状态", async () => {
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析当前表", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "分析当前表", summary: "分析", provider: "model", turnId: "excel-paused-turn" });
    vi.mocked(streamAssistantResponse).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
      hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    const running: Awaited<ReturnType<typeof getPlanningTaskStatus>> = {
      kind: "planning_pause", turnId: "excel-paused-turn", status: "running", revision: 7, lastSequence: 12,
      canContinue: false, reason: "原任务仍在运行", activeElapsedSeconds: 3, approvedWallTimeSeconds: 60,
      extensionSeconds: 0, extensionCount: 0, maximumExtensions: 3, requiresExtension: false,
      pauseType: "interrupted", modelSwitchCount: 0, canSwitchModel: false, availableModelConnections: [],
      continuation: { available: false, code: "PLANNING_TASK_STILL_RUNNING", message: "原任务仍在运行", stage: null }
    };
    vi.mocked(getPlanningTaskStatus).mockResolvedValue(running);
    const view = await send("分析当前表");
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    fireEvent.click(await view.findByRole("button", { name: "暂停任务" }));
    await waitFor(() => expect(view.getByRole("button", { name: "核对任务状态" })).toBeTruthy());
    vi.mocked(getPlanningTaskStatus).mockImplementationOnce(async () => {
      vi.mocked(inspectOfficeAuthorizedTarget).mockResolvedValue({ ...fixtures.target, fileKey: "office:another-file" });
      return running;
    });
    fireEvent.click(await view.findByRole("button", { name: "核对任务状态" }));
    await waitFor(() => expect(view.container.querySelector(".target-scope-notice")).not.toBeNull());
    await waitFor(() => expect((view.getByRole("button", { name: "核对任务状态" }) as HTMLButtonElement).disabled).toBe(false));
    expect(getPlanningTaskEvents).not.toHaveBeenCalled();
    expect(continuePlanningTask).not.toHaveBeenCalled();
    expect(checkIntent).toHaveBeenCalledOnce();
  });

  it("pausing a continued stream automatically rechecks the same task, without recursive resume or draft refill", async () => {
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析当前表", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "分析当前表", summary: "分析", provider: "model", turnId: "excel-paused-turn" });
    const waitForPause = (_request: unknown, hooks?: { signal?: AbortSignal }) => new Promise<never>((_resolve, reject) => {
      hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    });
    vi.mocked(streamAssistantResponse).mockImplementation(waitForPause);
    vi.mocked(getPlanningTaskStatus).mockResolvedValue({ kind: "planning_pause", turnId: "excel-paused-turn", status: "interrupted",
      revision: 7, lastSequence: 12, canContinue: true, pauseType: "interrupted", reason: "已核实检查点", extensionSeconds: 0,
      activeElapsedSeconds: 3, currentModelId: "model-1", canSwitchModel: false, availableModelConnections: [],
      // 这份现场没读过来源 ⇒ 继续前不需要重读数据；缺了这个布尔，客户端会按
      // fail-closed 要求重读来源，jsdom 里读不到就会正确地停下来。
      continuation: { available: true, code: null, message: "", stage: "after_tool", sourceVerificationRequired: false } } as never);
    vi.mocked(continuePlanningTask).mockImplementationOnce(waitForPause)
      .mockResolvedValueOnce({ kind: "answer", provider: "model", message: "再次暂停后完成" });
    const view = await send("分析当前表");
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    fireEvent.click(await view.findByRole("button", { name: "暂停任务" }));
    await waitFor(() => expect(view.getByRole("button", { name: "继续分析" })).toBeTruthy());
    fireEvent.click(view.getByRole("button", { name: "继续分析" }));
    await waitFor(() => expect(continuePlanningTask).toHaveBeenCalledOnce());
    fireEvent.click(await view.findByRole("button", { name: "暂停任务" }));
    await waitFor(() => expect(getPlanningTaskStatus).toHaveBeenCalledTimes(3));
    await waitFor(() => expect(view.getByRole("button", { name: "继续分析" })).toBeTruthy());
    expect(continuePlanningTask).toHaveBeenCalledOnce();
    expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value).toBe("");
    fireEvent.click(view.getByRole("button", { name: "继续分析" }));
    await waitFor(() => expect(view.container.textContent).toContain("再次暂停后完成"));
    expect(continuePlanningTask).toHaveBeenCalledTimes(2);
    expect(vi.mocked(continuePlanningTask).mock.calls.every(([request]) => request.turnId === "excel-paused-turn")).toBe(true);
    expect(checkIntent).toHaveBeenCalledOnce();
    expect(understandMessage).toHaveBeenCalledOnce();
  });

  it.each(["button", "继续", "接着刚才暂停的任务往下做"] as const)("manual pause continues via %s after checking the original task without replanning", async (via) => {
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析当前表", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "分析当前表", summary: "分析", provider: "model", turnId: "excel-paused-turn" });
    vi.mocked(streamAssistantResponse).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
      hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    vi.mocked(getPlanningTaskStatus).mockResolvedValue({ kind: "planning_pause", turnId: "excel-paused-turn", status: "interrupted",
      revision: 7, lastSequence: 12, canContinue: true, pauseType: "interrupted", reason: "已核实检查点", extensionSeconds: 60,
      activeElapsedSeconds: 3, currentModelId: "model-1", canSwitchModel: false, availableModelConnections: [],
      // 这份现场没读过来源 ⇒ 继续前不需要重读数据；缺了这个布尔，客户端会按
      // fail-closed 要求重读来源，jsdom 里读不到就会正确地停下来。
      continuation: { available: true, code: null, message: "", stage: "after_tool", sourceVerificationRequired: false } } as never);
    vi.mocked(continuePlanningTask).mockResolvedValue({ kind: "answer", provider: "model", message: "Excel 从原进度完成" });
    const view = await send("分析当前表");
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    fireEvent.click(await view.findByRole("button", { name: "暂停任务" }));
    await waitFor(() => expect(view.getByRole("button", { name: /继续分析/ })).toBeTruthy());
    const originalTimelineCount = view.container.querySelectorAll(".task-progress-timeline").length;
    const originalStepCount = view.container.querySelectorAll(".activity-steps > li").length;
    const beforeSegments = segmentSnapshots(view.container);
    expect(getPlanningTaskStatus).toHaveBeenCalledOnce();
    expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value).toBe("");
    if (via === "button") {
      fireEvent.click(await view.findByRole("button", { name: "继续任务" }));
    } else {
      vi.mocked(understandMessage).mockResolvedValueOnce({ kind: "task_recovery", turnId: "excel-paused-turn", action: "continue", provider: "model" });
      fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: via } });
      fireEvent.click(view.getByRole("button", { name: "发送" }));
    }
    await waitFor(() => expect(view.container.textContent).toContain("Excel 从原进度完成"));
    // 继续落在**新的回复段**里；旧段保留暂停时的样子，不被就地改写。
    const after = segmentSnapshots(view.container);
    expect(after).toHaveLength(originalTimelineCount + 1);
    expect(after.slice(0, originalTimelineCount).map((segment) => segment.steps.length))
      .toEqual(beforeSegments.map((segment) => segment.steps.length));
    /*
     * 新段正常结束 ⇒ 过程框**收起成一行**（阶段结论 + 耗时）。
     * 所以"这一步带来了新动作"不能再靠 DOM 里现成的 `<li>` 数量证明，
     * 要先把它展开：收起是呈现，不是内容被删。
     */
    const newCard = view.container
      .querySelectorAll<HTMLElement>(".task-progress-timeline")[originalTimelineCount];
    expect(newCard?.getAttribute("data-expanded")).toBe("false");
    fireEvent.click(newCard!.querySelector<HTMLButtonElement>(".task-progress-toggle")!);
    expect(view.container.querySelectorAll(".activity-steps > li").length)
      .toBeGreaterThan(originalStepCount);
    expect(liveSegmentCount(view.container)).toBeLessThanOrEqual(1);
    // 新段必须在最新一条用户消息之后：用户按下继续之后，反馈就在他这句话下面。
    expect(segmentRowIndex(view.container, originalTimelineCount))
      .toBeGreaterThan(lastUserRowIndex(view.container));
    expect(continuePlanningTask).toHaveBeenCalledOnce();
    expect(understandMessage).toHaveBeenCalledTimes(via === "button" ? 1 : 2);
    expect(vi.mocked(continuePlanningTask).mock.calls[0][0]).toMatchObject({ turnId: "excel-paused-turn", afterSequence: 0 });
    expect(vi.mocked(continuePlanningTask).mock.calls[0][0]).not.toHaveProperty("modelId");
    expect(vi.mocked(continuePlanningTask).mock.calls[0][0]).not.toHaveProperty("prompt");
    expect(streamAssistantResponse).toHaveBeenCalledOnce();
    expect(checkIntent).toHaveBeenCalledOnce();
  });

  it.each(["谢谢帮助", "你是谁"])("ordinary message %s replies without schema, field intent or authorization checks", async (question) => {
    vi.mocked(understandMessage).mockResolvedValue({ kind: "conversation", message: "不客气，随时告诉我。", provider: "model" });
    const view = await send(question);
    await waitFor(() => expect(view.container.textContent).toContain("不客气，随时告诉我。"));
    expect(captureWorkbookStructure).not.toHaveBeenCalled();
    expect(checkIntent).not.toHaveBeenCalled();
    expect(inspectOfficeAuthorizedTarget).not.toHaveBeenCalled();
    expect(view.container.textContent).not.toContain("读取已授权表结构");
    // 结束的过程框收起成一行，展开入口的措辞是「查看过程」（不再是「展开」）。
    expect(view.getByRole("button", { name: "查看过程" }).getAttribute("aria-expanded")).toBe("false");
    fireEvent.click(view.getByRole("button", { name: "查看过程" }));
    // 完成后的普通聊天只保留真实完成说明，不重复准备阶段文案。
    expect(view.container.textContent).toContain("已理解并直接回复；未读取工作表");
    expect(view.container.textContent).not.toContain("本次没有需要你处理的后续操作");
    expect(view.container.querySelector(".task-progress-timeline")?.closest(".message-row")?.querySelector(".message-author")).toBeNull();
    // 名字让给了结论行，而且**这句问话之后**的回复只印一次：过程行与结论行各印一次时，
    // 一段回复会读成两件事（对标的 ExcelMaster 是一段一个抬头）。欢迎语是独立一组，
    // 所以只能按"最后一条用户消息之后"计数。
    const userRow = Array.from(view.container.querySelectorAll(".message-row.user")).pop()!;
    const replyRows = Array.from(view.container.querySelectorAll(".message-row")).filter(
      (row) => userRow.compareDocumentPosition(row) & Node.DOCUMENT_POSITION_FOLLOWING
    );
    expect(replyRows.filter((row) => row.querySelector(".message-author"))).toHaveLength(1);
    expect(replyRows.at(-1)?.querySelector(".message-author")).not.toBeNull();
    expect(view.container.textContent).not.toContain("先判断是否需要工作簿上下文");
    expect(captureWorkbookStructure).not.toHaveBeenCalled();
  });

  it("understands a task, checks authorization, then reads schema and routes fields", async () => {
    vi.mocked(understandMessage).mockImplementation(async () => {
      fixtures.order.push("message");
      return { kind: "needs_workbook_context", summary: "分析当前表", provider: "model" };
    });
    vi.mocked(checkIntent).mockImplementation(async () => {
      fixtures.order.push("field-intent");
      return { kind: "conversation", message: "字段已核对。", provider: "model" };
    });
    const view = await send("分析当前表");
    await waitFor(() => expect(view.container.textContent).toContain("字段已核对。"));
    expect(fixtures.order[0]).toBe("message");
    expect(fixtures.order.indexOf("authorization")).toBeGreaterThan(0);
    expect(fixtures.order.indexOf("schema")).toBeGreaterThan(fixtures.order.indexOf("authorization"));
    expect(fixtures.order.indexOf("field-intent")).toBeGreaterThan(fixtures.order.indexOf("schema"));
    expect(vi.mocked(understandMessage).mock.calls[0][0]).not.toHaveProperty("scope");
  });

  it("message routing failure leaves the composer intact and reads zero fields", async () => {
    vi.mocked(understandMessage).mockRejectedValue(new Error("消息理解失败"));
    const view = await send("检查当前表");
    await waitFor(() => expect(view.container.textContent).toContain("消息理解失败"));
    expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value).toBe("检查当前表");
    expect(captureWorkbookStructure).not.toHaveBeenCalled();
    expect(checkIntent).not.toHaveBeenCalled();
  });

  /*
   * 「发送」的判据是**受理那一刻**就送达，不是等「理解消息」跑完。
   *
   * 这里三条读的是各自独立的现场：① 输入框的值；② 对话里的用户气泡；
   * ③ 运行中发送键的可访问名。旧实现里三样都要等 `understandMessage` 回来才有，
   * 于是运行中的按钮停在「打断并补充」（↑）——用户再按一次 Enter 就会拿着
   * **同一句**去转向重跑。判据必须在"理解还没回来"的窗口里取证。
   */
  it("delivers the message the moment it is sent, without waiting for message understanding", async () => {
    let resolve!: (value: Awaited<ReturnType<typeof understandMessage>>) => void;
    vi.mocked(understandMessage).mockImplementation(() => new Promise((done) => { resolve = done; }));
    const view = await send("检查当前表");
    await waitFor(() => expect(understandMessage).toHaveBeenCalledOnce());

    // ① 输入框立刻空出来。
    expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value).toBe("");
    // ② 这句话已经进了对话。
    expect(view.container.querySelector(".message-row.user")?.textContent).toContain("检查当前表");
    // ③ 运行中的发送键是「暂停」（■）而不是「打断并补充」（↑）。
    expect(view.getByRole("button", { name: "暂停任务" })).toBeTruthy();
    expect(view.queryByRole("button", { name: "打断并补充" })).toBeNull();

    await act(async () => resolve({ kind: "conversation", message: "已回复", provider: "model" }));
    await waitFor(() => expect(view.container.textContent).toContain("已回复"));
    // 送达只发生一次：收尾后仍然只有一条用户气泡。
    expect(view.container.querySelectorAll(".message-row.user")).toHaveLength(1);
  });

  it("selecting a range does not replay a blocked request or duplicate its message", async () => {
    let resolve!: (value: Awaited<ReturnType<typeof understandMessage>>) => void;
    vi.mocked(understandMessage).mockImplementation(() => new Promise((done) => { resolve = done; }));
    const view = await send("检查当前表");
    await waitFor(() => expect(understandMessage).toHaveBeenCalledOnce());
    // 受理即送达：还没理解完就已经有一条用户气泡。
    expect(view.container.querySelectorAll(".message-row.user")).toHaveLength(1);

    // 理解期间文件身份变了 ⇒ 停在"请先核对范围"，并把请求放回输入框。
    vi.mocked(inspectOfficeAuthorizedTarget).mockResolvedValue({ ...fixtures.target, fileKey: "office:another-file" });
    await act(async () => resolve({ kind: "needs_workbook_context", summary: "检查当前表", provider: "model" }));
    await waitFor(() => expect(view.getByRole("button", { name: "检查数据范围" })).toBeTruthy());
    expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value).toBe("检查当前表");

    // 走真实确认链路：点「检查数据范围」→ 确认 → `retrySendRef` 带着
    // `targetConfirmedRetry` 重入 `sendMessage`。
    vi.mocked(inspectOfficeAuthorizedTarget).mockResolvedValue(fixtures.target);
    vi.mocked(checkIntent).mockResolvedValue({ kind: "conversation", message: "已按所选范围继续。", provider: "model" });
    fireEvent.click(view.getByRole("button", { name: "检查数据范围" }));
    fireEvent.click(view.getByRole("button", { name: /跟随当前/ }));
    await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
    fireEvent.click(view.container.querySelector(".scope-close")!);
    expect(view.container.textContent).not.toContain("已按所选范围继续。");

    // 判据：这句话只送达一次 —— 重入不能再补一条气泡。
    expect(view.container.querySelectorAll(".message-row.user")).toHaveLength(1);
    expect(understandMessage).toHaveBeenCalledOnce();
  });

  it("a file change while understanding stops before any schema or field-intent request", async () => {
    let resolve!: (value: Awaited<ReturnType<typeof understandMessage>>) => void;
    vi.mocked(understandMessage).mockImplementation(() => new Promise((done) => { resolve = done; }));
    const view = await send("检查当前表");
    await waitFor(() => expect(understandMessage).toHaveBeenCalledOnce());
    vi.mocked(inspectOfficeAuthorizedTarget).mockResolvedValue({ ...fixtures.target, fileKey: "office:another-file" });
    await act(async () => resolve({ kind: "needs_workbook_context", summary: "检查当前表", provider: "model" }));
    await waitFor(() => expect(view.container.querySelector('.target-scope-notice')).not.toBeNull());
    expect(view.container.querySelector('[role="dialog"]')).toBeNull();
    expect(view.getByRole("button", { name: "检查数据范围" })).toBeTruthy();
    expect(view.container.querySelectorAll(".target-scope-notice button")).toHaveLength(1);
    expect(view.queryByRole("button", { name: "暂不操作" })).toBeNull();
    await act(async () => { window.dispatchEvent(new Event("focus")); });
    expect(view.container.querySelector(".target-scope-notice")).not.toBeNull();
    expect(captureWorkbookStructure).not.toHaveBeenCalled();
    expect(checkIntent).not.toHaveBeenCalled();
    expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value).toBe("检查当前表");
    vi.mocked(checkIntent).mockResolvedValue({ kind: "conversation", message: "已按所选范围继续。", provider: "model" });
    fireEvent.click(view.getByRole("button", { name: "检查数据范围" }));
    await waitFor(() => expect(view.container.querySelector(".scope-popover")).not.toBeNull());
    expect(view.container.querySelector(".target-scope-notice")).toBeNull();
    expect(view.queryByRole("dialog")).toBeNull();
    expect(captureWorkbookStructure).not.toHaveBeenCalled();
    fireEvent.click(view.getByRole("button", { name: /跟随当前/ }));
    await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
    fireEvent.click(view.container.querySelector(".scope-close")!);
    expect(view.container.textContent).not.toContain("已按所选范围继续。");
    expect(checkIntent).not.toHaveBeenCalled();
    expect(captureWorkbookStructure).not.toHaveBeenCalled();
    await waitFor(() => expect(view.container.querySelector(".scope-popover")).toBeNull());
    expect(view.queryByRole("dialog")).toBeNull();
  });

  /*
   * 「继续分析」必须把服务端的读表指针交回本机桥（2026-10-20）。
   *
   * 缺口：首次规划会 `begin`，续跑路径却只把进度交给文案。上一轮收尾时
   * `end()` 已经把桥关掉（绑定还留着给结果卡按引用取值），于是续跑里的
   * `bridgeTool` 指针被交给一条已收尾的桥 —— 谁也不再轮询，最后只能等超时。
   *
   * 这里断言的是**归属身份与次数**：轮询与回执都用了**原任务**的 turnId，
   * 同一个 operationId 只执行一次。
   */
  it("continuing re-opens the local tool bridge for the original task", async () => {
    const OP = "op_" + "9".repeat(32);
    const FP = "d".repeat(64);
    const bridgeCalls: Array<{ url: string; body: Record<string, any> }> = [];
    let markReceipted: (() => void) | null = null;
    const receipted = new Promise<void>((resolve) => { markReceipted = resolve; });
    vi.stubGlobal("fetch", vi.fn(async (url: string, init?: RequestInit) => {
      const target = String(url);
      bridgeCalls.push({ url: target, body: JSON.parse(String(init?.body ?? "{}")) });
      if (target.includes("/poll")) {
        return new Response(JSON.stringify({ calls: [{ operationId: OP, tool: "query_table",
          arguments: { mode: "values", scope: "selected" }, requestFingerprint: FP, expiresAt: 0 }], generation: 1 }),
          { status: 200, headers: { "Content-Type": "application/json" } });
      }
      if (target.includes("/receipt")) {
        markReceipted?.();
        return new Response(JSON.stringify({ status: "accepted", code: null }),
          { status: 200, headers: { "Content-Type": "application/json" } });
      }
      // 其它请求（健康检查等）给一个良性空响应：本用例只关心桥的两条路径，
      // 让无关请求抛错会把暂停卡的渲染一起带走。
      return new Response(JSON.stringify({}), {
        status: 200,
        headers: { "Content-Type": "application/json" }
      });
    }));

    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析当前表", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "分析当前表", summary: "分析", provider: "model", turnId: "excel-paused-turn" });
    vi.mocked(streamAssistantResponse).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
      hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    vi.mocked(getPlanningTaskStatus).mockResolvedValue({ kind: "planning_pause", turnId: "excel-paused-turn", status: "interrupted",
      revision: 7, lastSequence: 12, canContinue: true, pauseType: "interrupted", reason: "已核实检查点", extensionSeconds: 60,
      activeElapsedSeconds: 3, currentModelId: "model-1", canSwitchModel: false, availableModelConnections: [],
      continuation: { available: true, code: null, message: "", stage: "after_tool", sourceVerificationRequired: false } } as never);
    vi.mocked(continuePlanningTask).mockImplementation(async (_request, hooks) => {
      hooks?.onProgress?.({ sequence: 13, stepId: "bridge.query_table", title: "读取授权表",
        bridgeTool: { operationId: OP, tool: "query_table" } } as never);
      // 与生产一致的时序：服务端在等回执，流要等到回执之后才结束。
      await receipted;
      return { kind: "answer", provider: "model", message: "Excel 从原进度完成" };
    });

    const view = await send("分析当前表");
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    fireEvent.click(await view.findByRole("button", { name: "暂停任务" }));
    await waitFor(() => expect(view.getByRole("button", { name: /继续分析/ })).toBeTruthy());
    fireEvent.click(await view.findByRole("button", { name: /继续分析/ }));

    await waitFor(() => expect(bridgeCalls.some((call) => call.url.includes("/receipt"))).toBe(true));
    const poll = bridgeCalls.find((call) => call.url.includes("/poll"));
    expect(poll?.body.taskId).toBe("excel-paused-turn");
    const receipt = bridgeCalls.find((call) => call.url.includes("/receipt"));
    expect(receipt?.body).toMatchObject({ taskId: "excel-paused-turn", operationId: OP });
    expect(bridgeCalls.filter((call) => call.url.includes("/receipt"))).toHaveLength(1);
    expect(vi.mocked(continuePlanningTask)).toHaveBeenCalledTimes(1);
  });

  /*
   * ===================================================================
   * 继续之后的暂停：请求取消 / 界面停止必须分开断言（2026-10-23）
   *
   * 与 WPS 端同一组判据。abort 只能证明"连接断了"，界面必须自己收尾：
   * 活动段必须归零、耗时必须冻住、任何入口都不许再显示"可暂停"。
   * ===================================================================
   */
  it.each(["composer", "card"])(
    "pausing a continued stream via %s leaves no live card and freezes the elapsed time",
    async (via) => {
      vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析当前表", provider: "model" });
      vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "分析当前表", summary: "分析", provider: "model", turnId: "excel-continued-turn" });
      vi.mocked(streamAssistantResponse).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
        hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
      }));
      vi.mocked(getPlanningTaskStatus).mockResolvedValue({ kind: "planning_pause", turnId: "excel-continued-turn", status: "interrupted",
        revision: 7, lastSequence: 12, canContinue: true, pauseType: "interrupted", reason: "已核实检查点", extensionSeconds: 60,
        activeElapsedSeconds: 3, currentModelId: "model-1", canSwitchModel: false, availableModelConnections: [],
        continuation: { available: true, code: null, message: "", stage: "after_tool", sourceVerificationRequired: false } } as never);
      let continueSignal: AbortSignal | undefined;
      vi.mocked(continuePlanningTask).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
        continueSignal = hooks?.signal ?? undefined;
        hooks?.onProgress?.({ sequence: 13, stepId: "plan.continue", title: "继续规划中" } as never);
        hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
      }));

      const view = await send("分析当前表");
      await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
      fireEvent.click(await view.findByRole("button", { name: "暂停任务" }));
      await waitFor(() => expect(view.getByRole("button", { name: /继续分析/ })).toBeTruthy());
      fireEvent.click(await view.findByRole("button", { name: /继续分析/ }));
      await waitFor(() => expect(continuePlanningTask).toHaveBeenCalledOnce());
      await waitFor(() => expect(liveSegmentCount(view.container)).toBe(1));
      await waitFor(() => expect(view.getByRole("button", { name: via === "composer" ? "暂停任务" : /^暂停$/ })).toBeTruthy());

      fireEvent.click(view.getByRole("button", { name: via === "composer" ? "暂停任务" : /^暂停$/ }));

      // ① 请求取消：只证明连接断了。
      await waitFor(() => expect(continueSignal?.aborted).toBe(true));

      // ② 界面停止：活动段必须为 0（不是"≤1"）。
      await waitFor(() => expect(liveSegmentCount(view.container)).toBe(0));

      // ③ 耗时冻结：最后一张卡的标题行文本在真实时间 1.1s 后逐字不变。
      const headers = () => [...view.container.querySelectorAll<HTMLElement>(".task-progress-heading strong")];
      const frozen = headers().at(-1)?.textContent ?? "";
      await act(async () => { await new Promise<void>((resolve) => setTimeout(resolve, 1100)); });
      expect(headers().at(-1)?.textContent ?? "").toBe(frozen);

      // ④ 不再提供任何"可暂停"入口（输入框与卡片都算）。
      expect(view.queryByRole("button", { name: "暂停任务" })).toBeNull();
      expect(view.queryByRole("button", { name: /^暂停$/ })).toBeNull();
    }
  );

  /*
   * 「请求已在飞」不等于「界面已提交可暂停状态」—— 这条用例把那一帧钉住。
   *
   * 真实证据（outputs/pause-fix-2026-10-23/evidence/race-probe.log）：在第二轮
   * `checkIntent` **派发的那一刻**同步读 DOM，只有 `▶ / aria-label=继续任务`、
   * 活动卡 0 张，一个暂停按钮都没有；一个宏任务之后才变成 `■` + 1 张活动卡。
   *
   * 所以点按钮之前必须等"按钮真的可交互"，而不是等"请求已经发出"。这条用例
   * 直接断言：同一帧里 `getByRole`（立即查）抛，`findByRole`（等待式）拿到且可点。
   * 谁把这里改回立即查，这条会立刻红。
   */
  it("a pause click after continue waits for the button to be painted, not just for the request", async () => {
    const taskId = "excel-race-window";
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析当前表", provider: "model" });
    vi.mocked(getPlanningTaskStatus).mockResolvedValue({ kind: "planning_pause", turnId: taskId,
      status: "interrupted", revision: 3, lastSequence: 9, canContinue: true, pauseType: "interrupted",
      reason: "已核实检查点", extensionSeconds: 60, activeElapsedSeconds: 2, currentModelId: "model-1",
      canSwitchModel: false, availableModelConnections: [],
      continuation: { available: true, code: null, message: "", stage: "after_tool", sourceVerificationRequired: false } } as never);
    const probe: Array<{ at: number; immediate: boolean }> = [];
    let view: ReturnType<typeof render> | null = null;
    let calls = 0;
    vi.mocked(checkIntent).mockImplementation(() => {
      calls += 1;
      if (view) {
        // 这一瞬就是"请求已在飞、界面还没提交"的那一帧：立即查必须查不到。
        probe.push({ at: calls, immediate: Boolean(view.queryByRole("button", { name: "暂停任务" })) });
      }
      // 一直挂着：暂停要靠 abort 把它打断，不是靠它自己回包。
      return new Promise(() => {});
    });

    view = await send("分析当前表");
    await waitFor(() => expect(checkIntent).toHaveBeenCalledOnce());
    // 第一轮：等到按钮真的出现再点（这一轮派发时界面已提交）。
    fireEvent.click(await view.findByRole("button", { name: "暂停任务" }));
    await waitFor(() => expect(view!.queryByRole("button", { name: "暂停任务" })).toBeNull());
    /*
     * 已暂停（且有真检查点）⇒ 主按钮就是「继续任务」，不是"灰掉的 ■"、也不是"发送"。
     * 按钮只有一颗、位置只有一个：右下角。
     */
    const continueButton = view.getByRole("button", { name: "继续任务" });
    expect(continueButton.hasAttribute("disabled")).toBe(false);
    expect(continueButton.closest(".composer")).not.toBeNull();
    /*
     * 而回车**不会**顺手继续：回车只表达"把这段话发出去"，此时输入框是空的，
     * 所以什么也不该发生。谁把分派改回"运行中/暂停中按回车就继续"，这条会立刻红。
     */
    fireEvent.keyDown(view.container.querySelector("textarea")!, { key: "Enter" });
    await act(async () => {});
    expect(checkIntent).toHaveBeenCalledOnce();
    // 输入框上方只剩一句状态说明，没有第二颗按钮。
    expect(view.getByRole("button", { name: "继续任务" }).closest(".task-execution-notice")).toBeNull();
    // 第二轮：点「继续任务」触发的这次派发，界面还没提交。
    fireEvent.click(await view.findByRole("button", { name: "继续任务" }));
    await waitFor(() => expect(checkIntent).toHaveBeenCalledTimes(2));

    expect(probe.map((item) => item.at)).toEqual([1, 2]);
    // 判别点：第二轮派发那一刻，暂停按钮**确实**还不在 DOM 里。
    expect(probe[1].immediate).toBe(false);
    // 而等待式查找在同一帧之后立刻拿到它，并且可点。
    const pauseButton = await view.findByRole("button", { name: "暂停任务" });
    expect(pauseButton.hasAttribute("disabled")).toBe(false);
    expect(pauseButton.textContent?.trim()).toBe("■");
  });

  /*
   * 连续三轮「暂停 → 继续」，**每一轮**都断言，不看最后一轮交差。
   *
   * 为什么必须逐轮：第一轮的收尾最容易挂在后来的 controller 上 —— 只验最后一轮
   * 时，"第一轮其实没停"会被第三轮的正常掩盖。
   *
   * 每一轮验七件事：
   *   ① 运行中输入框按钮就是 ■（可暂停），不是 ↑/…；
   *   ② 按下之后**当前这一轮**的 signal 已 abort（请求取消）；
   *   ③ 暂停收尾完成的那一刻，活动卡**严格等于 0**（不是 ≤1）；
   *   ④ 冻结段的文本此后逐字不变（后续事件不改写它）；
   *   ⑤ 点继续之后**恰好 1 张**活动卡（恢复运行）；
   *   ⑥ 新进展落在被冻结那一段**之后**，不是就地改写；
   *   ⑦ 仍是同一个任务：理解/意图/首次流各只一次，continuePlanningTask 每次同 turnId。
   */
  it.each(["composer", "card"])("three pause/continue rounds via %s settle to zero live cards each time and keep one task identity", async (entry) => {
    const taskId = "excel-three-rounds";
    let resolvePause!: (value: { turnId: string; pauseRequested: boolean }) => void;
    vi.mocked(pausePlanningTask).mockImplementationOnce(() => new Promise((resolve) => { resolvePause = resolve; }));
    const statusPayload = { kind: "planning_pause", turnId: taskId, status: "interrupted",
      revision: 7, lastSequence: 12, canContinue: true, pauseType: "interrupted", reason: "已核实检查点",
      extensionSeconds: 60, activeElapsedSeconds: 3, currentModelId: "model-1", canSwitchModel: false,
      availableModelConnections: [],
      continuation: { available: true, code: null, message: "", stage: "after_tool", sourceVerificationRequired: false } };
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析当前表", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "分析当前表", summary: "分析", provider: "model", turnId: taskId });
    vi.mocked(getPlanningTaskStatus).mockResolvedValue(statusPayload as never);
    const signals: AbortSignal[] = [];
    vi.mocked(streamAssistantResponse).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
      if (hooks?.signal) signals.push(hooks.signal);
      hooks?.onProgress?.({ operationId: taskId, sequence: 1, stepId: "plan.start", title: "规划中" } as never);
      hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    vi.mocked(continuePlanningTask).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
      if (hooks?.signal) signals.push(hooks.signal);
      hooks?.onProgress?.({ operationId: taskId, sequence: 13, stepId: "plan.continue", title: "继续规划中" } as never);
      hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));

    const view = await send("分析当前表");
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());

    for (const round of [1, 2, 3]) {
      // —— 运行中：输入框按钮就是「暂停任务」（可暂停），且本轮请求已在飞 ——
      const pauseButton = await view.findByRole("button", { name: entry === "composer" ? "暂停任务" : /^暂停$/ });
      expect(pauseButton.hasAttribute("disabled")).toBe(false);
      await waitFor(() => expect(liveSegmentCount(view.container)).toBe(1));

      fireEvent.click(pauseButton);

      // —— ② 请求取消：本轮请求的 signal 已 abort ——
      await waitFor(() => expect(signals.length).toBe(round));
      expect(signals[round - 1].aborted).toBe(true);

      // —— ③ 暂停收尾完成：活动卡**严格 0**（不是"≤1"）——
      await waitFor(() => expect(liveSegmentCount(view.container)).toBe(0));
      // 收尾结论写在这一段上（不是留在"进行中"）。
      expect(view.container.textContent).toContain("已暂停");

      // —— ④ 冻结：执行耗时与步骤行此后逐字不变（后续事件不改写这一段）——
      /*
       * 冻结的是**执行耗时**（标题行 `… · N 步 · 91ms`）与每一行步骤。
       * 标题下那行 `暂停 X · 等待进度 Y` 里的"等待进度"会继续走，这是
       * `finishTimelineAt` 的既定口径：有服务端任务时本端只确认"停止接收"，
       * 后台可能仍在跑，所以那一段记"等待"而不是"暂停"（卡片上也有 tooltip）。
       * 把它当缺陷改掉，等于把"后台仍在执行"谎报成"已暂停"。
       */
      const titles = () => [...view.container.querySelectorAll<HTMLElement>(".task-progress-heading strong")]
        .map((node) => node.textContent ?? "");
      const frozenTitles = titles();
      const frozenSteps = segmentSnapshots(view.container).map((segment) => segment.steps);
      await act(async () => { await new Promise<void>((resolve) => setTimeout(resolve, 1100)); });
      expect(titles()).toEqual(frozenTitles);
      expect(segmentSnapshots(view.container).map((segment) => segment.steps)).toEqual(frozenSteps);

      // —— ⑤⑥ 继续：恰好 1 张活动卡，且落在被冻结那一段之后 ——
      const continueButton = await view.findByRole("button", { name: /继续分析/ });
      fireEvent.click(continueButton);
      await waitFor(() => expect(continuePlanningTask).toHaveBeenCalledTimes(round));
      await waitFor(() => expect(liveSegmentCount(view.container)).toBe(1));
      const after = segmentSnapshots(view.container);
      // 每轮「继续」各开一段（2/3/4 段），不是在旧段里就地改写。
      expect(after).toHaveLength(round + 1);
      const liveRow = segmentRowIndex(view.container, after.length - 1);
      const frozenRow = segmentRowIndex(view.container, after.length - 2);
      expect(liveRow).toBeGreaterThanOrEqual(0);
      expect(liveRow).toBeGreaterThan(frozenRow);
      expect(liveRow).toBeGreaterThan(lastUserRowIndex(view.container));
      if (round === 1) {
        await act(async () => { resolvePause({ turnId: taskId, pauseRequested: true }); });
        expect(view.container.querySelector(".composer-pause-note")).toBeNull();
      }
      // 恢复运行**没有改写**刚才冻结的那几段（新进展只写在新的一段里）。
      expect(after.slice(0, frozenSteps.length).map((segment) => segment.steps)).toEqual(frozenSteps);

      // —— ⑦ 同一任务，没有第二个任务、没有重复恢复请求 ——
      expect(vi.mocked(continuePlanningTask).mock.calls.at(-1)?.[0]?.turnId).toBe(taskId);
    }

    // 三轮之后：没有第二套任务身份，活动段仍然恰好一张。
    expect(understandMessage).toHaveBeenCalledOnce();
    expect(checkIntent).toHaveBeenCalledOnce();
    expect(streamAssistantResponse).toHaveBeenCalledOnce();
    expect(continuePlanningTask).toHaveBeenCalledTimes(3);
    expect(new Set(vi.mocked(continuePlanningTask).mock.calls.map((call) => call[0]?.turnId))).toEqual(new Set([taskId]));
    expect(liveSegmentCount(view.container)).toBe(1);
  });

  /*
   * 迟到进度：暂停之后，被取消的那条流**仍可能**再送一条事件（读表回执、
   * 步骤更新都还在路上）。它只能算历史 —— 不得把"运行中 / 可暂停"重新点亮，
   * 也不得改写已经冻结的那一段。
   */
  it("a late progress after pause never re-opens the running state", async () => {
    const taskId = "excel-late-progress";
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析当前表", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "分析当前表", summary: "分析", provider: "model", turnId: taskId });
    let hooks: { onProgress?: (event: unknown) => void } | undefined;
    vi.mocked(streamAssistantResponse).mockImplementation((_request, streamHooks) => new Promise((_resolve, reject) => {
      hooks = streamHooks as { onProgress?: (event: unknown) => void };
      streamHooks?.onProgress?.({ operationId: taskId, sequence: 1, stepId: "plan.start", title: "规划中" } as never);
      streamHooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));

    const view = await send("分析当前表");
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    fireEvent.click(await view.findByRole("button", { name: "暂停任务" }));
    await waitFor(() => expect(liveSegmentCount(view.container)).toBe(0));
    const frozenTitles = [...view.container.querySelectorAll<HTMLElement>(".task-progress-heading strong")]
      .map((node) => node.textContent ?? "");
    const frozenSteps = segmentSnapshots(view.container).map((segment) => segment.steps);

    await act(async () => {
      hooks?.onProgress?.({ operationId: taskId, sequence: 2, stepId: "plan.late", title: "迟到的读表回执" });
      hooks?.onProgress?.({ operationId: taskId, sequence: 3, stepId: "plan.late-2", title: "又一条迟到进度" });
    });

    expect(liveSegmentCount(view.container)).toBe(0);
    expect(view.queryByRole("button", { name: "暂停任务" })).toBeNull();
    expect(view.queryByRole("button", { name: /^暂停$/ })).toBeNull();
    expect([...view.container.querySelectorAll<HTMLElement>(".task-progress-heading strong")]
      .map((node) => node.textContent ?? "")).toEqual(frozenTitles);
    expect(segmentSnapshots(view.container).map((segment) => segment.steps)).toEqual(frozenSteps);
  });

  /*
   * 暂停请求的三种回执必须说三种话。
   *
   * 关键红线：只有 `pauseRequested: true` 才能说"服务端已接受暂停请求"；
   * 失败/没有在跑的任务都不许被说成"已暂停" —— 本端停止接收 ≠ 服务端已暂停。
   */
  it("pause timeout stays unconfirmed even if its acknowledgement arrives late", async () => {
    const taskId = "excel-pause-timeout";
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "分析", summary: "分析", provider: "model", turnId: taskId });
    let resolvePause!: (value: { turnId: string; pauseRequested: boolean }) => void;
    vi.mocked(pausePlanningTask).mockImplementationOnce(() => new Promise(resolve => { resolvePause = resolve; }));
    vi.mocked(streamAssistantResponse).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
      hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    const view = await send("分析当前表");
    const button = await view.findByRole("button", { name: "暂停任务" });
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    vi.useFakeTimers();
    try {
      await act(async () => { fireEvent.click(button); });
      await act(async () => { await vi.advanceTimersByTimeAsync(7999); });
      expect(view.container.querySelector(".composer-pause-note")?.getAttribute("data-phase")).toBe("requesting");
      await act(async () => { await vi.advanceTimersByTimeAsync(1); });
      const note = view.container.querySelector(".composer-pause-note");
      expect(note?.getAttribute("data-phase")).toBe("unconfirmed");
      expect(note?.textContent).toContain("8 秒内没有确认");
      expect(liveSegmentCount(view.container)).toBe(0);
      await act(async () => { resolvePause({ turnId: taskId, pauseRequested: true }); });
      expect(note?.getAttribute("data-phase")).toBe("unconfirmed");
    } finally { vi.useRealTimers(); }
  });

  it.each([
    ["失败", () => Promise.reject(new Error("network down")), "unconfirmed", "未能确认后台暂停"],
    ["被接受", () => Promise.resolve({ turnId: "excel-pause-api", pauseRequested: true }), "acknowledged", "服务端已接受暂停请求"],
    ["发现没有在跑", () => Promise.resolve({ turnId: "excel-pause-api", pauseRequested: false }), "local-stopped", "此刻不在运行"]
  ])("pause request outcome %# never claims more than the server confirmed", async (_label, impl, phase, wording) => {
    const taskId = "excel-pause-api";
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析当前表", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "分析当前表", summary: "分析", provider: "model", turnId: taskId });
    vi.mocked(pausePlanningTask).mockImplementation(impl as never);
    vi.mocked(streamAssistantResponse).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
      hooks?.onProgress?.({ operationId: taskId, sequence: 1, stepId: "plan.start", title: "规划中" } as never);
      hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));

    const view = await send("分析当前表");
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    fireEvent.click(await view.findByRole("button", { name: "暂停任务" }));

    await waitFor(() => expect(pausePlanningTask).toHaveBeenCalledOnce());
    expect(vi.mocked(pausePlanningTask).mock.calls[0][0]).toBe(taskId);
    const note = view.container.querySelector<HTMLElement>(".composer-pause-note");
    await waitFor(() => expect(note?.getAttribute("data-phase")).toBe(phase));
    expect(note?.textContent).toContain(wording);
    if (phase !== "acknowledged") {
      expect(note?.textContent).not.toContain("服务端已接受暂停请求");
    }
    // 无论回执是什么，本端都已经停了（"请求取消"是本端唯一能确定的事实）。
    expect(liveSegmentCount(view.container)).toBe(0);
  });

});

/**
 * 读表桥没立起来时，**请求不许发出去**（2026-10-20）。
 *
 * 缺口很具体：面板 `await` 了 `begin` / `ensureLive`，却从来不看返回值。桥初始化
 * 期间被取消、被收尾、或被另一轮取代时，返回值是 `false`，请求照样发出去 ——
 * 服务端一发出 `bridgeTool` 指针就交给一条不存在的桥，模型只能等到超时。表现和
 * "继续分析之后一动不动"一模一样，但根因在**发请求之前**。
 *
 * 这两条用例断言的是调用层：请求真的没发（`streamAssistantResponse` /
 * `continuePlanningTask` 未被调用），并且界面如实说明停在哪一步。
 */
describe("读表桥没立起来时不发请求", () => {
  it("首次规划：begin 返回 false ⇒ 不发规划请求，并如实说明停在哪一步", async () => {
    vi.mocked(understandMessage).mockResolvedValue({
      kind: "needs_workbook_context", summary: "分析当前表", provider: "model"
    });
    vi.mocked(checkIntent).mockResolvedValue({
      kind: "proceed", confirmedPrompt: "分析当前表", summary: "分析",
      provider: "model", turnId: "turn-plan-1"
    });
    bridgeStub.begin = true;

    const view = await send("分析当前表");

    await waitFor(() => expect(view.container.textContent).toContain("没有发送规划请求"));
    // 关键断言：一次都没有发出去（发了就是一次注定没人服务的调用）。
    expect(streamAssistantResponse).not.toHaveBeenCalled();
  });

  it("继续分析：ensureLive 返回 false ⇒ 不发继续请求，恢复卡留在可重试的状态", async () => {
    vi.mocked(understandMessage).mockResolvedValue({
      kind: "needs_workbook_context", summary: "分析当前表", provider: "model"
    });
    vi.mocked(checkIntent).mockResolvedValue({
      kind: "proceed", confirmedPrompt: "分析当前表", summary: "分析",
      provider: "model", turnId: "excel-paused-turn"
    });
    vi.mocked(streamAssistantResponse).mockImplementation(
      (_request, hooks) => new Promise((_resolve, reject) => {
        hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
      })
    );
    vi.mocked(getPlanningTaskStatus).mockResolvedValue({
      kind: "planning_pause", turnId: "excel-paused-turn", status: "interrupted",
      revision: 7, lastSequence: 12, canContinue: true, pauseType: "interrupted",
      reason: "已核实检查点", extensionSeconds: 0, activeElapsedSeconds: 3,
      currentModelId: "model-1", canSwitchModel: false, availableModelConnections: [],
      continuation: { available: true, code: null, message: "", stage: "after_tool", sourceVerificationRequired: false }
    } as never);
    // 只有"继续"这一次建立会失败；之前的首次规划仍然用真实返回值。
    bridgeStub.ensureLive = true;

    const view = await send("分析当前表");
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    fireEvent.click(await view.findByRole("button", { name: "暂停任务" }));
    await waitFor(() => expect(view.getByRole("button", { name: "继续分析" })).toBeTruthy());

    fireEvent.click(view.getByRole("button", { name: "继续分析" }));

    await waitFor(() => expect(view.container.textContent).toContain("没有发送继续请求"));
    expect(continuePlanningTask).not.toHaveBeenCalled();
    // 恢复卡还在：用户可以再点一次，而不是被吞掉一次尝试。
    expect(view.getByRole("button", { name: "继续分析" })).toBeTruthy();
  });
});


it("places a continued write plan in the latest reply segment rather than the old pause message", async () => {
  const taskId = "excel-continued-write-preview";
  vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "生成预览", provider: "model" });
  vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "生成预览", summary: "分析", provider: "model", turnId: taskId });
  vi.mocked(streamAssistantResponse).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
    hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
  }));
  vi.mocked(getPlanningTaskStatus).mockResolvedValue({ kind: "planning_pause", turnId: taskId, status: "interrupted",
    revision: 7, lastSequence: 12, canContinue: true, pauseType: "interrupted", reason: "已核实检查点", extensionSeconds: 60,
    activeElapsedSeconds: 3, currentModelId: "model-1", canSwitchModel: false, availableModelConnections: [],
    continuation: { available: true, code: null, message: "", stage: "after_tool", sourceVerificationRequired: false } } as never);
  vi.mocked(continuePlanningTask).mockResolvedValue({ kind: "plan", provider: "model", turnId: taskId, plan: {
    id: "continued-preview", title: "续跑后的写入预览", summary: "等待用户确认", assumptions: [], warnings: [],
    actions: [{ type: "writeValues", sheet: "Sheet1", range: "A1", values: [[42]] }], acceptanceCriteria: []
  } } as never);
  const view = await send("生成预览");
  await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
  fireEvent.click(await view.findByRole("button", { name: "暂停任务" }));
  fireEvent.click(await view.findByRole("button", { name: /继续分析/ }));
  const title = await waitFor(() => {
    const element = view.container.querySelector(".inline-plan-title strong");
    expect(element?.textContent).toBe("续跑后的写入预览");
    return element!;
  });
  await waitFor(() => expect(liveSegmentCount(view.container)).toBe(0));
  const latestCard = [...view.container.querySelectorAll(".task-progress-timeline")].at(-1);
  expect(title.closest(".message-row")).toBe(latestCard?.closest(".message-row"));
  expect(view.container.querySelectorAll(".inline-plan")).toHaveLength(1);
  expect(continuePlanningTask).toHaveBeenCalledOnce();
});
