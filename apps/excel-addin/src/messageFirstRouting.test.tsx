// @vitest-environment jsdom
import { act, cleanup, fireEvent, render, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import { checkIntent, understandMessage, streamAssistantResponse, getPlanningTaskStatus, continuePlanningTask } from "./api";
import { captureWorkbookCatalog, captureWorkbookStructure } from "./excel";
import { inspectOfficeAuthorizedTarget } from "./officeAuthorizedTarget";

const fixtures = vi.hoisted(() => ({
  order: [] as string[],
  snapshot: { name: "book.xlsx", capturedAt: "2026-10-08T00:00:00Z", activeWorksheet: "Sheet1", worksheets: [
    { name: "Sheet1", usedRange: "A1:B3", rowCount: 3, columnCount: 2, headers: ["Category", "Level"], dataRows: [], truncated: false }
  ] },
  target: { fileKey: "office:book-1", name: "book.xlsx", sheets: [{ id: "sheet-1", name: "Sheet1" }], selected: ["sheet-1"] }
}));

vi.mock("./api", async (importOriginal) => ({
  ...await importOriginal<typeof import("./api")>(),
  checkHealth: vi.fn(async () => ({ status: "ok", modelConfigured: true })),
  listModels: vi.fn(async () => ({ defaultModelId: "model-1", models: [{ id: "model-1", label: "模型", modelId: "test-model", baseUrl: "http://localhost", supportsVision: true, available: true }] })),
  getModelSettings: vi.fn(async () => ({ connectionCount: 1 })),
  understandMessage: vi.fn(), checkIntent: vi.fn(), streamAssistantResponse: vi.fn(), getPlanningTaskStatus: vi.fn(), continuePlanningTask: vi.fn()
}));
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

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
  vi.spyOn(console, "log").mockImplementation(() => {});
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
  it("a submitted answer survives an intent network error; changing that answer replaces it on retry", async () => {
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "填写等级", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValueOnce({ kind: "clarification", provider: "model", clarification: {
      id: "retry-mode", summary: "填写口径", question: "如何填写？", reason: "需要口径", scopeLabel: "Sheet1", options: []
    } }).mockRejectedValueOnce(new TypeError("确认连接中断"))
      .mockResolvedValue({ kind: "proceed", confirmedPrompt: "根据类别填写等级", summary: "填写等级", provider: "model" });
    vi.mocked(streamAssistantResponse).mockResolvedValue({ kind: "answer", provider: "model", message: "收到新口径" });
    const view = await send("根据类别填写等级");
    await waitFor(() => expect(view.container.querySelector(".clarification-custom-input textarea")).not.toBeNull());
    fireEvent.change(view.container.querySelector(".clarification-custom-input textarea")!, { target: { value: "逐行填写" } });
    fireEvent.click(view.getByRole("button", { name: "提交补充" }));
    await waitFor(() => expect(view.container.textContent).toContain("确认连接中断"));
    const history = JSON.parse(localStorage.getItem("excel-bro.chat.v4")!);
    const card = history.conversations.flatMap((item: { messages: unknown[] }) => item.messages)
      .find((item: { clarification?: { id: string } }) => item.clarification?.id === "retry-mode");
    expect(card.clarification.taskContext.confirmations).toEqual([
      { questionId: "retry-mode", label: "填写口径", value: "逐行填写" }
    ]);
    fireEvent.change(view.container.querySelector(".clarification-custom-input textarea")!, { target: { value: "分组汇总" } });
    fireEvent.click(view.getByRole("button", { name: "提交补充" }));
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
    await waitFor(() => expect(view.getByRole("button", { name: /逐行填写/ })).toBeTruthy());
    fireEvent.click(view.getByRole("button", { name: /逐行填写/ }));
    await waitFor(() => expect(view.getByRole("button", { name: /空值不写入/ })).toBeTruthy());
    fireEvent.click(view.getByRole("button", { name: /空值不写入/ }));
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
      fireEvent.click(view.getByRole("button", { name: "更多" }));
      fireEvent.click(view.getByRole("menuitem", { name: /^历史对话/ }));
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
    fireEvent.click(view.getByRole("button", { name: /^暂停$/ }));
    await waitFor(() => expect(getPlanningTaskStatus).toHaveBeenCalledOnce());
    await waitFor(() => expect(view.container.querySelector('[aria-label="暂停任务"]')).toBeNull());
    expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value).toBe("");
    expect(continuePlanningTask).not.toHaveBeenCalled();
    expect(checkIntent).toHaveBeenCalledOnce();
    expect(understandMessage).toHaveBeenCalledOnce();
    if (kind === "running" || kind === "unavailable") expect(view.getByRole("button", { name: "核对任务状态" })).toBeTruthy();
    if (kind === "no_checkpoint") expect(view.container.textContent).toContain("不能从断点继续");
    if (kind === "completed") expect(view.container.textContent).toContain("原任务结果已取回");
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
    fireEvent.click(view.getByRole("button", { name: /^暂停$/ }));
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
    fireEvent.click(view.getByRole("button", { name: /^暂停$/ }));
    await waitFor(() => expect(view.container.textContent).toContain("尚未形成可续跑检查点"));
    expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value).toBe("");
    expect(view.container.querySelector(".message-row.user")?.textContent).toContain("分析当前表");
    expect(getPlanningTaskStatus).not.toHaveBeenCalled();
    expect(captureWorkbookStructure).not.toHaveBeenCalled();
    expect(checkIntent).not.toHaveBeenCalled();
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
    fireEvent.click(view.getByRole("button", { name: "暂停任务" }));
    await waitFor(() => expect(view.container.querySelector(".target-scope-notice")).not.toBeNull());
    await waitFor(() => expect((view.getByRole("button", { name: "核对任务状态" }) as HTMLButtonElement).disabled).toBe(false));
    expect(getPlanningTaskStatus).not.toHaveBeenCalled();
    expect(continuePlanningTask).not.toHaveBeenCalled();
    expect(understandMessage).toHaveBeenCalledOnce();
    expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value).toBe("");
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
      activeElapsedSeconds: 3, currentModelId: "model-1", canSwitchModel: false, availableModelConnections: [] } as never);
    vi.mocked(continuePlanningTask).mockImplementationOnce(waitForPause)
      .mockResolvedValueOnce({ kind: "answer", provider: "model", message: "再次暂停后完成" });
    const view = await send("分析当前表");
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    fireEvent.click(view.getByRole("button", { name: "暂停任务" }));
    await waitFor(() => expect(view.getByRole("button", { name: "继续任务" })).toBeTruthy());
    fireEvent.click(view.getByRole("button", { name: "继续任务" }));
    await waitFor(() => expect(continuePlanningTask).toHaveBeenCalledOnce());
    fireEvent.click(view.getByRole("button", { name: "暂停任务" }));
    await waitFor(() => expect(getPlanningTaskStatus).toHaveBeenCalledTimes(3));
    await waitFor(() => expect(view.getByRole("button", { name: "继续任务" })).toBeTruthy());
    expect(continuePlanningTask).toHaveBeenCalledOnce();
    expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value).toBe("");
    fireEvent.click(view.getByRole("button", { name: "继续任务" }));
    await waitFor(() => expect(view.container.textContent).toContain("再次暂停后完成"));
    expect(continuePlanningTask).toHaveBeenCalledTimes(2);
    expect(vi.mocked(continuePlanningTask).mock.calls.every(([request]) => request.turnId === "excel-paused-turn")).toBe(true);
    expect(checkIntent).toHaveBeenCalledOnce();
    expect(understandMessage).toHaveBeenCalledOnce();
  });

  it("manual pause checks the original task before an explicit continue without replanning", async () => {
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析当前表", provider: "model" });
    vi.mocked(checkIntent).mockResolvedValue({ kind: "proceed", confirmedPrompt: "分析当前表", summary: "分析", provider: "model", turnId: "excel-paused-turn" });
    vi.mocked(streamAssistantResponse).mockImplementation((_request, hooks) => new Promise((_resolve, reject) => {
      hooks?.signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    vi.mocked(getPlanningTaskStatus).mockResolvedValue({ kind: "planning_pause", turnId: "excel-paused-turn", status: "interrupted",
      revision: 7, lastSequence: 12, canContinue: true, pauseType: "interrupted", reason: "已核实检查点", extensionSeconds: 60,
      activeElapsedSeconds: 3, currentModelId: "model-1", canSwitchModel: false, availableModelConnections: [] } as never);
    vi.mocked(continuePlanningTask).mockResolvedValue({ kind: "answer", provider: "model", message: "Excel 从原进度完成" });
    const view = await send("分析当前表");
    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    fireEvent.click(view.getByRole("button", { name: "暂停任务" }));
    await waitFor(() => expect(view.getByRole("button", { name: /继续处理/ })).toBeTruthy());
    expect(getPlanningTaskStatus).toHaveBeenCalledOnce();
    expect(view.container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value).toBe("");
    fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
    fireEvent.keyDown(view.container.querySelector(".composer textarea")!, { key: "Enter" });
    expect(understandMessage).toHaveBeenCalledOnce();
    expect(continuePlanningTask).not.toHaveBeenCalled();
    fireEvent.click(view.getByRole("button", { name: /继续处理/ }));
    await waitFor(() => expect(view.container.textContent).toContain("Excel 从原进度完成"));
    expect(continuePlanningTask).toHaveBeenCalledOnce();
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
    expect(view.getByRole("button", { name: "展开" }).getAttribute("aria-expanded")).toBe("false");
    fireEvent.click(view.getByRole("button", { name: "展开" }));
    expect(view.container.textContent).toContain("理解消息");
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

  it("the retry after target confirmation does not deliver the same message twice", async () => {
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
    await waitFor(() => expect((view.getByRole("button", { name: "确认" }) as HTMLButtonElement).disabled).toBe(false));
    fireEvent.click(view.getByRole("button", { name: "确认" }));
    await waitFor(() => expect(view.container.textContent).toContain("已按所选范围继续。"));

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
    await waitFor(() => expect((view.getByRole("button", { name: "确认" }) as HTMLButtonElement).disabled).toBe(false));
    fireEvent.click(view.getByRole("button", { name: "确认" }));
    await waitFor(() => expect(view.container.textContent).toContain("已按所选范围继续。"));
    expect(checkIntent).toHaveBeenCalledOnce();
    expect(captureWorkbookStructure).toHaveBeenCalledOnce();
    expect(view.container.querySelector(".scope-popover")).toBeNull();
    expect(view.queryByRole("dialog")).toBeNull();
  });
});
