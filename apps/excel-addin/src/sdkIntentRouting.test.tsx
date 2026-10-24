// @vitest-environment jsdom
/**
 * Excel 面板级接线：SDK 适配器下的意图路径。
 *
 * 这里驱动的是**真实面板**（`<App />` + 真实 composer 点击），不是纯函数：
 * 目标是证明"模型选了 SDK 时，本地不再预置计算结果，读表推迟到工具桥指针"。
 * 预置的 `dataResults` 会被服务端"不上传数据行"守卫拒绝，新工具桥就永远
 * 拿不到执行权 —— 这正是本轮要修的缺陷。
 */
import { cleanup, fireEvent, render, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import { checkIntent, understandMessage, streamAssistantResponse } from "./api";
import {
  captureWorkbookCatalog,
  captureWorkbookSourceFingerprint,
  captureSelectionContext,
  currentOfficeWorkbookIdentity
} from "./excel";
import { inspectOfficeAuthorizedTarget } from "./officeAuthorizedTarget";
import { executeQueryTableTool } from "./dataTools";

const fixtures = vi.hoisted(() => ({
  snapshot: {
    name: "book.xlsx",
    workbookIdentity: "url:https://contoso.sharepoint.com/book.xlsx",
    sourceFingerprint: "fp-source-1",
    capturedAt: "2026-10-09T00:00:00Z",
    activeWorksheet: "Sheet1",
    worksheets: [
      {
        name: "Sheet1",
        sourceSheet: "Sheet1",
        usedRange: "A1:B3",
        rowCount: 3,
        columnCount: 2,
        headers: ["Category", "Level"],
        dataRows: [],
        truncated: false
      }
    ]
  },
  target: {
    fileKey: "office:book-1",
    name: "book.xlsx",
    sheets: [{ id: "sheet-1", name: "Sheet1" }],
    selected: ["sheet-1"]
  }
}));

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
  continuePlanningTask: vi.fn()
}));

vi.mock("./excel", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./excel")>()),
  isRunningInExcel: () => true,
  captureWorkbookCatalog: vi.fn(async () => ({
    ...fixtures.snapshot,
    catalogOnly: true,
    worksheets: fixtures.snapshot.worksheets.map((sheet) => ({
      ...sheet,
      headers: [],
      dataRows: []
    }))
  })),
  captureWorkbookStructure: vi.fn(async () => fixtures.snapshot),
  captureWorkbook: vi.fn(async () => fixtures.snapshot),
  captureWorkbookSourceFingerprint: vi.fn(async () => "fp-source-1"),
  captureSelectionContext: vi.fn(async () => ({
    activeWorksheet: "Sheet1",
    selectedRange: "A1:B3"
  })),
  watchWorkbookStructureChanges: vi.fn(async () => () => {}),
  currentOfficeWorkbookIdentity: vi.fn(
    () => "url:https://contoso.sharepoint.com/book.xlsx"
  )
}));

vi.mock("./officeAuthorizedTarget", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./officeAuthorizedTarget")>()),
  inspectOfficeAuthorizedTarget: vi.fn(async () => fixtures.target),
  inspectOfficeScopeAdmission: vi.fn(async (names: readonly string[]) => ({
    items: names.map((id) => ({ id, status: "ready", reason: "none" })),
    target: await inspectOfficeAuthorizedTarget(names)
  })),
  confirmOfficeAuthorizedScope: vi.fn(async (_names, confirm, expected) => confirm(expected)),
  officeAnalysisTargetEvidence: vi.fn(async () => ({ key: "target-1", conclusive: true }))
}));

vi.mock("./dataTools", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./dataTools")>()),
  executeQueryTableTool: vi.fn()
}));

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
  vi.spyOn(console, "log").mockImplementation(() => {});
  vi.mocked(inspectOfficeAuthorizedTarget).mockResolvedValue(fixtures.target);
  vi.mocked(currentOfficeWorkbookIdentity).mockReturnValue(
    "url:https://contoso.sharepoint.com/book.xlsx"
  );
  vi.mocked(captureSelectionContext).mockResolvedValue({
    activeWorksheet: "Sheet1",
    selectedRange: "A1:B3"
  });
  releaseStream = () => {};
  vi.stubGlobal("Office", {
    context: { document: { url: "https://contoso.sharepoint.com/book.xlsx" } },
    onReady: (cb: () => void) => cb()
  });
  Element.prototype.scrollIntoView = vi.fn();
});
afterEach(() => {
  releaseStream();
  cleanup();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

async function send(text: string) {
  const view = render(<App />);
  await waitFor(() => expect(vi.mocked(captureWorkbookCatalog)).toHaveBeenCalled());
  await waitFor(() =>
    expect(
      view.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title
    ).toContain("当前模型")
  );
  fireEvent.change(view.container.querySelector(".composer textarea")!, {
    target: { value: text }
  });
  fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
  return view;
}

const toolRequestIntent = {
  kind: "tool_request" as const,
  provider: "model" as const,
  summary: "统计记录",
  confirmedPrompt: "有多少条记录",
  request: {
    id: "q1",
    tool: "query_table" as const,
    arguments: {
      mode: "aggregate" as const,
      metrics: [{ operation: "countRows" as const, outputName: "记录数" }]
    }
  }
};

describe("Excel 面板：SDK 适配器的意图路径", () => {
  it.each(["read_only", "write_preview"] as const)("tool_request 保留目标 %s，不预置 dataResults", async (planningGoal) => {
    vi.mocked(understandMessage).mockResolvedValue({
      kind: "needs_workbook_context",
      summary: "统计记录",
      provider: "model"
    });
    vi.mocked(checkIntent).mockResolvedValue({ ...toolRequestIntent, planningGoal });
    vi.mocked(streamAssistantResponse).mockResolvedValue({
      kind: "answer",
      provider: "model",
      message: "已按工具桥结果作答"
    });

    await send("有多少条记录");

    await waitFor(() => expect(streamAssistantResponse).toHaveBeenCalledOnce());
    const request = vi.mocked(streamAssistantResponse).mock.calls[0][0] as {
      dataResults?: unknown[];
      planningReadPolicy?: string;
      planningGoal?: string;
    };
    expect(request.dataResults).toEqual([]);
    expect(request.planningReadPolicy).toBe("none");
    expect(request.planningGoal).toBe(planningGoal);
    // 本地预置计算与全表读取都不该发生：读表要等到桥指针回来那一刻。
    expect(executeQueryTableTool).not.toHaveBeenCalled();
  });
});

/**
 * 缺陷 2：桥的 live context 曾经返回**规划启动时**的闭包快照。
 *
 * 那样一来，"等待期间宿主/来源变了"根本测不出来 —— 复核拿的是同一个静止现场。
 * 这里用面板真实链路证明：绑定之后改了宿主身份或源指纹，桥**一个宿主单元格
 * 都不读**，只回固定错误码。
 */
const OP = "op_" + "7".repeat(32);
const FP = "a".repeat(64);
const RESULT_REF = "res_" + "5".repeat(32);

let releaseStream: () => void = () => {};

function stubBridgeFetch(
  options: {
    /** `/api/sdk/tool-bridge/result` 的响应体；不给则给一份最小可用的真值。 */
    result?: unknown;
    /** result 端点的 HTTP 状态（>400 用来模拟"已过保留期/授权不符"）。 */
    resultStatus?: number;
    /** 轮询一旦发出就回调 —— 表示"已经进入等待期"，此刻改宿主现场最贴切。 */
    onPoll?: () => void;
  } = {}
) {
  const captured: Array<{ url: string; body: Record<string, unknown> }> = [];
  vi.stubGlobal(
    "fetch",
    vi.fn(async (url: string, init?: RequestInit) => {
      const target = String(url);
      captured.push({
        url: target,
        body: JSON.parse(String(init?.body ?? "{}"))
      });
      if (target.includes("/receipt")) {
        return jsonResponse({ status: "accepted", code: null }, 200);
      }
      if (target.includes("/result")) {
        const status = options.resultStatus ?? 200;
        if (status >= 400) return jsonResponse({ detail: { code: "RESULT_EXPIRED" } }, status);
        return jsonResponse(
          options.result ?? {
            resultRef: RESULT_REF,
            title: "记录数",
            headers: ["记录数"],
            rows: [[3]],
            sourceSheets: ["Sheet1"],
            warnings: [],
            complete: true,
            calculation: "计数",
            scannedRows: 3
          },
          200
        );
      }
      options.onPoll?.();
      return jsonResponse(
        {
          calls: [
            {
              operationId: OP,
              tool: "query_table",
              arguments: {
                mode: "aggregate",
                scope: "selected",
                metrics: [{ operation: "countRows", outputName: "记录数" }]
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

function jsonResponse(payload: unknown, status: number): Response {
  return new Response(JSON.stringify(payload), {
    status,
    headers: { "Content-Type": "application/json" }
  });
}

/** 本地查询成功的最小结果（真实链路里由确定性计算产出）。 */
const queryToolResult = {
  id: "q1",
  tool: "query_table" as const,
  title: "记录数",
  headers: ["记录数"],
  rows: [[3]],
  sourceSheets: ["Sheet1"],
  scannedRows: 3,
  complete: true,
  calculation: "计数",
  warnings: []
};

/**
 * 走到"桥指针已到达、回执已发出"为止；`streamAssistantResponse` 会挂住，
 * 让本轮规划停在进行中 —— 只有这样才能在**等待期间**改变宿主现场。
 * 用完调用 `releaseStream()` 收尾。
 */
async function startBridgeTurn(
  captured: Array<{ url: string; body: Record<string, unknown> }>
) {
  vi.mocked(understandMessage).mockResolvedValue({
    kind: "needs_workbook_context",
    summary: "统计记录",
    provider: "model"
  });
  vi.mocked(checkIntent).mockResolvedValue(toolRequestIntent);
  vi.mocked(streamAssistantResponse).mockImplementation(async (_request, hooks) => {
    hooks?.onProgress?.({
      status: "started",
      stepId: "bridge.query_table",
      bridgeTool: { operationId: OP, tool: "query_table" }
    } as never);
    await new Promise<void>((resolve) => {
      releaseStream = resolve;
    });
    return { kind: "answer", provider: "model", message: "完成" };
  });
  const view = await send("有多少条记录");
  await waitFor(() =>
    expect(captured.some((c) => c.url.includes("/receipt"))).toBe(true)
  );
  return view;
}

describe("Excel 面板：桥执行前重读宿主现场（不是启动快照）", () => {
  it("等待期间宿主身份被换 → 拒读，只回 source_changed", async () => {
    let identity = "url:https://contoso.sharepoint.com/book.xlsx";
    vi.mocked(currentOfficeWorkbookIdentity).mockImplementation(() => identity);
    vi.mocked(executeQueryTableTool).mockResolvedValue(queryToolResult as never);
    /*
     * 轮询一发出就说明绑定已经建好、正等宿主执行；此刻换掉宿主身份。
     * "绑定时的身份"和"执行前复核读到的身份"是两次真读，差一次就必须拒 ——
     * 旧实现拿的是规划启动时的闭包快照，它会当成"没变化"继续读表。
     */
    const captured = stubBridgeFetch({
      onPoll: () => {
        identity = "url:https://contoso.sharepoint.com/other.xlsx";
      }
    });

    await startBridgeTurn(captured);
    releaseStream();

    const receipt = captured.find((c) => c.url.includes("/receipt"))?.body
      .receipt as Record<string, unknown>;
    expect(receipt).toEqual({ errorCode: "source_changed" });
    expect(executeQueryTableTool).not.toHaveBeenCalled();
  });

  /*
   * 身份与**源值**是两条独立判据。只测身份会漏掉"同一个文件、同一张表，
   * 但用户把数字改掉了" —— 那正是最危险的一种：身份没变，答案已经错了。
   */
  it("等待期间源值变化（身份不变）→ 同样拒读，只回 source_changed", async () => {
    let fingerprint = "fp-source-1";
    const captured = stubBridgeFetch({
      // 轮询一发出就说明进了等待期；此刻源值被改动。
      onPoll: () => {
        fingerprint = "fp-source-2";
      }
    });
    vi.mocked(captureWorkbookSourceFingerprint).mockImplementation(async () =>
      fingerprint
    );
    vi.mocked(executeQueryTableTool).mockResolvedValue(queryToolResult as never);

    await startBridgeTurn(captured);
    releaseStream();

    const receipt = captured.find((c) => c.url.includes("/receipt"))?.body
      .receipt as Record<string, unknown>;
    expect(receipt).toEqual({ errorCode: "source_changed" });
    // 负向断言要有正向锚点：指纹确实被真读了三次 —— 本轮来源基线一次、
    // 绑定一次、执行前复核一次（前两次必须同值，第三次变了才拒）。
    expect(captureWorkbookSourceFingerprint).toHaveBeenCalledTimes(3);
    expect(executeQueryTableTool).not.toHaveBeenCalled();
  });

  it("运行途中宿主身份读不到时 fail-closed：不读表，回 unauthorized_scope", async () => {
    const captured = stubBridgeFetch({ onPoll: () => {
      vi.mocked(currentOfficeWorkbookIdentity).mockImplementation(() => {
        throw new Error("Office 上下文不可用：D:\\敏感\\x.xlsx");
      });
    } });
    vi.mocked(executeQueryTableTool).mockResolvedValue(queryToolResult as never);
    await startBridgeTurn(captured);
    releaseStream();

    const receipt = captured.find((c) => c.url.includes("/receipt"))?.body;
    expect((receipt?.receipt as Record<string, unknown>)?.errorCode).toBeTruthy();
    expect(executeQueryTableTool).not.toHaveBeenCalled();
    expect(JSON.stringify(receipt)).not.toContain("敏感");
  });

  it("建立桥时宿主读不到：不发送规划，不等待工具指针", async () => {
    const captured = stubBridgeFetch();
    vi.mocked(understandMessage).mockResolvedValue({
      kind: "needs_workbook_context", summary: "统计记录", provider: "model"
    });
    vi.mocked(checkIntent).mockResolvedValue(toolRequestIntent);
    vi.mocked(currentOfficeWorkbookIdentity).mockImplementation(() => {
      throw new Error("Office 上下文不可用：D:\\敏感\\x.xlsx");
    });
    const view = await send("有多少条记录");
    await waitFor(() => expect(view.container.textContent).toContain("没有发送规划请求"));
    expect(streamAssistantResponse).not.toHaveBeenCalled();
    expect(captured).toHaveLength(0);
    expect(executeQueryTableTool).not.toHaveBeenCalled();
    expect(view.container.textContent).not.toContain("敏感");
  });
});

/**
 * 任务结束、桥已停止轮询之后，结果卡仍要能按引用把值取回来。
 *
 * 这条通路在会话关闭后走的是**授权归属**校验（服务端不再看活跃会话），
 * 所以请求体里必须回带**产生该结果那一轮**的绑定。结果过了保留期或授权不符时，
 * 界面要如实说"不可用"，绝不渲染一张空表冒充成功。
 */
describe("Excel 面板：完成后按引用取回结果 / 结果不可用", () => {
  const localResultPointer = {
    resultRef: RESULT_REF,
    title: "记录数",
    headers: ["记录数"],
    rowCount: 1,
    columnCount: 1,
    sourceSheets: ["Sheet1"],
    complete: true
  };

  async function finishTurnWithLocalResult(
    captured: Array<{ url: string; body: Record<string, unknown> }>, progressOnly = false
  ) {
    vi.mocked(understandMessage).mockResolvedValue({
      kind: "needs_workbook_context",
      summary: "统计记录",
      provider: "model"
    });
    vi.mocked(checkIntent).mockResolvedValue(toolRequestIntent);
    vi.mocked(streamAssistantResponse).mockImplementation(async (request, hooks) => {
      if (progressOnly) hooks?.onProgress?.({ schemaVersion: 1, operationId: request.turnId!,
        sequence: 1, phase: "planning", stepId: "python.completed", status: "completed",
        title: "隔离分析完成", timestamp: new Date().toISOString(), pythonResult: localResultPointer });
      return {
      kind: "answer",
      provider: "model",
      message: "共 3 条记录",
      localResult: progressOnly ? undefined : localResultPointer
    };
    });
    const view = await send("有多少条记录");
    await waitFor(() =>
      expect(captured.some((c) => c.url.includes("/result"))).toBe(true)
    );
    return view;
  }

  it.each([false, true])("桥已停：按引用取回该轮绑定并渲染真实值（仅进度事件=%s）", async (progressOnly) => {
    const captured = stubBridgeFetch();
    const view = await finishTurnWithLocalResult(captured, progressOnly);

    await waitFor(() =>
      expect(view.container.querySelector(".local-result-card-table")).not.toBeNull()
    );
    const table = view.container.querySelector(".local-result-card-table");
    expect(table?.textContent).toContain("记录数");
    expect(table?.textContent).toContain("3");
    expect(
      view.container.querySelector(".local-result-card-calculation")?.textContent
    ).toContain("计数");

    const body = captured.find((c) => c.url.includes("/result"))?.body;
    expect(body).toMatchObject({
      hostId: "excel-office",
      workbookName: "book.xlsx",
      workbookIdentity: "url:https://contoso.sharepoint.com/book.xlsx",
      sourceFingerprint: "fp-source-1",
      resultRef: RESULT_REF
    });
    // 任务身份必须是**这一轮**的 turnId，不是随便一个占位。
    expect(String(body?.taskId)).toMatch(/^turn-/);
  });

  it("结果已过保留期 / 授权不符：如实说不可用，不渲染空表冒充成功", async () => {
    const captured = stubBridgeFetch({ resultStatus: 410 });
    const view = await finishTurnWithLocalResult(captured);

    await waitFor(() =>
      expect(
        view.container.querySelector(".local-result-card-unavailable")
      ).not.toBeNull()
    );
    // 负向断言配正向锚点：卡片本身在（说明确实走到了这一步），只有表没渲染。
    expect(view.container.querySelector(".local-result-card")).not.toBeNull();
    expect(view.container.querySelector(".local-result-card-table")).toBeNull();
    expect(
      view.container.querySelector(".local-result-card-unavailable")?.textContent
    ).toContain("保留期");
  });

  it("面板重开（新的加载项会话）：结果卡靠**该轮**绑定按引用再取回，不依赖活跃会话", async () => {
    const captured = stubBridgeFetch();
    const first = await finishTurnWithLocalResult(captured);
    await waitFor(() =>
      expect(first.container.querySelector(".local-result-card-table")).not.toBeNull()
    );
    const firstCalls = captured.filter((item) => item.url.includes("/result"));
    expect(firstCalls.length).toBe(1);

    // 重开：当前加载项会话里已没有任何桥状态，只剩对话历史（localStorage）。
    first.unmount();
    const second = render(<App />);
    await waitFor(() =>
      expect(second.container.querySelector(".local-result-card-table")).not.toBeNull()
    );
    expect(
      second.container.querySelector(".local-result-card-table")?.textContent
    ).toContain("3");
    // 正向证明"是**重新**取回来的"：/result 又发了一次，且带着同一轮的绑定。
    const resultCalls = captured.filter((item) => item.url.includes("/result"));
    expect(resultCalls.length).toBe(2);
    expect(resultCalls[1].body).toMatchObject({
      taskId: firstCalls[0].body.taskId,
      hostId: "excel-office",
      workbookIdentity: "url:https://contoso.sharepoint.com/book.xlsx",
      resultRef: RESULT_REF
    });
  });
});
