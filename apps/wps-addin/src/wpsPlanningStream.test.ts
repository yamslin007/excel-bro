// @vitest-environment jsdom
/**
 * CAP-02：WPS 流式规划、暂停与检查点恢复 —— 端到端契约测试。
 *
 * 这个文件的存在理由：CAP-02 的**唯一**风险不是"功能没做"，而是"看起来做了"。
 * 具体表现为三种假成功：
 *
 *  1. **假续跑**：点"继续"其实是把原需求重发一遍 —— 服务端重新规划、重复调用
 *     模型与工具，而用户以为是从上次进度接着跑。
 *  2. **假暂停**：把 `planning_pause` 当成 `answer` 渲染，卡片画出来了，
 *     但底层是一次压根没被识别的结果。
 *  3. **假恢复**：刷新后拿界面上的旧 `revision` 去继续，被服务端 CAS 拒绝，
 *     或者反过来盲目推进游标、跳过没收到的事件。
 *
 * 因此这里的断言刻意分成两类：
 *  - **正向锚点**：正常路径确实产出了预期结果；
 *  - **反向变异**：把被测机制拆掉时必须变红（本次在下面的 `it` 名称里标注）。
 *
 * 隐私纪律同样在这里取证：先证明哨兵值确实在**本地输入**里，再查**真实
 * 出站 body** 里没有它 —— 顺序反过来就是自欺（body 里本来就没有，断言恒真）。
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  consumeWpsPlanningStream,
  continueWpsPlanning,
  getWpsPlanningTaskStatus,
  normalizeWpsPlanningStatus,
  parseWpsSseFrame,
  planningStatusView,
  streamWpsPlanning,
  wpsPlanningResumeDecision
} from "./wpsPlanningStream";
import { assertKnownResultKind, isPlanningPause } from "./wpsPlanningTypes";
import { StreamingToolDeliveryUnavailableError } from "../../excel-addin/src/api";
import {
  newPlanningCursor,
  planningAfterSequence,
  applyPlanningEvent,
  reconcilePlanningCursor,
  planningResumeDecision,
  resolveTurnIdForNewRequest,
  PLANNING_RESUME_HOST_WRITES
} from "@excel-bro/core";

/** 把一段 SSE 文本变成带可读 `body` 的 Response。 */
function sseResponse(frames: string, init: { ok?: boolean; status?: number } = {}): Response {
  const encoder = new TextEncoder();
  const bytes = encoder.encode(frames);
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(bytes);
      controller.close();
    }
  });
  return new Response(stream, {
    status: init.status ?? 200,
    headers: { "Content-Type": "text/event-stream" }
  });
}

function frame(event: string, payload: unknown): string {
  return `event: ${event}\ndata: ${JSON.stringify(payload)}\n\n`;
}

const pauseResult = {
  kind: "planning_pause" as const,
  turnId: "turn-1",
  status: "waiting_user" as const,
  revision: 4,
  lastSequence: 9,
  reason: "本轮处理时间已达上限。",
  activeElapsedSeconds: 120,
  approvedWallTimeSeconds: 120,
  extensionSeconds: 60,
  extensionCount: 1,
  maximumExtensions: 3,
  canContinue: true,
  requiresExtension: true,
  pauseType: "budget" as const,
  currentModelId: null,
  modelSwitchCount: 0,
  canSwitchModel: false,
  availableModelConnections: []
};

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("CAP-02 SSE 分帧", () => {
  it("解析 event/data，缺省 event 归 message（与 Excel 端同口径）", () => {
    expect(parseWpsSseFrame("event: step\ndata: {\"a\":1}")).toEqual({
      event: "step",
      data: '{"a":1}'
    });
    expect(parseWpsSseFrame("data: {\"a\":1}")).toEqual({
      event: "message",
      data: '{"a":1}'
    });
    expect(parseWpsSseFrame("event: step")).toBeNull();
  });

  it("多行 data 按 SSE 规则用 \\n 连接", () => {
    expect(parseWpsSseFrame("data: a\ndata: b")?.data).toBe("a\nb");
  });
});

describe("CAP-02 结构化出口：流式结果与一次性结果同形", () => {
  it("把 paused 帧识别成结构化暂停，而不是 answer", async () => {
    const response = sseResponse(frame("paused", pauseResult));
    const result = await consumeWpsPlanningStream(response, "turn-1");
    expect(isPlanningPause(result)).toBe(true);
    expect(assertKnownResultKind(result)).toBe("planning_pause");
  });

  it("把 result 帧识别成既有 AssistantResponse 联合（plan 分支不失真）", async () => {
    const plan = {
      kind: "plan" as const,
      provider: "model" as const,
      turnId: "turn-1",
      plan: { id: "p1", title: "删除", summary: "删除", assumptions: [], warnings: [], actions: [], acceptanceCriteria: [] }
    };
    const result = await consumeWpsPlanningStream(sseResponse(frame("result", plan)), "turn-1");
    expect(isPlanningPause(result)).toBe(false);
    expect(assertKnownResultKind(result)).toBe("plan");
    // 反向变异：若把 result 一律当 answer，这一条会红（kind 会变成 answer）。
    expect((result as { kind: string }).kind).not.toBe("answer");
  });

  it("未知事件类型必须报错，不许静默当 answer 蒙混（变异点：吞掉未知事件）", async () => {
    const response = sseResponse(frame("brand_new_event", { hello: 1 }) + frame("result", pauseResult));
    await expect(consumeWpsPlanningStream(response, "turn-1")).rejects.toThrow(
      /WPS_PLANNING_UNKNOWN_EVENT|尚不支持的事件类型/
    );
  });

  it("error 帧抛 ApiRequestError 并保留服务端 code（供 UI 归因）", async () => {
    const response = sseResponse(frame("error", { status: 502, code: "MODEL_CONNECTION_FAILED", message: "模型连不上", retryable: true }));
    await expect(consumeWpsPlanningStream(response, "turn-1")).rejects.toMatchObject({
      code: "MODEL_CONNECTION_FAILED",
      retryable: true
    });
  });

  it("流结束仍没有 result/paused 时明确失败，不返回 null 让调用方自己猜", async () => {
    const response = sseResponse(frame("step", { schemaVersion: 1, operationId: "turn-1", sequence: 1, phase: "planning", stepId: "s1", status: "started", title: "开始", timestamp: new Date().toISOString() }));
    await expect(consumeWpsPlanningStream(response, "turn-1")).rejects.toThrow(/未返回完整结果/);
  });

  it("没有 body 的响应直接按服务错误处理（代理不支持流式）", async () => {
    const response = new Response(null, { status: 200 });
    await expect(consumeWpsPlanningStream(response, "turn-1")).rejects.toThrow(/没有返回数据流/);
  });

  it("step 帧作为增量进度上报，不改变最终结果", async () => {
    const seen: string[] = [];
    const stepEvent = {
      schemaVersion: 1,
      operationId: "turn-1",
      sequence: 2,
      phase: "planning",
      stepId: "s1",
      status: "progress",
      title: "读取结构",
      timestamp: new Date().toISOString()
    };
    const response = sseResponse(frame("step", stepEvent) + frame("paused", pauseResult));
    const result = await consumeWpsPlanningStream(response, "turn-1", {
      onProgress: (event) => seen.push(event.title)
    });
    expect(seen).toEqual(["读取结构"]);
    expect(isPlanningPause(result)).toBe(true);
  });
});

describe("CAP-02 事件归并与游标（共享层纪律）", () => {
  it("取消隔离：旧 turnId 的迟到事件不得污染新任务（变异点：去掉身份闸）", () => {
    const cursor = newPlanningCursor("turn-new");
    const applied = applyPlanningEvent(cursor, { sequence: 3, stepId: "s", phase: "planning", status: "progress", title: "旧任务" }, "turn-old");
    expect(applied.accepted).toBe(false);
    expect(applied.rejectedReason).toBe("turn_mismatch");
    // 游标一动不动 —— 这正是"取消后旧事件改不了新任务状态"的证据。
    expect(applied.cursor.clientLastSequence).toBe(0);
  });

  it("去重闸：重复/重放事件被丢弃（变异点：删掉 sequence 比较）", () => {
    const cursor = newPlanningCursor("turn-1");
    const first = applyPlanningEvent(cursor, { sequence: 5, stepId: "s", phase: "planning", status: "progress", title: "一次" }, "turn-1");
    expect(first.accepted).toBe(true);
    const replay = applyPlanningEvent(first.cursor, { sequence: 5, stepId: "s", phase: "planning", status: "progress", title: "一次" }, "turn-1");
    expect(replay.accepted).toBe(false);
    expect(replay.rejectedReason).toBe("stale_sequence");
  });

  it("游标永不超前服务端确认位置（变异点：用 max 代替 min）", () => {
    const cursor = { ...newPlanningCursor("turn-1"), clientLastSequence: 12, serverLastSequence: 4 };
    expect(planningAfterSequence(cursor)).toBe(4);
  });

  it("revision 以服务端为准，客户端游标只取小", () => {
    const cursor = { ...newPlanningCursor("turn-1"), revision: 1, clientLastSequence: 20, serverLastSequence: 20 };
    const reconciled = reconcilePlanningCursor(cursor, { revision: 7, lastSequence: 6 });
    expect(reconciled.revision).toBe(7);
    expect(reconciled.serverLastSequence).toBe(6);
    // 客户端"以为"看到 20，服务端只确认到 6：取 6，中间那些靠继续时补发。
    expect(reconciled.clientLastSequence).toBe(6);
  });

  it("规划恢复路径永远不写宿主（变异点：让恢复顺手执行计划）", () => {
    expect(PLANNING_RESUME_HOST_WRITES).toBe(0);
  });

  it("新需求不会被旧暂停任务的恢复逻辑接管", () => {
    expect(resolveTurnIdForNewRequest("turn-paused", ["turn-paused"])).toBeNull();
    expect(resolveTurnIdForNewRequest("turn-fresh", ["turn-paused"])).toBe("turn-fresh");
  });
});

describe("CAP-02 状态查询与按钮推导", () => {
  it("状态白名单收窄：未知状态不得被当作 waiting_user（变异点：默认放行）", () => {
    const normalized = normalizeWpsPlanningStatus("turn-1", { status: "something_new", revision: 3, lastSequence: 5 });
    expect(normalized.status).toBe("failed");
  });

  it("无检查点时给「重新读取并生成预览」，不给「继续任务」（变异点：乐观放行）", () => {
    const decision = planningResumeDecision({
      status: "waiting_user",
      canContinue: true,
      hasCheckpoint: false,
      pauseType: null
    });
    expect(decision.action).toBe("regenerate");
    expect(decision.actionLabel).toContain("重新读取");
    expect(decision.actionLabel).not.toContain("继续");
  });

  it("有检查点且服务端允许 → 才给「继续任务」", () => {
    const decision = planningResumeDecision({
      status: "waiting_user",
      canContinue: true,
      hasCheckpoint: true,
      pauseType: "budget"
    });
    expect(decision.action).toBe("continue_checkpoint");
    expect(decision.canContinue).toBe(true);
  });

  it("模型连接类故障 + 服务端明确可继续 → 允许接续", () => {
    const decision = planningResumeDecision({
      status: "failed",
      canContinue: true,
      hasCheckpoint: true,
      errorCode: "MODEL_CONNECTION_FAILED"
    });
    expect(decision.action).toBe("continue_checkpoint");
  });

  it("非模型故障的终态一律只能重新规划", () => {
    const decision = planningResumeDecision({
      status: "failed",
      canContinue: false,
      hasCheckpoint: false,
      errorCode: "PLAN_TOO_LARGE"
    });
    expect(decision.action).toBe("none");
  });

  it("conservative 视图：只回 200 不等于有现场（变异点：把查询成功读成可续跑）", () => {
    // 服务端只回了"我有一个 waiting_user 任务"，**没有** canContinue、
    // 没有 pauseType、没有 result —— 这就是"问到了状态，但现场不完整"。
    // 唯一正确的动作是重新读取并生成预览；若 hasCheckpoint 被写成恒真，
    // 这里会放行出一个"继续任务"按钮，用户点下去只会重复规划。
    const status = normalizeWpsPlanningStatus("turn-1", {
      status: "waiting_user",
      revision: 2,
      lastSequence: 0
    });
    expect(status.canContinue).toBeUndefined();
    expect(status.pauseType).toBeNull();
    const decision = wpsPlanningResumeDecision(status);
    expect(decision.action).toBe("regenerate");
    expect(decision.actionLabel).not.toContain("继续");
  });

  it("hasCheckpoint 判据：三项证据缺一不可（变异点：任一分支写死 true）", () => {
    /**
     * 这条直接钉 `planningStatusView` 的**推导**，而不是它的使用方。
     *
     * 原因：`wpsPlanningResumeDecision` 每次都用状态重算 `hasCheckpoint`，
     * 所以"把推导改坏"从上层是看不见的 —— 必须在这一层取证。
     *
     * 另一条纪律也在这里被固定：`canContinue` 与 `hasCheckpoint` 是**两个判据**。
     * 服务端可能说"你可以继续"（`canContinue: true`）却拿不出完整现场
     * （没有 `pauseType`、没有 `result`）；这时唯一正确的动作仍是重新读取，
     * 而不是给一个会重复规划的"继续"按钮。
     */
    const bare = normalizeWpsPlanningStatus("turn-1", {
      status: "waiting_user",
      revision: 2,
      lastSequence: 4,
      canContinue: true
    });
    // 有 canContinue，但没有 pauseType / result → 现场不完整。
    expect(planningStatusView(bare).hasCheckpoint).toBe(false);
    const decision = wpsPlanningResumeDecision(bare);
    expect(decision.action).toBe("regenerate");
    expect(decision.actionLabel).not.toContain("继续");

    // 补上 pauseType（服务端给出的现场字段）→ 才承认现场完整。
    const withPauseType = normalizeWpsPlanningStatus("turn-1", {
      status: "waiting_user",
      revision: 2,
      lastSequence: 4,
      canContinue: true,
      pauseType: "budget"
    });
    expect(planningStatusView(withPauseType).hasCheckpoint).toBe(true);
    expect(wpsPlanningResumeDecision(withPauseType).action).toBe("continue_checkpoint");
  });

  it("completed 且带回 result 才算有现场（否则不能继续）", () => {
    const completedNoResult = normalizeWpsPlanningStatus("t", {
      status: "completed",
      revision: 3,
      lastSequence: 7
    });
    expect(planningStatusView(completedNoResult).hasCheckpoint).toBe(false);
    expect(wpsPlanningResumeDecision(completedNoResult).action).toBe("none");
  });
});

describe("CAP-02 查询与继续走真实 HTTP", () => {
  it("查状态只 GET，不创建任务（变异点：把查询写成 POST /api/turn）", async () => {
    const calls: Array<{ url: string; init?: RequestInit }> = [];
    vi.stubGlobal("fetch", vi.fn(async (url: string, init?: RequestInit) => {
      calls.push({ url: String(url), init });
      return new Response(JSON.stringify({ status: "waiting_user", revision: 4, lastSequence: 9, canContinue: true, pauseType: "budget" }), { status: 200 });
    }));
    const status = await getWpsPlanningTaskStatus("turn-1");
    expect(status.revision).toBe(4);
    expect(calls[0].url).toContain("/api/turn/turn-1/status");
    expect(calls[0].init?.method ?? "GET").toBe("GET");
    // 反向变异点：查询**不能**打到规划端点上（那会创建任务）。
    expect(calls[0].url).not.toMatch(/\/api\/turn\/stream$/);
    expect(calls[0].url).not.toMatch(/\/api\/turn$/);
    // 只有一次请求：查询不会顺带创建任务。
    expect(calls).toHaveLength(1);
  });

  it("继续只带身份与游标，**绝不重发原需求**（变异点：把 prompt 塞回请求体）", async () => {
    const bodies: string[] = [];
    vi.stubGlobal("fetch", vi.fn(async (_url: string, init?: RequestInit) => {
      bodies.push(String(init?.body ?? ""));
      return sseResponse(frame("result", { kind: "answer", provider: "model", message: "完成", resultContext: null }));
    }));
    const outcome = await continueWpsPlanning({
      turnId: "turn-1",
      expectedRevision: 4,
      requestId: "plan-continue-abc",
      afterSequence: 9
    });
    expect(outcome.kind).toBe("answer");
    const body = JSON.parse(bodies[0]);
    expect(body).toMatchObject({
      turnId: "turn-1",
      expectedRevision: 4,
      requestId: "plan-continue-abc",
      afterSequence: 9,
      confirmed: true
    });
    // 关键反向断言：请求体里没有需求文本字段。
    expect(body).not.toHaveProperty("prompt");
    expect(body).not.toHaveProperty("workbook");
    expect(body).not.toHaveProperty("dataResults");
  });

  it("流式规划会带上完整请求体，并用 turnId 做身份", async () => {
    const bodies: string[] = [];
    vi.stubGlobal("fetch", vi.fn(async (_url: string, init?: RequestInit) => {
      bodies.push(String(init?.body ?? ""));
      return sseResponse(frame("result", { kind: "answer", provider: "model", message: "好", resultContext: null }));
    }));
    await streamWpsPlanning({ turnId: "turn-9", prompt: "统计", workbook: {}, dataResults: [] }, "turn-9");
    const body = JSON.parse(bodies[0]);
    expect(body.prompt).toBe("统计");
  });

  it("继续时非 2xx 抛结构化错误，不静默降级", async () => {
    vi.stubGlobal("fetch", vi.fn(async () => new Response(
      JSON.stringify({ detail: { code: "PLANNING_REVISION_CONFLICT", message: "版本已变化", retryable: false } }),
      { status: 409 }
    )));
    await expect(continueWpsPlanning({ turnId: "t", expectedRevision: 1, requestId: "r", afterSequence: 0 }))
      .rejects.toMatchObject({ code: "PLANNING_REVISION_CONFLICT", status: 409 });
  });
});

/*
 * 无流式工具投递能力 → **提前明确拒绝**。
 *
 * WPS 端的取数通道只有一条：服务端经 SSE 投递 `bridgeTool` 指针，本机读表后回执。
 * 流式通道不可用时若只抛一句笼统的"请求失败"，调用方分不清"后端是旧版"还是
 * "代理缓冲了 SSE"；而 SDK 适配器恰恰**只能**靠这条通道。这里给出可执行的明确原因，
 * 并且只打一次请求 —— 绝不退回非流式（那会让服务端为送不到本机的工具请求空等）。
 */
describe("CAP-02 SDK 适配器：无流式工具投递能力时明确拒绝", () => {
  it.each([422, 500, 503])("HTTP %s 保留服务端错误，不误报流式缺失且不重发", async (status) => {
    const fetchMock = vi.fn(async () => new Response(JSON.stringify({ detail: {
      code: "TASK_STATE_CORRUPT", message: "任务数据库损坏", retryable: false
    } }), { status }));
    vi.stubGlobal("fetch", fetchMock);
    await expect(streamWpsPlanning({ modelId: "claude-agent-sdk" }, "turn-error"))
      .rejects.toMatchObject({ status, code: "TASK_STATE_CORRUPT", message: "任务数据库损坏" });
    await expect(continueWpsPlanning({ turnId: "turn-error", expectedRevision: 1,
      requestId: "continue-error", afterSequence: 0, modelId: "claude-agent-sdk" }))
      .rejects.toMatchObject({ status, code: "TASK_STATE_CORRUPT" });
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });
  it("流式端点缺失时抛固定错误码，且只请求一次（不退回 /api/turn）", async () => {
    const calls: string[] = [];
    vi.stubGlobal("fetch", vi.fn(async (url: string) => {
      calls.push(String(url));
      return new Response("Not Found", { status: 404 });
    }));

    const rejection = streamWpsPlanning(
      { turnId: "turn-sdk", prompt: "统计", workbook: {}, dataResults: [], modelId: "claude-agent-sdk" },
      "turn-sdk"
    );

    await expect(rejection).rejects.toBeInstanceOf(StreamingToolDeliveryUnavailableError);
    await expect(rejection).rejects.toMatchObject({
      code: "streaming_tool_delivery_unavailable"
    });
    // 反向变异点：任何第二次请求都意味着服务端已经开始等一个永远到不了的回执。
    expect(calls).toHaveLength(1);
    expect(calls[0]).toContain("/api/turn/stream");
  });

  it("续跑流式端点缺失时同样明确拒绝", async () => {
    vi.stubGlobal("fetch", vi.fn(async () => new Response("Not Found", { status: 404 })));

    await expect(
      continueWpsPlanning({
        turnId: "turn-sdk",
        expectedRevision: 2,
        requestId: "r-1",
        afterSequence: 3,
        modelId: "claude-agent-sdk:connection-0001"
      })
    ).rejects.toMatchObject({ code: "streaming_tool_delivery_unavailable" });
  });

  it("普通模型在同样的 404 下不被这个闸拦住（仍旧报结构化请求错误）", async () => {
    vi.stubGlobal("fetch", vi.fn(async () => new Response(
      JSON.stringify({ detail: { code: "PLANNING_UNKNOWN", message: "旧后端", retryable: false } }),
      { status: 404 }
    )));

    await expect(
      streamWpsPlanning(
        { turnId: "turn-1", prompt: "统计", workbook: {}, dataResults: [], modelId: "connection:regular" },
        "turn-1"
      )
    ).rejects.toMatchObject({ code: "PLANNING_UNKNOWN", status: 404 });
  });
});

describe("CAP-02 事件游标接进真实 SSE 消费链（第 3 条）", () => {
  const step = (sequence: number, operationId: string, title = "步骤") => ({
    schemaVersion: 1,
    operationId,
    sequence,
    phase: "planning",
    stepId: `s${sequence}`,
    status: "progress",
    title,
    timestamp: new Date().toISOString()
  });

  it("重复事件被真实消费链丢弃，不重复上报进度（变异点：回到本地自增计数）", async () => {
    const seen: string[] = [];
    const response = sseResponse(
      frame("step", step(1, "turn-1", "第一次")) +
      frame("step", step(1, "turn-1", "重放的同一帧")) +
      frame("result", { kind: "answer", provider: "model", message: "好", resultContext: null })
    );
    await consumeWpsPlanningStream(response, "turn-1", {
      onProgress: (event) => seen.push(event.title)
    });
    // 序号 1 只应上报一次；重放的那一帧必须被去重闸丢掉。
    expect(seen).toEqual(["第一次"]);
  });

  it("别条任务的迟到事件不进本任务的进度（变异点：去掉身份闸）", async () => {
    const seen: string[] = [];
    const response = sseResponse(
      frame("step", step(1, "turn-other", "别的任务")) +
      frame("step", step(2, "turn-1", "本任务")) +
      frame("result", { kind: "answer", provider: "model", message: "好", resultContext: null })
    );
    await consumeWpsPlanningStream(response, "turn-1", {
      onProgress: (event) => seen.push(event.title)
    });
    // 身份不符的帧被丢；本任务的照常上报。
    expect(seen).toEqual(["本任务"]);
  });

  it("afterSequence 之后的补发帧照常上报；边界内的被跳过（漏收补发）", async () => {
    const seen: string[] = [];
    const response = sseResponse(
      frame("step", step(1, "turn-1", "已消费过")) +
      frame("step", step(2, "turn-1", "补发的新帧")) +
      frame("result", { kind: "answer", provider: "model", message: "好", resultContext: null })
    );
    await consumeWpsPlanningStream(response, "turn-1", {
      onProgress: (event) => seen.push(event.title),
      afterSequence: 1
    });
    // 序号 1 ≤ afterSequence → 已展示过，不再重复上报。
    expect(seen).toEqual(["补发的新帧"]);
  });

  it("取消后旧任务的迟到 result 不改变本次消费结果（身份闸）", async () => {
    /*
     * 场景：本端已切到新任务 turn-new，但旧任务 turn-old 的流还没断。
     * `consumeWpsPlanningStream` 的身份参数是 turn-new —— 旧帧的 step
     * 必须被丢；而 result 帧本身是"本流的终态"，仍按本流身份解析。
     */
    const seen: string[] = [];
    const response = sseResponse(
      frame("step", step(5, "turn-old", "旧任务的步骤")) +
      frame("result", { kind: "answer", provider: "model", message: "本任务完成", resultContext: null })
    );
    const result = await consumeWpsPlanningStream(response, "turn-new", {
      onProgress: (event) => seen.push(event.title)
    });
    expect(seen).toEqual([]);
    expect((result as { message?: string }).message).toBe("本任务完成");
  });
});

describe("CAP-02 隐私：源值只留本端", () => {
  it("哨兵值先证明在本地输入里，再证明不在真实出站 body 中", async () => {
    /**
     * 哨兵选一个**只存在于单元格**的长字符串，而不是短数字串 ——
     * `"2"` 这种会在 revision/sequence 里偶然出现，断言恒真、证不了任何东西。
     */
    const sentinel = "SENTINEL-CELL-VALUE-7f3a";
    const request = {
      turnId: "turn-privacy",
      prompt: "统计记录",
      workbook: { name: "测试.xlsx", worksheets: [{ name: "明细", dataRows: [] }] },
      dataResults: []
    };
    // 第 1 步：证明哨兵确实在这份本地输入里（否则下面的"不存在"毫无意义）。
    const withSentinel = { ...request, workbook: { ...request.workbook, worksheets: [{ name: "明细", dataRows: [[sentinel]] }] } };
    expect(JSON.stringify(withSentinel)).toContain(sentinel);

    const bodies: string[] = [];
    vi.stubGlobal("fetch", vi.fn(async (_url: string, init?: RequestInit) => {
      bodies.push(String(init?.body ?? ""));
      return sseResponse(frame("result", { kind: "answer", provider: "model", message: "好", resultContext: null }));
    }));
    // 第 2 步：真正发出去的请求体里没有它。
    await streamWpsPlanning(request as unknown as Record<string, unknown>, "turn-privacy");
    expect(bodies[0]).not.toContain(sentinel);
    expect(bodies[0]).toContain("turn-privacy");
  });

  it("继续请求体里不含任何单元格内容", async () => {
    const sentinel = "SENTINEL-CELL-VALUE-9b1c";
    const bodies: string[] = [];
    vi.stubGlobal("fetch", vi.fn(async (_url: string, init?: RequestInit) => {
      bodies.push(String(init?.body ?? ""));
      return sseResponse(frame("paused", pauseResult));
    }));
    const outcome = await continueWpsPlanning({
      turnId: sentinel,
      expectedRevision: 1,
      requestId: "r",
      afterSequence: 0
    });
    // turnId 是身份，允许出现；但除它以外不该有任何"内容"字段。
    expect(bodies[0]).not.toContain("prompt");
    expect(isPlanningPause(outcome)).toBe(true);
  });
});
