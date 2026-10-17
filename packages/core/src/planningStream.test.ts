import { describe, expect, it } from "vitest";
import {
  MODEL_ERROR_CODE_PREFIX,
  PLANNING_RESUME_HOST_WRITES,
  applyPlanningEvent,
  closeDanglingSteps,
  isSamePlanningResume,
  newPlanningCursor,
  newPlanningRequestId,
  planningAfterSequence,
  planningPausedEntriesFromHistory,
  planningResumeDecision,
  reconcilePlanningCursor,
  resolveTurnIdForNewRequest,
  type PlanningStepEventView
} from "./planningStream";

function event(sequence: number, overrides: Partial<PlanningStepEventView> = {}): PlanningStepEventView {
  return {
    sequence,
    stepId: `planning-${sequence}`,
    phase: "planning",
    status: "progress",
    title: `步骤 ${sequence}`,
    ...overrides
  };
}

describe("规划流游标：身份、去重与顺序", () => {
  it("归并一条同身份事件并推进游标", () => {
    const cursor = newPlanningCursor("turn-1");
    const applied = applyPlanningEvent(cursor, event(1), "turn-1");
    expect(applied.accepted).toBe(true);
    expect(applied.cursor.clientLastSequence).toBe(1);
    expect(applied.events).toHaveLength(1);
    expect(applied.events[0].sequence).toBe(1);
  });

  it("拒绝身份不匹配的事件（取消后旧事件不得污染新任务）", () => {
    const cursor = newPlanningCursor("turn-new");
    const applied = applyPlanningEvent(cursor, event(5), "turn-old");
    expect(applied.accepted).toBe(false);
    expect(applied.rejectedReason).toBe("turn_mismatch");
    // 游标**不得**被旧任务推进 —— 这是"取消隔离"的取证点。
    expect(applied.cursor).toEqual(cursor);
    expect(applied.events).toEqual([]);
  });

  it("拒绝重复与重放事件（sequence 不严格递增）", () => {
    let cursor = newPlanningCursor("turn-1");
    cursor = applyPlanningEvent(cursor, event(3), "turn-1").cursor;
    const replay = applyPlanningEvent(cursor, event(3), "turn-1");
    expect(replay.accepted).toBe(false);
    expect(replay.rejectedReason).toBe("stale_sequence");
    const older = applyPlanningEvent(cursor, event(2), "turn-1");
    expect(older.accepted).toBe(false);
    expect(older.rejectedReason).toBe("stale_sequence");
    // 正向锚点：更大的序号仍然被接受，否则上面的负向断言可能是"什么都不收"。
    expect(applyPlanningEvent(cursor, event(4), "turn-1").accepted).toBe(true);
  });

  it("拒绝非数值 sequence，不把它当成 0 或 NaN 收下来", () => {
    const cursor = newPlanningCursor("turn-1");
    const applied = applyPlanningEvent(
      cursor,
      event(Number.NaN),
      "turn-1"
    );
    expect(applied.accepted).toBe(false);
    expect(applied.rejectedReason).toBe("stale_sequence");
  });

  it("afterSequence 永不超前于服务端确认位置", () => {
    const cursor = {
      turnId: "turn-1",
      revision: 2,
      serverLastSequence: 5,
      clientLastSequence: 9
    };
    // 客户端自以为看到 9，但服务端只提交到 5 —— 只能从 5 之后补发。
    expect(planningAfterSequence(cursor)).toBe(5);
    expect(planningAfterSequence({ ...cursor, clientLastSequence: 3 })).toBe(3);
  });

  it("用服务端序号校准游标：revision 以服务端为准，client 游标只取小", () => {
    const cursor = { turnId: "turn-1", revision: 1, serverLastSequence: 2, clientLastSequence: 7 };
    const next = reconcilePlanningCursor(cursor, { revision: 4, lastSequence: 3 });
    expect(next.revision).toBe(4);
    expect(next.serverLastSequence).toBe(3);
    // min(7, 3) = 3：不能因为服务端更大就假装中间事件都看到了。
    expect(next.clientLastSequence).toBe(3);
  });

  it("服务端给出非法序号时保守保留原值，不写进 NaN/Infinity", () => {
    const cursor = { turnId: "turn-1", revision: 3, serverLastSequence: 5, clientLastSequence: 4 };
    const next = reconcilePlanningCursor(cursor, {
      revision: Number.NaN,
      lastSequence: Number.POSITIVE_INFINITY
    });
    expect(next.revision).toBe(cursor.revision);
    // Infinity 会被 `Number.isFinite` 挡下 —— 保留原值，绝不写进非有限数，
    // 否则后续 `Math.min` / `afterSequence` 计算会污染成 Infinity。
    expect(next.serverLastSequence).toBe(5);
    expect(Number.isFinite(next.serverLastSequence)).toBe(true);
    expect(next.clientLastSequence).toBe(4);
    // 原游标未被原地修改。
    expect(cursor.serverLastSequence).toBe(5);
  });

  it("空的 turnId 直接拒绝，不生成一个可用身份", () => {
    expect(() => newPlanningCursor("   ")).toThrow();
  });
});

describe("规划继续方案：按钮只由结构化状态推导", () => {
  it("运行中不给继续按钮", () => {
    const decision = planningResumeDecision({ status: "running" });
    expect(decision.action).toBe("none");
    expect(decision.canContinue).toBe(false);
  });

  it("已完成不给继续按钮", () => {
    const decision = planningResumeDecision({ status: "completed" });
    expect(decision.action).toBe("none");
    expect(decision.canContinue).toBe(false);
  });

  it("有检查点的 waiting_user 走精确续跑，文案说'从上次保存的进度继续'", () => {
    const decision = planningResumeDecision({
      status: "waiting_user",
      canContinue: true,
      hasCheckpoint: true,
      reason: "本轮处理时间已用完，当前进度已安全保存。"
    });
    expect(decision.action).toBe("continue_checkpoint");
    expect(decision.actionLabel).toBe("继续任务");
    expect(decision.detail).toContain("从上次保存的进度继续");
  });

  it("没有检查点时只能重新规划，文案说'重新读取并生成预览'而不是'继续'", () => {
    const decision = planningResumeDecision({
      status: "waiting_user",
      canContinue: true,
      hasCheckpoint: false
    });
    expect(decision.action).toBe("regenerate");
    expect(decision.actionLabel).toBe("重新读取并生成预览");
    // 关键：不得出现"从上次保存的进度继续"这类精确续跑承诺。
    expect(decision.actionLabel).not.toBe("继续任务");
  });

  it("过期/不可继续的任务只允许真实降级，不提供假继续按钮", () => {
    const decision = planningResumeDecision({
      status: "waiting_user",
      canContinue: false,
      hasCheckpoint: true
    });
    expect(decision.action).toBe("regenerate");
    expect(decision.canContinue).toBe(true);
    expect(decision.detail).toContain("没有可验证的完整规划现场");
  });

  it("失败的普通任务不提供继续按钮", () => {
    const decision = planningResumeDecision({
      status: "failed",
      errorCode: "PLANNING_STOPPED",
      errorMessage: "规划已停止。",
      canContinue: false,
      hasCheckpoint: false
    });
    expect(decision.action).toBe("none");
    expect(decision.reason).toBe("规划已停止。");
  });

  it("模型连接故障 + 可继续 + 有检查点，才允许换模型接续", () => {
    const recoverable = planningResumeDecision({
      status: "failed",
      errorCode: `${MODEL_ERROR_CODE_PREFIX}TIMEOUT`,
      canContinue: true,
      hasCheckpoint: true,
      errorMessage: "模型连接中断。"
    });
    expect(recoverable.action).toBe("continue_checkpoint");
    // 缺任一条件都不成立 —— 逐项拆掉对照。
    expect(
      planningResumeDecision({
        status: "failed",
        errorCode: `${MODEL_ERROR_CODE_PREFIX}TIMEOUT`,
        canContinue: true,
        hasCheckpoint: false
      }).action
    ).toBe("none");
    expect(
      planningResumeDecision({
        status: "failed",
        errorCode: `${MODEL_ERROR_CODE_PREFIX}TIMEOUT`,
        canContinue: false,
        hasCheckpoint: true
      }).action
    ).toBe("none");
  });

  it("总时长已用完（budget_exhausted）不提供续跑按钮，也不承诺必然可继续", () => {
    // 服务端会把这种暂停的 canContinue 置为 false；这里如实核对下游不给按钮。
    const decision = planningResumeDecision({
      status: "waiting_user",
      canContinue: false,
      hasCheckpoint: true,
      pauseType: "budget_exhausted",
      reason: "本轮有效运行总时长已用完，不再启动新的模型或工具调用。"
    });
    expect(decision.action).toBe("regenerate");
    expect(decision.actionLabel).not.toBe("继续任务");
    // 原因原样透出，不被"进度已安全保存"之类的兜底文案顶掉。
    expect(decision.detail).not.toContain("从上次保存的进度继续");
  });

  it("真实原因原样透出，不吞掉服务端的 error.message", () => {
    const decision = planningResumeDecision({
      status: "failed",
      errorMessage: "所选备用模型不支持本任务所需的图片输入。",
      canContinue: false,
      hasCheckpoint: true
    });
    expect(decision.reason).toBe("所选备用模型不支持本任务所需的图片输入。");
  });
});

describe("继续请求身份与点击去重", () => {
  it("同一 turnId + 同一 requestId 视为同一次继续", () => {
    expect(
      isSamePlanningResume(
        { turnId: "turn-1", requestId: "req-a" },
        { turnId: "turn-1", requestId: "req-a" }
      )
    ).toBe(true);
  });

  it("不同 requestId 或不同 turnId 不算同一次", () => {
    expect(
      isSamePlanningResume(
        { turnId: "turn-1", requestId: "req-a" },
        { turnId: "turn-1", requestId: "req-b" }
      )
    ).toBe(false);
    expect(
      isSamePlanningResume(
        { turnId: "turn-1", requestId: "req-a" },
        { turnId: "turn-2", requestId: "req-a" }
      )
    ).toBe(false);
  });

  it("requestId 稳定且非空，便于服务端幂等领取", () => {
    const first = newPlanningRequestId();
    expect(first.startsWith("plan-continue-")).toBe(true);
    expect(first.length).toBeGreaterThan(20);
    // 两次调用互不相同（否则幂等键会碰撞，把两次真实继续当成同一次）。
    expect(first).not.toBe(newPlanningRequestId());
  });

  it("新需求不得复用已知可恢复任务的 turnId", () => {
    expect(resolveTurnIdForNewRequest("turn-paused", ["turn-paused"])).toBeNull();
    // 未知 turnId 可正常使用；空值返回 null。
    expect(resolveTurnIdForNewRequest("turn-fresh", ["turn-paused"])).toBe("turn-fresh");
    expect(resolveTurnIdForNewRequest(null, ["turn-paused"])).toBeNull();
    expect(resolveTurnIdForNewRequest("  ", ["turn-paused"])).toBeNull();
  });
});

describe("终态收尾：不留'正在执行'的父步骤", () => {
  it("成功收尾把仍在 progress 的步骤标为 completed", () => {
    const closed = closeDanglingSteps(
      [
        { stepId: "a", status: "completed" as const },
        { stepId: "b", status: "progress" as const }
      ],
      "completed"
    );
    expect(closed).toEqual([
      { stepId: "a", status: "completed" },
      { stepId: "b", status: "completed" }
    ]);
  });

  it("失败收尾把仍在 progress 的步骤标为 failed", () => {
    const closed = closeDanglingSteps([{ stepId: "a", status: "progress" as const }], "failed");
    expect(closed[0].status).toBe("failed");
  });

  it("取消收尾按 completed 处理，不谎报步骤失败", () => {
    const closed = closeDanglingSteps([{ stepId: "a", status: "progress" as const }], "cancelled");
    expect(closed[0].status).toBe("completed");
  });

  it("不原地修改输入数组", () => {
    const input = [{ stepId: "a", status: "progress" as const }];
    closeDanglingSteps(input, "completed");
    expect(input[0].status).toBe("progress");
  });
});

describe("规划恢复不得触碰宿主写入", () => {
  it("共享常量声明规划恢复路径的宿主写入次数为 0", () => {
    expect(PLANNING_RESUME_HOST_WRITES).toBe(0);
  });
});

describe("面板重开：从历史挑出需要重新核对的暂停", () => {
  const paused = (turnId: string, reason = "本轮处理时间已达上限") => ({
    turnId,
    pauseType: "budget" as const,
    reason,
    canContinue: true,
    action: "continue_checkpoint" as const,
    actionLabel: "继续任务",
    detail: "从上次保存的进度继续"
  });

  const payloadOf = (turnId: string) => (turnId ? { turnId } : null);

  it("只挑带 planningPause 的消息", () => {
    const entries = planningPausedEntriesFromHistory(
      [
        { id: "m1", text: "普通回答" },
        { id: "m2", planningPause: paused("turn-a") },
        { id: "m3", clarification: {} }
      ] as never,
      payloadOf
    );
    expect(entries.map((entry) => entry.turnId)).toEqual(["turn-a"]);
  });

  it("同一 turnId 被多次暂停时取历史里**最后**一条", () => {
    const entries = planningPausedEntriesFromHistory(
      [
        { id: "m1", planningPause: paused("turn-a", "第一次暂停") },
        { id: "m2", planningPause: paused("turn-a", "第二次暂停") }
      ] as never,
      payloadOf
    );
    expect(entries).toHaveLength(1);
    expect(entries[0].messageId).toBe("m2");
  });

  it("不同 turnId 各挑一条，顺序按首次出现", () => {
    const entries = planningPausedEntriesFromHistory(
      [
        { id: "m1", planningPause: paused("turn-a") },
        { id: "m2", planningPause: paused("turn-b") }
      ] as never,
      payloadOf
    );
    expect(entries.map((entry) => entry.turnId)).toEqual(["turn-a", "turn-b"]);
  });

  it("载荷取不到（服务端记录已不在）时整条不挑，不画没依据的卡", () => {
    const entries = planningPausedEntriesFromHistory(
      [{ id: "m1", planningPause: paused("turn-a") }] as never,
      () => null
    );
    expect(entries).toEqual([]);
  });

  it("turnId 为空白字符串时跳过，不拿空身份去查服务端", () => {
    const entries = planningPausedEntriesFromHistory(
      [{ id: "m1", planningPause: paused("   ") }] as never,
      payloadOf
    );
    expect(entries).toEqual([]);
  });
});
