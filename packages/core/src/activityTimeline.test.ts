import { describe, expect, it } from "vitest";
import {
  applyTimelineEvent,
  finishTimelineAt,
  markTimelineBoundary,
  resumeTimeline,
  timelineDurations,
  type TimelineLogLike
} from "./activityTimeline";

/** 测试用的最小步骤：只带时间线关心的字段。 */
type Step = { id?: string; label: string; elapsedMs: number; status?: string; startedAt?: number };
type Log = TimelineLogLike<Step>;

function log(steps: Step[], extra: Partial<Log> = {}): Log {
  return { steps, totalMs: 0, ...extra };
}

function step(id: string, status = "started", startedAt = 1_000): Step {
  return { id, label: id, elapsedMs: 0, status, startedAt };
}

/** 把一条事件按"服务端会怎么发"的方式喂进去。 */
function feed(target: Log, event: {
  operationId?: string;
  stepId: string;
  sequence: number;
  status: string;
  label?: string;
}, now: number) {
  return applyTimelineEvent(
    target,
    {
      operationId: event.operationId ?? "turn-1",
      stepId: event.stepId,
      sequence: event.sequence,
      status: event.status,
      title: event.label
    },
    (existing) => ({
      ...(existing ?? {}),
      id: event.stepId,
      label: event.label ?? event.stepId,
      status: event.status,
      elapsedMs: Math.max(0, now - (existing?.startedAt ?? now)),
      startedAt: existing?.startedAt ?? now
    }),
    now
  );
}

describe("时间线：步骤身份与去重", () => {
  it("同一次调用的重放只更新不新增（同一 stepId、序号更大）", () => {
    let current = log([step("bridge.query_table:aaa")]);
    current = feed(current, { stepId: "bridge.query_table:aaa", sequence: 1, status: "progress" }, 1_500).log;
    const replay = feed(current, { stepId: "bridge.query_table:aaa", sequence: 2, status: "progress" }, 1_800);
    expect(replay.ignored).toBeNull();
    expect(replay.inserted).toBe(false);
    expect(replay.log.steps).toHaveLength(1);
  });

  it("同一个工具调用两次是两个独立步骤（stepId 不同就新增）", () => {
    let current = log([]);
    current = feed(current, { stepId: "bridge.query_table:aaa", sequence: 1, status: "started" }, 1_000).log;
    current = feed(current, { stepId: "bridge.query_table:aaa", sequence: 2, status: "completed" }, 1_200).log;
    const second = feed(current, { stepId: "bridge.query_table:bbb", sequence: 3, status: "started" }, 1_300);
    expect(second.inserted).toBe(true);
    expect(second.log.steps.map((item) => item.id)).toEqual([
      "bridge.query_table:aaa",
      "bridge.query_table:bbb"
    ]);
    // 第一次调用的终态没有被第二次调用改写。
    expect(second.log.steps[0].status).toBe("completed");
  });

  it("序号不增大的事件一律丢弃（迟到 / 重复 / 乱序）", () => {
    let current = feed(log([]), { stepId: "a", sequence: 5, status: "progress" }, 1_000).log;
    for (const sequence of [5, 4, 1]) {
      const result = feed(current, { stepId: "b", sequence, status: "progress" }, 1_200);
      expect(result.ignored).toBe("stale");
      expect(result.log).toBe(current);
    }
  });

  it("别的任务的事件不修改这条时间线", () => {
    const current = log([step("a")], { operationId: "turn-1", sequence: 3 });
    const result = feed(current, { operationId: "turn-2", stepId: "a", sequence: 9, status: "completed" }, 1_400);
    expect(result.ignored).toBe("operation");
    expect(result.log).toBe(current);
  });
});

describe("时间线：终态不回退", () => {
  it("已完成的步骤不会被迟到的'进行中'事件重新激活", () => {
    const current = log([
      { ...step("query", "completed", 1_000), elapsedMs: 200 }
    ]);
    const late = feed(current, { stepId: "query", sequence: 9, status: "progress" }, 5_000);
    expect(late.ignored).toBe("terminal-replay");
    expect(late.log.steps[0].status).toBe("completed");
    // 耗时也没有被"重新开始"污染。
    expect(late.log.steps[0].elapsedMs).toBe(200);
  });

  it("已失败的步骤同样不会被拉回进行中，但可以被另一条终态事件更正为完成", () => {
    const current = log([{ ...step("query", "failed", 1_000), elapsedMs: 100 }]);
    expect(feed(current, { stepId: "query", sequence: 2, status: "progress" }, 2_000).ignored)
      .toBe("terminal-replay");
    const corrected = feed(current, { stepId: "query", sequence: 3, status: "completed" }, 2_000);
    expect(corrected.ignored).toBeNull();
    expect(corrected.log.steps[0].status).toBe("completed");
  });
});

describe("时间线：收尾覆盖所有活动步骤", () => {
  it("收尾时当前步骤拿活动终态，其他没结束的步骤拿 warning 并写明原因", () => {
    const current = log([
      { ...step("a", "completed", 1_000), elapsedMs: 500 },
      { ...step("b", "progress", 1_500) },
      { ...step("c", "started", 1_800) }
    ]);
    const finished = finishTimelineAt(current, "completed", "c", 2_000, 1_000);
    expect(finished.outcome).toBe("completed");
    expect(finished.steps.map((item) => item.status)).toEqual([
      "completed",
      "warning",
      "completed"
    ]);
    const warning = finished.steps[1] as Step & { note?: string };
    expect(warning.note).toContain("未收到结束确认");
  });

  it("不假定最后一行就是当前步骤：按 startedAt 找最近开始的那一步", () => {
    // 数组顺序把"当前步骤"放在中间（迟到事件会让它不在末位）。
    const current = log([
      step("early", "started", 1_000),
      step("current", "started", 2_500),
      step("late-inserted", "started", 1_200)
    ]);
    const finished = finishTimelineAt(current, "completed", undefined, 3_000, 2_000);
    const byId = new Map(finished.steps.map((item) => [item.id, item.status]));
    expect(byId.get("current")).toBe("completed");
    expect(byId.get("early")).toBe("warning");
    expect(byId.get("late-inserted")).toBe("warning");
  });

  it("整体结束后每个步骤的计时都被冻结，不再随 now 增长", () => {
    const current = log([{ ...step("a", "progress", 1_000) }]);
    const finished = finishTimelineAt(current, "completed", "a", 2_000, 1_000);
    expect(finished.steps[0].elapsedMs).toBe(1_000);
    // 再往后推进时间，冻结值不变（这就是"完成事件到达后停止该步骤计时"）。
    const frozen = finished.steps[0].elapsedMs;
    expect(frozen).toBe(1_000);
    expect(finished.steps[0].startedAt).toBeUndefined();
  });
});

describe("时间线：恢复与计时分账", () => {
  it("恢复时**每一行**都按自己累计的时间重新定点，不是只修最后一行", () => {
    const current = log([
      { ...step("a", "completed", 1_000), elapsedMs: 400 },
      { ...step("b", "progress", 1_400), elapsedMs: 600 },
      { ...step("c", "started", 2_000), elapsedMs: 200 }
    ]);
    const resumed = resumeTimeline(current, 10_000);
    const [a, b, c] = resumed.steps;
    // 已完成的行一个都不动。
    expect(a).toMatchObject({ status: "completed", elapsedMs: 400 });
    // 未完成的行：累计时间保留，起点往前推同样多 ⇒ 再过一个 tick 不会从 0 长。
    expect(b.elapsedMs).toBe(600);
    expect(b.startedAt).toBe(10_000 - 600);
    expect(c.elapsedMs).toBe(200);
    expect(c.startedAt).toBe(10_000 - 200);
    // 恢复后**不新增**业务步骤、不重编号。
    expect(resumed.steps.map((item) => item.id)).toEqual(["a", "b", "c"]);
    expect(resumed.outcome).toBeUndefined();
  });

  it("暂停时间与断线等待时间分开记，后台执行不会被算成暂停", () => {
    let current: Log = { steps: [], totalMs: 0 };
    // 0..1000 正常执行；1000..3000 断线等待；3000..5000 用户暂停；5000 之后恢复。
    current = markTimelineBoundary(current, "waiting", 1_000);
    current = markTimelineBoundary(current, null, 3_000);
    current = markTimelineBoundary(current, "pause", 3_000);
    current = markTimelineBoundary(current, null, 5_000);

    const durations = timelineDurations(current, 0, 6_000);
    expect(durations.waitingMs).toBe(2_000);
    expect(durations.pausedMs).toBe(2_000);
    expect(durations.totalMs).toBe(6_000);
    // 实际执行 = 墙钟 − 等待 − 暂停。
    expect(durations.activeMs).toBe(2_000);
  });

  it("同一种边界重复标记不会重复开区间", () => {
    let current: Log = { steps: [], totalMs: 0 };
    current = markTimelineBoundary(current, "pause", 1_000);
    current = markTimelineBoundary(current, "pause", 2_000);
    const durations = timelineDurations(current, 0, 3_000);
    expect(durations.pausedMs).toBe(2_000);
    expect(durations.activeMs).toBe(1_000);
  });
});

describe("暂停后的继续：被暂停关掉的步骤要能重新接上", () => {
  it("暂停顺手关掉的那一步，恢复后还能接收进度（不是已完成）", () => {
    // 第 7 步正在跑，用户按了暂停：活动结束把它标成 warning + note。
    const running = feed(log([step("s7", "started", 1_000)]), {
      stepId: "s7", sequence: 9, status: "progress", label: "第 7 步"
    }, 1_500);
    const paused = finishTimelineAt(running.log, "awaiting_input", "s7", 3_000, 1_500);
    expect(paused.steps[0].status).toBe("warning");
    expect(paused.closedByFinish).toEqual(["s7"]);

    // 恢复：那一步被重新打开。
    const resumed = resumeTimeline(paused, 10_000);
    expect(resumed.steps[0].status).toBe("progress");
    expect(resumed.closedByFinish).toEqual([]);

    // 恢复后的第一条进度**必须**落到同一步上（而不是被"已完成不许重激活"丢掉）。
    const next = feed(resumed, {
      stepId: "s7", sequence: 31, status: "progress", label: "第 7 步（继续）"
    }, 10_400);
    expect(next.ignored).toBeNull();
    expect(next.log.steps).toHaveLength(1);
    expect(next.log.steps[0].label).toBe("第 7 步（继续）");
    expect(next.log.steps[0].status).toBe("progress");
  });

  it("真正收到过结束事件的步骤，恢复后仍是终态（不许被重新激活）", () => {
    const done = feed(log([step("s1", "started", 1_000)]), {
      stepId: "s1", sequence: 3, status: "completed", label: "第 1 步"
    }, 1_200);
    const paused = finishTimelineAt(done.log, "awaiting_input", undefined, 2_000, 300);
    // 真完成的那一步没有被"结束活动"改写，因此也不在 closedByFinish 里。
    expect(paused.steps[0].status).toBe("completed");
    expect(paused.closedByFinish).toEqual([]);

    const resumed = resumeTimeline(paused, 5_000);
    expect(resumed.steps[0].status).toBe("completed");
    const replay = feed(resumed, {
      stepId: "s1", sequence: 9, status: "progress", label: "第 1 步"
    }, 5_100);
    // 迟到的"进行中"事件不许把已完成步骤拉回来。
    expect(replay.ignored).toBe("terminal-replay");
    expect(replay.log.steps[0].status).toBe("completed");
  });

  /**
   * `warning` 有**两个来源**，恢复时的处理必须相反：
   *
   *  - 被结束活动**强制**关成 warning（上面那条，进 `closedByFinish`）→ 要重开；
   *  - 这一步**自己报了** warning 结束（"这一步没有完成"是真实结论）→ 不许重开。
   *
   * 只按状态判会把两者混成一个，于是自己报过 warning 的步骤恢复后又开始接收进度，
   * 一个已经声明失败的步骤被迟到的"进行中"事件拉回运行态。
   */
  it("自己报了 warning 结束的步骤，恢复后不重开", () => {
    const warned = feed(log([step("s3", "started", 1_000)]), {
      stepId: "s3", sequence: 5, status: "warning", label: "第 3 步"
    }, 1_400);
    expect(warned.log.steps[0].status).toBe("warning");
    const paused = finishTimelineAt(warned.log, "awaiting_input", undefined, 2_000, 400);
    // 它自己报了终态 ⇒ 不是被结束活动关掉的，不进 closedByFinish。
    expect(paused.closedByFinish).toEqual([]);

    const resumed = resumeTimeline(paused, 9_000);
    expect(resumed.steps[0].status).toBe("warning");
    const replay = feed(resumed, {
      stepId: "s3", sequence: 12, status: "progress", label: "第 3 步"
    }, 9_100);
    expect(replay.ignored).toBe("terminal-replay");
    expect(replay.log.steps[0].status).toBe("warning");
  });

  /**
   * 收尾那一刻"这一步自己已经做完"（活动级 outcome 是"带着待确认结果收尾"，
   * 但这一步真的做完了）：步骤终态必须是调用方给的 `currentStatus`，
   * 而且**不能**被记成"被结束活动关掉的" —— 否则恢复时它会被重新打开。
   */
  it("收尾时这一步自己已做完：恢复后仍是终态，不被重新打开", () => {
    const running = feed(log([step("s5", "started", 1_000)]), {
      stepId: "s5", sequence: 4, status: "progress", label: "第 5 步"
    }, 1_500);
    const finished = finishTimelineAt(
      running.log, "awaiting_input", "s5", 3_000, 2_000, undefined, "completed"
    );
    expect(finished.steps[0].status).toBe("completed");
    expect(finished.closedByFinish).toEqual([]);

    const resumed = resumeTimeline(finished, 9_000);
    expect(resumed.steps[0].status).toBe("completed");
    const replay = feed(resumed, {
      stepId: "s5", sequence: 21, status: "progress", label: "第 5 步"
    }, 9_100);
    expect(replay.ignored).toBe("terminal-replay");
    expect(replay.log.steps[0].status).toBe("completed");
  });
});


describe("恢复后的终态和耗时", () => {
  it("新序号的重复终态冻结原结果与耗时，同时接受补充说明", () => {
    const completed = feed(log([step("s")]), { stepId: "s", sequence: 1, status: "completed" }, 1500).log;
    const repeated = feed(completed, { stepId: "s", sequence: 2, status: "completed", label: "补充说明" }, 10000).log;
    expect(repeated.sequence).toBe(2);
    expect(repeated.steps[0]).toMatchObject({ status: "completed", elapsedMs: 500, label: "补充说明" });
    expect(repeated.steps[0].startedAt).toBeUndefined();
    expect(feed(repeated, { stepId: "s", sequence: 3, status: "failed" }, 11000).log.steps[0].status).toBe("failed");
  });
  it("恢复先结算等待桶，保持墙钟起点和步骤累计时间", () => {
    const stopped = finishTimelineAt(log([step("s")], { operationId: "turn-1" }), "awaiting_input", "s", 1500, 500);
    expect(stopped.excludedKind).toBe("waiting");
    const resumed = resumeTimeline(stopped, 6500);
    expect(resumed.waitingMs).toBe(5000);
    expect(resumed.totalMs).toBe(5500);
    const done = feed(resumed, { stepId: "s", sequence: 1, status: "completed" }, 7000).log;
    expect(done.steps[0].elapsedMs).toBe(1000);
    expect(timelineDurations(done, 1000, 7000)).toMatchObject({ activeMs: 1000, waitingMs: 5000, totalMs: 6000 });
  });
});
