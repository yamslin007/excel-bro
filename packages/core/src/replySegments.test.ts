import { describe, expect, it } from "vitest";
import type { TimelineLogLike, TimelineStepLike } from "./activityTimeline";
import {
  activeReplySegment,
  closeReplySegment,
  continuationLabel,
  freezeSegmentLog,
  latestReplySegmentForTask,
  openReplySegment,
  planSegmentContinuation,
  replySegmentsForTask,
  segmentAcceptsProgress,
  type ReplySegment
} from "./replySegments";

describe("回复段身份", () => {
  it("打开第一段时成为活动段，且不带上一段", () => {
    const segments = openReplySegment([], { id: "seg-1", taskId: "turn-a" });
    expect(segments).toHaveLength(1);
    expect(segments[0]).toMatchObject({ id: "seg-1", status: "live", taskId: "turn-a" });
    expect(segments[0].previousSegmentId).toBeUndefined();
    expect(activeReplySegment(segments)?.id).toBe("seg-1");
  });

  it("继续时打开新段：旧段降为 superseded 并接上下一段，活动段只有一个", () => {
    const first = openReplySegment([], { id: "seg-1", taskId: "turn-a" });
    const second = openReplySegment(first, {
      id: "seg-2",
      taskId: "turn-a",
      continuesTask: true
    });
    expect(second.map((segment) => segment.status)).toEqual(["superseded", "live"]);
    expect(second[0].nextSegmentId).toBe("seg-2");
    expect(second[1].previousSegmentId).toBe("seg-1");
    expect(second[1].continuesTask).toBe(true);
    expect(second.filter((segment) => segment.status === "live")).toHaveLength(1);
    expect(activeReplySegment(second)?.id).toBe("seg-2");
    // 任务身份没有被换掉：展示段不重造执行身份。
    expect(second.every((segment) => segment.taskId === "turn-a")).toBe(true);
  });

  it("同一个段 id 重复打开是幂等的（重放不产生第二段）", () => {
    const once = openReplySegment([], { id: "seg-1" });
    const twice = openReplySegment(once, { id: "seg-1" });
    expect(twice).toHaveLength(1);
    expect(twice[0].status).toBe("live");
  });

  it("连续三次暂停/继续后仍只有一个活动段，历史段各接下一段", () => {
    let segments: ReplySegment[] = [];
    for (const id of ["s1", "s2", "s3", "s4"]) {
      segments = openReplySegment(segments, { id, taskId: "turn-a" });
    }
    expect(segments.map((segment) => segment.status)).toEqual([
      "superseded",
      "superseded",
      "superseded",
      "live"
    ]);
    expect(segments[0].nextSegmentId).toBe("s2");
    expect(segments[2].nextSegmentId).toBe("s4");
    expect(segments).toHaveLength(replySegmentsForTask(segments, "turn-a").length);
  });

  it("历史段不接受进度：只有活动段能落事件", () => {
    const segments = openReplySegment(
      openReplySegment([], { id: "s1", taskId: "turn-a" }),
      { id: "s2", taskId: "turn-a" }
    );
    expect(segmentAcceptsProgress(segments, "s2", "turn-a")).toBe(true);
    // 旧段即使任务身份相同也不能抢回展示位置。
    expect(segmentAcceptsProgress(segments, "s1", "turn-a")).toBe(false);
    // 别的任务无论如何都不落。
    expect(segmentAcceptsProgress(segments, "s2", "turn-b")).toBe(false);
  });

  it("收尾只处理活动段，已经 superseded 的段不被改写", () => {
    const segments = openReplySegment(
      openReplySegment([], { id: "s1" }),
      { id: "s2" }
    );
    const closed = closeReplySegment(segments, "s1");
    expect(closed[0].status).toBe("superseded");
    expect(closed[1].status).toBe("live");
    const closedLive = closeReplySegment(segments, "s2");
    expect(closedLive[1].status).toBe("closed");
    expect(activeReplySegment(closedLive)).toBeUndefined();
  });

  it("续跑文案只用真实的前序步骤数，不编编号", () => {
    expect(continuationLabel(3)).toContain("3 步");
    expect(continuationLabel(0)).toBe("继续（此前未记录步骤）");
  });
});

describe("冻结历史段时间线", () => {
  // 步骤形态跟宿主一致（`ActivityStep` 就是 `TimelineStepLike + note`），
  // 否则测试里读不到 `note` —— 那正是"交代未结束步骤"的判据。
  type TestStep = TimelineStepLike & { note?: string };
  const base: TimelineLogLike<TestStep> = {
    operationId: "turn-a",
    totalMs: 0,
    timelineStartedAt: 1000,
    steps: [
      { id: "a", label: "读取表结构", elapsedMs: 500, status: "completed" },
      { id: "b", label: "生成代码", elapsedMs: 200, status: "progress", startedAt: 1200 }
    ]
  };

  it("未结束的步骤被交代为 warning，已完成的步骤原样保留", () => {
    const frozen = freezeSegmentLog(base, 5000);
    const [done, open] = frozen.steps;
    expect(done.status).toBe("completed");
    expect(done.elapsedMs).toBe(500);
    // 当前步骤（最晚开始的那个未结束步骤）拿活动级终态，不带"未收到结束确认"。
    expect(open.status).toBe("warning");
    expect(frozen.outcome).toBe("awaiting_input");
  });

  it("更早的未结束步骤额外标明「没有收到结束确认」", () => {
    const twoOpen: TimelineLogLike<TestStep> = {
      ...base,
      steps: [
        { id: "a", label: "读取表结构", elapsedMs: 500, status: "progress", startedAt: 1100 },
        { id: "b", label: "生成代码", elapsedMs: 200, status: "progress", startedAt: 1200 }
      ]
    };
    const frozen = freezeSegmentLog(twoOpen, 5000);
    expect(frozen.steps[0].status).toBe("warning");
    expect(frozen.steps[0].note).toContain("未收到结束确认");
    expect(frozen.steps[1].note).toBeUndefined();
  });

  it("已经 failed 的步骤不会被改写成 warning", () => {
    const frozen = freezeSegmentLog(
      {
        ...base,
        steps: [{ id: "a", label: "执行", elapsedMs: 10, status: "failed" }]
      },
      5000
    );
    expect(frozen.steps[0].status).toBe("failed");
  });

  it("冻结后的耗时停在冻结那一刻，不再随墙钟增长", () => {
    const frozen = freezeSegmentLog(base, 5000);
    expect(frozen.totalMs).toBe(4000);
    expect(frozen.steps[1].elapsedMs).toBe(3800);
  });
});

describe("连续继续的段链归属", () => {
  const three = (): ReplySegment[] => [
    { id: "A", taskId: "turn-a", status: "superseded", nextSegmentId: "B" },
    { id: "B", taskId: "turn-a", status: "superseded", nextSegmentId: "C" },
    { id: "C", taskId: "turn-a", status: "live" }
  ];

  it("新段接在该任务**最新**的段之后，不是最早那段", () => {
    expect(latestReplySegmentForTask(three(), "turn-a")?.id).toBe("C");
    expect(planSegmentContinuation(three(), "turn-a")).toEqual({
      previousSegmentId: "C",
      supersededSegmentId: "C"
    });
  });

  it("A → B → C → D 每一段只指向自己的直接后继", () => {
    const before = three();
    const after = openReplySegment(before, {
      id: "D",
      taskId: "turn-a",
      continuesTask: true,
      previousSegmentId: "C"
    });
    expect(after.map((segment) => segment.id)).toEqual(["A", "B", "C", "D"]);
    // A 的后继没有被改成 D，B 也没有丢后继 —— 这就是旧实现里丢掉的定位入口。
    expect(after.find((segment) => segment.id === "A")?.nextSegmentId).toBe("B");
    expect(after.find((segment) => segment.id === "B")?.nextSegmentId).toBe("C");
    expect(after.find((segment) => segment.id === "C")?.nextSegmentId).toBe("D");
    expect(after.find((segment) => segment.id === "D")?.previousSegmentId).toBe("C");
    expect(after.filter((segment) => segment.status === "live")).toHaveLength(1);
  });

  it("多任务交错时不会把新段接到别的任务后面", () => {
    const segments: ReplySegment[] = [
      { id: "A", taskId: "turn-a", status: "superseded", nextSegmentId: "B" },
      { id: "B", taskId: "turn-a", status: "live" },
      { id: "X", taskId: "turn-x", status: "closed" }
    ];
    const after = openReplySegment(segments, {
      id: "C",
      taskId: "turn-a",
      continuesTask: true,
      previousSegmentId: latestReplySegmentForTask(segments, "turn-a")?.id
    });
    expect(after.find((segment) => segment.id === "C")?.previousSegmentId).toBe("B");
    expect(after.find((segment) => segment.id === "B")?.nextSegmentId).toBe("C");
    // 另一个任务的段完全不动。
    expect(after.find((segment) => segment.id === "X")).toEqual(segments[2]);
  });

  it("同一个 id 重复打开不会造出第二段，也不会改已有链", () => {
    const before = three();
    const after = openReplySegment(before, { id: "C", taskId: "turn-a", continuesTask: true });
    expect(after.map((segment) => segment.id)).toEqual(["A", "B", "C"]);
    expect(after).toEqual(before);
  });
});

describe("段能不能接收进度", () => {
  const segments: ReplySegment[] = [
    { id: "A", taskId: "turn-a", status: "superseded", nextSegmentId: "B" },
    { id: "B", taskId: "turn-a", status: "live" }
  ];

  it("只有活动段接收，同任务的历史段也不接收", () => {
    expect(segmentAcceptsProgress(segments, "B", "turn-a")).toBe(true);
    expect(segmentAcceptsProgress(segments, "A", "turn-a")).toBe(false);
  });

  it("老消息没有段身份时按可接收处理（兼容），不因缺字段判成冻结", () => {
    expect(segmentAcceptsProgress(segments, null, "turn-a")).toBe(true);
    expect(segmentAcceptsProgress(segments, undefined, null)).toBe(true);
  });

  it("别的任务的段不接收本任务进度", () => {
    expect(segmentAcceptsProgress(segments, "B", "turn-other")).toBe(false);
  });
});

// A/B are different execution tasks, not consecutive replies of one task.
describe("跨任务继续保持双向定位", () => {
  it.each(["superseded", "closed"] as const)("继续 %s 的父段也补后继，冻结状态不变", (status) => {
    const parent: ReplySegment = { id: "A1", taskId: "task-a", status };
    const other: ReplySegment = { id: "B1", taskId: "task-b", status: "live" };
    const before = [parent, other];
    const next = openReplySegment(before, { id: "A2", taskId: "task-a", continuesTask: true,
      previousSegmentId: planSegmentContinuation(before, "task-a").previousSegmentId });
    expect(next[0]).toEqual({ ...parent, nextSegmentId: "A2" });
    expect(next[1].nextSegmentId).toBeUndefined();
    expect(next[2].previousSegmentId).toBe("A1");
    expect(next.filter(s => s.status === "live")).toHaveLength(1);
    expect(parent.nextSegmentId).toBeUndefined();
  });
});
