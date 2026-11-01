import { describe, expect, it } from "vitest";
import { verifyResumeScene } from "./executionResumeScene";

describe("续跑前的现场判据：写入过的目标还是不是执行后的样子", () => {
  it("还没写任何动作时无事可核", () => {
    expect(verifyResumeScene({ completedActionCount: 0, coveredActionIndexes: [], observations: [] }))
      .toEqual({ ok: true, checked: 0 });
  });

  it("每个已完成动作都读到且一致才放行", () => {
    expect(verifyResumeScene({
      completedActionCount: 2,
      coveredActionIndexes: [0, 1],
      observations: [
        { actionIndex: 0, state: "matched" },
        { actionIndex: 1, state: "matched" }
      ]
    })).toEqual({ ok: true, checked: 2 });
  });

  it("目标已被撤销/被改：阻断，并指向撤销与“不能按旧进度补记”", () => {
    const outcome = verifyResumeScene({
      completedActionCount: 2,
      coveredActionIndexes: [0, 1],
      observations: [
        { actionIndex: 0, state: "matched" },
        { actionIndex: 1, state: "changed" }
      ]
    });
    expect(outcome.ok).toBe(false);
    if (outcome.ok) throw new Error("unreachable");
    expect(outcome.reason).toBe("changed");
    expect(outcome.detail).toContain("第 2 个动作的目标");
    expect(outcome.detail).toContain("撤销");
  });

  it("少一个动作的写后摘要就是没证据：不放行", () => {
    const outcome = verifyResumeScene({
      completedActionCount: 2,
      coveredActionIndexes: [0],
      observations: [{ actionIndex: 0, state: "matched" }]
    });
    expect(outcome.ok).toBe(false);
    if (outcome.ok) throw new Error("unreachable");
    expect(outcome.reason).toBe("no_evidence");
    expect(outcome.detail).toContain("第 2 个动作没有留下执行后校验");
  });

  it("少读到一格也算读不到（读不到不等于没变）", () => {
    const outcome = verifyResumeScene({
      completedActionCount: 2,
      coveredActionIndexes: [0, 1],
      observations: [{ actionIndex: 0, state: "matched" }]
    });
    expect(outcome.ok).toBe(false);
    if (outcome.ok) throw new Error("unreachable");
    expect(outcome.reason).toBe("unreadable");
  });

  it("读失败与对不上分开报：处置方法不同（重试 vs 撤销）", () => {
    const unreadable = verifyResumeScene({
      completedActionCount: 1,
      coveredActionIndexes: [0],
      observations: [{ actionIndex: 0, state: "unreadable" }]
    });
    const changed = verifyResumeScene({
      completedActionCount: 1,
      coveredActionIndexes: [0],
      observations: [{ actionIndex: 0, state: "changed" }]
    });
    expect(unreadable.ok).toBe(false);
    expect(changed.ok).toBe(false);
    if (unreadable.ok || changed.ok) throw new Error("unreachable");
    expect(unreadable.detail).toContain("读不到");
    expect(changed.detail).toContain("已经不是这次执行写入后的样子");
    expect(unreadable.detail).not.toBe(changed.detail);
  });

  it("进度是坏的（负数）也不当成“已完成”放行", () => {
    expect(verifyResumeScene({
      completedActionCount: -2, coveredActionIndexes: [], observations: []
    })).toEqual({ ok: true, checked: 0 });
  });
});
