import { describe, expect, it } from "vitest";
import { reviewStatusBarSuppressedByRecovery } from "./ResultReviewBar";

/**
 * 「本次执行需要核对」状态卡要不要因恢复操作区接管而收起 —— 唯一判据。
 *
 * 核心是**按执行身份匹配**（`executionId === operationId`），不是"有任意恢复操作区就收"：
 * 换了一条任务的旧告警（不同执行身份）照旧要留着。两端拿不到身份时（`null`）也不收 ——
 * 拿不准就别替用户藏告警。
 */
describe("reviewStatusBarSuppressedByRecovery", () => {
  it("失败执行仅归属于创建它的审阅任务，不能借用上一轮身份", () => {
    const args = {
      recoveryComposerActive: true, reviewExecutionId: null,
      recoveryOperationId: "execution-1",
      failedExecution: { key: "task-1", operationId: "execution-1" }
    };
    expect(reviewStatusBarSuppressedByRecovery({ ...args, reviewKey: "task-1" })).toBe(true);
    expect(reviewStatusBarSuppressedByRecovery({ ...args, reviewKey: "task-2" })).toBe(false);
    expect(reviewStatusBarSuppressedByRecovery({ ...args, reviewKey: null })).toBe(false);
    expect(reviewStatusBarSuppressedByRecovery({ ...args, reviewKey: "task-1", reviewExecutionId: "execution-2" })).toBe(false);
  });
  it("恢复操作区未接管时不收，即使身份一致", () => {
    expect(reviewStatusBarSuppressedByRecovery({
      recoveryComposerActive: false,
      reviewExecutionId: "execution-1",
      recoveryOperationId: "execution-1"
    })).toBe(false);
  });

  it("接管且身份一致：收（同一次执行，操作区已经说清了）", () => {
    expect(reviewStatusBarSuppressedByRecovery({
      recoveryComposerActive: true,
      reviewExecutionId: "wps-execution-abc",
      recoveryOperationId: "wps-execution-abc"
    })).toBe(true);
  });

  it("接管但身份不同：不收（不能因为任意操作区在场就藏别的任务的告警）", () => {
    expect(reviewStatusBarSuppressedByRecovery({
      recoveryComposerActive: true,
      reviewExecutionId: "execution-A",
      recoveryOperationId: "execution-B"
    })).toBe(false);
  });

  it("任一侧身份拿不到（null）：不收（拿不准不藏告警）", () => {
    expect(reviewStatusBarSuppressedByRecovery({
      recoveryComposerActive: true,
      reviewExecutionId: null,
      recoveryOperationId: "execution-B"
    })).toBe(false);
    expect(reviewStatusBarSuppressedByRecovery({
      recoveryComposerActive: true,
      reviewExecutionId: "execution-A",
      recoveryOperationId: null
    })).toBe(false);
  });
});
