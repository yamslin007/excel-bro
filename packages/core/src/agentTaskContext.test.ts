import { describe, expect, it } from "vitest";
import { createAgentTaskContext, confirmAgentTaskDecision } from "./agentTaskContext";

describe("同一任务的显式语义上下文", () => {
  it("同一道题的新回答替换旧回答，且不携带授权、源值或计划", () => {
    const context = createAgentTaskContext("整理当前表", [
      { questionId: "blank", label: "空值", value: "保留原值" },
      { questionId: "blank", label: "空值", value: "清空目标" }
    ]);
    expect(context.confirmations).toEqual([{ questionId: "blank", label: "空值", value: "清空目标" }]);
    const extended = { ...context, sourceSnapshots: [["PRIVATE_ROW"]], authorized: true };
    expect(JSON.stringify(confirmAgentTaskDecision(extended, { questionId: "mode", label: "口径", value: "逐行" })))
      .not.toContain("PRIVATE_ROW");
    expect(Object.keys(confirmAgentTaskDecision(extended, { label: "口径", value: "逐行" })))
      .toEqual(["schemaVersion", "originalPrompt", "confirmations"]);
  });
  it("替换任务从零开始，不继承上件任务的确认", () => {
    const previous = createAgentTaskContext("旧任务", [{ label: "口径", value: "汇总" }]);
    expect(confirmAgentTaskDecision(previous, { label: "新任务", value: "按名称排序" }, "replace"))
      .toEqual(createAgentTaskContext("按名称排序"));
  });
  it("超限拒绝，不静默截断用户已经确认的业务规则", () => {
    expect(() => createAgentTaskContext("当前任务", [{ label: "规则", value: "甲".repeat(1001) }])).toThrow();
  });
});
