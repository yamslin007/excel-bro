import { describe, expect, it } from "vitest";
import type { ChatMessage } from "./types/chat";
import type { AnalysisPlan } from "./contracts";
import { retireTargetMessage } from "./authorizedTargetArtifacts";
const plan: AnalysisPlan = { id: "p", title: "旧预览", summary: "", assumptions: [], warnings: [],
  actions: [{ type: "writeValues", sheet: "Sheet1", range: "E2", values: [["PRIVATE_LOCAL_VALUE"]] }] };
describe("new authorization retires previews without erasing conversation", () => {
  it.each([undefined, "p"])("removes old executable plans, including previously executed ones (%s)", (executedPlanId) => {
    const message: ChatMessage = { id: "m", role: "assistant", createdAt: "now", text: "已保留需求", plan,
      executedPlanId, executedExecutionId: "execution", staleSourceConfirm: { plan, bakedActionCount: 1 } };
    expect(JSON.stringify(message)).toContain("PRIVATE_LOCAL_VALUE");
    const retired = retireTargetMessage(message);
    expect(retired.plan).toBeUndefined(); expect(retired.staleSourceConfirm).toBeUndefined();
    expect(retired.text).toContain("已保留需求"); expect(retired.text).toContain("旧预览");
    expect(retired.executedExecutionId).toBe("execution");
    expect(JSON.stringify(retired)).not.toContain("PRIVATE_LOCAL_VALUE");
  });
  it("leaves already confirmed semantic answers intact", () => {
    const message: ChatMessage = { id: "answer", role: "user", createdAt: "now", text: "已确认的回答" };
    expect(retireTargetMessage(message)).toEqual(message);
  });
});
