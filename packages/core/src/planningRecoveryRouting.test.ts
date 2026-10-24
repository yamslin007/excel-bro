import { describe, expect, it } from "vitest";
import { planningRecoveryCandidates, isOfferedPlanningRecovery } from "./planningRecoveryRouting";

describe("planning recovery routing", () => {
  it("offers at most six distinct caller-owned tasks in chronological order", () => {
    expect(planningRecoveryCandidates([null, "", "a", "a", "b", "c", "d", "e", "f", "g"]))
      .toEqual(["b", "c", "d", "e", "f", "g"]);
  });
  it("does not infer a task from a model-invented ID", () => {
    expect(isOfferedPlanningRecovery("other-conversation", ["original"])).toBe(false);
    expect(isOfferedPlanningRecovery("original", ["original"])).toBe(true);
  });
});
