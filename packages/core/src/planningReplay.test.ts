import { describe, expect, it } from "vitest";
import { acceptsPlanningReplay, planningReceivedSequence, waitForPlanning } from "./planningReplay";

describe("planning recovery ownership", () => {
  it("keeps the greatest consumed cursor across reply segments without borrowing another task", () => {
    expect(planningReceivedSequence("a", [
      { taskId: "a", sequence: 4 }, { taskId: "b", sequence: 999 },
      { taskId: "a", sequence: 17 }, { taskId: "a", sequence: 8 },
      { taskId: "a", sequence: NaN }
    ])).toBe(17);
    expect(acceptsPlanningReplay("a", 17, { operationId: "a", sequence: 17 })).toBe(false);
    expect(acceptsPlanningReplay("a", 17, { operationId: "b", sequence: 18 })).toBe(false);
    expect(acceptsPlanningReplay("a", 17, { operationId: "a", sequence: 18 })).toBe(true);
  });
  it("stops waiting even if the host transport ignores abort, and ignores its late result", async () => {
    const controller = new AbortController();
    let finish!: (value: string) => void;
    const pending = waitForPlanning(new Promise<string>((resolve) => { finish = resolve; }), controller.signal);
    controller.abort();
    await expect(pending).rejects.toMatchObject({ name: "AbortError" });
    finish("late");
    await expect(pending).rejects.toMatchObject({ name: "AbortError" });
  });
});
