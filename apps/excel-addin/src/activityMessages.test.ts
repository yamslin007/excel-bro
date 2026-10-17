import { describe, expect, it } from "vitest";
import type { ActivityLog } from "./hooks/useActivityProgress";
import type { ChatMessage } from "./types/chat";
import { parseTaskProgressEvent } from "./contracts";
import {
  insertChatMessage,
  persistActivityMessage
} from "./activityMessages";

function message(
  id: string,
  role: ChatMessage["role"],
  overrides: Partial<ChatMessage> = {}
): ChatMessage {
  return {
    id,
    role,
    createdAt: "2026-09-05T00:00:00Z",
    ...overrides
  };
}

const log: ActivityLog = {
  outcome: "awaiting_input",
  totalMs: 1200,
  steps: [
    {
      id: "intent",
      label: "模型判断需要先澄清需求",
      elapsedMs: 1200,
      status: "completed"
    }
  ]
};

describe("activityMessages", () => {
  const card = { resultRef: "res_" + "a".repeat(32), title: "Python 分析结果",
    headers: ["列1"], rowCount: 2, columnCount: 1, sourceSheets: ["Sheet1"], complete: true };
  const binding = { taskId: "turn-python", hostId: "excel-office" as const,
    workbookName: "test.xlsx", workbookIdentity: "url:test.xlsx", sourceFingerprint: null };

  it("preserves Python metadata through the shared parser without carrying values", () => {
    const event = parseTaskProgressEvent({ schemaVersion: 1, operationId: "turn-python",
      sequence: 1, phase: "planning", stepId: "python.completed", status: "completed",
      title: "完成", timestamp: "2026-10-11T00:00:00Z", pythonResult: { ...card, values: [["SECRET"]] } });
    expect(event?.pythonResult).toEqual(card);
    expect(JSON.stringify(event)).not.toContain("SECRET");
  });

  it("keeps existing result and binding when an activity has no Python result", () => {
    const result = persistActivityMessage([message("activity", "assistant", {
      localResult: card, localResultBinding: binding })], log, "activity", message("fallback", "assistant"));
    expect(result[0].localResult).toEqual(card);
    expect(result[0].localResultBinding).toEqual(binding);
  });

  it("attaches the newest Python result to fallback messages too", () => {
    const newer = { ...card, resultRef: "res_" + "b".repeat(32) };
    const result = persistActivityMessage([], { ...log, steps: [
      { label: "first", elapsedMs: 1, pythonResult: card },
      { label: "second", elapsedMs: 1, pythonResult: newer }] }, null,
      message("fallback", "assistant"), binding);
    expect(result[0].localResult).toEqual(newer);
    expect(result[0].localResultBinding).toEqual(binding);
  });

  it("does not attach an unbound result", () => {
    const result = persistActivityMessage([], { ...log, steps: [
      { label: "python", elapsedMs: 1, pythonResult: card }] }, null, message("fallback", "assistant"));
    expect(result[0].localResult).toBeUndefined();
  });
  it("places a late user message immediately before the active process message", () => {
    const result = insertChatMessage(
      [
        message("old", "assistant"),
        message("activity", "assistant", { activityPending: true })
      ],
      message("user", "user"),
      "activity"
    );

    expect(result.map((item) => item.id)).toEqual([
      "old",
      "user",
      "activity"
    ]);
  });

  it("keeps later assistant output after the active process message", () => {
    const result = insertChatMessage(
      [message("activity", "assistant", { activityPending: true })],
      message("clarification", "assistant"),
      "activity"
    );

    expect(result.map((item) => item.id)).toEqual([
      "activity",
      "clarification"
    ]);
  });

  it("persists a log to its stable target even when another assistant message is last", () => {
    const result = persistActivityMessage(
      [
        message("activity", "assistant", { activityPending: true }),
        message("clarification", "assistant")
      ],
      log,
      "activity",
      message("fallback", "assistant", { activityLog: log })
    );

    expect(result[0]).toMatchObject({
      id: "activity",
      activityPending: false,
      activityLog: log
    });
    expect(result[1].id).toBe("clarification");
    expect(result[1].activityLog).toBeUndefined();
    expect(result).toHaveLength(2);
  });

  it("does not insert the same process message id twice", () => {
    const original = message("activity", "assistant", {
      activityPending: true
    });

    const result = insertChatMessage(
      [original],
      message("activity", "assistant", { activityPending: true }),
      "activity"
    );

    expect(result).toEqual([original]);
  });

  it("coalesces duplicate process placeholders while persisting the log", () => {
    const result = persistActivityMessage(
      [
        message("activity", "assistant", { activityPending: true }),
        message("clarification", "assistant"),
        message("activity", "assistant", { activityPending: true })
      ],
      log,
      "activity",
      message("activity", "assistant", { activityLog: log })
    );

    expect(result.map((item) => item.id)).toEqual([
      "activity",
      "clarification"
    ]);
    expect(result[0]).toMatchObject({
      activityPending: false,
      activityLog: log
    });
  });
});
