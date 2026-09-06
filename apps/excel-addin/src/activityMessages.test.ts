import { describe, expect, it } from "vitest";
import type { ActivityLog } from "./hooks/useActivityProgress";
import type { ChatMessage } from "./types/chat";
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
