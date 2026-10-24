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

  /*
   * 段身份守卫（缺陷 1 的结构性兜底）：历史段已经冻结在它自己的那一刻，
   * 迟到的进度不许把它重新激活；活动段照常接收。
   */
  function segment(id: string, status: "live" | "superseded" | "closed", taskId?: string) {
    return { id, status, continuesTask: Boolean(taskId),
      ...(taskId ? { taskId } : {}) } as NonNullable<ChatMessage["replySegment"]>;
  }

  it("refuses to write progress into a frozen segment", () => {
    const frozen = { ...log, operationId: "turn-a", sequence: 9 };
    const frozenLog: ActivityLog = { ...frozen, connectionMessage: "暂停那一刻" };
    const messages = [
      message("activity", "assistant", { activityLog: frozenLog,
        replySegment: segment("activity", "superseded", "turn-a") })
    ];

    const result = persistActivityMessage(messages,
      { ...frozen, connectionMessage: "迟到的执行进度" },
      "activity", message("fallback", "assistant"));

    expect(result).toBe(messages);
    expect(result[0].activityLog).toBe(frozenLog);
    expect(result[0].activityLog?.connectionMessage).toBe("暂停那一刻");
  });

  it("still writes progress into the live segment", () => {
    const result = persistActivityMessage(
      [message("activity", "assistant", { replySegment: segment("activity", "live", "turn-a") })],
      { ...log, operationId: "turn-a", sequence: 9, connectionMessage: "正在执行" },
      "activity",
      message("fallback", "assistant")
    );

    expect(result[0].activityLog?.connectionMessage).toBe("正在执行");
    expect(result[0].activityPending).toBe(false);
  });

  it("keeps accepting progress for messages that never had a segment", () => {
    const result = persistActivityMessage([message("activity", "assistant")],
      { ...log, operationId: "turn-a", sequence: 9, connectionMessage: "老消息照旧" },
      "activity", message("fallback", "assistant"));

    expect(result[0].activityLog?.connectionMessage).toBe("老消息照旧");
  });
});
