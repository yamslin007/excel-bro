import { describe, expect, it } from "vitest";
import {
  canAcknowledgeStaleSource,
  deleteConversationFromHistory,
  normalizePetVisibility,
  normalizeStoredMessage,
  normalizeStoredVerification
} from "./conversation";
import type { AnalysisPlan } from "./contracts";

function conversation(id: string, updatedAt: string) {
  return {
    id,
    title: `对话 ${id}`,
    messages: [],
    createdAt: updatedAt,
    updatedAt
  };
}

describe("history deletion", () => {
  it("removes an inactive conversation without changing the active one", () => {
    const next = deleteConversationFromHistory(
      {
        activeConversationId: "active",
        conversations: [
          conversation("active", "2026-07-28T00:00:00.000Z"),
          conversation("old", "2026-07-27T00:00:00.000Z")
        ]
      },
      "old"
    );

    expect(next.activeConversationId).toBe("active");
    expect(next.conversations.map((item) => item.id)).toEqual(["active"]);
  });

  it("selects a remaining conversation after deleting the active one", () => {
    const next = deleteConversationFromHistory(
      {
        activeConversationId: "active",
        conversations: [
          conversation("active", "2026-07-28T00:00:00.000Z"),
          conversation("remaining", "2026-07-27T00:00:00.000Z")
        ]
      },
      "active"
    );

    expect(next.activeConversationId).toBe("remaining");
    expect(next.conversations.map((item) => item.id)).toEqual(["remaining"]);
  });

  it("creates a fresh conversation after deleting the last one", () => {
    const next = deleteConversationFromHistory(
      {
        activeConversationId: "only",
        conversations: [
          conversation("only", "2026-07-28T00:00:00.000Z")
        ]
      },
      "only"
    );

    expect(next.conversations).toHaveLength(1);
    expect(next.activeConversationId).toBe(next.conversations[0].id);
    expect(next.activeConversationId).not.toBe("only");
    expect(next.conversations[0].messages).toHaveLength(1);
  });
});

describe("history verification migration", () => {
  it("normalizes verification reports saved before status was introduced", () => {
    const migrated = normalizeStoredVerification({
      passed: true,
      checks: []
    } as never);

    expect(migrated).toMatchObject({
      status: "verified",
      passed: true,
      unverifiedActions: []
    });
  });
});

describe("stale source plan preview", () => {
  const plan = {
    id: "stale-plan",
    title: "写入汇总结果",
    summary: "把预览时计算的汇总值写入结果表",
    assumptions: [],
    warnings: [],
    actions: [
      {
        type: "writeValues",
        sheet: "结果",
        range: "A1",
        values: [[42]]
      }
    ]
  } as AnalysisPlan;

  it("restores the stored plan as a visible preview before stale execution", () => {
    const normalized = normalizeStoredMessage({
      id: "message-1",
      role: "system",
      text: "数据来源已变化",
      staleSourceConfirm: { plan, bakedActionCount: 1 },
      createdAt: "2026-09-01T00:00:00.000Z"
    });

    expect(normalized.plan).toBe(plan);
    expect(normalized.staleSourceConfirm?.status).toBe("pending");
    expect(canAcknowledgeStaleSource(normalized)).toBe(true);
  });

  it("blocks stale execution without the matching visible preview or after completion", () => {
    const message = normalizeStoredMessage({
      id: "message-2",
      role: "system",
      staleSourceConfirm: { plan, bakedActionCount: 1 },
      createdAt: "2026-09-01T00:00:00.000Z"
    });

    expect(canAcknowledgeStaleSource({ ...message, plan: undefined })).toBe(
      false
    );
    expect(
      canAcknowledgeStaleSource({
        ...message,
        staleSourceConfirm: {
          ...message.staleSourceConfirm!,
          status: "completed"
        }
      })
    ).toBe(false);
  });
});

describe("pet visibility preference", () => {
  it("shows the pet by default", () => {
    expect(normalizePetVisibility(null)).toBe(true);
  });

  it("restores an explicitly hidden pet", () => {
    expect(normalizePetVisibility("hidden")).toBe(false);
    expect(normalizePetVisibility("visible")).toBe(true);
  });
});
