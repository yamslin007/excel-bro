import { describe, expect, it } from "vitest";
import {
  deleteConversationFromHistory,
  normalizePetVisibility,
  normalizeStoredVerification
} from "./App";

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

describe("pet visibility preference", () => {
  it("shows the pet by default", () => {
    expect(normalizePetVisibility(null)).toBe(true);
  });

  it("restores an explicitly hidden pet", () => {
    expect(normalizePetVisibility("hidden")).toBe(false);
    expect(normalizePetVisibility("visible")).toBe(true);
  });
});
