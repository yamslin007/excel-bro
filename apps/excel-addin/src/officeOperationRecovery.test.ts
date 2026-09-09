import { describe, expect, it } from "vitest";
import {
  clearPendingOfficeExecution,
  loadPendingOfficeExecution,
  PENDING_OFFICE_EXECUTION_STORAGE_KEY,
  savePendingOfficeExecution,
  type PendingOfficeExecution
} from "./officeOperationRecovery";

function memoryStorage() {
  const values = new Map<string, string>();
  return {
    getItem: (key: string) => values.get(key) ?? null,
    setItem: (key: string, value: string) => values.set(key, value),
    removeItem: (key: string) => values.delete(key)
  };
}

function pending(operationId = "office-execution-1"): PendingOfficeExecution {
  return {
    schemaVersion: 1,
    operationId,
    conversationId: "conversation-1",
    activityMessageId: "activity-1",
    workbookIdentity: "url:c:/work/report.xlsx",
    createdAt: "2026-09-09T00:00:00.000Z"
  };
}

describe("office operation recovery discovery storage", () => {
  it("stores only the discovery identity, not the authoritative checkpoint", () => {
    const storage = memoryStorage();
    expect(savePendingOfficeExecution(pending(), storage)).toBe(true);
    expect(loadPendingOfficeExecution(storage)).toEqual(pending());
    expect(storage.getItem(PENDING_OFFICE_EXECUTION_STORAGE_KEY)).not.toContain(
      "beforeSnapshot"
    );
  });

  it("does not clear a newer Office operation from a stale completion", () => {
    const storage = memoryStorage();
    savePendingOfficeExecution(pending("office-execution-new"), storage);
    clearPendingOfficeExecution("office-execution-old", storage);
    expect(loadPendingOfficeExecution(storage)?.operationId).toBe(
      "office-execution-new"
    );
  });

  it("drops malformed or unstable workbook identities", () => {
    const storage = memoryStorage();
    storage.setItem(
      PENDING_OFFICE_EXECUTION_STORAGE_KEY,
      JSON.stringify({ ...pending(), workbookIdentity: "unsaved" })
    );
    expect(loadPendingOfficeExecution(storage)).toBeNull();
  });
});
