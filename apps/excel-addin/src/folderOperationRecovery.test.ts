import { describe, expect, it } from "vitest";
import {
  clearPendingFolderOperation,
  loadPendingFolderOperation,
  PENDING_FOLDER_OPERATION_STORAGE_KEY,
  savePendingFolderOperation,
  type PendingFolderOperation
} from "./folderOperationRecovery";

function memoryStorage() {
  const values = new Map<string, string>();
  return {
    getItem: (key: string) => values.get(key) ?? null,
    setItem: (key: string, value: string) => values.set(key, value),
    removeItem: (key: string) => values.delete(key),
    values
  };
}

function pending(operationId = "execution-test-1"): PendingFolderOperation {
  return {
    schemaVersion: 1,
    operationId,
    sessionId: "session-1",
    folderPath: "C:\\work",
    conversationId: "conversation-1",
    activityMessageId: "message-1",
    createdAt: "2026-09-08T00:00:00.000Z",
    plan: {
      id: "plan-1",
      title: "整理报表",
      summary: "整理报表",
      assumptions: [],
      warnings: [],
      actions: []
    }
  };
}

describe("folder operation recovery storage", () => {
  it("round-trips the stable execution identity and plan", () => {
    const storage = memoryStorage();
    expect(savePendingFolderOperation(pending(), storage)).toBe(true);
    expect(loadPendingFolderOperation(storage)).toEqual(pending());
  });

  it("does not clear a newer operation with a stale completion", () => {
    const storage = memoryStorage();
    savePendingFolderOperation(pending("execution-newer"), storage);
    clearPendingFolderOperation("execution-older", storage);
    expect(loadPendingFolderOperation(storage)?.operationId).toBe(
      "execution-newer"
    );
  });

  it("discards malformed storage without creating an operation", () => {
    const storage = memoryStorage();
    storage.setItem(PENDING_FOLDER_OPERATION_STORAGE_KEY, "{broken");
    expect(loadPendingFolderOperation(storage)).toBeNull();
    expect(storage.getItem(PENDING_FOLDER_OPERATION_STORAGE_KEY)).toBeNull();
  });
});
