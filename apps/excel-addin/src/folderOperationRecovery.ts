import type { AnalysisPlan } from "./contracts";

export const PENDING_FOLDER_OPERATION_STORAGE_KEY =
  "excel-bro.pending-folder-operation.v1";

export interface PendingFolderOperation {
  schemaVersion: 1;
  operationId: string;
  sessionId: string;
  folderPath: string;
  conversationId: string;
  activityMessageId: string;
  plan: AnalysisPlan;
  createdAt: string;
}

interface StorageLike {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
}

function browserStorage(): StorageLike | null {
  try {
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

function isPendingFolderOperation(value: unknown): value is PendingFolderOperation {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<PendingFolderOperation>;
  return (
    candidate.schemaVersion === 1 &&
    typeof candidate.operationId === "string" &&
    candidate.operationId.length >= 8 &&
    typeof candidate.sessionId === "string" &&
    Boolean(candidate.sessionId) &&
    typeof candidate.folderPath === "string" &&
    Boolean(candidate.folderPath) &&
    typeof candidate.conversationId === "string" &&
    Boolean(candidate.conversationId) &&
    typeof candidate.activityMessageId === "string" &&
    Boolean(candidate.activityMessageId) &&
    typeof candidate.createdAt === "string" &&
    Boolean(candidate.plan) &&
    typeof candidate.plan === "object" &&
    typeof candidate.plan.id === "string" &&
    Array.isArray(candidate.plan.actions)
  );
}

export function savePendingFolderOperation(
  value: PendingFolderOperation,
  storage: StorageLike | null = browserStorage()
): boolean {
  if (!storage) return false;
  try {
    storage.setItem(PENDING_FOLDER_OPERATION_STORAGE_KEY, JSON.stringify(value));
    return true;
  } catch {
    return false;
  }
}

export function loadPendingFolderOperation(
  storage: StorageLike | null = browserStorage()
): PendingFolderOperation | null {
  if (!storage) return null;
  try {
    const parsed: unknown = JSON.parse(
      storage.getItem(PENDING_FOLDER_OPERATION_STORAGE_KEY) ?? "null"
    );
    if (isPendingFolderOperation(parsed)) return parsed;
    storage.removeItem(PENDING_FOLDER_OPERATION_STORAGE_KEY);
    return null;
  } catch {
    try {
      storage.removeItem(PENDING_FOLDER_OPERATION_STORAGE_KEY);
    } catch {
      // 存储被宿主禁用时只放弃界面自动挂接，不影响服务端持久任务。
    }
    return null;
  }
}

export function clearPendingFolderOperation(
  operationId: string,
  storage: StorageLike | null = browserStorage()
): void {
  if (!storage) return;
  try {
    const current = loadPendingFolderOperation(storage);
    if (!current || current.operationId === operationId) {
      storage.removeItem(PENDING_FOLDER_OPERATION_STORAGE_KEY);
    }
  } catch {
    // 清理失败会在下次加载时由 operationId 去重，不得影响执行结果。
  }
}
