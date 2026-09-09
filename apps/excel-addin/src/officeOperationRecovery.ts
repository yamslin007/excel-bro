export const PENDING_OFFICE_EXECUTION_STORAGE_KEY =
  "excel-bro.pending-office-execution.v1";

export interface PendingOfficeExecution {
  schemaVersion: 1;
  operationId: string;
  conversationId: string;
  activityMessageId: string;
  workbookIdentity: string;
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

function isPendingOfficeExecution(value: unknown): value is PendingOfficeExecution {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<PendingOfficeExecution>;
  return (
    candidate.schemaVersion === 1 &&
    typeof candidate.operationId === "string" &&
    candidate.operationId.length >= 8 &&
    typeof candidate.conversationId === "string" &&
    Boolean(candidate.conversationId) &&
    typeof candidate.activityMessageId === "string" &&
    Boolean(candidate.activityMessageId) &&
    typeof candidate.workbookIdentity === "string" &&
    candidate.workbookIdentity.startsWith("url:") &&
    typeof candidate.createdAt === "string"
  );
}

export function savePendingOfficeExecution(
  value: PendingOfficeExecution,
  storage: StorageLike | null = browserStorage()
): boolean {
  if (!storage) return false;
  try {
    storage.setItem(PENDING_OFFICE_EXECUTION_STORAGE_KEY, JSON.stringify(value));
    return true;
  } catch {
    return false;
  }
}

export function loadPendingOfficeExecution(
  storage: StorageLike | null = browserStorage()
): PendingOfficeExecution | null {
  if (!storage) return null;
  try {
    const parsed: unknown = JSON.parse(
      storage.getItem(PENDING_OFFICE_EXECUTION_STORAGE_KEY) ?? "null"
    );
    if (isPendingOfficeExecution(parsed)) return parsed;
    storage.removeItem(PENDING_OFFICE_EXECUTION_STORAGE_KEY);
    return null;
  } catch {
    try {
      storage.removeItem(PENDING_OFFICE_EXECUTION_STORAGE_KEY);
    } catch {
      // 本地索引只负责重载发现；服务端记录仍是执行现场的权威来源。
    }
    return null;
  }
}

export function clearPendingOfficeExecution(
  operationId: string,
  storage: StorageLike | null = browserStorage()
): void {
  if (!storage) return;
  try {
    const current = loadPendingOfficeExecution(storage);
    if (!current || current.operationId === operationId) {
      storage.removeItem(PENDING_OFFICE_EXECUTION_STORAGE_KEY);
    }
  } catch {
    // 下次加载仍会按 operationId 查询服务端，不影响写入幂等状态。
  }
}
