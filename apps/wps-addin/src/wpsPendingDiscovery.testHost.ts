/**
 * 测试用的**服务端账本发现**替身（`GET /api/{wps|office}/executions/pending`）。
 *
 * 为什么要有它：以前面板只有一个发现入口 —— 浏览器 `localStorage` 里的单条索引。
 * 于是"完整退出 WPS 后重开"这类用例可以在**保留索引**的前提下通过，把真正的缺口
 * （索引丢了就没入口）挡在测试之外。现在索引只是缓存，权威发现来自账本，用例就必须
 * 有一个"账本侧"的替身，否则只能手工塞 `operationId` —— 那等于把被测的发现路径整段
 * 跳过，通过与否都不说明问题。
 *
 * 语义与服务端（`task_ledger.pending_office_executions` / `main._pending_executions`）
 * 对齐：
 *
 * - 只按 `workbookIdentity` **完整相等**过滤 —— 同名不同路径互不可见；
 * - 只认 `running / needs_reconciliation / failed`（`completed` 不是待续写任务）；
 * - 已处置（`dismiss`）的记录不再出现，但**账本状态、进度、材料一律不动**；
 * - 排序与截断的判据在共享层（`orderPendingExecutionCandidates` / `limit+1`），
 *   替身不另写一套，只做"服务端那一层的过滤与投影"。
 */

import { orderPendingExecutionCandidates, PENDING_DISCOVERY_LIMIT, PENDING_DISCOVERY_SCHEMA_VERSION, type PendingExecutionCandidate, type PendingExecutionDiscoveryResponse } from "@excel-bro/core/pendingExecutionDiscovery";

export type PendingDiscoveryJob = {
  operationId: string;
  taskId: string;
  workbookIdentity: string;
  status: PendingExecutionCandidate["status"];
  nextActionIndex: number;
  totalActions: number;
  inFlightActionIndex?: number | null;
  ownerActive?: boolean;
  requiresReconciliation?: boolean;
  canResume?: boolean;
  createdAt?: number;
  /** 账本最后一次变动：处置元数据的判据就是"处置时间 vs 这一条时间"。 */
  updatedAt?: number;
};

type Stored = Required<Omit<PendingDiscoveryJob, "inFlightActionIndex">> & {
  inFlightActionIndex: number | null;
  dismissedAt: number | null;
};

export type PendingDiscoveryBackend = {
  /** 账本里登记/更新一次执行（`registerWpsExecution` 的那一侧）。 */
  publish(job: PendingDiscoveryJob): void;
  /** `GET …/executions/pending`：只读，调用不改变任何记录。 */
  list(workbookIdentity: string): Promise<PendingExecutionDiscoveryResponse>;
  /** `POST …/executions/{id}/dismiss-discovery`：只写处置时间戳。 */
  dismiss(operationId: string): void;
  /** 服务不可达：下一次 `list` 直接拒绝（**不是**返回空名单）。 */
  failNextList(error?: Error): void;
  /** 服务一直不可达：之后每次 `list` 都拒绝，直到 `clearFailure()`。 */
  failLists(error?: Error): void;
  clearFailure(): void;
  /** 扣住下一次 `list` 的响应，用来固定"旧查询迟到"的时序。 */
  holdNextList(): void;
  /** 放行被扣住的那一次（返回被扣住的次数）。 */
  releaseHeldLists(): number;
  /** 每一次 `list` 收到的身份，按调用顺序。 */
  listCalls(): string[];
  /** 每一次 `dismiss` 收到的执行身份。 */
  dismissCalls(): string[];
  /** 当前账本里还剩哪些执行（含已完成/已处置），用于"处置不改账本"的断言。 */
  entries(): Stored[];
  /** 按执行身份取一条记录的**快照**（不存在就是 `null`）。 */
  snapshot(operationId: string): Stored | null;
  reset(): void;
};

function normalize(job: PendingDiscoveryJob): Stored {
  const now = Date.now();
  const createdAt = job.createdAt ?? now;
  return {
    operationId: job.operationId,
    taskId: job.taskId,
    workbookIdentity: job.workbookIdentity,
    status: job.status,
    nextActionIndex: job.nextActionIndex,
    totalActions: job.totalActions,
    inFlightActionIndex: job.inFlightActionIndex ?? null,
    ownerActive: job.ownerActive ?? false,
    requiresReconciliation: job.requiresReconciliation ?? job.status === "needs_reconciliation",
    canResume: job.canResume ?? true,
    createdAt,
    updatedAt: job.updatedAt ?? createdAt,
    dismissedAt: null
  };
}

function project(entry: Stored): PendingExecutionCandidate {
  // 只投影"定位 + 时间 + 进度"：不含执行权令牌、单元格值、计划与恢复材料。
  return {
    operationId: entry.operationId,
    taskId: entry.taskId,
    status: entry.status,
    createdAt: entry.createdAt,
    updatedAt: entry.updatedAt,
    nextActionIndex: entry.nextActionIndex,
    totalActions: entry.totalActions,
    inFlightActionIndex: entry.inFlightActionIndex,
    ownerActive: entry.ownerActive,
    requiresReconciliation: entry.requiresReconciliation,
    canResume: entry.canResume
  };
}

export function createPendingDiscoveryBackend(
  options: { limit?: number; namespace?: "wps" | "office" } = {}
): PendingDiscoveryBackend {
  const limit = options.limit ?? PENDING_DISCOVERY_LIMIT;
  const namespace = options.namespace ?? "wps";
  const entries = new Map<string, Stored>();
  const listed: string[] = [];
  const dismissed: string[] = [];
  let failWith: Error | null = null;
  let stickyFailure: Error | null = null;
  let held = 0;
  const pendingReleases: Array<() => void> = [];

  const backend: PendingDiscoveryBackend = {
    publish(job) {
      entries.set(job.operationId, normalize(job));
    },
    list(workbookIdentity) {
      listed.push(workbookIdentity);
      if (stickyFailure) return Promise.reject(stickyFailure);
      if (failWith) {
        const error = failWith;
        failWith = null;
        return Promise.reject(error);
      }
      const visible = [...entries.values()].filter((entry) =>
        entry.workbookIdentity === workbookIdentity &&
        (entry.status === "running" || entry.status === "needs_reconciliation" || entry.status === "failed") &&
        // 处置只压住"这一版"：账本之后又有变动（`updatedAt` 更新）就重新出现。
        (entry.dismissedAt === null || entry.dismissedAt < entry.updatedAt)
      );
      const ordered = orderPendingExecutionCandidates(visible.map(project));
      const truncated = ordered.length > limit;
      const response: PendingExecutionDiscoveryResponse = {
        schemaVersion: PENDING_DISCOVERY_SCHEMA_VERSION,
        namespace,
        workbookIdentity,
        limit,
        truncated,
        candidates: ordered.slice(0, limit)
      };
      if (held > 0) {
        held -= 1;
        return new Promise((resolve) => { pendingReleases.push(() => resolve(response)); });
      }
      return Promise.resolve(response);
    },
    dismiss(operationId) {
      dismissed.push(operationId);
      const entry = entries.get(operationId);
      // 只写处置时间戳：状态、进度、材料、执行权一概不动（与服务端一致）。
      if (entry) entries.set(operationId, { ...entry, dismissedAt: Date.now() });
    },
    failNextList(error = new Error("本机服务不可达")) {
      failWith = error;
    },
    failLists(error = new Error("本机服务不可达")) {
      stickyFailure = error;
    },
    clearFailure() {
      stickyFailure = null;
      failWith = null;
    },
    holdNextList() {
      held += 1;
    },
    releaseHeldLists() {
      const releases = pendingReleases.splice(0);
      releases.forEach((release) => release());
      return releases.length;
    },
    listCalls() {
      return [...listed];
    },
    dismissCalls() {
      return [...dismissed];
    },
    entries() {
      return [...entries.values()].map((entry) => ({ ...entry }));
    },
    snapshot(operationId) {
      const entry = entries.get(operationId);
      return entry ? { ...entry } : null;
    },
    reset() {
      entries.clear();
      listed.length = 0;
      dismissed.length = 0;
      failWith = null;
      stickyFailure = null;
      held = 0;
      pendingReleases.length = 0;
    }
  };
  return backend;
}

/** 整个测试文件共享的"那台账本服务"（`beforeEach` 里 reset）。 */
let backend = createPendingDiscoveryBackend();

export function pendingDiscoveryBackend(): PendingDiscoveryBackend {
  return backend;
}

export function resetPendingDiscoveryBackend(
  options?: { limit?: number; namespace?: "wps" | "office" }
): PendingDiscoveryBackend {
  backend = createPendingDiscoveryBackend(options);
  return backend;
}

/**
 * 给"只关心别的功能"的测试文件用的三件套替身。
 *
 * 为什么必须显式给：面板挂载后就会**按当前工作簿身份发现未完成执行**，那是真实网络调用。
 * 不替身的话，一次全量测试会按用户当前打开的工作簿去查他那台 `127.0.0.1:8765`
 * （2026-10-27 实际发生过）。默认账本是空的 ⇒ 这些用例看到的就是"没有未完成任务"。
 *
 * 需要"真的发现出东西"的用例请直接用 `pendingDiscoveryBackend().publish(...)`，
 * 不要另写一份假响应。
 */
export function pendingDiscoveryApiStubs(): {
  listPendingWpsExecutions: (workbookIdentity: string) => Promise<PendingExecutionDiscoveryResponse>;
  listPendingOfficeExecutions: (workbookIdentity: string) => Promise<PendingExecutionDiscoveryResponse>;
  dismissWpsExecutionDiscovery: (operationId: string) => Promise<void>;
} {
  return {
    listPendingWpsExecutions: (workbookIdentity) => pendingDiscoveryBackend().list(workbookIdentity),
    listPendingOfficeExecutions: (workbookIdentity) => pendingDiscoveryBackend().list(workbookIdentity),
    dismissWpsExecutionDiscovery: async (operationId) => { pendingDiscoveryBackend().dismiss(operationId); }
  };
}
