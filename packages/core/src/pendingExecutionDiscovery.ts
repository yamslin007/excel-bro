/**
 * 「重新发现未完成执行」的共享判据（Excel / WPS 两端与服务端同一份口径）。
 *
 * 存在理由（2026-10-27 真实 WPS 现场）：浏览器 `localStorage` 里的**单条**待恢复索引
 * 曾经是唯一发现入口。完整退出 WPS 再打开后，那个键没了 —— 服务端账本里的进度
 * （`nextActionIndex` / 在途动作）与原表写入现场都还在，界面却再也找不到这次执行，
 * 只剩一条历史任务的"已撤销"。
 *
 * 修法不是"再存一份索引"，而是把浏览器索引**降级为缓存**：由后端按**当前宿主的
 * 当前工作簿稳定身份**（WPS = 已保存 `.xlsx` 的小写磁盘完整路径 `wps-path:`；
 * Excel = `url:`）精确查找未完成执行。本模块只管"拿到响应后怎么判"，不碰网络、
 * 不碰宿主：
 *
 * - **身份必须完整相等**：不许按文件名、活动表名、表结构、相同单元格值或"最近一条"
 *   推断归属；不同路径的同名簿互不发现。
 * - **发现只恢复入口，不恢复写入资格**：认领、核对、续跑仍走既有状态查询与材料水合。
 * - **多候选不自动选**：多于一条（或结果被截断）时只提供选择入口，界面与判据都不得
 *   替用户挑"最新的一条"。
 * - **已完成不是待续写**：`completed` 一律排除。
 * - **响应只许带定位与计数**：执行权令牌、单元格内容、恢复包正文、完整值计划出现即拒绝。
 * - **查不到 ≠ 没有**：失败与"没有未完成任务"是两个结论，调用方必须分开表达。
 */

import capabilities from "../../../config/capabilities.json";

/** 发现协议版本。读到不认识的版本必须如实说"无法发现"，不许猜。 */
export const PENDING_DISCOVERY_SCHEMA_VERSION = 1;

/** 单次发现最多带回多少条候选（超出即截断，且**必须**明确说被截断）。 */
export const PENDING_DISCOVERY_LIMIT = capabilities.taskRecovery.pendingDiscoveryMaxCandidates;

/** 身份串长度上限（与服务端登记时的 `workbookIdentity` 上限同量级）。 */
export const PENDING_DISCOVERY_MAX_IDENTITY_CHARS =
  capabilities.taskRecovery.pendingDiscoveryMaxIdentityChars;

/** 两个宿主的身份命名空间。前缀不同即不同命名空间，互不冒充。 */
export type PendingExecutionNamespace = "wps" | "office";

/** 未完成执行的状态。`completed` 刻意不在此列 —— 它不该作为待续写任务出现。 */
export type PendingExecutionStatus = "running" | "needs_reconciliation" | "failed";

/**
 * 一条未完成执行的**最小**元数据。
 *
 * 只有"定位 + 时间 + 进度"三样：`operationId` 用来继续查真实状态，时间与进度用来让
 * 用户分辨是哪一次。**没有** 单元格值、写前/写后矩阵、计划正文、恢复包正文，也没有
 * `ownerToken`（执行权只能通过认领接口取得）。
 */
export interface PendingExecutionCandidate {
  operationId: string;
  taskId: string;
  status: PendingExecutionStatus;
  /** 登记时间（epoch 毫秒）。 */
  createdAt: number;
  /** 最后一次账本变动（epoch 毫秒）。 */
  updatedAt: number;
  nextActionIndex: number;
  totalActions: number;
  inFlightActionIndex: number | null;
  ownerActive: boolean;
  requiresReconciliation: boolean;
  canResume: boolean;
}

export interface PendingExecutionDiscoveryResponse {
  schemaVersion: 1;
  namespace: PendingExecutionNamespace;
  /** 查询用的工作簿身份；与请求的身份**完整相等**才认。 */
  workbookIdentity: string;
  limit: number;
  /** 还有超出 `limit` 的候选没有带回。截断时必须让用户知道。 */
  truncated: boolean;
  candidates: PendingExecutionCandidate[];
}

/**
 * 发现结论。
 *
 * `mismatch` 与 `none` 是两件事：前者是"服务端给的不是这个工作簿的记录"，后者才是
 * "这个工作簿确实没有未完成执行"。把它们合并会让一次身份错配被读成"任务已经没了"。
 */
export type PendingExecutionSelection =
  | { kind: "none" }
  | { kind: "single"; candidate: PendingExecutionCandidate }
  | { kind: "choose"; candidates: PendingExecutionCandidate[]; truncated: boolean }
  | { kind: "mismatch"; detail: string };

const RESPONSE_KEYS = [
  "schemaVersion",
  "namespace",
  "workbookIdentity",
  "limit",
  "truncated",
  "candidates"
] as const;

const CANDIDATE_KEYS = [
  "operationId",
  "taskId",
  "status",
  "createdAt",
  "updatedAt",
  "nextActionIndex",
  "totalActions",
  "inFlightActionIndex",
  "ownerActive",
  "requiresReconciliation",
  "canResume"
] as const;

const IDENTITY_PREFIX: Record<PendingExecutionNamespace, string> = {
  wps: "wps-path:",
  office: "url:"
};

function isNonEmptyText(value: unknown, max: number): value is string {
  return typeof value === "string" && value.length > 0 && value.length <= max;
}

/** epoch 毫秒：必须是有限、非负的整数秒级时间戳（服务端由账本的秒值换算）。 */
function isTimestamp(value: unknown): value is number {
  return typeof value === "number" && Number.isSafeInteger(value) && value > 0;
}

/**
 * 白名单准入闸：形状不对、夹带执行权令牌/矩阵/恢复材料、进度自相矛盾，一律拒绝。
 *
 * "多带了几个字段"不算兼容 —— 出站闸的另一侧（服务端）就是按 `extra="forbid"` 写的，
 * 客户端如果宽松收下，泄漏过的那一版服务端会被读成"正常返回"。
 */
export function assertPendingExecutionDiscovery(
  value: unknown
): asserts value is PendingExecutionDiscoveryResponse {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("服务返回的未完成执行发现结果不是有效对象");
  }
  const record = value as Record<string, unknown>;
  if (Object.keys(record).some((key) => !(RESPONSE_KEYS as readonly string[]).includes(key))) {
    throw new Error("未完成执行发现结果含未知字段（不允许携带执行权令牌、单元格内容或恢复材料）");
  }
  if (record.schemaVersion !== PENDING_DISCOVERY_SCHEMA_VERSION) {
    throw new Error("未完成执行发现结果的格式版本无法识别，不能用于恢复入口");
  }
  const namespace = record.namespace;
  if (namespace !== "wps" && namespace !== "office") {
    throw new Error("未完成执行发现结果的宿主命名空间无效");
  }
  const identity = record.workbookIdentity;
  if (!isNonEmptyText(identity, PENDING_DISCOVERY_MAX_IDENTITY_CHARS) ||
      !identity.startsWith(IDENTITY_PREFIX[namespace])) {
    throw new Error("未完成执行发现结果的工作簿身份无效");
  }
  if (typeof record.limit !== "number" || !Number.isInteger(record.limit) ||
      record.limit < 1 || record.limit > PENDING_DISCOVERY_LIMIT) {
    throw new Error("未完成执行发现结果的限额无效");
  }
  if (typeof record.truncated !== "boolean") {
    throw new Error("未完成执行发现结果必须明确说明是否截断");
  }
  if (!Array.isArray(record.candidates) || record.candidates.length > record.limit) {
    throw new Error("未完成执行发现结果的候选列表无效");
  }
  for (const candidate of record.candidates) {
    assertPendingExecutionCandidate(candidate);
  }
}

function assertPendingExecutionCandidate(value: unknown): asserts value is PendingExecutionCandidate {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("未完成执行候选不是有效对象");
  }
  const candidate = value as Record<string, unknown>;
  if (Object.keys(candidate).some((key) => !(CANDIDATE_KEYS as readonly string[]).includes(key))) {
    throw new Error("未完成执行候选含未知字段（不允许携带单元格内容或完整计划）");
  }
  if (!isNonEmptyText(candidate.operationId, 200) || !isNonEmptyText(candidate.taskId, 200)) {
    throw new Error("未完成执行候选缺少执行身份");
  }
  if (!["running", "needs_reconciliation", "failed"].includes(String(candidate.status))) {
    throw new Error("已完成或未知状态的执行不能作为待续写任务");
  }
  if (!isTimestamp(candidate.createdAt) || !isTimestamp(candidate.updatedAt)) {
    throw new Error("未完成执行候选的时间无效");
  }
  const total = candidate.totalActions;
  const next = candidate.nextActionIndex;
  if (typeof total !== "number" || !Number.isInteger(total) || total < 1 ||
      typeof next !== "number" || !Number.isInteger(next) || next < 0 || next > total) {
    throw new Error("未完成执行候选的动作进度无效");
  }
  const inFlight = candidate.inFlightActionIndex;
  if (inFlight !== null &&
      (typeof inFlight !== "number" || !Number.isInteger(inFlight) || inFlight < 0 || inFlight >= total)) {
    throw new Error("未完成执行候选的在途动作位置无效");
  }
  for (const flag of ["ownerActive", "requiresReconciliation", "canResume"] as const) {
    if (typeof candidate[flag] !== "boolean") {
      throw new Error("未完成执行候选缺少执行状态标记");
    }
  }
}

/**
 * 排序：**有在途动作的优先，写前失败的最后**，同档按更新时间倒序、再按 `operationId` 定序。
 *
 * 为什么不是"最新优先"：一条刚失败的**写前**记录（一个动作都没推进、也没有在途动作）
 * 如果排在真正的在途执行前面，用户会先看到那条、以为没写，而实际有一步已经落表。
 * 所以写前失败固定排最后；有在途动作（现场可能已被改过）的固定排最前。这里也不替
 * 用户选，只保证最危险的那条不会被名单尾部或截断盖住。
 */
export function orderPendingExecutionCandidates(
  candidates: readonly PendingExecutionCandidate[]
): PendingExecutionCandidate[] {
  const rank = (candidate: PendingExecutionCandidate): number => {
    // 第一档：有在途动作 —— 现场可能已经被改过，必须先核对。
    if (candidate.inFlightActionIndex != null || candidate.status === "needs_reconciliation") return 0;
    // 第三档：写前失败，一个动作都没推进 —— 最不该挡住别的记录。
    if (candidate.status === "failed" &&
        candidate.nextActionIndex === 0 &&
        candidate.inFlightActionIndex == null) return 2;
    return 1;
  };
  return [...candidates].sort((left, right) =>
    rank(left) - rank(right) ||
    right.updatedAt - left.updatedAt ||
    (left.operationId < right.operationId ? -1 : left.operationId > right.operationId ? 1 : 0)
  );
}

/**
 * 由发现响应得出"接下来怎么处理"。
 *
 * 纪律：**只有恰好一条且没有被截断**时才直接采用；其余情况一律交给用户选择。
 * 去重按 `operationId`（同一条被带回两次不该变成"两个任务"），已完成的一律丢弃。
 */
export function pendingExecutionSelection(
  response: PendingExecutionDiscoveryResponse,
  workbookIdentity: string
): PendingExecutionSelection {
  if (response.workbookIdentity !== workbookIdentity) {
    return {
      kind: "mismatch",
      detail: "服务端返回的不是当前这个工作簿的执行记录，不据此恢复；请确认打开的是原工作簿。"
    };
  }
  const seen = new Set<string>();
  const visible: PendingExecutionCandidate[] = [];
  for (const candidate of response.candidates) {
    /*
     * `completed` 在类型上已经进不来（`PendingExecutionStatus` 不含它，准入闸也拒绝），
     * 但选择函数可能被"还没过闸"的调用方直接用 —— 运行时这一句让它照样不许出现在
     * 待续写名单里，而不是只靠类型保证。
     */
    const completed = (candidate as { status: string }).status === "completed";
    if (completed || seen.has(candidate.operationId)) continue;
    seen.add(candidate.operationId);
    visible.push(candidate);
  }
  const ordered = orderPendingExecutionCandidates(visible).slice(0, response.limit);
  const truncated = response.truncated || ordered.length < visible.length;
  if (ordered.length === 0) return { kind: "none" };
  if (ordered.length === 1 && !truncated) return { kind: "single", candidate: ordered[0] };
  return { kind: "choose", candidates: ordered, truncated };
}

/** 候选行的一句话说明：只有进度与状态，没有任何单元格内容。 */
export function pendingExecutionCandidateLabel(candidate: PendingExecutionCandidate): string {
  const inFlight = candidate.inFlightActionIndex;
  const parts = [
    inFlight != null ? `第 ${inFlight + 1} 步待核对` : `已写 ${candidate.nextActionIndex} 步`,
    `共 ${candidate.totalActions} 步`
  ];
  if (candidate.ownerActive) parts.push("另一个窗口仍持执行权");
  else if (candidate.status === "failed") parts.push("写前失败，未推进动作");
  return parts.join(" · ");
}

/** 多候选/截断时给用户的说明（两端共用，不各写一套）。 */
export const PENDING_DISCOVERY_CHOOSE_DETAIL =
  "这个工作簿有多条未完成的写入。请选择要核对的那一次；不会自动选最新的一条，也不会写入任何单元格。";

export const PENDING_DISCOVERY_TRUNCATED_DETAIL =
  "候选较多，只列出了最需要先处理的一部分；处理或清除后重新核对可以看到其余记录。";

/**
 * 查询失败时**唯一**允许的说法。
 *
 * "查不到"不是"没有"：服务没起来、端口换了、账本忙，都会走到这里。把它说成"没有未完成
 * 的任务"会让用户以为写入从未发生。
 */
export const PENDING_DISCOVERY_UNAVAILABLE_DETAIL =
  "暂时读不到服务端的执行记录；这不等于没有未完成的写入，也不代表原表没被改过。";
