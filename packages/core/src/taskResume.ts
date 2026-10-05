/**
 * 「任务中断后继续」的共享口径（Excel / WPS 两端共用，L2 纯语义层）。
 *
 * 设计边界（与项目的既有纪律一致，不得在此放宽）：
 *
 * 1. **这不是第二套恢复账本。** 真实执行身份仍由既有机制提供 ——
 *    规划走服务端 `planning_tasks`（`turnId`），写入走 `operationId` /
 *    `office_executions`（`executionId`）。本模块只描述"用户看到的这张中断卡
 *    是什么状态、下一步能做什么"，并把可以安全持久化的那部分收敛成一个
 *    **版本化、显式白名单**的结构。它不复制服务端的检查点内容。
 *
 * 2. **阶段、阻塞原因、能否安全继续是三件不同的事**，必须分开表达：
 *    - `phase` 说任务走到哪一步；
 *    - `blocker` 说它**为什么停**（断网 / 等回答 / 需重新授权 / 结果不确定 …）；
 *    - `resumability` 说**下一步能安全做什么**。
 *    把"网络中断 / 等待回答 / 需要重新授权 / 执行结果不确定"一律归为
 *    "澄清卡已失效"，会让用户在范围根本没变时被告知"数据已变化"，
 *    或者在写入结果未知时被诱导重跑 —— 两者都是对用户的假话。
 *
 * 3. **白名单是唯一的持久化准入。** `sanitizeTaskResumeRecord` 只重建
 *    显式列出的字段，任何多余键（`plan`、`values`、`matrix`、
 *    `sourceSnapshots`、宿主对象引用、完整路径 …）都会被丢弃。
 *    因此"原始数据行 / 完整源值 / 目标矩阵"结构上无法经由本结构落盘。
 *
 * 4. **没有原需求就不算可恢复。** `originalPrompt` 为空时消毒返回 `null` ——
 *    宁可让调用方明确反馈"无法恢复"，也不用问题摘要冒充完整任务。
 */

/** 任务走到哪一步。与"为什么停"正交。 */
export type TaskPhase =
  | "awaiting_clarification"
  | "answer_submitted"
  | "reading_source"
  | "planning"
  | "preview_ready"
  | "executing";

/** 任务**为什么**停在当前阶段。 */
export type TaskBlocker =
  | "none"
  | "network"
  | "user_input"
  | "authorization"
  | "source_changed"
  | "execution_unknown"
  | "storage";

/** 下一步**能安全做什么**。`unavailable` 表示不可续跑，需明确反馈。 */
export type TaskResumability =
  | "checkpoint"
  | "replan"
  | "reauthorize"
  | "reconcile"
  | "unavailable";

/**
 * 原表的写入状态。**未知就是未知**，不得默认成"未修改"：
 * 一次执行可能在服务端已写、回执丢失，此时说"原表未修改"是假话。
 */
export type TaskWriteState =
  | "not_written"
  | "confirmed_result"
  | "needs_review"
  | "unknown";

/** 用户显式确认过的一条信息（不是本端从工作表读出来的来源值）。 */
export interface TaskResumeConfirmation {
  /** 当时确认的标签，例如澄清卡的选项文案。 */
  label: string;
  /** 用户补充/确认的内容本体。 */
  value: string;
  /**
   * 这道题的**稳定身份**。
   *
   * 带上它，用户改主意时新的确认才会**替换**旧的；不带就只能追加，
   * 于是提示词里同时出现"按 A 处理"与"按 B 处理"两条冲突确认。
   * 旧记录没有这个字段，按"无从判断"处理（追加），不猜。
   */
  questionId?: string;
}

/**
 * 可持久化的任务恢复记录（`schemaVersion: 1`）。
 *
 * 字段全是"定位 + 阶段 + 原需求 + 用户确认"，**没有任何单元格内容**。
 * 授权是"版本"而不是"永久有效凭据"：`authorizationRevision` 只表示
 * "用户当时确认的是第几版范围"，恢复时必须重新读取真实宿主元数据并重新确认。
 */
export interface TaskResumeRecord {
  schemaVersion: 1;
  /** 用户任务的稳定身份。重新规划可换 `turnId`，但 `taskId` 不变。 */
  taskId: string;
  phase: TaskPhase;
  blocker: TaskBlocker;
  resumability: TaskResumability;
  writeState: TaskWriteState;
  /** 原需求全文。缺失即不可恢复，不拿摘要冒充。 */
  originalPrompt: string;
  confirmations: TaskResumeConfirmation[];
  /** 用户当时确认的数据授权版本。**不是**许可凭据。 */
  authorizationRevision: number;
  /** 某次规划尝试的身份（可选）。 */
  turnId?: string;
  /** 真实执行身份（可选），重试/重连不得新建。 */
  operationId?: string;
  /** 一次继续动作的幂等身份（可选）。 */
  requestId?: string;
  createdAt: string;
  updatedAt: string;
}

export const TASK_RESUME_SCHEMA_VERSION = 1 as const;

/** 记录条数与文本长度上限：中断卡只需定位与说明，不需要正文搬运。 */
export const TASK_RESUME_MAX_CONFIRMATIONS = 8;
export const TASK_RESUME_MAX_TEXT = 4000;

export const TASK_PHASE_LABEL: Record<TaskPhase, string> = {
  awaiting_clarification: "等待你回答问题",
  answer_submitted: "已提交回答，尚未生成预览",
  reading_source: "读取数据或在本机计算",
  planning: "正在规划",
  preview_ready: "预览已生成，等待你确认",
  executing: "正在写入"
};

export const TASK_BLOCKER_LABEL: Record<TaskBlocker, string> = {
  none: "未中断",
  network: "连接中断（断网 / 服务重启 / 任务窗格重载）",
  user_input: "在等你回答",
  authorization: "需要你重新确认数据范围",
  source_changed: "数据来源发生变化，需要重新读取",
  execution_unknown: "写入结果无法确认，需要先核对",
  storage: "本机保存恢复记录失败"
};

export const TASK_RESUMABILITY_LABEL: Record<TaskResumability, string> = {
  checkpoint: "可从上次保存的进度继续",
  replan: "保留需求后重新读取并重新生成预览",
  reauthorize: "确认数据范围后重新读取",
  reconcile: "先核对原表状态",
  unavailable: "无法继续"
};

export const TASK_WRITE_STATE_LABEL: Record<TaskWriteState, string> = {
  not_written: "原表未写入",
  confirmed_result: "已确认写入结果",
  needs_review: "需要核对原表",
  unknown: "写入状态未知，需要核对"
};

export type TaskResumeAction =
  | "continue_checkpoint"
  | "regenerate_preview"
  | "reauthorize"
  | "reconcile"
  | "unavailable";

export interface TaskResumeDecision {
  action: TaskResumeAction;
  /** 主按钮文案。 */
  actionLabel: string;
  /**
   * 对"这次继续会做什么"的诚实说明。
   * **只有 `checkpoint` 才允许说"从上次保存的进度继续"**；其余一律说明会
   * 重新读取 / 重新生成预览。
   */
  detail: string;
  canContinue: boolean;
}

/**
 * 由记录推导"下一步能做什么"。
 *
 * 判据顺序刻意与用户能看懂的顺序一致：不可恢复 > 需要先核对 > 需要重新授权 >
 * 有检查点 > 退回重新规划。
 */
export function taskResumeDecision(record: TaskResumeRecord): TaskResumeDecision {
  if (record.blocker === "storage" || record.resumability === "unavailable") {
    return {
      action: "unavailable",
      actionLabel: "无法继续",
      detail:
        "本机没能保存这次任务需要的恢复信息，也没有可用的服务端检查点。请重新发送需求。",
      canContinue: false
    };
  }
  if (
    record.blocker === "execution_unknown" ||
    record.resumability === "reconcile"
  ) {
    return {
      action: "reconcile",
      actionLabel: "核对待续跑状态",
      detail:
        "上一次写入的结果无法确认。继续之前会先核对原表当前状态，核对不通过就保持挂起，不会盲目重跑。",
      canContinue: true
    };
  }
  if (
    record.blocker === "authorization" ||
    record.resumability === "reauthorize"
  ) {
    return {
      action: "reauthorize",
      actionLabel: "确认此范围并继续",
      detail:
        "先读取当前工作簿与工作表的真实情况并重新确认数据范围，再重新读取、重新生成预览；旧的范围确认不再生效。",
      canContinue: true
    };
  }
  if (record.resumability === "checkpoint") {
    return {
      action: "continue_checkpoint",
      actionLabel: "继续任务",
      detail:
        "从上次保存的进度继续，不会重做已完成的步骤；写入前仍然要经过预览与你的确认。",
      canContinue: true
    };
  }
  return {
    action: "regenerate_preview",
    actionLabel: "继续任务",
    detail:
      "保留原需求与你已确认的信息，重新读取并重新生成预览。这是重新生成，不是从旧步骤精确续跑。",
    canContinue: true
  };
}

function isTaskPhase(value: unknown): value is TaskPhase {
  return typeof value === "string" && value in TASK_PHASE_LABEL;
}

function isTaskBlocker(value: unknown): value is TaskBlocker {
  return typeof value === "string" && value in TASK_BLOCKER_LABEL;
}

function isTaskResumability(value: unknown): value is TaskResumability {
  return typeof value === "string" && value in TASK_RESUMABILITY_LABEL;
}

function isTaskWriteState(value: unknown): value is TaskWriteState {
  return typeof value === "string" && value in TASK_WRITE_STATE_LABEL;
}

function boundedText(value: unknown, max = TASK_RESUME_MAX_TEXT): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;
  return trimmed.length > max ? trimmed.slice(0, max) : trimmed;
}

/** 可选字符串字段：只有真的是非空字符串才保留，否则整个字段丢掉。 */
function optionalText(value: unknown): string | undefined {
  const text = boundedText(value);
  return text ?? undefined;
}

function optionalTimestamp(value: unknown, fallback: string): string {
  return typeof value === "string" && value.trim() ? value : fallback;
}

function sanitizeConfirmations(value: unknown): TaskResumeConfirmation[] {
  if (!Array.isArray(value)) return [];
  const seen = new Set<string>();
  const result: TaskResumeConfirmation[] = [];
  for (const entry of value) {
    if (result.length >= TASK_RESUME_MAX_CONFIRMATIONS) break;
    if (!entry || typeof entry !== "object") continue;
    const label = boundedText((entry as { label?: unknown }).label, 200);
    const content = boundedText((entry as { value?: unknown }).value);
    if (!label || !content) continue;
    const key = `${label}\u0000${content}`;
    if (seen.has(key)) continue;
    seen.add(key);
    // `questionId` 也在白名单里：它是"同一道题"的定位键，不是单元格内容。
    // 丢掉它，用户改过的答案就会退化成"追加"，提示词里留下互相冲突的两条。
    const questionId = boundedText(
      (entry as { questionId?: unknown }).questionId,
      120
    );
    result.push(
      questionId ? { label, value: content, questionId } : { label, value: content }
    );
  }
  return result;
}

/**
 * 消毒一条"可能来自存储"的恢复记录。
 *
 * **显式白名单**：重建一个只有下述键的新对象，多余键一律丢弃 —— 这是
 * "原始数据行 / 完整源值 / 目标矩阵 / 宿主对象引用 / 完整路径不落盘"的
 * 结构性保证，而不是靠调用方自觉。
 *
 * 任一项不合法就返回 `null`（调用方据此明确反馈"无法恢复"），不做"尽力猜"：
 * 缺原需求时用摘要冒充完整任务，比直接说恢复不了更糟。
 */
export function sanitizeTaskResumeRecord(
  value: unknown
): TaskResumeRecord | null {
  if (!value || typeof value !== "object") return null;
  const raw = value as Record<string, unknown>;
  if (raw.schemaVersion !== TASK_RESUME_SCHEMA_VERSION) return null;
  const taskId = boundedText(raw.taskId, 200);
  if (!taskId) return null;
  if (!isTaskPhase(raw.phase)) return null;
  if (!isTaskBlocker(raw.blocker)) return null;
  if (!isTaskResumability(raw.resumability)) return null;
  if (!isTaskWriteState(raw.writeState)) return null;
  const originalPrompt = boundedText(raw.originalPrompt);
  if (!originalPrompt) return null;
  const authorizationRevision =
    typeof raw.authorizationRevision === "number" &&
    Number.isFinite(raw.authorizationRevision) &&
    raw.authorizationRevision >= 0
      ? Math.floor(raw.authorizationRevision)
      : 0;
  const now = new Date(0).toISOString();
  const createdAt = optionalTimestamp(raw.createdAt, now);
  const record: TaskResumeRecord = {
    schemaVersion: TASK_RESUME_SCHEMA_VERSION,
    taskId,
    phase: raw.phase,
    blocker: raw.blocker,
    resumability: raw.resumability,
    writeState: raw.writeState,
    originalPrompt,
    confirmations: sanitizeConfirmations(raw.confirmations),
    authorizationRevision,
    createdAt,
    updatedAt: optionalTimestamp(raw.updatedAt, createdAt)
  };
  const turnId = optionalText(raw.turnId);
  if (turnId) record.turnId = turnId;
  const operationId = optionalText(raw.operationId);
  if (operationId) record.operationId = operationId;
  const requestId = optionalText(raw.requestId);
  if (requestId) record.requestId = requestId;
  return record;
}

/** 类型守卫，便于调用方在 `unknown` 上早早收窄。 */
export function isTaskResumeRecord(value: unknown): value is TaskResumeRecord {
  return sanitizeTaskResumeRecord(value) !== null;
}

/** 稳定的任务身份生成器。同一任务的多次规划尝试只调用一次。 */
export function newTaskId(): string {
  const cryptoRef = (globalThis as { crypto?: Crypto }).crypto;
  const uuid = cryptoRef?.randomUUID?.();
  if (uuid) return `task-${uuid}`;
  return `task-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

/** 由记录复制出新的一版（保持 `taskId`，换阶段/阻塞/可续跑性）。 */
export function withTaskResumeState(
  record: TaskResumeRecord,
  patch: Partial<
    Pick<
      TaskResumeRecord,
      | "phase"
      | "blocker"
      | "resumability"
      | "writeState"
      | "authorizationRevision"
      | "turnId"
      | "operationId"
      | "requestId"
    >
  >
): TaskResumeRecord {
  return { ...record, ...patch, updatedAt: new Date().toISOString() };
}

/** 在已有记录上追加一条用户确认（去重、限量），供续答/补充参数使用。 */
export function withTaskResumeConfirmation(
  record: TaskResumeRecord,
  confirmation: TaskResumeConfirmation
): TaskResumeRecord {
  /*
   * 同一 `questionId` 的后答**替换**先答。
   *
   * 用户改了答案，旧答案就必须消失 —— 留着它，新一轮的提示词里会是
   * "按 A 处理"和"按 B 处理"两条互相冲突的确认，模型只能替用户挑一个，
   * 而那正是用户刚刚做过、不该被再代做一次的决定。
   */
  const questionId = confirmation.questionId?.trim();
  const base =
    questionId && record.confirmations.some(
      (item) => item.questionId?.trim() === questionId
    )
      ? record.confirmations.map((item) =>
          item.questionId?.trim() === questionId ? confirmation : item
        )
      : [...record.confirmations, confirmation];
  const merged = sanitizeConfirmations(base);
  return { ...record, confirmations: merged, updatedAt: new Date().toISOString() };
}
