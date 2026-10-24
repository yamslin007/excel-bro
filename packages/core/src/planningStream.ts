/**
 * 规划流式消费与检查点续跑的**共享口径**（Excel / WPS 两端共用，L2 纯语义层）。
 *
 * 设计边界（与项目既有纪律一致，不得在此放宽）：
 *
 * 1. **这不是第二套任务账本。** 权威现场在服务端 `planning_tasks`（`turnId` 身份）。
 *    本模块只做三件事：把 SSE 帧归并成一份可展示状态、判定客户端事件游标、
 *    以及从服务端状态推导"下一步能做什么"。它**不复制检查点内容**。
 *
 * 2. **身份必须沿用，不能新造。** 继续一次暂停的规划，必须带**原 `turnId`** +
 *    服务端返回的 `expectedRevision` + 稳定 `requestId` + 客户端 `afterSequence`。
 *    任何一项缺失或自造，都会让服务端产生**重复规划**或**重复工具调用** ——
 *    服务端侧靠 `turnId` 去重、`revision` 做 CAS、`requestId` 做幂等领取。
 *
 * 3. **事件游标只能前进到服务端已确认的位置。** `clientLastSequence` 与服务端的
 *    `lastSequence` 取较小值；盲目相信本地更大的游标会在继续时**跳过**尚未收到的
 *    事件（服务端只补发 `> afterSequence` 的），造成步骤静默缺失。
 *
 * 4. **取消与断线是两种状态。** 用户主动取消后，旧任务的迟到事件**不得**改变
 *    新任务的状态（由 `applyPlanningEvent` 的 `turnId` 闸保证）。
 *
 * 5. **不得用助手文本或业务关键词决定卡片与按钮。** 按钮只由服务端的结构化
 *    状态（`status` / `canContinue` / `pauseType` / `error.code`）推导。
 */

/**
 * 规划任务在**客户端视角**下的归并状态。
 *
 * 与 `TaskPhase` / `TaskBlocker`（`taskResume.ts`）刻意分开：那一套描述"用户看到
 * 的那张中断卡"，这一套描述"规划流本身跑到哪了"。混用会让"读取数据"和
 * "等用户回答"被压成同一个值。
 */
export type PlanningStreamPhase =
  | "idle"
  | "streaming"
  | "paused"
  | "completed"
  | "failed"
  | "cancelled";

/** 服务端 `GET /api/turn/{turnId}/status` 的 `status` 取值。 */
export type PlanningTaskStatus =
  | "running"
  | "waiting_user"
  | "interrupted"
  | "completed"
  | "failed"
  | "cancelled";

/**
 * 暂停成因，来自服务端 `pauseType`。
 *
 * `budget_exhausted`（2026-10-10 新增）：累计活跃执行时间已达总时限，不再启动
 * 新的模型或工具调用；它与 `budget`（本轮预算暂停、可再续）的恢复能力不同，
 * **不能**合并成一个值 —— 否则界面会把"总时长用完了"说成"按一下就能继续"。
 */
export type PlanningPauseType =
  | "budget"
  | "budget_exhausted"
  | "model_error"
  | "interrupted";

/** 服务端 `error.code` 里表示"模型连接类故障、可换模型接续"的前缀。 */
export const MODEL_ERROR_CODE_PREFIX = "MODEL_";

/**
 * 规划流在工作簿侧**永远不应该写单元格**。
 *
 * 这条断言存在的意义：恢复功能一旦接错线，最危险的失败模式是"续跑顺手把
 * 计划也执行了"。调用方在恢复路径上应显式声明写入次数，并断言为 0。
 */
export const PLANNING_RESUME_HOST_WRITES = 0 as const;

/** 规划事件（`step`）里与父/子步骤相关的字段，做保守白名单解析。 */
export interface PlanningStepEventView {
  sequence: number;
  stepId: string;
  phase: string;
  status: string;
  title: string;
  summary?: string;
}

/** 客户端持有的规划任务游标与身份。**不落盘单元格内容**。 */
export interface PlanningCursor {
  turnId: string;
  /** 服务端账本 revision，继续时必须原样带回做 CAS。 */
  revision: number;
  /** 服务端已提交的最大事件序号。 */
  serverLastSequence: number;
  /** 客户端已消费到的最大事件序号。 */
  clientLastSequence: number;
}

/** 事件归并结果：状态 + 是否应当采纳这条事件。 */
export interface PlanningEventApplication<TEvent extends PlanningStepEventView> {
  accepted: boolean;
  /** 拒绝原因，便于测试断言与用户可见说明。 */
  rejectedReason?: "turn_mismatch" | "stale_sequence" | "terminal";
  cursor: PlanningCursor;
  /** 已接受的、按 sequence 去重排序后的事件列表（新数组，不原地改）。 */
  events: TEvent[];
}

/**
 * 新建一个初始游标。
 *
 * `turnId` 由调用方从服务端响应或请求里取，**不允许在本函数里凭空生成** ——
 * 自造的 turnId 会让服务端认为这是一次新任务，从而重复规划。
 */
export function newPlanningCursor(turnId: string): PlanningCursor {
  const trimmed = turnId.trim();
  if (!trimmed) throw new Error("规划任务身份不能为空。");
  return {
    turnId: trimmed,
    revision: 0,
    serverLastSequence: 0,
    clientLastSequence: 0
  };
}

/** 游标可安全提交给服务端的 `afterSequence`：永不超前于服务端确认位置。 */
export function planningAfterSequence(cursor: PlanningCursor): number {
  return Math.min(
    Math.max(0, Math.floor(cursor.clientLastSequence)),
    Math.max(0, Math.floor(cursor.serverLastSequence))
  );
}

/**
 * 归并一条规划事件。
 *
 * 三条闸，顺序不可调换：
 *  1. **身份闸**：事件 `operationId` 必须等于当前 `turnId`。不等即拒 —— 这是
 *     "取消后旧事件不污染新任务"的唯一守门点。
 *  2. **终态闸**：已到终态（completed / failed / cancelled）后不再采纳展示事件。
 *  3. **去重闸**：`sequence` 必须严格大于已消费的最大值。重放/重复事件在此被丢弃。
 */
export function applyPlanningEvent<TEvent extends PlanningStepEventView>(
  cursor: PlanningCursor,
  event: TEvent,
  turnIdOfEvent: string
): PlanningEventApplication<TEvent> {
  if (turnIdOfEvent.trim() !== cursor.turnId) {
    return { accepted: false, rejectedReason: "turn_mismatch", cursor, events: [] };
  }
  const sequence = Number(event.sequence);
  if (!Number.isFinite(sequence) || sequence <= cursor.clientLastSequence) {
    return { accepted: false, rejectedReason: "stale_sequence", cursor, events: [] };
  }
  const next: PlanningCursor = {
    ...cursor,
    clientLastSequence: sequence,
    serverLastSequence: Math.max(cursor.serverLastSequence, sequence)
  };
  return { accepted: true, cursor: next, events: [event] };
}

/**
 * 用服务端返回的真实序号校准游标。
 *
 * 服务端是权威：`revision` 与 `lastSequence` 一律以它为准。但 `clientLastSequence`
 * 只能**取小**（`min`），不能因为服务端更大就跳过去 —— 中间那些事件客户端没收过，
 * 靠继续时的 `afterSequence` 补发，而不是假装已经看到。
 */
export function reconcilePlanningCursor(
  cursor: PlanningCursor,
  server: { revision: number; lastSequence: number }
): PlanningCursor {
  const revision = Number.isFinite(server.revision)
    ? Math.max(1, Math.floor(server.revision))
    : cursor.revision;
  const lastSequence = Number.isFinite(server.lastSequence)
    ? Math.max(0, Math.floor(server.lastSequence))
    : cursor.serverLastSequence;
  return {
    ...cursor,
    revision,
    serverLastSequence: lastSequence,
    clientLastSequence: Math.min(cursor.clientLastSequence, lastSequence)
  };
}

/** 从服务端状态推导出的"能做什么"。按钮文案只由这里给，不由助手文本决定。 */
export interface PlanningResumeDecision {
  /** 主按钮动作。`none` 表示不画按钮。 */
  action: "continue_checkpoint" | "regenerate" | "none";
  actionLabel: string;
  detail: string;
  canContinue: boolean;
  /** 真实降级原因（服务端给的原因或本模块归因），直接展示给用户。 */
  reason: string;
}

/** 服务端给出的「继续资格」事实（`planning_context` 契约 + 执行账本算出）。 */
export interface PlanningContinuationView {
  available: boolean;
  /** 固定拒绝码；`available` 时为 null。服务端闸门码不属于检查点契约，按字符串处理。 */
  code: string | null;
  message: string;
  stage: "before_model" | "after_tool" | null;
  /**
   * 原任务是否依赖来源数据 ⇒ 继续前必须重读并比对来源指纹。
   *
   * 由服务端从检查点判定，客户端**不自己猜**：猜"需要"会把纯结构任务挡在门外，
   * 猜"不需要"会放过数据已经变过的现场。缺省按"需要"处理（fail-closed）。
   */
  sourceVerificationRequired?: boolean;
}

/** 服务端状态的保守视图：只读结构化字段，不读助手文本。 */
export interface PlanningStatusView {
  status: PlanningTaskStatus;
  canContinue?: boolean;
  pauseType?: PlanningPauseType | null;
  reason?: string | null;
  errorCode?: string | null;
  errorMessage?: string | null;
  hasCheckpoint?: boolean;
  /**
   * 能不能继续的**权威事实**。
   *
   * `hasCheckpoint` 只是"暂停过"的同义词（旧实现直接拿 `pauseType != null` 当它），
   * 不能单独决定按钮 —— 一份损坏的、别的任务的、或已经写过原表的现场同样"有暂停"。
   */
  continuation?: PlanningContinuationView | null;
}

/**
 * 由服务端状态推导继续方案。
 *
 * 判据顺序：终态 → 运行中 → 无检查点 → 可继续。
 *
 * **只能重新规划时必须说"重新读取并生成预览"**，不得显示"继续" —— 否则用户会
 * 以为可以从上次进度精确续跑，而实际上重做了一遍读取。
 */
export function planningResumeDecision(
  view: PlanningStatusView
): PlanningResumeDecision {
  if (view.status === "completed") {
    return {
      action: "none",
      actionLabel: "",
      detail: "任务已经完成，结果可直接使用。",
      canContinue: false,
      reason: view.reason ?? "任务已经完成。"
    };
  }
  if (view.status === "failed" || view.status === "cancelled") {
    const modelFailure =
      typeof view.errorCode === "string" &&
      view.errorCode.startsWith(MODEL_ERROR_CODE_PREFIX);
    // 模型连接类故障且服务端明确标了可继续 → 允许接续；其余终态一律只能重新规划。
    const recoverable =
      modelFailure && view.canContinue === true && view.hasCheckpoint === true;
    return {
      action: recoverable ? "continue_checkpoint" : "none",
      actionLabel: recoverable ? "继续分析" : "",
      detail: recoverable
        ? "将按当前授权重新读取并分析；已确认的口径会保留，写入前仍需预览与你的确认。"
        : "该任务已经停止，没有可精确续跑的现场。请重新发送需求。",
      canContinue: recoverable,
      reason:
        view.errorMessage ?? view.reason ?? "该规划任务已经停止，不能继续。"
    };
  }
  if (view.status === "running") {
    return {
      action: "none",
      actionLabel: "",
      detail: "任务仍在后台处理中，无需重复继续。",
      canContinue: false,
      reason: view.reason ?? "任务仍在处理中，请稍后再查看状态。"
    };
  }
  // waiting_user / interrupted
  /*
   * 总预算用完是**另一种事实**：现场本身可能是好的，但不该再承诺"按一下就能
   * 接着跑"。此时服务端把 `canContinue` 置假而 `continuation.available` 仍为真，
   * 若按"没有现场"处理会被读成"重新生成预览"，等于绕开总时限。
   */
  if (view.pauseType === "budget_exhausted") {
    return {
      action: "none",
      actionLabel: "",
      detail:
        "本轮有效运行总时长已用完，不再启动新的模型或工具调用。已完成的步骤与现场已保存，请重新发起任务。",
      canContinue: false,
      reason: view.reason ?? "本轮有效运行总时长已用完。"
    };
  }
  /*
   * 服务端明确说不能继续时，**原因照抄**，不自己拼一句"没有可验证的现场"。
   * 它可能知道本端看不到的事实（例如这个任务其实已经登记过写入），
   * 换成通用措辞会把用户导向"重发一遍"，而不是去核对原表。
   */
  const blocked =
    view.continuation && view.continuation.available === false
      ? view.continuation
      : null;
  if (blocked) {
    const message = blocked.message.trim() || "该任务当前不能继续分析。";
    return {
      action: "none",
      actionLabel: "",
      detail: message,
      canContinue: false,
      reason: message
    };
  }
  if (view.canContinue !== true || view.hasCheckpoint !== true) {
    return {
      action: "regenerate",
      actionLabel: "重新读取并生成预览",
      detail:
        "该任务没有可验证的完整规划现场，不能精确续跑。将保留原需求重新读取并重新生成预览。",
      canContinue: true,
      reason: view.reason ?? "没有可验证的完整规划现场。"
    };
  }
  return {
    action: "continue_checkpoint",
    actionLabel: "继续分析",
    /*
     * 不许写"不会重做已完成的步骤"：检查点保存的是引用与已确认口径，服务端会按
     * 当前授权重新读取、重新计算，引用失效就重读 —— 这正是"不沿用过期结果"。
     */
    detail: "将按当前授权重新读取并分析；已确认的口径会保留，写入前仍需预览与你的确认。",
    canContinue: true,
    reason: view.reason ?? "本轮处理暂停，当前进度已安全保存。"
  };
}

/**
 * 两个"看起来像同一件事"的请求是否是同一次继续。
 *
 * 用于点击去重：同一个 `turnId` + `requestId` 的并发点击只能产生一次服务端领取。
 * `requestId` 由调用方在**首次点击时生成并固定**，不得每次点击重新生成 ——
 * 否则服务端看到的是多个不同请求，幂等领取形同虚设。
 */
export function isSamePlanningResume(
  a: { turnId: string; requestId: string },
  b: { turnId: string; requestId: string }
): boolean {
  return a.turnId === b.turnId && a.requestId === b.requestId;
}

/**
 * 新需求必须拿到**新的** `turnId`，不能被旧任务的恢复逻辑接管。
 *
 * 判据：如果传入的 `candidateTurnId` 属于一个**已知的暂停/失败任务**，
 * 则本次请求是新需求，必须换新 turnId。
 */
export function resolveTurnIdForNewRequest(
  candidateTurnId: string | null | undefined,
  knownResumableTurnIds: readonly string[]
): string | null {
  const candidate = candidateTurnId?.trim() ?? "";
  if (!candidate) return null;
  const isKnownResumable = knownResumableTurnIds.some(
    (turnId) => turnId.trim() === candidate
  );
  if (isKnownResumable) return null;
  return candidate;
}

/** 生成一个稳定的继续请求身份。调用方在首次点击时调用一次并缓存。 */
export function newPlanningRequestId(): string {
  const cryptoRef = (globalThis as { crypto?: Crypto }).crypto;
  const uuid = cryptoRef?.randomUUID?.();
  if (uuid) return `plan-continue-${uuid}`;
  return `plan-continue-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

/** 父步骤在终态到达时必须收尾：不得遗留"正在执行"的父步骤。 */
export interface PlanningStepLifecycle {
  stepId: string;
  status: "progress" | "completed" | "failed";
}

/**
 * 终态到达后，把仍处于 `progress` 的步骤归为终态。
 *
 * 没有这一步，界面会在任务完成后仍显示一行"A 正在执行"，与结果卡自相矛盾。
 * `failed` 只给终态本身是失败时；成功收尾一律 `completed`。
 */
export function closeDanglingSteps<TStep extends PlanningStepLifecycle>(
  steps: readonly TStep[],
  terminal: "completed" | "failed" | "cancelled"
): TStep[] {
  const closingStatus = terminal === "failed" ? "failed" : "completed";
  return steps.map((step) =>
    step.status === "progress" ? { ...step, status: closingStatus } : step
  );
}

/** 面板重开后，从历史里筛出的"值得回服务端核一遍"的暂停条目。 */
export interface PlanningPausedEntry<TPayload> {
  /** 服务端账本主键。用于 `GET /api/turn/{turnId}/status`。 */
  turnId: string;
  /** 落盘的那份暂停快照（含 `pauseType` / `reason` / `canContinue`）。 */
  payload: TPayload;
  /** 这条暂停所在的消息 id，用于把结果挂回原卡片。 */
  messageId: string;
}

/**
 * 面板重开后，从聊天历史里挑出需要重新核对的**规划暂停**。
 *
 * 为什么必须在重载时重新核对：历史里那条暂停卡是"当时那一刻"的快照。
 * 重开期间服务端可能已经把它继续完了、被别的窗口接手了、或者超时失败了。
 * 直接拿旧快照画"继续任务"按钮，用户会点到一个基于过期 `revision` 的动作
 * （服务端 CAS 会拒，界面上看起来却像"按钮坏了"）。
 *
 * 三条纪律：
 *
 *  1. **只挑暂停**，不挑答案 / 计划 / 澄清 —— 只有暂停是"等用户决定"的中间态，
 *     其余几类各自的恢复通路（`taskResume`）已经覆盖，重复处理会抢走它们的卡。
 *  2. **`turnId` 去重取最后一条**。同一任务可能被多次暂停（续跑又暂停），
 *     应当以**最新**的那次为准；重复渲染两张卡会让用户面对两个"继续"。
 *  3. **只读历史，不查服务端**。本函数是纯函数：把"要查谁"筛出来交给调用方，
 *     网络与宿主副作用一律留在调用方 —— 这样它能被确定性测试钉住。
 */
export function planningPausedEntriesFromHistory<
  TMessage extends {
    id: string;
    planningPause?: { turnId: string } | undefined;
  },
  TPayload
>(
  messages: readonly TMessage[],
  payloadOf: (turnId: string) => TPayload | null
): PlanningPausedEntry<TPayload>[] {
  const byTurn = new Map<string, PlanningPausedEntry<TPayload>>();
  for (const message of messages) {
    const turnId = message.planningPause?.turnId?.trim();
    if (!turnId) continue;
    const payload = payloadOf(turnId);
    if (!payload) continue;
    // 后写覆盖先写 = 取历史里最后出现的那次暂停（同一任务被多次暂停时）。
    byTurn.set(turnId, { turnId, payload, messageId: message.id });
  }
  return [...byTurn.values()];
}
