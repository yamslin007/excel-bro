/**
 * 任务进度时间线的共享纯语义层（Excel / WPS 两端共用）。
 *
 * 这一层只回答四件事，全部是纯函数，**不碰宿主、不碰 React、不碰时间源**
 * （`now` 一律由调用方传入，测试因此可确定地推进）：
 *
 * 1. **事件怎么落到某一步上**（`applyTimelineEvent`）
 * 2. **整段活动结束时，那些没有收到结束事件的步骤怎么办**（`finishTimeline`）
 * 3. **暂停之后继续时，时间线怎么接回去**（`resumeTimeline`）
 * 4. **耗时怎么分账**（`timelineDurations` / `markTimelineBoundary`）
 *
 * ---------------------------------------------------------------------------
 * 为什么这四件事必须只有一份实现
 * ---------------------------------------------------------------------------
 * 两端各自"照着写一遍"的结果是三条已经发生过的真实缺陷：
 *
 * - WPS 把每一条更新**强制**写成 `progress`，于是"某一步已经完成"永远显示成
 *   "还在进行"，只能靠"下一步出现了"去猜上一步完成 —— 猜错的代价是：最后一步
 *   永远停在半途，而中间某一步可能被提前宣布完成。
 * - WPS 结束时只处理数组**最后一行**，而"当前步骤"是按 `stepId` 定位的，
 *   不保证在末尾（迟到事件、乱序补发都会让当前步骤不在末位）。
 * - Excel 的 `resumeActivityLog` 只修正**最后一行**的计时，其余行走的是
 *   `startedAt` 缺失的旧值，恢复后每一行的耗时会从 0 重新长。
 *
 * ---------------------------------------------------------------------------
 * 三条不变量（改动前先读）
 * ---------------------------------------------------------------------------
 * A. **状态一律采信事件，不许推导。** 事件说 `completed` 才是完成；"后面又来了
 *    一步"不能作为上一步完成的证据 —— 那是猜测，而猜测会写在用户看得见的界面上。
 * B. **终态回不去。** 已经 `completed` / `failed` / `cancelled` / `warning` 的步骤，
 *    不会被同一 `stepId` 的非终态事件"重新激活"（重放、迟到、断线补发都会这样）。
 * C. **身份分层。** `operationId` 是任务身份，`stepId` 是步骤身份。同一次工具调用
 *    的重放共用同一个 `stepId`（因此只更新不新增）；**新的一次调用必须由服务端
 *    生成新的 `stepId`**（见 `server/app/claude_sdk_agent.py` 与 `excel_agent.py`
 *    里的 per-invocation 后缀）—— 本层不会替服务端"猜"是不是同一次调用。
 */

export type TimelineStatus =
  | "started"
  | "progress"
  | "completed"
  | "warning"
  | "failed"
  | "cancelled";

export type TimelineOutcome =
  | "completed"
  | "awaiting_input"
  | "warning"
  | "failed"
  | "cancelled"
  | "interrupted";

/** 与"为什么停"无关的连接状态；只用于计时分账与文案。 */
export type TimelineConnectionState =
  | "connected"
  | "reconnecting"
  | "interrupted";

/** 计时的两个"不算执行"的桶。 */
export type TimelineExcludedKind = "pause" | "waiting";

const TERMINAL: ReadonlySet<string> = new Set([
  "completed",
  "warning",
  "failed",
  "cancelled"
]);

export function isTerminalStatus(status: string | undefined | null): boolean {
  return typeof status === "string" && TERMINAL.has(status);
}

function isActiveStatus(status: string | undefined | null): boolean {
  return !isTerminalStatus(status);
}

/**
 * 步骤的最小子集。
 *
 * 刻意**不用** `[key: string]: unknown` 索引签名：那样会让 `...step` 的展开在
 * 类型上丢掉两端丰富的字段（`facts` / `trace` / `pythonResult` …）。
 * 用泛型约束 + 展开，调用方的类型原样保留。
 */
export interface TimelineStepLike {
  id?: string;
  label: string;
  elapsedMs: number;
  status?: string;
  startedAt?: number;
}

/** 事件的最小子集；多余的字段由调用方在合并前自行取用。 */
export interface TimelineEventLike {
  operationId: string;
  stepId: string;
  sequence: number;
  status: string;
  title?: string | null;
}

export interface TimelineLogLike<S extends TimelineStepLike> {
  operationId?: string | null;
  sequence?: number;
  steps: S[];
  totalMs: number;
  /** Stable wall-clock origin, including pauses and disconnected intervals. */
  timelineStartedAt?: number;
  outcome?: string;
  connectionState?: string;
  /** 累计的"用户主动暂停"时间。 */
  pausedMs?: number;
  /** 累计的"断线等待"时间（任务在后台仍可能执行，因此不是暂停）。 */
  waitingMs?: number;
  /** 当前正在被计入排除桶的起点；`undefined` 表示不在任何排除桶里。 */
  excludedSince?: number;
  excludedKind?: TimelineExcludedKind;
  /**
   * 上一次"结束活动"顺手关掉的步骤 id。
   *
   * 它们没有收到真实结束事件（是被暂停/断线关掉的），所以恢复时应当重新打开：
   * 否则那一步的后续进度会被"已完成步骤不许被重新激活"挡住。
   */
  closedByFinish?: string[];
}

export type TimelineIgnoreReason =
  /** 事件属于另一条时间线（`operationId` 不同）。 */
  | "operation"
  /** 序号不比已归并的更大：重放 / 迟到 / 乱序。 */
  | "stale"
  /** 目标步骤已经终态，而这条事件想把它拉回进行中。 */
  | "terminal-replay";

export interface TimelineApplyResult<S extends TimelineStepLike> {
  log: TimelineLogLike<S>;
  /** 非 `null` 表示这条事件**没有**被接受；调用方据此写诊断，不要当成更新成功。 */
  ignored: TimelineIgnoreReason | null;
  /** 本次事件是否新建了一步（用于"新调用 vs 同次重放"的判别）。 */
  inserted: boolean;
}

/**
 * 该步骤此刻的耗时。
 *
 * 口径统一成"**单段**"：`startedAt` 是这一段（含恢复时带过来的累计）的起点，
 * 进行中的耗时一律 `now - startedAt` 现算；进入终态时把这一刻的值冻进 `elapsedMs`，
 * 之后不再增长。恢复时把 `startedAt` 往前推 `elapsedMs` 那么久，因此累计时间不会丢。
 */
function stepElapsed(step: TimelineStepLike, now: number): number {
  if (isTerminalStatus(step.status)) return Math.max(0, step.elapsedMs || 0);
  if (typeof step.startedAt === "number") return Math.max(0, now - step.startedAt);
  return Math.max(0, step.elapsedMs || 0);
}

/**
 * 把一条结构化进度事件归并进时间线。
 *
 * `patch` 是调用方给的"这一步除了状态之外还有什么"（标题、明细、目标、事实…）。
 * 本函数只负责**身份、终态、序号**这三件容易被写错的事，不解释语义。
 */
export function applyTimelineEvent<S extends TimelineStepLike>(
  log: TimelineLogLike<S>,
  event: TimelineEventLike,
  patch: (existing: S | undefined) => S,
  now: number
): TimelineApplyResult<S> {
  if (log.operationId && log.operationId !== event.operationId) {
    return { log, ignored: "operation", inserted: false };
  }
  if (typeof log.sequence === "number" && event.sequence <= log.sequence) {
    return { log, ignored: "stale", inserted: false };
  }
  const index = log.steps.findIndex((step) => step.id === event.stepId);
  const existing = index >= 0 ? log.steps[index] : undefined;

  /*
   * 终态保护：这条事件想把一个已经结束的步骤拉回"进行中"。
   *
   * 三种来源都真实发生过：断线重连后的补发、用户重复点击导致的请求重试、
   * 以及服务端把本轮的最后一条事件又发了一次。三种都不该让"已完成"变回"进行中"。
   * 注意这里**不是**按序号判断 —— 序号可能是新的，但语义仍然是旧的。
   */
  if (existing && isTerminalStatus(existing.status) && isActiveStatus(event.status)) {
    return { log, ignored: "terminal-replay", inserted: false };
  }

  const patched = patch(existing);
  // A later terminal receipt may enrich the detail, but cannot restart the clock
  // while terminal corrections retain the established elapsed time.
  const next = existing && isTerminalStatus(existing.status)
    ? { ...patched, elapsedMs: existing.elapsedMs, startedAt: undefined }
    : isTerminalStatus(patched.status)
      ? { ...patched, startedAt: undefined }
      : patched;
  const steps = log.steps.slice();
  if (index >= 0) steps[index] = next;
  else steps.push(next);

  return {
    ignored: null,
    inserted: index < 0,
    // 时间线内容归本函数；`totalMs` / `excluded*` 是耗时账，归调用方
    // （它才知道 `startedAt`），这里刻意不动。
    log: {
      ...log,
      operationId: event.operationId,
      sequence: event.sequence,
      steps
    }
  };
}

/**
 * 一段活动结束：**所有**没有收到结束事件的步骤都要有交代，不是只处理最后一行。
 *
 * - 当前步骤（`activeStepId` 指向的那一步）拿活动级终态；
 * - 其余没结束的步骤拿 `warning` + `note`，明确"没有收到结束确认，不代表已完成"；
 * - 已经终态的步骤一个都不动（不重置、不改写耗时）。
 */
export function finishTimeline<S extends TimelineStepLike>(
  log: TimelineLogLike<S>,
  outcome: TimelineOutcome,
  activeStepId: string | undefined,
  now: number,
  note = "此步骤未收到结束确认；活动已停止，不代表该步骤已完成。",
  /**
   * 当前步骤自己的终态。缺省按 `outcome` 推（WPS 就是这种口径）；
   * 但两者并不总是一回事：一次活动可以"带着待确认的结果收尾"（`outcome` 是
   * `awaiting_input`），而这一步本身**确实做完了**（`completed`）。混用会把
   * 真正完成的步骤写成"未收到结束确认"。
   */
  currentStatus?: TimelineStatus
): TimelineLogLike<S> {
  const terminalStatus: TimelineStatus =
    outcome === "failed"
      ? "failed"
      : outcome === "cancelled"
        ? "cancelled"
        : outcome === "warning" ||
            outcome === "interrupted" ||
            outcome === "awaiting_input"
          ? "warning"
          : "completed";

  const openSteps = log.steps.filter((step) => isActiveStatus(step.status));
  /*
   * 没显式指定"当前步骤"时，用**开始时间最晚**的那个未结束步骤，而不是数组末行。
   *
   * 数组顺序只是插入顺序：迟到事件、乱序补发、恢复后重新归并都会让"当前步骤"
   * 不在末尾。`startedAt` 是这条时间线自己写下来的证据，比位置可靠。
   */
  const latestOpen = openSteps.reduce<TimelineStepLike | undefined>((latest, step) => {
    if (!latest) return step;
    const a = typeof step.startedAt === "number" ? step.startedAt : -Infinity;
    const b = typeof latest.startedAt === "number" ? latest.startedAt : -Infinity;
    return a >= b ? step : latest;
  }, undefined);
  const resolvedActive = activeStepId ?? (latestOpen?.id as string | undefined);

  const closedByFinish: string[] = [];
  const steps = log.steps.map((step) => {
    if (isTerminalStatus(step.status)) return step;
    const isActive = resolvedActive !== undefined && step.id === resolvedActive;
    const status: TimelineStatus = isActive
      ? currentStatus ?? terminalStatus
      : terminalStatus === "failed" || terminalStatus === "cancelled"
        ? terminalStatus
        : "warning";
    /*
     * 记下"这些步骤是被**结束活动**关掉的"，不是收到了真实结束事件。
     *
     * 区别很要命：用户按暂停、断线、预算用完时，那一步其实没做完 —— 恢复后它还会
     * 继续发声。若把它当成真实终态，恢复后的第一条进度就会因为"已完成步骤不许被
     * 重新激活"被丢掉，界面停在一句"未收到结束确认"上再也不动。
     *
     * 只记 `warning`：它是"没有结论"的那一档。被收尾判成 `completed` 的步骤（调用方
     * 明确给了 `currentStatus`）是**这一步确实做完了**，被收尾判成 `failed` /
     * `cancelled` 的也是真实结论 —— 这三种一旦写进这里，恢复时就会被重新打开，
     * 等于让迟到的"进行中"事件把已经结束的步骤拉回来。
     */
    if (status === "warning") closedByFinish.push(step.id as string);
    return {
      ...step,
      status,
      elapsedMs: stepElapsed(step, now),
      startedAt: undefined,
      ...(isActive ? {} : { note })
    } as S;
  });

  return { ...log, steps, outcome, closedByFinish };
}

/** `finishTimeline` 的耗时口径由调用方给（有些调用方只有墙钟，没有起点）。 */
export function finishTimelineAt<S extends TimelineStepLike>(
  log: TimelineLogLike<S>,
  outcome: TimelineOutcome,
  activeStepId: string | undefined,
  now: number,
  totalMs: number,
  note?: string,
  currentStatus?: TimelineStatus
): TimelineLogLike<S> {
  const finished = finishTimeline(log, outcome, activeStepId, now, note, currentStatus);
  const timed = { ...markTimelineBoundary(finished, null, now),
    timelineStartedAt: log.timelineStartedAt ?? now - Math.max(0, totalMs),
    totalMs: Math.max(0, totalMs) };
  // No server task yet: local analysis is stopped. Once a server task exists,
  // only reception is known to have stopped; account for that as waiting.
  return finished.closedByFinish?.length &&
    (outcome === "awaiting_input" || outcome === "interrupted" || outcome === "warning")
    ? markTimelineBoundary(timed, log.operationId ? "waiting" : "pause", now)
    : timed;
}

/**
 * 暂停之后继续：把时间线**接回去**，不新增业务步骤、不重编号。
 *
 * 关键点：
 *
 * 1. **每一行**的计时都要按自己已经累计的时间重新定点（`startedAt = now - 已累计`），
 *    而不是只修最后一行 —— 只修最后一行的结果就是其余行在恢复后从 0 重新计时。
 * 2. **被"结束活动"关掉的步骤要重新打开**（`closedByFinish`）。它们不是做完了，
 *    是被暂停/断线打断的；不打开的话，恢复后的第一条进度会被"已完成步骤不许被
 *    重新激活"丢掉，界面永远停在"未收到结束确认"。
 * 3. 真正收到过结束事件的步骤（completed/failed/cancelled/warning）原样保留，
 *    它们的时间不再长 —— 这是"已完成步骤不得被重新激活"的正面情形。
 */
export function resumeTimeline<S extends TimelineStepLike>(
  log: TimelineLogLike<S>,
  now: number
): TimelineLogLike<S> {
  const interrupted = new Set(log.closedByFinish ?? []);
  /*
   * `closedByFinish` **缺省**（`undefined`）才是历史日志：那时没有这条记录，
   * 只能按 `warning` 猜。写成 `?? []` 会让"新日志里明确为空"和"老日志没这字段"
   * 变成同一件事 —— 于是自己报了 warning 结束的步骤会被重新打开。
   */
  const legacyLog = log.closedByFinish === undefined;
  return {
    ...markTimelineBoundary(log, null, now),
    totalMs: log.timelineStartedAt !== undefined ? Math.max(0, now - log.timelineStartedAt)
      : log.totalMs + (log.excludedSince === undefined ? 0 : Math.max(0, now - log.excludedSince)),
    outcome: undefined,
    closedByFinish: [],
    connectionState: "connected",
    excludedSince: undefined,
    excludedKind: undefined,
    steps: log.steps.map((step) => {
      /*
       * 重新打开的条件只有一条，含义是："这一步**没有收到真实的结束事件**"。
       *
       *  - 新日志：完全以 `closedByFinish` 为准。它是被结束活动关掉的才重开。
       *  - 历史日志（没有这个字段）：只能按 `warning` 判 —— warning 本身就是
       *    "没收到结束确认"的结论。
       *
       * `completed` / `failed` / `cancelled` 是真终态，一律原样保留 —— 这是
       * "已结束的步骤不得被迟到的进行中事件重新激活"的正面情形。
       */
      const wasInterrupted =
        interrupted.has(step.id as string) || (legacyLog && step.status === "warning");
      if (isTerminalStatus(step.status) && !wasInterrupted) return step;
      const carried = Math.max(0, step.elapsedMs || 0);
      return {
        ...step,
        status: step.status === "started" ? "started" : "progress",
        elapsedMs: carried,
        startedAt: now - carried,
        // "未收到结束确认"这句话只对**这一次**暂停成立；重新接上就该消失。
        note: undefined
      } as S;
    })
  };
}

/**
 * 排除桶的边界（暂停 / 断线等待）。
 *
 * `kind === null` 表示"回到正常执行"。切换时把**已经过去的那一段**记进对应的桶：
 * - 用户主动暂停 → `pausedMs`（这段时间任务不推进）
 * - 断线等待 → `waitingMs`（任务在后台**仍可能继续执行**，所以它不是暂停）
 *
 * 这就是"后台仍执行时不能把它误算为真正暂停"的落点：两种时间分开记，
 * 谁也不能冒充谁。
 */
export function markTimelineBoundary<S extends TimelineStepLike>(
  log: TimelineLogLike<S>,
  kind: TimelineExcludedKind | null,
  now: number
): TimelineLogLike<S> {
  const openKind = log.excludedKind;
  const openSince = log.excludedSince;
  /*
   * 同一个桶重复标记：**原样返回**，既不再累加、也不重开区间。
   *
   * 先累加再判断会把它算两遍：第二次标记时把"已经过去的一段"记进桶，
   * 却仍保留原来的起点，于是 `timelineDurations` 又按起点算一遍 —— 暂停时间翻倍。
   */
  if (kind !== null && openKind === kind && typeof openSince === "number") {
    return log;
  }
  if (kind === null && !openKind) return log;
  let pausedMs = Math.max(0, log.pausedMs ?? 0);
  let waitingMs = Math.max(0, log.waitingMs ?? 0);
  if (openKind && typeof openSince === "number") {
    const elapsed = Math.max(0, now - openSince);
    if (openKind === "pause") pausedMs += elapsed;
    else waitingMs += elapsed;
  }
  const steps = log.steps.map((step) => {
    if (isTerminalStatus(step.status)) return step;
    const elapsedMs = openKind ? step.elapsedMs : stepElapsed(step, now);
    return { ...step, elapsedMs, startedAt: kind === null ? now - elapsedMs : undefined };
  });
  if (kind === null) {
    return { ...log, steps, pausedMs, waitingMs, excludedKind: undefined, excludedSince: undefined };
  }
  return { ...log, steps, pausedMs, waitingMs, excludedKind: kind, excludedSince: now };
}

export interface TimelineDurations {
  /** 从开始到现在的墙钟时间。 */
  totalMs: number;
  /** 前台接收进度的耗时（= total − 暂停 − 等待），不是服务端计算耗时。 */
  activeMs: number;
  pausedMs: number;
  waitingMs: number;
  /** 当前是否正处在某个排除桶里，以及是哪一个。 */
  excluded: TimelineExcludedKind | null;
}

/** 四个数一次算清楚，避免各处各算一套导致标题与正文对不上。 */
export function timelineDurations<S extends TimelineStepLike>(
  log: TimelineLogLike<S>,
  startedAt: number,
  now: number
): TimelineDurations {
  const totalMs = Math.max(0, now - startedAt);
  let pausedMs = Math.max(0, log.pausedMs ?? 0);
  let waitingMs = Math.max(0, log.waitingMs ?? 0);
  const open = log.excludedKind;
  const openSince = log.excludedSince;
  if (open && typeof openSince === "number") {
    const carried = Math.max(0, now - openSince);
    if (open === "pause") pausedMs += carried;
    else waitingMs += carried;
  }
  return {
    totalMs,
    activeMs: Math.max(0, totalMs - pausedMs - waitingMs),
    pausedMs,
    waitingMs,
    excluded: open ?? null
  };
}
