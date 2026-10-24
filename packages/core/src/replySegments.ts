/**
 * 「回复段」—— 用户看到的展示身份（Excel / WPS 共用，L2 纯语义层）。
 *
 * ---------------------------------------------------------------------------
 * 为什么要多一个身份
 * ---------------------------------------------------------------------------
 * 执行身份只有一套，且**不许**为了界面重造：
 *   - 规划：服务端 `planning_tasks.turn_id`
 *   - 写入：`operationId` / `executionId`
 *   - 单次继续动作：`requestId`
 *
 * 旧口径把这些身份和"那张过程卡"绑死（一张任务只能有一张卡，继续时原地复用
 * 同一条消息）。后果是真实发生过的：用户暂停、再说"接着完成"，界面上的进度仍然
 * 更新在上方那条旧消息里，最新消息下面什么都没有 —— 用户看不到自己刚发的那句话
 * 换来了什么。
 *
 * 本模块只定义**展示段**：
 *   - 同一任务可以有多个回复段（多轮"暂停 → 继续"）；
 *   - 同一时刻只有一个 `live` 段（当前活动的展示位置）；
 *   - 段之间是 `previousSegmentId` → `nextSegmentId` 的单链，用于轻量定位；
 *   - 段身份**不参与**任何执行判定：步骤身份、事件序号、结果引用、写入回执
 *     仍然只认上面那套执行身份。
 *
 * 不变量（改动前先读）：
 *   A. `openReplySegment` 幂等：同一个 id 重复打开不会产生两段。
 *   B. 任何时刻 `live` 段数量 ≤ 1；打开新段会把旧段的 `live` 摘掉并接到链上。
 *   C. 段状态只会 `live → superseded|closed`，不会反向。历史段的状态是**当时**的
 *      事实，迟到的进度不能把它改写成"还在跑"。
 */

import {
  finishTimelineAt,
  type TimelineLogLike,
  type TimelineOutcome,
  type TimelineStepLike
} from "./activityTimeline";

export type ReplySegmentStatus = "live" | "superseded" | "closed";

export interface ReplySegment {
  /** 展示段身份（本端生成，不替代任何执行身份）。 */
  id: string;
  /**
   * 这一段在展示**哪个执行任务**（规划 turnId / 写入 operationId）。
   * 只用于把段和任务对上账，不用于判定执行归属。
   */
  taskId?: string | null;
  status: ReplySegmentStatus;
  /** 上一段（同一条链）。 */
  previousSegmentId?: string;
  /** 下一段（同一条任务链）；历史段可以在稍后继续时补上此定位链接。 */
  nextSegmentId?: string;
  /**
   * 是不是"接着原任务继续"的段。
   *
   * 续跑段**不能**把自己当成一个新任务从"步骤 1"讲起：它没有重做前面那些步骤，
   * 也没有权利改写它们。界面据此只展示阶段名、不重排编号。
   */
  continuesTask?: boolean;
}

export interface OpenReplySegmentOptions {
  id: string;
  taskId?: string | null;
  continuesTask?: boolean;
  /**
   * 显式指定上一段。
   *
   * 不传时按"当前活动段"接（同一任务继续的常态）。**不能**退化成"列表里最后
   * 一个元素"：多任务交错时那会把新段接到另一个任务的段后面，于是 A 的后继被
   * 反复改写成最新段，中间段永远没有"查看后续进展"。
   */
  previousSegmentId?: string | null;
}

/**
 * 打开新的回复段。
 *
 * 幂等：`id` 已存在时原样返回（同一个 id 被重复打开只可能是一次重放）。
 * 旧的 `live` 段一律降为 `superseded` 并接上 `nextSegmentId`。
 */
export function openReplySegment(
  segments: readonly ReplySegment[],
  options: OpenReplySegmentOptions
): ReplySegment[] {
  if (segments.some((segment) => segment.id === options.id)) {
    return segments.slice();
  }
  const parent =
    options.previousSegmentId === undefined
      ? activeReplySegment(segments) ?? segments[segments.length - 1]
      : options.previousSegmentId
        ? segments.find((segment) => segment.id === options.previousSegmentId)
        : undefined;
  return [
    ...segments.map((segment) => {
      const isParent = segment.id === parent?.id;
      // 接链与摘掉活动身份是两件事。跨任务继续时，父段已经被冻结，仍须补上
      // 后继入口；它的历史状态、日志与耗时不重开。别的任务只关闭展示身份。
      if (!isParent && segment.status !== "live") return segment;
      return {
        ...segment,
        ...(segment.status === "live"
          ? { status: "superseded" as const }
          : {}),
        ...(isParent ? { nextSegmentId: options.id } : {})
      };
    }),
    {
      id: options.id,
      status: "live" as const,
      ...(options.taskId ? { taskId: options.taskId } : {}),
      ...(options.continuesTask ? { continuesTask: true } : {}),
      ...(parent ? { previousSegmentId: parent.id } : {})
    }
  ];
}

/**
 * 该任务**最新的**展示段：新段接在它后面，而不是接在这个任务最早的段后面。
 *
 * 旧实现用 `messages.find(...)` 从前往后找，于是第二次继续仍然改第一段：
 * A 的 `nextSegmentId` 被反复覆盖，B 永远没有后继入口（A→C、A→D，B 孤立）。
 */
export function latestReplySegmentForTask(
  segments: readonly ReplySegment[],
  taskId: string
): ReplySegment | undefined {
  for (let index = segments.length - 1; index >= 0; index -= 1) {
    if (segments[index].taskId === taskId) return segments[index];
  }
  return undefined;
}

export interface SegmentContinuationPlan {
  /** 新段的上一段（该任务最新的段）；`null` 表示这个任务还没有历史段。 */
  previousSegmentId: string | null;
  /** 需要降级为 `superseded` 的段（当前仍 `live` 的那个）；`null` 表示没有。 */
  supersededSegmentId: string | null;
}

/**
 * 「继续」时该接在哪一段后面。
 *
 * 两端都必须用它来决定父子关系：早期分析恢复与规划续跑走**同一个**判据，
 * 否则一处按任务找段、一处按活动消息找段，链就会断。
 */
export function planSegmentContinuation(
  segments: readonly ReplySegment[],
  taskId: string | null | undefined
): SegmentContinuationPlan {
  const previous = taskId ? latestReplySegmentForTask(segments, taskId) : undefined;
  const live = activeReplySegment(segments);
  return {
    previousSegmentId: previous?.id ?? null,
    supersededSegmentId: live?.id ?? null
  };
}

/** 当前活动的展示段：`live` 段只应有最后一个，取最后一个仍然正确（旧日志可能残留多个）。 */
export function activeReplySegment(
  segments: readonly ReplySegment[]
): ReplySegment | undefined {
  for (let index = segments.length - 1; index >= 0; index -= 1) {
    if (segments[index].status === "live") return segments[index];
  }
  return undefined;
}

/**
 * 段是否能接收这条事件的展示。
 *
 * 判据只有"是不是当前活动段"，**不是**"任务身份是否相同"：同一任务的历史段
 * 不许抢回活动位置（这正是旧实现里"继续后仍然更新上方旧消息"的根因）。
 *
 * 兼容：老消息没有段（`null`/`undefined`）时按"可以接收"处理 —— 那段历史没有
 * 段身份，不能因为缺字段就把它判成冻结。
 */
export function segmentAcceptsProgress(
  segments: readonly ReplySegment[],
  segmentId: string | null | undefined,
  operationId: string | null | undefined
): boolean {
  if (!segmentId) return true;
  const segment = segments.find((item) => item.id === segmentId);
  if (!segment) return true;
  if (segment.status !== "live") return false;
  if (!operationId) return true;
  return !segment.taskId || segment.taskId === operationId;
}

/** 同一任务的所有段（按打开顺序），用于"查看后续进展"之类的定位。 */
export function replySegmentsForTask(
  segments: readonly ReplySegment[],
  taskId: string
): ReplySegment[] {
  return segments.filter((segment) => segment.taskId === taskId);
}

/** 把一段显式收尾（没有下一段时的终态）。已 `superseded` 的段不再改写。 */
export function closeReplySegment(
  segments: readonly ReplySegment[],
  id: string
): ReplySegment[] {
  return segments.map((segment) =>
    segment.id === id && segment.status === "live"
      ? { ...segment, status: "closed" as const }
      : segment
  );
}

/**
 * 冻结一个历史段的时间线。
 *
 * 只做一件事：**所有没有收到结束事件的步骤都要有交代**，并且耗时停在冻结这一刻。
 * `completed` / `failed` / `cancelled` 原样保留 —— 它们是真结论，不能被"这一段
 * 已经过去了"改写成 `warning`。
 *
 * 这不是"把结果丢掉"：权威任务记录（服务端账本 + 最新段）仍然保留全部事件；
 * 冻结只影响**这一段**的展示。
 */
export function freezeSegmentLog<S extends TimelineStepLike>(
  log: TimelineLogLike<S>,
  now: number,
  options: {
    outcome?: TimelineOutcome;
    note?: string;
    activeStepId?: string;
  } = {}
): TimelineLogLike<S> {
  const totalMs = Math.max(0, now - (log.timelineStartedAt ?? now));
  return finishTimelineAt(
    log,
    options.outcome ?? "awaiting_input",
    options.activeStepId,
    now,
    totalMs,
    options.note ??
      "未收到结束确认；此回复段已停止接收进度，后续进展在新的回复中。",
    undefined
  );
}

/**
 * 续跑段里展示"接着第几步"的文案。
 *
 * 只在**有真实前序步骤**时才说"继续"；没有前序记录时如实说"此前未记录步骤"，
 * 不编一个编号出来。
 */
export function continuationLabel(previousStepCount: number): string {
  return previousStepCount > 0
    ? `继续（此前已完成 ${previousStepCount} 步）`
    : "继续（此前未记录步骤）";
}
