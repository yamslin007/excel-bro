import type {
  ActivityLog,
  ActivityOutcome,
  ActivityStep
} from "../../excel-addin/src/hooks/useActivityProgress";
import type { TaskProgressStatus } from "../../excel-addin/src/contracts";
import {
  applyTimelineEvent,
  finishTimelineAt,
  isTerminalStatus,
  markTimelineBoundary,
  resumeTimeline,
  type TimelineExcludedKind,
  type TimelineLogLike
} from "@excel-bro/core/activityTimeline";
import type { WpsReadonlyTurnProgress } from "./wpsReadonlyAgent";
import type { WpsExecutionProgress } from "./wpsPlanExecutor";
import { segmentAcceptsProgress, type ReplySegment } from "@excel-bro/core/replySegments";
import {
  applyPythonCodeEvent,
  DEFAULT_PYTHON_CODE_LIMITS,
  type PythonCodeEvent
} from "@excel-bro/core/pythonCodeArtifact";

/** 两条生产写回路径共用：在调用日志归并/收尾之前检查段身份。 */
export function updateWpsActivityMessage<M extends {
  id: string; replySegment?: ReplySegment; activityLog?: ActivityLog;
}>(messages: M[], id: string, update: (log: ActivityLog) => ActivityLog,
   operationId?: string): M[] {
  const segments = messages.flatMap((item) => item.replySegment ? [item.replySegment] : []);
  return messages.map((message) => message.id === id && message.activityLog &&
    segmentAcceptsProgress(segments, id, operationId ?? message.activityLog.operationId)
    ? { ...message, activityLog: update(message.activityLog) } : message);
}

export type WpsActivityUpdate = Pick<
  ActivityStep,
  "id" | "label" | "phase" | "detail" | "toolName" | "actionType" |
  "target" | "current" | "total" | "unit" | "facts" | "note" | "pythonResult"
> & {
  id: string;
  label: string;
  /** Python 代码工件的**事件**（不是工件本体）：由共享层归并进已有工件。 */
  pythonCode?: PythonCodeEvent;
  operationId?: string;
  sequence?: number;
  /**
   * 这一步**真实**的状态。
   *
   * 以前这个字段不存在，`advanceWpsActivity` 一律写 `progress`，于是"完成"
   * 在界面上根本表达不出来 —— 那是 WPS 端"步骤永远转圈、最后一步永不收尾"的根因。
   * 缺省仍然按"刚开始"处理，但**调用方有真状态就必须传**。
   */
  status?: TaskProgressStatus;
  /**
   * **生产方声明**：报出这一条更新时，这些步骤已经确实做完了。
   *
   * 与"看到下一步就猜上一步完成"的关键区别是**谁在做判断**：这里由生产方
   * （执行器适配层）按它自己真实走过的顺序声明，时间线层只是照办。
   * 时间线层不允许自己推导这件事 —— 推导出来的"完成"会写在用户看得见的界面上。
   */
  completes?: string[];
};

/**
 * 服务端事件 → 时间线更新。
 *
 * 有 `event` 时**按事件整体派生**：阶段、状态、计数、目标、工具名、事实全部来自
 * 同一条结构化事件，不是逐字段手挑。手挑的写法已经漏过一次字段（`bridgeTool`），
 * 那种缺陷在界面上看不出原因，只能靠"整条带着"从结构上消除。
 */
export function readonlyActivityUpdate(progress: WpsReadonlyTurnProgress): WpsActivityUpdate {
  const labels = {
    reading: "读取已授权数据",
    intent: "理解需求",
    calculating: "本地确定性计算",
    answering: "生成回答或计划"
  };
  if (progress.event) {
    const event = progress.event;
    return {
      operationId: progress.turnId ?? event.operationId,
      sequence: event.sequence,
      id: event.stepId,
      label: event.title || labels[progress.stage],
      phase: event.phase,
      status: event.status,
      detail: event.detail ?? event.summary ?? undefined,
      toolName: event.toolName ?? undefined,
      actionType: event.actionType ?? undefined,
      target: event.target ?? undefined,
      current: event.current ?? undefined,
      total: event.total ?? undefined,
      unit: event.unit ?? undefined,
      pythonResult: event.pythonResult ?? undefined,
      pythonCode: event.pythonCode ?? undefined,
      facts: event.facts ?? undefined
    };
  }
  return {
    operationId: progress.turnId,
    sequence: progress.sequence,
    id: progress.stepId ?? progress.stage,
    label: progress.title ?? labels[progress.stage],
    phase: progress.stage === "intent" ? "intent"
      : progress.stage === "answering" ? "planning"
        : progress.stage === "calculating" ? "executing" : "preflight",
    status: progress.status,
    detail: progress.message,
    toolName: progress.toolName,
    target: progress.sheetName ? { sheet: progress.sheetName } : undefined,
    current: progress.rowsRead,
    total: progress.totalRows,
    unit: progress.totalRows == null ? undefined : "rows"
  };
}

export function executionActivityUpdate(progress: WpsExecutionProgress): WpsActivityUpdate {
  const index = progress.actionIndex ?? 0;
  /*
   * 执行器是**顺序**推进的，它自己知道"上一步已经做完"。
   *
   * 这个事实以前被时间线层用"又来了新的一步"倒推出来 —— 猜对是运气，猜错
   * 就是把没写完的一步标成完成。现在由执行器适配层**显式声明**。
   */
  const completes: string[] = [];
  if (progress.stage === "writing") {
    if (index === 0) completes.push("snapshot");
    else completes.push(`writing:${index - 1}`);
  } else if (progress.stage === "verifying") {
    completes.push(`writing:${index}`);
    if (index > 0) completes.push(`verifying:${index - 1}`);
  } else if (progress.stage === "undoing" && index > 0) {
    completes.push(`undoing:${index - 1}`);
  }
  const facts = progress.stage === "snapshot"
    ? [
        ...(progress.actionCount == null ? [] : [{ label: "操作", value: `${progress.actionCount} 步` }]),
        ...(progress.cellCount == null ? [] : [{ label: "目标", value: `${progress.cellCount} 个单元格` }])
      ]
    : [
        ...(progress.actionCount == null ? [] : [{ label: "进度", value: `${index + 1}/${progress.actionCount} 步` }]),
        ...(progress.cellCount == null ? [] : [{ label: "目标", value: `${progress.cellCount} 个单元格` }])
      ];
  const phase = progress.stage === "snapshot" ? "backup"
    : progress.stage === "verifying" ? "verifying" : "executing";
  const label = progress.stage === "snapshot" ? "保存写前恢复点"
    : progress.stage === "writing" ? `执行第 ${index + 1} 个写入动作`
      : progress.stage === "verifying" ? `回读核对第 ${index + 1} 个动作`
        : `撤销第 ${index + 1} 个动作`;
  return {
    id: progress.stage === "snapshot" ? "snapshot" : `${progress.stage}:${index}`,
    label,
    phase,
    detail: progress.message,
    actionType: progress.actionType,
    target: progress.sheetName ? { sheet: progress.sheetName, range: progress.range } : undefined,
    facts,
    current: progress.completedRows,
    total: progress.totalRows,
    unit: progress.totalRows == null ? undefined : "rows",
    ...(completes.length ? { completes } : {})
  };
}

/** WPS 的宿主回调只有安全摘要；不把单元格值写入对话历史。 */
export function startWpsActivity(update: WpsActivityUpdate, now = Date.now()): ActivityLog {
  const {
    operationId,
    sequence,
    status,
    completes: _completes,
    pythonCode: codeEvent,
    ...step
  } = update;
  const pythonCode = codeEvent
    ? applyPythonCodeEvent(null, codeEvent, DEFAULT_PYTHON_CODE_LIMITS).artifact
    : undefined;
  return {
    operationId,
    sequence,
    steps: [{
      ...step,
      ...(pythonCode ? { pythonCode } : {}),
      status: status ?? "started",
      elapsedMs: 0,
      startedAt: now
    }],
    totalMs: 0,
    timelineStartedAt: now
  };
}

/** 把一个步骤标成完成并冻结它的耗时；已经终态的不动。 */
function closeStep(step: ActivityStep, now: number): ActivityStep {
  if (isTerminalStatus(step.status)) return step;
  return {
    ...step,
    status: "completed",
    elapsedMs: Math.max(0, now - (step.startedAt ?? now)),
    startedAt: undefined
  };
}

function asTimeline(log: ActivityLog): TimelineLogLike<ActivityStep> {
  return log as unknown as TimelineLogLike<ActivityStep>;
}

function fromTimeline(log: TimelineLogLike<ActivityStep>): ActivityLog {
  return log as unknown as ActivityLog;
}

/**
 * 归并一条更新。
 *
 * 三条纪律（与共享层一致，这里不重复实现）：
 * 1. **状态照抄**，不再一律写 `progress`；
 * 2. **不靠"下一步出现"猜上一步完成** —— 完成必须由事件本身说出来；
 * 3. 同一 `stepId` 的重放只更新不新增；终态不会被拉回进行中。
 *
 * 本地阶段（读表结构、本地计算）没有服务端序号，因此走 "no-sequence" 分支：
 * 只按 `id` 定位、照抄状态、同样受终态保护。
 */
export function advanceWpsActivity(
  log: ActivityLog,
  update: WpsActivityUpdate,
  startedAt: number,
  now = Date.now()
): ActivityLog {
  const timeline = asTimeline(log);
  const { operationId, sequence, status, completes, pythonCode: codeEvent, ...stepUpdate } = update;
  const incoming = status ?? "progress";

  const withDeclared = (next: ActivityLog): ActivityLog => {
    if (!completes?.length) return next;
    const declared = new Set(completes);
    return {
      ...next,
      steps: next.steps.map((step) =>
        step.id && declared.has(step.id) ? closeStep(step, now) : step
      )
    };
  };

  if (operationId && sequence != null) {
    const result = applyTimelineEvent(
      timeline,
      { operationId, stepId: update.id, sequence, status: incoming, title: update.label },
      (existing) => ({
        ...(existing ?? {}),
        ...stepUpdate,
        id: update.id,
        status: incoming,
        // 代码工件由共享层归并：同一 stepId 的"校验 → 执行 → 完成"只更新一个
        // 工件；重放不会把同一段代码粘两遍，也不会丢掉"被执行的是哪一版"。
        pythonCode: codeEvent
          ? applyPythonCodeEvent(
              existing?.pythonCode ?? null,
              codeEvent,
              DEFAULT_PYTHON_CODE_LIMITS
            ).artifact
          : existing?.pythonCode,
        // 单段口径：`startedAt` 是这一段（含恢复时带来的累计）的起点。
        elapsedMs: Math.max(0, now - (existing?.startedAt ?? now)),
        startedAt: existing?.startedAt ?? now
      }),
      now
    );
    if (result.ignored) return log;
    return withDeclared(fromTimeline({ ...result.log, totalMs: Math.max(0, now - startedAt) }));
  }

  const index = timeline.steps.findIndex((step) => step.id === update.id);
  const existing = index >= 0 ? timeline.steps[index] : undefined;
  // 终态保护：已经结束的步骤不接受"重新开始"。
  if (existing && isTerminalStatus(existing.status) && !isTerminalStatus(incoming)) {
    return withDeclared(log);
  }
  const next: ActivityStep = {
    ...(existing ?? {}),
    ...stepUpdate,
    id: update.id,
    status: incoming,
    elapsedMs: existing && isTerminalStatus(existing.status)
      ? existing.elapsedMs : Math.max(0, now - (existing?.startedAt ?? now)),
    startedAt: isTerminalStatus(incoming) ? undefined : existing?.startedAt ?? now
  };
  const steps = timeline.steps.slice();
  if (index >= 0) steps[index] = next;
  else steps.push(next);
  return withDeclared({
    ...log,
    steps,
    sequence: sequence ?? log.sequence,
    totalMs: Math.max(0, now - startedAt)
  });
}

/**
 * 收尾：**所有**没结束的步骤都要有交代。
 *
 * 原来只处理数组最后一行，而"当前步骤"是按 `stepId` 定位的 —— 迟到事件、乱序
 * 补发都会让当前步骤不在末位，于是它永远停在半途，末位那条却被改写。
 */
export function finishWpsActivity(
  log: ActivityLog,
  outcome: ActivityOutcome,
  startedAt: number,
  now = Date.now(),
  activeStepId?: string
): ActivityLog {
  const timeline = asTimeline(log);
  // 不传 `activeStepId` 时由共享层按 `startedAt` 选"当前步骤" —— 不取数组末行，
  // 因为末行只是插入顺序，迟到事件会让它和当前步骤不是同一条。
  const finished = finishTimelineAt(
    timeline,
    outcome,
    activeStepId,
    now,
    Math.max(0, now - startedAt)
  );
  return fromTimeline(finished);
}

/** 把活动标记为"正在等待用户"或"断线重连"，用于把时间记进正确的桶。 */
export function markWpsActivityBoundary(
  log: ActivityLog,
  kind: TimelineExcludedKind | null,
  now = Date.now()
): ActivityLog {
  return fromTimeline(markTimelineBoundary(asTimeline(log), kind, now));
}

/** 恢复：**每一行**按自己已累计的时间重新定点，不是只修最后一行。 */
export function resumeWpsActivity(log: ActivityLog, now = Date.now()): ActivityLog {
  return fromTimeline(resumeTimeline(asTimeline(log), now));
}
