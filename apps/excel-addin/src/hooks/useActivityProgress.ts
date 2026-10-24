import { useCallback, useEffect, useRef, useState } from "react";
import {
  applyTimelineEvent,
  finishTimelineAt,
  isTerminalStatus,
  resumeTimeline,
  markTimelineBoundary,
  type TimelineLogLike
} from "@excel-bro/core/activityTimeline";
import type {
  TaskProgressEvent,
  TaskProgressFact,
  TaskProgressPhase,
  TaskProgressStatus,
  TaskProgressTarget,
  TaskProgressTrace,
  LocalResultCard
} from "../contracts";
import { PYTHON_CODE_LIMITS } from "../contracts";
import {
  applyPythonCodeEvent,
  type PythonCodeArtifact
} from "@excel-bro/core/pythonCodeArtifact";
import { recordProgressDiagnosticEvent } from "../diagnostics";

/**
 * 代码工件的展示上限。
 *
 * 与 `config/capabilities.json` 的 `pythonCode` 同源（`PYTHON_CODE_LIMITS`），
 * 不在这一层另写一组数字 —— 服务端按同一组上限截断，两边数字不一致就会出
 * "服务端以为给了全文、界面以为被截了"的假象。
 */
const pythonCodeLimits = PYTHON_CODE_LIMITS;

export type ActivityConnectionState =
  | "connected"
  | "reconnecting"
  | "interrupted";

export type ActivityOutcome =
  | "completed"
  | "awaiting_input"
  | "warning"
  | "failed"
  | "cancelled"
  | "interrupted";

export interface ActivityStep {
  id?: string;
  label: string;
  elapsedMs: number;
  phase?: TaskProgressPhase;
  status?: TaskProgressStatus;
  current?: number | null;
  total?: number | null;
  unit?: TaskProgressEvent["unit"];
  // note：模型对这一步的真实判断（第一层，默认可见），例如「它把需求理解成了什么」。
  note?: string;
  // detail：原始明细（第二层，展开「查看详情」才显示），例如锁定的执行指令、结构化返回。
  detail?: string;
  toolName?: string | null;
  actionType?: string | null;
  target?: TaskProgressTarget | null;
  facts?: TaskProgressFact[] | null;
  trace?: (TaskProgressTrace & { id: string }) | null;
  /** Python 分析完成时的结果卡（只含引用与结构，值由前端另行查询）。 */
  pythonResult?: LocalResultCard | null;
  /**
   * 这一次分析生成的 Python 代码工件（状态 + 正文）。
   *
   * 与 `pythonResult` 分开：结果卡是"分析产出了什么"，工件是"跑了哪段代码"。
   * 两者混在一起就无法回答"屏幕上这段代码是不是真的被运行过"。
   */
  pythonCode?: PythonCodeArtifact | null;
  startedAt?: number;
}

interface ActivityTiming {
  timelineStartedAt?: number;
  pausedMs?: number;
  waitingMs?: number;
  excludedSince?: number;
  excludedKind?: "pause" | "waiting";
}

function timingFields(value: ActivityTiming | null | undefined): ActivityTiming {
  return { timelineStartedAt: value?.timelineStartedAt, pausedMs: value?.pausedMs,
    waitingMs: value?.waitingMs, excludedSince: value?.excludedSince, excludedKind: value?.excludedKind };
}

function activityBoundary(current: ActivityProgress, kind: "waiting" | null): ActivityProgress {
  const marked = markTimelineBoundary({ ...current, totalMs: 0 }, kind, Date.now());
  return { ...current, ...timingFields(marked), steps: marked.steps };
}

export interface ActivityProgress extends ActivityTiming {
  persistenceKey?: string;
  operationId: string | null;
  executionId?: string | null;
  sequence: number;
  phase: TaskProgressPhase;
  stepId: string;
  status: TaskProgressStatus;
  title: string;
  detail: string;
  summary?: string;
  current?: number | null;
  total?: number | null;
  unit?: TaskProgressEvent["unit"];
  connectionState: ActivityConnectionState;
  connectionMessage?: string;
  traces: Array<TaskProgressTrace & { id: string }>;
  steps: ActivityStep[];
  // 兼容仍读取 completed 的旧调用点；新时间线一律使用 steps。
  completed: ActivityStep[];
  startedAt: number;
  lastStepAt: number;
}

export interface ActivityLog extends ActivityTiming {
  outcome?: ActivityOutcome;
  /** Local presentation only: completed ordinary chat keeps details available but closed. */
  defaultCollapsed?: boolean;
  operationId?: string | null;
  /** 前端最后成功归并的服务端事件序号；旧日志可能没有。 */
  sequence?: number;
  steps: ActivityStep[];
  traces?: Array<TaskProgressTrace & { id: string }>;
  totalMs: number;
  connectionState?: ActivityConnectionState;
  connectionMessage?: string;
  /**
   * 上一次收尾顺手关掉的步骤 id（不是收到真实结束事件的）。
   *
   * 恢复时按它重新打开：暂停/断线打断的那一步还没做完，后续进度必须还能落到
   * 同一行上。**不能**只靠 `status === "warning"` 判 —— 那只是历史日志的兜底。
   */
  closedByFinish?: string[];
}

interface StartActivityOptions {
  persistenceKey?: string;
  operationId?: string | null;
  phase?: TaskProgressPhase;
  stepId?: string;
}

interface ResumeActivityOptions extends StartActivityOptions {
  previousLog?: ActivityLog | null;
  initialSequence?: number;
}

interface UseActivityProgressOptions {
  onPersistLog: (log: ActivityLog, persistenceKey?: string) => void;
}

/**
 * Resume presentation without inventing an extra numbered business step.
 *
 * **每一行**都按自己已经累计的耗时重新定点，不是只修最后一行。
 * 只修最后一行的后果是：恢复后其余行从 0 重新计时，用户看到"第一步又花了 0 秒"，
 * 而标题上的总耗时却还是旧的 —— 两个数字互相矛盾。
 * 已完成的步骤原样保留（它们的时间不再增长）。
 */
export function resumeActivityLog(log: ActivityLog, now = Date.now()): ActivityLog {
  return {
    ...(resumeTimeline(
      log as unknown as TimelineLogLike<ActivityStep>,
      now
    ) as unknown as ActivityLog),
    defaultCollapsed: false,
    connectionState: "connected",
    connectionMessage: undefined
  };
}

const TERMINAL_STATUSES = new Set<TaskProgressStatus>([
  "completed",
  "warning",
  "failed",
  "cancelled"
]);

function terminalSteps(steps: ActivityStep[]): ActivityStep[] {
  return steps.filter(
    (step) => step.status && TERMINAL_STATUSES.has(step.status)
  );
}

function persistableStep(step: ActivityStep): ActivityStep {
  const { startedAt: _startedAt, ...persisted } = step;
  return persisted;
}

function progressTrace(
  event: TaskProgressEvent
): (TaskProgressTrace & { id: string }) | null {
  return event.trace
    ? {
        ...event.trace,
        id: event.stepId + ":" + event.trace.kind
      }
    : null;
}

function fallbackSummary(status: TaskProgressStatus): string {
  if (status === "started" || status === "progress") return "正在处理";
  if (status === "completed") return "已完成";
  if (status === "warning") return "出现警告";
  if (status === "cancelled") return "已取消";
  return "执行失败";
}

/**
 * 活动进度管理 Hook
 *
 * 职责：
 * - 管理当前执行/规划进度的标题、明细和时间线步骤
 * - 管理进度计时器（按秒刷新）
 * - 将当前 activity 的步骤固化为日志并通知上层
 */
export function useActivityProgress({
  onPersistLog
}: UseActivityProgressOptions) {
  const [activity, setActivity] = useState<ActivityProgress | null>(null);
  const activityRef = useRef<ActivityProgress | null>(null);
  const [activitySeconds, setActivitySeconds] = useState(0);

  const updateActivity = useCallback(
    (
      updater: (
        current: ActivityProgress | null
      ) => ActivityProgress | null
    ) => {
      const next = updater(activityRef.current);
      activityRef.current = next;
      setActivity(next);
    },
    []
  );

  const startActivity = useCallback(
    (
      title: string,
      detail: string,
      options: StartActivityOptions = {}
    ) => {
      const now = Date.now();
      const stepId = options.stepId ?? "local";
      const firstStep: ActivityStep = {
        id: stepId,
        label: title,
        elapsedMs: 0,
        phase: options.phase ?? "intent",
        status: "started",
        detail: detail || undefined,
        startedAt: now
      };
      const nextActivity: ActivityProgress = {
        timelineStartedAt: now,
        persistenceKey: options.persistenceKey,
        operationId: options.operationId ?? null,
        sequence: -1,
        phase: options.phase ?? "intent",
        stepId,
        status: "started",
        title,
        detail,
        connectionState: "connected",
        traces: [],
        steps: [firstStep],
        completed: [],
        startedAt: now,
        lastStepAt: now
      };
      activityRef.current = nextActivity;
      setActivity(nextActivity);
      setActivitySeconds(0);
    },
    []
  );

  const bindActivityOperation = useCallback(
    (operationId: string) => {
      updateActivity((current) =>
        current ? { ...current, operationId } : current
      );
    },
    [updateActivity]
  );

  const getActivitySnapshot = useCallback(() => activityRef.current, []);

  const resumeActivity = useCallback(
    (
      title: string,
      detail: string,
      options: ResumeActivityOptions
    ) => {
      const now = Date.now();
      const current = activityRef.current;
      const canReuseCurrent =
        current?.operationId === (options.operationId ?? null) &&
        current.persistenceKey === options.persistenceKey;
      const previousLog = canReuseCurrent
        ? {
            ...timingFields(current),
            operationId: current.operationId,
            sequence: current.sequence,
            steps: current.steps,
            traces: current.traces,
            totalMs: Math.max(0, now - current.startedAt),
            connectionState: current.connectionState,
            connectionMessage: current.connectionMessage
          }
        : options.previousLog;
      const resumedLog = previousLog ? resumeActivityLog(previousLog, now) : null;
      const lastStep = resumedLog?.steps.at(-1);
      const stepId = lastStep?.id ?? options.stepId ?? "planning.resume";
      const priorSteps = resumedLog?.steps ?? [];
      const existingIndex = priorSteps.findIndex((step) => step.id === stepId);
      const resumeStep: ActivityStep = {
        id: stepId,
        label: title,
        elapsedMs: 0,
        phase: options.phase ?? "planning",
        status: "started",
        detail: detail || undefined,
        startedAt: now
      };
      const steps =
        existingIndex >= 0
          ? priorSteps.map((step, index) =>
              index === existingIndex ? (lastStep ?? resumeStep) : step
            )
          : [...priorSteps, resumeStep];
      const carriedMs = Math.max(0, resumedLog?.totalMs ?? 0);
      const nextActivity: ActivityProgress = {
        ...timingFields(resumedLog),
        timelineStartedAt: resumedLog?.timelineStartedAt ?? now - carriedMs,
        persistenceKey: options.persistenceKey,
        operationId: options.operationId ?? previousLog?.operationId ?? null,
        sequence:
          options.initialSequence ?? previousLog?.sequence ?? current?.sequence ?? -1,
        phase: options.phase ?? "planning",
        stepId,
        status: "started",
        title,
        detail,
        connectionState: "connected",
        traces: [...(previousLog?.traces ?? [])],
        steps,
        completed: terminalSteps(steps),
        startedAt: now - carriedMs,
        lastStepAt: now
      };
      activityRef.current = nextActivity;
      setActivity(nextActivity);
      setActivitySeconds(Math.floor(carriedMs / 1000));
    },
    []
  );

  const advanceActivity = useCallback(
    (
      title: string,
      detail: string,
      completedStep?: string,
      stepInsight?: { note?: string; detail?: string }
    ) => {
      updateActivity((current) => {
        const now = Date.now();
        const startedAt = current?.startedAt ?? now;
        let steps = current?.steps ?? [];
        let nextStepId = current?.stepId ?? "legacy";
        let nextStatus = current?.status ?? "progress";
        if (steps.length === 0) {
          steps = [
            {
              id: nextStepId,
              label: title,
              elapsedMs: 0,
              phase: current?.phase ?? "planning",
              status: "started",
              detail: detail || undefined,
              startedAt: now
            }
          ];
        }
        if (completedStep && completedStep.trim()) {
          const activeIndex = Math.max(
            0,
            steps.findIndex((step) => step.id === current?.stepId)
          );
          steps = steps.map((step, index) =>
            index === activeIndex
              ? {
                  ...step,
                  label: completedStep,
                  status: "completed",
                  note: stepInsight?.note,
                  detail: stepInsight?.detail ?? step.detail,
                  elapsedMs: Math.max(0, now - (step.startedAt ?? now))
                }
              : step
          );
          if (title.trim() !== completedStep.trim()) {
            nextStepId = "legacy-" + (steps.length + 1);
            steps = [
              ...steps,
              {
                id: nextStepId,
                label: title,
                elapsedMs: 0,
                phase: current?.phase ?? "planning",
                status: "started",
                detail: detail || undefined,
                startedAt: now
              }
            ];
            nextStatus = "started";
          } else {
            nextStatus = "completed";
          }
        } else {
          const activeIndex = steps.findIndex(
            (step) => step.id === current?.stepId
          );
          if (activeIndex >= 0) {
            const activeStatus = steps[activeIndex].status ?? "progress";
            if (TERMINAL_STATUSES.has(activeStatus)) {
              nextStepId = "legacy-" + (steps.length + 1);
              steps = [
                ...steps,
                {
                  id: nextStepId,
                  label: title,
                  elapsedMs: 0,
                  phase: current?.phase ?? "planning",
                  status: "started",
                  detail: detail || undefined,
                  startedAt: now
                }
              ];
              nextStatus = "started";
            } else {
              steps = steps.map((step, index) =>
                index === activeIndex
                  ? {
                      ...step,
                      label: title,
                      detail: detail || undefined,
                      status: "progress",
                      elapsedMs: Math.max(0, now - (step.startedAt ?? now))
                    }
                  : step
              );
              nextStatus = "progress";
            }
          }
        }
        return {
          ...timingFields(current),
          timelineStartedAt: current?.timelineStartedAt ?? startedAt,
          persistenceKey: current?.persistenceKey,
          operationId: current?.operationId ?? null,
          executionId: current?.executionId ?? null,
          sequence: current?.sequence ?? -1,
          phase: current?.phase ?? "planning",
          stepId: nextStepId,
          status: nextStatus,
          title,
          detail,
          connectionState: current?.connectionState ?? "connected",
          connectionMessage: current?.connectionMessage,
          traces: current?.traces ?? [],
          steps,
          completed: terminalSteps(steps),
          startedAt,
          lastStepAt:
            completedStep && completedStep.trim()
              ? now
              : current?.lastStepAt ?? now
        };
      });
    },
    [updateActivity]
  );

  const completeCurrentStep = useCallback(
    (
      label: string,
      stepInsight?: { note?: string; detail?: string }
    ) => {
      updateActivity((current) => {
        if (!current) return current;
        const now = Date.now();
        const activeIndex = current.steps.findIndex(
          (step) => step.id === current.stepId
        );
        if (activeIndex < 0) return current;
        const steps = current.steps.map((step, index) =>
          index === activeIndex && !isTerminalStatus(step.status)
            ? {
                ...step,
                label,
                status: "completed" as const,
                note: stepInsight?.note ?? step.note,
                detail: stepInsight?.detail ?? step.detail,
                elapsedMs: Math.max(0, now - (step.startedAt ?? now))
              }
            : step
        );
        return {
          ...current,
          title: label,
          detail: stepInsight?.detail ?? current.detail,
          summary: stepInsight?.note ?? current.summary,
          status: "completed",
          steps,
          completed: terminalSteps(steps),
          lastStepAt: now
        };
      });
    },
    [updateActivity]
  );

  const advanceActivityStage = useCallback((stepId: string, title: string, detail: string) => {
    updateActivity((current) => {
      if (!current) return current;
      const now = Date.now();
      const existingIndex = current.steps.findIndex((step) => step.id === stepId);
      const existing = current.steps[existingIndex];
      // Re-reading the host is a verification, not a restart of completed UI steps.
      if (existing?.status === "completed" && existingIndex < current.steps.length - 1) return current;
      const steps = current.steps.map((step) => step.id === current.stepId && step.id !== stepId &&
        (step.status === "started" || step.status === "progress")
        ? { ...step, status: "completed" as const, elapsedMs: now - (step.startedAt ?? now) } : step);
      const next: ActivityStep = { ...existing, id: stepId, label: title, detail,
        status: "progress", phase: "intent", startedAt: existing?.startedAt ?? now,
        elapsedMs: existing?.elapsedMs ?? 0 };
      if (existingIndex < 0) steps.push(next); else steps[existingIndex] = next;
      return { ...current, steps, completed: terminalSteps(steps), stepId, title, detail,
        status: "progress", lastStepAt: now };
    });
  }, [updateActivity]);

  const applyProgressEvent = useCallback((event: TaskProgressEvent) => {
    updateActivity((current) => {
      if (current?.operationId && current.operationId !== event.operationId) {
        recordProgressDiagnosticEvent(event, "ignored-operation");
        return current;
      }
      if (
        current?.operationId === event.operationId &&
        event.sequence <= current.sequence
      ) {
        recordProgressDiagnosticEvent(event, "ignored-stale");
        return current;
      }
      const now = Date.now();
      current = current ? activityBoundary(current, null) : current;
      const startedAt = current?.startedAt ?? now;
      const trace = progressTrace(event);
      /*
       * 身份、序号、终态三条判据走共享层（与 WPS 同一份实现）：
       * - 同一 `stepId` 的重放只更新不新增；
       * - 已完成的步骤不会被迟到的"进行中"事件重新激活；
       * - 属于另一条时间线的事件不改这个面板。
       */
      const result = applyTimelineEvent(
        {
          operationId: current?.operationId ?? undefined,
          sequence: current?.sequence,
          steps: current?.steps ?? [],
          totalMs: 0
        } as TimelineLogLike<ActivityStep>,
        {
          operationId: event.operationId,
          stepId: event.stepId,
          sequence: event.sequence,
          status: event.status,
          title: event.title
        },
        (existing) => {
          /*
           * 代码工件走共享层归并：同一 `stepId` 的多条状态事件（校验 → 执行 →
           * 完成）只更新**一个**工件，重放/乱序不会把同一段代码粘两遍，也不会
           * 让"已运行的是哪一版"失去依据。
           */
          const pythonCode = event.pythonCode
            ? applyPythonCodeEvent(existing?.pythonCode ?? null, event.pythonCode, pythonCodeLimits).artifact
            : existing?.pythonCode ?? null;
          return {
            id: event.stepId,
            label: event.title || fallbackSummary(event.status),
            elapsedMs: Math.max(0, now - (existing?.startedAt ?? now)),
            phase: event.phase,
            // 状态**照抄事件**。Excel 端一直是这么做的；WPS 端曾经强制写成
            // progress，两条链因此对同一份协议给出了不同结论。
            status: event.status,
            current: event.current,
            total: event.total,
            unit: event.unit,
            note: event.summary ?? event.trace?.summary ?? existing?.note,
            detail: event.detail ?? existing?.detail,
            toolName: event.toolName ?? existing?.toolName ?? null,
            actionType: event.actionType ?? existing?.actionType ?? null,
            target: event.target ?? existing?.target ?? null,
            facts: event.facts ?? existing?.facts ?? null,
            trace: trace ?? existing?.trace ?? null,
            pythonResult: event.pythonResult ?? existing?.pythonResult ?? null,
            pythonCode,
            startedAt: existing?.startedAt ?? now
          };
        },
        now
      );
      if (result.ignored) {
        recordProgressDiagnosticEvent(
          event,
          result.ignored === "stale" ? "ignored-stale"
            : result.ignored === "operation" ? "ignored-operation" : "ignored-terminal-replay"
        );
        return current;
      }
      const steps = result.log.steps;
      recordProgressDiagnosticEvent(event, result.inserted ? "inserted" : "updated");
      let traces = current?.traces ?? [];
      if (trace) {
        const traceIndex = traces.findIndex((item) => item.id === trace.id);
        traces =
          traceIndex === -1
            ? [...traces, trace]
            : traces.map((item, index) => (index === traceIndex ? trace : item));
      }
      return {
        ...timingFields(current),
        timelineStartedAt: current?.timelineStartedAt ?? startedAt,
        persistenceKey: current?.persistenceKey,
        operationId: event.operationId,
        executionId: event.executionId ?? current?.executionId ?? null,
        sequence: event.sequence,
        phase: event.phase,
        stepId: event.stepId,
        status: event.status,
        title: event.title,
        detail: event.detail ?? "",
        summary: event.summary ?? event.trace?.summary ?? undefined,
        current: event.current,
        total: event.total,
        unit: event.unit,
        connectionState: "connected",
        connectionMessage: undefined,
        traces,
        steps,
        completed: terminalSteps(steps),
        startedAt,
        lastStepAt: isTerminalStatus(event.status)
          ? now
          : current?.lastStepAt ?? startedAt
      };
    });
  }, [updateActivity]);

  const updateActivityDetail = useCallback((detail: string) => {
    updateActivity((current) => {
      if (!current) return current;
      const steps = current.steps.map((step) =>
        step.id === current.stepId && !isTerminalStatus(step.status) && !current.excludedKind
          ? {
              ...step,
              detail,
              elapsedMs: Math.max(
                0,
                Date.now() - (step.startedAt ?? current.lastStepAt)
              )
            }
          : step
      );
      return { ...current, detail, steps, completed: terminalSteps(steps) };
    });
  }, [updateActivity]);

  const markActivityInterrupted = useCallback((message: string) => {
    updateActivity((current) => {
      if (!current) return current;
      return {
        ...activityBoundary(current, "waiting"),
        summary: message,
        connectionState: "interrupted",
        connectionMessage: message
      };
    });
  }, [updateActivity]);

  const markActivityReconnecting = useCallback((message: string) => {
    updateActivity((current) =>
      current
        ? {
            ...activityBoundary(current, "waiting"),
            connectionState: "reconnecting",
            connectionMessage: message
          }
        : current
    );
  }, [updateActivity]);

  const markActivityConnected = useCallback(() => {
    updateActivity((current) =>
      current
        ? {
            ...activityBoundary(current, null),
            connectionState: "connected",
            connectionMessage: undefined
          }
        : current
    );
  }, [updateActivity]);

  const completeActivity = useCallback(
    (
      finalStatus: TaskProgressStatus = "completed",
      finalSummary?: string,
      outcome?: ActivityOutcome,
      presentation: { defaultCollapsed?: boolean } = {}
    ) => {
      const current = activityRef.current;
      if (!current) {
        setActivitySeconds(0);
        return;
      }
      const now = Date.now();
      const effectiveStatus =
        current.connectionState === "interrupted" &&
        finalStatus === "completed"
          ? "warning"
          : finalStatus;
      const hasUnfinishedEarlierStep = current.steps.some((step) =>
        step.id !== current.stepId &&
        !TERMINAL_STATUSES.has(step.status ?? "progress")
      );
      const effectiveOutcome =
        outcome ??
        (current.connectionState === "interrupted"
          ? "interrupted"
          : effectiveStatus === "warning" ||
              effectiveStatus === "failed" ||
              effectiveStatus === "cancelled"
            ? effectiveStatus
            : hasUnfinishedEarlierStep ? "warning" : "completed");
      /*
       * 收尾也走共享层：**所有**没结束的步骤都要有交代（当前步骤拿活动级终态，
       * 其余拿 warning + "未收到结束确认"），并且把耗时冻在收尾这一刻。
       * 这里不再自己写一遍 map —— 两端各写一遍正是"最后一行"那类缺陷的来源。
       */
      const finished = finishTimelineAt(
        {
          ...timingFields(current),
          operationId: current.operationId,
          steps: current.steps,
          totalMs: 0
        } as TimelineLogLike<ActivityStep>,
        effectiveOutcome,
        current.stepId,
        now,
        Math.max(0, now - current.startedAt),
        undefined,
        // 这一步自己的终态由调用方给（`finalStatus`），不按活动级 `outcome` 推：
        // "带着待确认结果收尾"的活动中，这一步可能确实已经做完了。
        effectiveStatus
      );
      const steps = finished.steps.map((step) =>
        step.id === current.stepId && finalSummary
          ? { ...step, note: finalSummary }
          : step
      );
      activityRef.current = null;
      setActivity(null);
      setActivitySeconds(0);
      if (steps.length > 0) {
        const log: ActivityLog = {
          ...timingFields(finished),
          outcome: effectiveOutcome,
          ...(effectiveOutcome === "completed" && presentation.defaultCollapsed === true ? { defaultCollapsed: true } : {}),
          operationId: current.operationId,
          sequence: current.sequence,
          steps: steps.map(persistableStep),
          traces: current.traces,
          totalMs: Math.max(0, now - current.startedAt),
          connectionState: current.connectionState,
          connectionMessage: current.connectionMessage,
          // 收尾"顺手关掉"的那些步骤要一起落盘：恢复时靠它把被打断的步骤重新打开。
          closedByFinish: finished.closedByFinish
        };
        if (current.persistenceKey) {
          onPersistLog(log, current.persistenceKey);
        } else {
          onPersistLog(log);
        }
      }
    },
    [onPersistLog]
  );

  useEffect(() => {
    if (!activity) return;
    const updateElapsed = () =>
      setActivitySeconds(
        Math.max(0, Math.floor((Date.now() - activity.startedAt) / 1000))
      );
    updateElapsed();
    const timer = window.setInterval(updateElapsed, 1000);
    return () => window.clearInterval(timer);
  }, [activity?.startedAt]);

  return {
    activity,
    activitySeconds,
    startActivity,
    resumeActivity,
    advanceActivityStage,
    bindActivityOperation,
    getActivitySnapshot,
    advanceActivity,
    completeCurrentStep,
    applyProgressEvent,
    updateActivityDetail,
    markActivityInterrupted,
    markActivityReconnecting,
    markActivityConnected,
    completeActivity
  };
}
