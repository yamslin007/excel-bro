import { useCallback, useEffect, useRef, useState } from "react";
import type {
  TaskProgressEvent,
  TaskProgressFact,
  TaskProgressPhase,
  TaskProgressStatus,
  TaskProgressTarget,
  TaskProgressTrace,
  LocalResultCard
} from "../contracts";
import { recordProgressDiagnosticEvent } from "../diagnostics";

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
  startedAt?: number;
}

export interface ActivityProgress {
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

export interface ActivityLog {
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
            operationId: current.operationId,
            sequence: current.sequence,
            steps: current.steps,
            traces: current.traces,
            totalMs: Math.max(0, now - current.startedAt),
            connectionState: current.connectionState,
            connectionMessage: current.connectionMessage
          }
        : options.previousLog;
      const stepId = options.stepId ?? "planning.resume";
      const priorSteps = (previousLog?.steps ?? []).map((step) => ({ ...step }));
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
              index === existingIndex ? resumeStep : step
            )
          : [...priorSteps, resumeStep];
      const carriedMs = Math.max(0, previousLog?.totalMs ?? 0);
      const nextActivity: ActivityProgress = {
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
          index === activeIndex
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
      const startedAt = current?.startedAt ?? now;
      const steps = [...(current?.steps ?? [])];
      const existingIndex = steps.findIndex((step) => step.id === event.stepId);
      const existing = existingIndex >= 0 ? steps[existingIndex] : undefined;
      const trace = progressTrace(event);
      const nextStep: ActivityStep = {
        id: event.stepId,
        label: event.title || fallbackSummary(event.status),
        elapsedMs: Math.max(0, now - (existing?.startedAt ?? now)),
        phase: event.phase,
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
        startedAt: existing?.startedAt ?? now
      };
      if (existingIndex === -1) {
        steps.push(nextStep);
        recordProgressDiagnosticEvent(event, "inserted");
      } else {
        steps[existingIndex] = nextStep;
        recordProgressDiagnosticEvent(event, "updated");
      }
      let traces = current?.traces ?? [];
      if (trace) {
        const traceIndex = traces.findIndex((item) => item.id === trace.id);
        traces =
          traceIndex === -1
            ? [...traces, trace]
            : traces.map((item, index) => (index === traceIndex ? trace : item));
      }
      return {
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
        lastStepAt: TERMINAL_STATUSES.has(event.status)
          ? now
          : current?.lastStepAt ?? startedAt
      };
    });
  }, [updateActivity]);

  const updateActivityDetail = useCallback((detail: string) => {
    updateActivity((current) => {
      if (!current) return current;
      const steps = current.steps.map((step) =>
        step.id === current.stepId
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
        ...current,
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
            ...current,
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
            ...current,
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
      const steps = current.steps.map((step) =>
        !TERMINAL_STATUSES.has(step.status ?? "progress")
          ? {
              ...step,
              status: step.id === current.stepId ? effectiveStatus :
                effectiveStatus === "failed" || effectiveStatus === "cancelled"
                  ? effectiveStatus : "warning" as const,
              note: step.id === current.stepId ? finalSummary ?? step.note :
                "此步骤未收到结束确认；活动已停止，不代表该步骤已完成。",
              elapsedMs: Math.max(0, now - (step.startedAt ?? now))
            }
          : step
      );
      activityRef.current = null;
      setActivity(null);
      setActivitySeconds(0);
      if (steps.length > 0) {
        const log: ActivityLog = {
          outcome: effectiveOutcome,
          ...(effectiveOutcome === "completed" && presentation.defaultCollapsed === true ? { defaultCollapsed: true } : {}),
          operationId: current.operationId,
          sequence: current.sequence,
          steps: steps.map(persistableStep),
          traces: current.traces,
          totalMs: now - current.startedAt,
          connectionState: current.connectionState,
          connectionMessage: current.connectionMessage
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
