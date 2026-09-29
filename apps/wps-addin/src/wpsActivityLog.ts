import type {
  ActivityLog,
  ActivityOutcome,
  ActivityStep
} from "../../excel-addin/src/hooks/useActivityProgress";
import type { WpsReadonlyTurnProgress } from "./wpsReadonlyAgent";
import type { WpsExecutionProgress } from "./wpsPlanExecutor";

export type WpsActivityUpdate = Pick<
  ActivityStep,
  "id" | "label" | "phase" | "detail" | "toolName" | "actionType" |
  "target" | "current" | "total" | "unit" | "facts" | "note"
> & { id: string; label: string };

export function readonlyActivityUpdate(progress: WpsReadonlyTurnProgress): WpsActivityUpdate {
  const labels = {
    reading: "读取已授权数据",
    intent: "理解需求",
    calculating: "本地确定性计算",
    answering: "生成回答或计划"
  };
  return {
    id: progress.stepId ?? progress.stage,
    label: progress.title ?? labels[progress.stage],
    phase: progress.stage === "intent" ? "intent"
      : progress.stage === "answering" ? "planning"
        : progress.stage === "calculating" ? "executing" : "preflight",
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
    unit: progress.totalRows == null ? undefined : "rows"
  };
}

/** WPS 的宿主回调只有安全摘要；不把单元格值写入对话历史。 */
export function startWpsActivity(update: WpsActivityUpdate, now = Date.now()): ActivityLog {
  return {
    steps: [{ ...update, status: "started", elapsedMs: 0, startedAt: now }],
    totalMs: 0
  };
}

export function advanceWpsActivity(
  log: ActivityLog,
  update: WpsActivityUpdate,
  startedAt: number,
  now = Date.now()
): ActivityLog {
  const steps = log.steps.map((step) => ({ ...step }));
  const previous = steps[steps.length - 1];
  if (previous && previous.id !== update.id &&
    (previous.status === "started" || previous.status === "progress")) {
    previous.status = "completed";
    previous.elapsedMs = Math.max(0, now - (previous.startedAt ?? now));
  }
  const existingIndex = steps.findIndex((step) => step.id === update.id);
  const existing = existingIndex >= 0 ? steps[existingIndex] : null;
  const next: ActivityStep = {
    ...(existing ?? {}),
    ...update,
    status: "progress",
    elapsedMs: Math.max(0, now - (existing?.startedAt ?? now)),
    startedAt: existing?.startedAt ?? now
  };
  if (existingIndex >= 0) steps[existingIndex] = next;
  else steps.push(next);
  return { ...log, steps, totalMs: Math.max(0, now - startedAt) };
}

export function finishWpsActivity(
  log: ActivityLog,
  outcome: ActivityOutcome,
  startedAt: number,
  now = Date.now()
): ActivityLog {
  const terminalStatus: NonNullable<ActivityStep["status"]> = outcome === "failed" ? "failed"
    : outcome === "cancelled" ? "cancelled"
      : outcome === "warning" || outcome === "interrupted" ? "warning" : "completed";
  const steps = log.steps.map((step, index) => index === log.steps.length - 1 &&
    (step.status === "started" || step.status === "progress")
    ? {
        ...step,
        status: terminalStatus,
        elapsedMs: Math.max(0, now - (step.startedAt ?? now))
      }
    : step);
  return { ...log, steps, outcome, totalMs: Math.max(0, now - startedAt) };
}
