export type DiagnosticPhase =
  | "scope_selection"
  | "scan"
  | "intent_model"
  | "planning_model"
  | "local_query"
  | "folder_query"
  | "execution"
  | "verification"
  | "saved_tool";

export interface DiagnosticEvent {
  timestamp: string;
  phase: DiagnosticPhase;
  durationMs: number;
  scannedRows?: number;
  modelCalls: number;
  status: "succeeded" | "failed" | "cancelled";
  errorCategory?: string;
  scopeSelection?: { mode: string; selectedCount: number; committedMode: string; committedCount: number };
}

export type ProgressDiagnosticDisposition =
  | "inserted"
  | "updated"
  | "ignored-stale"
  | "ignored-operation"
  /**
   * 目标步骤已经是终态，而这条事件想把它拉回进行中（重放 / 迟到 / 补发）。
   *
   * 单独一类而不是并进 `ignored-stale`：两者的原因不同 —— 序号更新但语义更旧，
   * 说明**生产方**又发了一遍旧状态，排查时要看的不是游标而是生产侧。
   */
  | "ignored-terminal-replay"
  | "rejected-invalid";

export interface ProgressDiagnosticEvent {
  receivedAt: string;
  operationId: string;
  executionId?: string | null;
  sequence: number;
  stepId: string;
  phase: string;
  status: string;
  toolName?: string | null;
  actionType?: string | null;
  disposition: ProgressDiagnosticDisposition;
}

const events: DiagnosticEvent[] = [];
const progressEvents: ProgressDiagnosticEvent[] = [];
let modelCalls = 0;

export function recordModelCall(): void {
  modelCalls += 1;
}

export function currentModelCallCount(): number {
  return modelCalls;
}

export function recordDiagnosticEvent(event: DiagnosticEvent): void {
  events.push(event);
  if (events.length > 500) events.splice(0, events.length - 500);
}

export function diagnosticEvents(): DiagnosticEvent[] {
  return events.map((event) => ({ ...event }));
}

export function recordProgressDiagnosticEvent(
  event: {
    operationId: string;
    executionId?: string | null;
    sequence: number;
    stepId: string;
    phase: string;
    status: string;
    toolName?: string | null;
    actionType?: string | null;
  },
  disposition: ProgressDiagnosticDisposition
): void {
  progressEvents.push({
    receivedAt: new Date().toISOString(),
    operationId: event.operationId,
    executionId: event.executionId,
    sequence: event.sequence,
    stepId: event.stepId,
    phase: event.phase,
    status: event.status,
    toolName: event.toolName,
    actionType: event.actionType,
    disposition
  });
  if (progressEvents.length > 200) {
    progressEvents.splice(0, progressEvents.length - 200);
  }
}

export function progressDiagnosticEvents(): ProgressDiagnosticEvent[] {
  return progressEvents.map((event) => ({ ...event }));
}

export function diagnosticReport(): {
  generatedAt: string;
  modelCalls: number;
  events: DiagnosticEvent[];
  progressEvents: ProgressDiagnosticEvent[];
} {
  return {
    generatedAt: new Date().toISOString(),
    modelCalls,
    events: diagnosticEvents(),
    progressEvents: progressDiagnosticEvents()
  };
}
