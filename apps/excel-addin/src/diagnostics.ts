export type DiagnosticPhase =
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
}

export type ProgressDiagnosticDisposition =
  | "inserted"
  | "updated"
  | "ignored-stale"
  | "ignored-operation"
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

export function exportDiagnosticReport(): void {
  const blob = new Blob([JSON.stringify(diagnosticReport(), null, 2)], {
    type: "application/json"
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `excel-bro-diagnostics-${Date.now()}.json`;
  anchor.click();
  URL.revokeObjectURL(url);
}
