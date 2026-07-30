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

const events: DiagnosticEvent[] = [];
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

export function diagnosticReport(): {
  generatedAt: string;
  modelCalls: number;
  events: DiagnosticEvent[];
} {
  return {
    generatedAt: new Date().toISOString(),
    modelCalls,
    events: diagnosticEvents()
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
