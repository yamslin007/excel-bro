import { describe, expect, it } from "vitest";
import {
  currentModelCallCount,
  diagnosticReport,
  recordDiagnosticEvent,
  recordModelCall,
  recordProgressDiagnosticEvent
} from "./diagnostics";

describe("diagnostics", () => {
  it("records only structured metrics and model-call counts", () => {
    const before = currentModelCallCount();
    recordModelCall();
    recordDiagnosticEvent({
      timestamp: "2026-07-29T00:00:00Z",
      phase: "local_query",
      durationMs: 12,
      scannedRows: 200,
      modelCalls: 0,
      status: "succeeded"
    });
    const report = diagnosticReport();
    expect(report.modelCalls).toBe(before + 1);
    expect(JSON.stringify(report)).not.toContain("apiKey");
    expect(report.events.at(-1)).toMatchObject({
      phase: "local_query",
      scannedRows: 200
    });
  });

  it("records progress correlation metadata without message contents", () => {
    recordProgressDiagnosticEvent(
      {
        operationId: "operation-safe",
        executionId: "execution-safe",
        sequence: 3,
        stepId: "planning.tool.1.1",
        phase: "planning",
        status: "completed",
        toolName: "read_range",
        actionType: null
      },
      "inserted"
    );

    const report = diagnosticReport();
    expect(report.progressEvents.at(-1)).toEqual(
      expect.objectContaining({
        operationId: "operation-safe",
        sequence: 3,
        stepId: "planning.tool.1.1",
        disposition: "inserted"
      })
    );
    expect(JSON.stringify(report.progressEvents.at(-1))).not.toContain(
      "summary"
    );
    expect(JSON.stringify(report.progressEvents.at(-1))).not.toContain(
      "facts"
    );
  });
});
