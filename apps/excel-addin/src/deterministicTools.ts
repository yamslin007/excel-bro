import type {
  DataToolRequest,
  DataToolResult,
  WorkbookSnapshot
} from "./contracts";
import {
  analyzeQueryToolCompatibility,
  instantiateQueryTool,
  type SavedQueryTool
} from "./storage";
import {
  currentModelCallCount,
  recordDiagnosticEvent
} from "./diagnostics";

export class SavedQueryToolFallbackError extends Error {
  constructor(public readonly reasons: string[]) {
    super(`固化工具需要重新确认：${reasons.join("；")}`);
    this.name = "SavedQueryToolFallbackError";
  }
}

export interface SavedQueryToolRunners {
  workbook: (request: DataToolRequest) => Promise<DataToolResult>;
  folder: (request: DataToolRequest) => Promise<DataToolResult>;
}

export async function executeSavedQueryTool(
  tool: SavedQueryTool,
  workbook: WorkbookSnapshot,
  runners: SavedQueryToolRunners
): Promise<DataToolResult> {
  const compatibility = analyzeQueryToolCompatibility(tool, workbook);
  if (!compatibility.runnable) {
    throw new SavedQueryToolFallbackError(compatibility.reasons);
  }
  const started = performance.now();
  const modelCallsBefore = currentModelCallCount();
  try {
    const request = instantiateQueryTool(tool);
    const result = await runners[tool.sourceMode](request);
    if (
      tool.expectedHeaders.length > 0 &&
      JSON.stringify(result.headers) !== JSON.stringify(tool.expectedHeaders)
    ) {
      throw new SavedQueryToolFallbackError([
        "查询结果字段已经变化，不能沿用原工具输出"
      ]);
    }
    const modelCalls = currentModelCallCount() - modelCallsBefore;
    recordDiagnosticEvent({
      timestamp: new Date().toISOString(),
      phase: "saved_tool",
      durationMs: performance.now() - started,
      scannedRows: result.scannedRows,
      modelCalls,
      status: "succeeded"
    });
    return result;
  } catch (reason) {
    recordDiagnosticEvent({
      timestamp: new Date().toISOString(),
      phase: "saved_tool",
      durationMs: performance.now() - started,
      modelCalls: currentModelCallCount() - modelCallsBefore,
      status: "failed",
      errorCategory:
        reason instanceof SavedQueryToolFallbackError
          ? "semantic_change"
          : "data_tool"
    });
    throw reason;
  }
}
