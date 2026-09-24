import {
  checkIntent,
  createAssistantResponse,
  type ModelOption
} from "../../excel-addin/src/api";
import type {
  AssistantResponse,
  DataToolRequest,
  DataToolResult,
  IntentCheckRequest,
  IntentCheckResponse,
  IntentScopeContext,
  ResultContext,
  WorkbookSnapshot
} from "../../excel-addin/src/contracts";
import {
  executeQueryTableData,
  type SheetValues
} from "../../excel-addin/src/dataTools";
import { detectSheetFields } from "../../excel-addin/src/tableSchema";
import { readWpsHostProbe, type WpsApplication } from "./wpsHost";
import {
  authorizeWpsSheets,
  listWpsWorksheets,
  readAuthorizedWpsSheets,
  type WpsSheetReadResult
} from "./wpsWorkbookReader";

export type WpsReadonlyTurnResult =
  | {
      kind: "answer";
      message: string;
      resultContext?: ResultContext | null;
      provider: "model" | "local" | "system";
    }
  | {
      kind: "clarification";
      response: Extract<IntentCheckResponse, { kind: "clarification" }>;
    }
  | {
      kind: "blocked";
      message: string;
    };

export type WpsReadonlyTurnProgress = {
  stage: "reading" | "intent" | "calculating" | "answering";
  message: string;
};

export type WpsReadonlyTurnInput = {
  application: WpsApplication;
  selectedSheetIndexes: readonly number[];
  prompt: string;
  modelId?: string | null;
  conversation?: IntentCheckRequest["conversation"];
  intentConfirmed?: boolean;
  clarificationRound?: number;
  signal?: AbortSignal;
  onProgress?: (progress: WpsReadonlyTurnProgress) => void;
};

type WpsReadonlyAgentDependencies = {
  checkIntent: typeof checkIntent;
  createAssistantResponse: typeof createAssistantResponse;
  executeQuery: (request: DataToolRequest, sheets: SheetValues[]) => DataToolResult;
  readSheets: typeof readAuthorizedWpsSheets;
};

const DEFAULT_DEPENDENCIES: WpsReadonlyAgentDependencies = {
  checkIntent,
  createAssistantResponse,
  executeQuery: executeQueryTableData,
  readSheets: readAuthorizedWpsSheets
};

function columnName(index: number): string {
  let value = Math.max(1, index);
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function usedRangeAddress(result: WpsSheetReadResult): string | null {
  if (result.rowCount <= 0 || result.columnCount <= 0) return null;
  const first = `${columnName(result.startColumn)}${result.startRow}`;
  const last = `${columnName(result.startColumn + result.columnCount - 1)}${result.startRow + result.rowCount - 1}`;
  return first === last ? first : `${first}:${last}`;
}

export function buildWpsReadonlyContexts(
  application: WpsApplication,
  results: WpsSheetReadResult[]
): { scope: IntentScopeContext; workbook: WorkbookSnapshot } {
  const metadata = listWpsWorksheets(application);
  const probe = readWpsHostProbe(application);
  const sheets = results.map((result) => ({
    name: result.sheet.name,
    usedRange: usedRangeAddress(result),
    rowCount: result.rowCount,
    columnCount: result.columnCount,
    headers: detectSheetFields(result.values)
  }));
  const workbook: WorkbookSnapshot = {
    name: metadata.name,
    capturedAt: new Date().toISOString(),
    activeWorksheet: probe.activeSheetName,
    selectedRange: probe.selectionAddress,
    worksheets: sheets.map((sheet) => ({
      ...sheet,
      dataRows: [],
      truncated: sheet.rowCount > 0
    }))
  };
  const scope: IntentScopeContext = {
    workbookName: metadata.name,
    sourceMode: "workbook",
    selectionMode: "manual",
    activeWorksheet: probe.activeSheetName,
    selectedRange: probe.selectionAddress,
    totalWorksheetCount: metadata.sheets.length,
    // WPS W2 is read-only; unselected names are deliberately not exposed to planning.
    worksheetNames: sheets.map((sheet) => sheet.name),
    sheets
  };
  return { scope, workbook };
}

function sheetsForRequest(
  request: DataToolRequest,
  results: WpsSheetReadResult[],
  activeWorksheet: string
): SheetValues[] {
  const selected =
    request.arguments.scope === "active"
      ? results.filter((result) => result.sheet.name === activeWorksheet)
      : results;
  if (selected.length === 0) {
    throw new Error("查询要求读取当前工作表，但当前表不在已授权范围内。");
  }
  return selected.map((result) => ({
    name: result.sheet.name,
    values: result.values
  }));
}

function finalizeResponse(response: AssistantResponse): WpsReadonlyTurnResult {
  if (response.kind === "answer") {
    return {
      kind: "answer",
      message: response.message,
      resultContext: response.resultContext,
      provider: response.provider
    };
  }
  return {
    kind: "blocked",
    message:
      "模型生成了工作簿写入方案，但 WPS 首期仅开放只读查询；本次没有预览、执行或修改任何工作表。"
  };
}

export async function runWpsReadonlyTurn(
  input: WpsReadonlyTurnInput,
  dependencies: WpsReadonlyAgentDependencies = DEFAULT_DEPENDENCIES
): Promise<WpsReadonlyTurnResult> {
  const prompt = input.prompt.trim();
  if (prompt.length < 2) throw new Error("请输入至少两个字符的问题。");
  input.onProgress?.({ stage: "reading", message: "正在读取本次明确授权的工作表…" });
  const authorization = authorizeWpsSheets(input.application, input.selectedSheetIndexes);
  const results = await dependencies.readSheets(input.application, authorization, {
    signal: input.signal,
    onProgress: ({ sheet, rowsRead, totalRows }) => {
      input.onProgress?.({
        stage: "reading",
        message: `正在读取「${sheet.name}」：${rowsRead} / ${totalRows} 行`
      });
    }
  });
  const { scope, workbook } = buildWpsReadonlyContexts(input.application, results);

  input.onProgress?.({ stage: "intent", message: "正在判断是否需要本地确定性查询…" });
  const intent = await dependencies.checkIntent(
    {
      prompt,
      scope,
      imageCount: 0,
      intentConfirmed: input.intentConfirmed ?? false,
      clarificationRound: input.clarificationRound ?? 0,
      conversation: input.conversation ?? [],
      modelId: input.modelId || null
    },
    input.signal
  );
  if (intent.kind === "clarification") return { kind: "clarification", response: intent };

  let dataResults: DataToolResult[] = [];
  let confirmedPrompt = intent.confirmedPrompt;
  if (intent.kind === "tool_request") {
    input.onProgress?.({ stage: "calculating", message: "正在本地计算完整授权数据…" });
    const querySheets = sheetsForRequest(intent.request, results, scope.activeWorksheet);
    dataResults = [dependencies.executeQuery(intent.request, querySheets)];
  }

  input.onProgress?.({ stage: "answering", message: "正在根据本地计算结果组织回答…" });
  const response = await dependencies.createAssistantResponse(
    {
      turnId: intent.turnId,
      prompt: confirmedPrompt,
      workbook,
      dataResults,
      planningReadPolicy: "none",
      modelId: input.modelId || null
    },
    input.signal
  );
  return finalizeResponse(response);
}

export type { ModelOption };
