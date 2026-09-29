import {
  checkIntent,
  createAssistantResponse,
  type ModelOption
} from "../../excel-addin/src/api";
import capabilities from "../../../config/capabilities.json";
import type {
  AssistantResponse,
  AnalysisPlan,
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
import {
  buildWpsColumnMappingPlan,
  sliceMappingSource,
  type MappingSourceSheet
} from "./wpsColumnMappingPlan";
import type { WpsSourceSnapshot } from "./wpsPlanExecutor";
import { wpsReportedActionTypes } from "./wpsHostAdapter";
import { readWpsHostProbe, type WpsApplication } from "./wpsHost";
import {
  authorizeWpsSheets,
  listWpsWorksheets,
  readAuthorizedWpsSheets,
  readAuthorizedWpsSheetSchemas,
  type WpsSheetReadResult
} from "./wpsWorkbookReader";

export type WpsTurnResult =
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
      kind: "plan";
      plan: AnalysisPlan;
      /**
       * 计划是否在本端算出（`local_plan_rule` 那一支）。
       *
       * 为 `true` 时**不得**送 `/api/plans/validate`：`writeValues` 的验收条件
       * 里带着完整的值矩阵（`models.py` 的 `RangeEqualsCriterion.expected`），
       * 上传等于把刚才刻意留在本地的源值送出去 —— 单射映射可逆。
       *
       * 跳过服务端校验在这条路径上不丢东西：WPS 的执行验证是自洽的（写入前后
       * 指纹比对 + 回读，见 `wpsPlanExecutor.ts`），全程不读 `acceptanceCriteria`；
       * 而授权与存在性由 `prepareWpsPlan` 对着活宿主查，比服务端拿名字列表比
       * 更严。区域也已在服务端下发规则时规范化过。
       */
      computedLocally?: boolean;
      sourceSnapshots?: WpsSourceSnapshot[];
    };

export type WpsReadonlyTurnProgress = {
  stage: "reading" | "intent" | "calculating" | "answering";
  stepId?: string;
  title?: string;
  sheetName?: string;
  rowsRead?: number;
  totalRows?: number;
  toolName?: string;
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
  readSchemas: typeof readAuthorizedWpsSheetSchemas;
};

const DEFAULT_DEPENDENCIES: WpsReadonlyAgentDependencies = {
  checkIntent,
  createAssistantResponse,
  executeQuery: executeQueryTableData,
  readSheets: readAuthorizedWpsSheets,
  readSchemas: readAuthorizedWpsSheetSchemas
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
    headers: detectSheetFields(result.values).slice(0, capabilities.intentContext.maxFieldsPerSheet)
  }));
  const authorizedNames = new Set(sheets.map((sheet) => sheet.name));
  const activeWorksheet = authorizedNames.has(probe.activeSheetName)
    ? probe.activeSheetName
    : sheets[0]?.name ?? "";
  const selectedRange = authorizedNames.has(probe.activeSheetName)
    ? probe.selectionAddress
    : null;
  const workbook: WorkbookSnapshot = {
    name: metadata.name,
    capturedAt: new Date().toISOString(),
    activeWorksheet,
    selectedRange,
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
    activeWorksheet,
    selectedRange,
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

function finalizeResponse(response: AssistantResponse): WpsTurnResult {
  if (response.kind === "answer") {
    return {
      kind: "answer",
      message: response.message,
      resultContext: response.resultContext,
      provider: response.provider
    };
  }
  if (response.kind === "local_plan_rule") {
    // 这一支由 runWpsTurn 在拿得到授权工作表值的地方处理，不该走到这里。
    // 明确抛错而不是静默当成计划：真走到了说明调用顺序被改坏了。
    throw new Error("本地计算规则需要先读取授权工作表，不能在此处直接转成计划。");
  }
  return { kind: "plan", plan: response.plan };
}

export async function runWpsTurn(
  input: WpsReadonlyTurnInput,
  dependencies: WpsReadonlyAgentDependencies = DEFAULT_DEPENDENCIES
): Promise<WpsTurnResult> {
  const prompt = input.prompt.trim();
  if (prompt.length < 2) throw new Error("请输入至少两个字符的问题。");
  input.onProgress?.({ stage: "reading", stepId: "read-schemas", title: "读取已授权表结构", message: "只读取字段结构，不向意图模型发送原始数据行。" });
  const authorization = authorizeWpsSheets(input.application, input.selectedSheetIndexes);
  const schemas = await dependencies.readSchemas(input.application, authorization, {
    signal: input.signal
  });
  const { scope, workbook } = buildWpsReadonlyContexts(input.application, schemas);

  input.onProgress?.({ stage: "intent", stepId: "intent", title: "理解需求", message: "判断是直接回答、澄清、查询还是写入计划。" });
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
  if (intent.kind === "conversation") {
    return { kind: "answer", message: intent.message, provider: intent.provider, resultContext: null };
  }
  if (intent.kind === "clarification") return { kind: "clarification", response: intent };

  let dataResults: DataToolResult[] = [];
  let confirmedPrompt = intent.confirmedPrompt;
  if (intent.kind === "tool_request") {
    input.onProgress?.({ stage: "reading", stepId: "read-data", title: "读取授权数据", message: "只为已确定的本地查询读取完整数据。" });
    const results = await dependencies.readSheets(input.application, authorization, {
      signal: input.signal,
      onProgress: ({ sheet, rowsRead, totalRows }) => {
        input.onProgress?.({
          stage: "reading",
          stepId: `read-data:${sheet.name}`,
          title: `读取「${sheet.name}」`,
          sheetName: sheet.name,
          rowsRead,
          totalRows,
          message: `正在读取「${sheet.name}」：${rowsRead} / ${totalRows} 行`
        });
      }
    });
    input.onProgress?.({ stage: "calculating", stepId: "query", title: "本地确定性计算", toolName: "query_table", message: "在本机计算完整授权数据，不以模型快照代替全表。" });
    const querySheets = sheetsForRequest(intent.request, results, scope.activeWorksheet);
    dataResults = [dependencies.executeQuery(intent.request, querySheets)];
  }

  input.onProgress?.({ stage: "answering", stepId: "answer", title: "生成回答或计划", message: "根据意图结果与本地计算摘要组织回答。" });
  const hostVersion = String(input.application.Version ?? "").trim() || null;
  const response = await dependencies.createAssistantResponse(
    {
      turnId: intent.turnId,
      prompt: confirmedPrompt,
      workbook,
      dataResults,
      planningReadPolicy: "none",
      modelId: input.modelId || null,
      // 规划前能力收窄（方案 §4.3）：服务端按可信配置算上界，与本端自报求交。
      // 自报只能收窄，不能扩大 —— 服务端不采信这段内容作为授权凭据。
      hostId: "wps-et",
      hostVersion,
      selfReportedActions: wpsReportedActionTypes(hostVersion)
    },
    input.signal
  );
  if (response.kind === "local_plan_rule") {
    // 服务端只给了规则（它手上没有源值）。到这里才读源区域：上面 schema 扫描
    // 只取字段名，不含数据行。读出的值**只在本端用**，算完不回传。
    input.onProgress?.({ stage: "reading", stepId: "read-mapping-source", title: "读取授权源列", message: "映射源值只留在本机，不上传服务端计划校验。" });
    const sourceSheet = await readMappingSource(
      input,
      authorization,
      response.rule.sourceSheet,
      dependencies
    );
    input.onProgress?.({ stage: "calculating", stepId: "map-column", title: "计算目标列", toolName: "map_column", message: "在本机按映射规则生成目标值。" });
    return {
      kind: "plan",
      computedLocally: true,
      sourceSnapshots: [{
        sheetName: sourceSheet.name,
        range: response.rule.sourceRange,
        values: sliceMappingSource(sourceSheet, response.rule.sourceRange)
      }],
      plan: buildWpsColumnMappingPlan(response.rule, sourceSheet, {
        summary: response.summary,
        warnings: response.warnings,
        assumptions: response.assumptions
      })
    };
  }
  return finalizeResponse(response);
}

/** 读出规则指定的那一张授权表。表不在授权范围内时拒绝，不自动扩大范围。 */
async function readMappingSource(
  input: WpsReadonlyTurnInput,
  authorization: ReturnType<typeof authorizeWpsSheets>,
  sheetName: string,
  dependencies: WpsReadonlyAgentDependencies
): Promise<MappingSourceSheet> {
  const results = await dependencies.readSheets(input.application, authorization, {
    signal: input.signal,
    onProgress: ({ sheet, rowsRead, totalRows }) => {
      input.onProgress?.({
        stage: "reading",
        stepId: `read-mapping-source:${sheet.name}`,
        title: `读取「${sheet.name}」`,
        sheetName: sheet.name,
        rowsRead,
        totalRows,
        message: `正在读取「${sheet.name}」：${rowsRead} / ${totalRows} 行`
      });
    }
  });
  const match = results.find((result) => result.sheet.name === sheetName);
  if (!match) {
    throw new Error(
      `映射规则要读取「${sheetName}」，但它不在本次已授权的工作表里。`
    );
  }
  return {
    name: match.sheet.name,
    startRow: match.startRow,
    startColumn: match.startColumn,
    columnCount: match.columnCount,
    values: match.values
  };
}

/** Compatibility alias for extensions/tests importing the W2 name. */
export const runWpsReadonlyTurn = runWpsTurn;
export type WpsReadonlyTurnResult = WpsTurnResult;

export type { ModelOption };
