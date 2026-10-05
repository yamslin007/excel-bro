import {
  checkIntent,
  createAssistantResponse,
  type ModelOption
} from "../../excel-addin/src/api";
import capabilities from "../../../config/capabilities.json";
import {
  BlankPolicyNeededError,
  NeedsClarificationError,
  UnmappedCategoryError
} from "@excel-bro/core";
import type {
  AssistantResponse,
  AnalysisPlan,
  CellValue,
  ColumnMappingRule,
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

/**
 * 需要在写入前澄清的一轮结果。
 *
 * `columns` 非空 = **列映射澄清**，来自服务端下发的规则 `clarification`。
 * 两者共用同一张澄清卡，但归因与可选项不同 —— 卡片的选项只由 `reason`
 * 推导（`@excel-bro/core` 的 `columnMappingClarificationCard`），本端不猜测。
 *
 * `columns` 为空 = 既有的需求澄清（模型说不清要做什么，走 `response`）。
 */
export type WpsTurnClarification = {
  /** 列映射澄清：服务端已判定规则缺参数/缺类别，**未读源值、未计算**。 */
  columns?: {
    clarification: NonNullable<
      Extract<ColumnMappingRule, { clarification?: unknown }>["clarification"]
    >;
    scopeLabel: string;
    summary: string;
  };
  /** 需求澄清（模型层）。 */
  response?: Extract<IntentCheckResponse, { kind: "clarification" }>;
  /** 本地读源列后发现的未覆盖值；只允许留在面板内存，不得进历史或请求。 */
  localMapping?: {
    rule: ColumnMappingRule;
    context: { summary: string; warnings: string[]; assumptions: string[] };
    values: readonly CellValue[];
    reason: "unmapped" | "separator_or_mapping" | "blank";
    /**
     * `reason === "blank"` 时**真实读到**的空白格数量。
     *
     * 它来自一次真实读取。没有读证据就不许声称"源区域有空白"。
     */
    blankCount?: number;
    /** 已声明拆分后缺的是"项"，不是整格；不能暗用映射声明顺序决定优先级。 */
    allowExactMapping: boolean;
  };
};

export type WpsTurnResult =
  | {
      kind: "answer";
      message: string;
      resultContext?: ResultContext | null;
      provider: "model" | "local" | "system";
      outcome?: "completed" | "failed";
      /**
       * 服务端这次规划的身份。**断线后要先拿它查真实状态**，
       * 有有效检查点才继续 —— 不能靠"重新发一遍需求"冒充续跑。
       */
      turnId?: string | null;
    }
  | {
      kind: "clarification";
      clarification: WpsTurnClarification;
      turnId?: string | null;
      /**
       * 结构扫描**真实读到的表头行**（每张已授权表一条，顺序与授权序一致）。
       *
       * 为什么要把它带出来：澄清卡需要一个"这份结构还是不是原来那份"的判据，
       * 而它必须落在**表头单元格的真实内容**上 —— 工作表名只是身份，表头是语义。
       * 之前面板拿 `listWpsWorksheets().sheets.map(([index, name]))` 冒充表头，
       * 于是"表头判据"退化成身份判据的副本，两条判据吃同一份数据、互相遮蔽。
       *
       * 只含表头行（`schemaOnly` 扫描天然只取第 1 行），数据行不进这里，
       * 因此它落到澄清卡上也不会把单元格值写进对话历史。
       */
      headerRows?: string[][];
    }
  | {
      kind: "plan";
      plan: AnalysisPlan;
      turnId?: string | null;
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

/**
 * 结构扫描**实际取到的表头行**（只取第 1 行）。
 *
 * `readAuthorizedWpsSheetSchemas` 的取值面已经被收窄到表头行，所以 `values[0]`
 * 就是真实表头；把它取出来交给澄清卡做指纹，才能让"表头判据"真的看表头内容，
 * 而不是去看工作表名（那是身份判据的事）。
 *
 * 之所以只取第 1 行、并且把 `null`/`undefined` 一律折成 `""`：
 * 指纹会被持久化进对话历史，必须稳定、可比较，且不含数据行。
 */
function sheetHeaderRows(results: WpsSheetReadResult[]): string[][] {
  return results.map((result) => {
    const header = result.values[0] ?? [];
    return header.map((cell) => (cell == null ? "" : String(cell)));
  });
}

/** 与 core/Python 的来源值键化一致；用户补的整格映射必须命中同一个键。 */
export function wpsMappingSourceKey(value: CellValue): string {
  if (value === true) return "True";
  if (value === false) return "False";
  return String(value);
}

type LocalMappingContext = { summary: string; warnings: string[]; assumptions: string[] };

function localMappingResult(
  rule: ColumnMappingRule,
  sourceSheet: MappingSourceSheet,
  context: LocalMappingContext,
  headerRows?: string[][]
): WpsTurnResult {
  try {
    const plan = buildWpsColumnMappingPlan(rule, sourceSheet, context);
    return {
      kind: "plan",
      computedLocally: true,
      sourceSnapshots: [{
        sheetName: sourceSheet.name,
        range: rule.sourceRange,
        values: sliceMappingSource(sourceSheet, rule.sourceRange)
      }],
      plan
    };
  } catch (error) {
    /*
     * 空白策略未定，且**这一轮真实读到了空白格**。
     *
     * 停在这是对的：它发生在读源值之后，不是"还没读就先问"。四个选项由
     * 界面给出，本端不替用户挑 —— 静默挑一个默认标签（曾经就是这么做的，
     * 空白一律写成「未分类」）会让用户在单元格层面看不出那是实现替他选的。
     */
    if (error instanceof BlankPolicyNeededError) {
      return {
        kind: "clarification",
        headerRows,
        clarification: {
          localMapping: {
            rule,
            context,
            // **不带任何来源值**：空白格没有内容可带，而"有几格空白"这个
            // 计数足够让用户判断规模。带值会违反"源值只留本端"。
            values: [],
            reason: "blank",
            blankCount: error.blankCount,
            allowExactMapping: false
          }
        }
      };
    }
    if (!(error instanceof UnmappedCategoryError) && !(error instanceof NeedsClarificationError)) throw error;
    return {
      kind: "clarification",
      headerRows,
      clarification: {
        localMapping: {
          rule,
          context,
          values: error instanceof UnmappedCategoryError ? error.values : error.cells,
          reason: error instanceof UnmappedCategoryError ? "unmapped" : "separator_or_mapping",
          allowExactMapping: error instanceof NeedsClarificationError || !rule.delimiters?.length
        }
      }
    };
  }
}

/** 用户在本机补全整格映射后重新读授权源表；不重发意图/规则或单元格值。 */
export async function resumeWpsLocalMapping(input: {
  application: WpsApplication;
  selectedSheetIndexes: readonly number[];
  rule: ColumnMappingRule;
  context: LocalMappingContext;
  additions: readonly { source: string; target: string }[];
  signal?: AbortSignal;
  onProgress?: WpsReadonlyTurnInput["onProgress"];
}, dependencies: WpsReadonlyAgentDependencies = DEFAULT_DEPENDENCIES): Promise<WpsTurnResult> {
  if (input.additions.some((item) => !item.source || !item.target.trim())) {
    throw new Error("请为每个未覆盖的来源值明确填写目标值。");
  }
  const existing = new Set(input.rule.mapping.map((item) => item.source));
  if (input.additions.some((item) => existing.has(item.source))) {
    throw new Error("补充映射与已有规则冲突，请重新生成预览。");
  }
  const authorization = authorizeWpsSheets(input.application, input.selectedSheetIndexes);
  const sourceSheet = await readMappingSource(
    { ...input, prompt: "本地映射复核" }, authorization, input.rule.sourceSheet, dependencies
  );
  const rule: ColumnMappingRule = {
    ...input.rule,
    mapping: [...input.rule.mapping, ...input.additions]
  };
  return localMappingResult(rule, sourceSheet, input.context);
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
  if (response.kind === "clarification") {
    return { kind: "clarification", clarification: { response }, turnId: response.turnId };
  }
  if (response.kind === "answer") {
    return {
      kind: "answer",
      message: response.message,
      resultContext: response.resultContext,
      provider: response.provider,
      outcome: response.outcome,
      turnId: response.turnId
    };
  }
  if (response.kind === "local_plan_rule") {
    // 这一支由 runWpsTurn 在拿得到授权工作表值的地方处理，不该走到这里。
    // 明确抛错而不是静默当成计划：真走到了说明调用顺序被改坏了。
    throw new Error("本地计算规则需要先读取授权工作表，不能在此处直接转成计划。");
  }
  return { kind: "plan", plan: response.plan, turnId: response.turnId };
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
  if (intent.kind === "clarification") {
    return {
      kind: "clarification",
      clarification: { response: intent },
      headerRows: sheetHeaderRows(schemas)
    };
  }

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
  if (response.kind === "clarification") {
    return {
      kind: "clarification", clarification: { response },
      headerRows: sheetHeaderRows(schemas)
    };
  }
  if (response.kind === "local_plan_rule") {
    // 服务端已经判定"这条规则还缺参数/缺类别，算不出正确答案"。
    // 这时**不要**读源值、更不要算：与 Excel 端同一判据 —— 算出来的矩阵必然
    // 是猜的，而猜错在单元格层面看不出来。把归因接进澄清卡，让用户补完再跑。
    //
    // 位置很关键：这段必须在 `readMappingSource` **之前**。放后面等于先读了
    // 整片源列才告诉用户"其实缺映射" —— 白读一次源值。
    if (response.rule.clarification) {
      return {
        kind: "clarification",
        clarification: {
          columns: {
            clarification: response.rule.clarification,
            scopeLabel: `${response.rule.sourceSheet}!${response.rule.sourceRange}`,
            summary: response.summary
          }
        },
        headerRows: sheetHeaderRows(schemas)
      };
    }
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
    const mapped = localMappingResult(response.rule, sourceSheet, {
      summary: response.summary,
      warnings: response.warnings,
      assumptions: response.assumptions
    }, sheetHeaderRows(schemas));
    // 带上这次规划的身份：断线后要先拿它查真实状态，而不是重新发一遍需求。
    return { ...mapped, turnId: response.turnId } as WpsTurnResult;
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
