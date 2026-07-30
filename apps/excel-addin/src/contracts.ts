export type CellValue = string | number | boolean | null;

export interface WorksheetSnapshot {
  name: string;
  sourceFileId?: string | null;
  sourceSheetId?: string | null;
  sourceFile?: string | null;
  sourceSheet?: string | null;
  usedRange: string | null;
  rowCount: number;
  columnCount: number;
  headers: CellValue[];
  previewRows?: CellValue[][];
  dataRows: CellValue[][];
  displayRows?: CellValue[][];
  truncated: boolean;
}

export interface WorkbookSnapshot {
  name: string;
  capturedAt: string;
  activeWorksheet: string;
  selectedRange?: string | null;
  sourceFingerprint?: string | null;
  sourceFingerprintSheets?: string[];
  worksheets: WorksheetSnapshot[];
}

export type ExcelAction =
  | { type: "createWorksheet"; sheet: string }
  | {
      type: "writeTable";
      sheet: string;
      startCell: string;
      headers: CellValue[];
      rows: CellValue[][];
    }
  | {
      type: "writeValues";
      sheet: string;
      range: string;
      values: CellValue[][];
    }
  | {
      type: "setFill";
      sheet: string;
      range: string;
      color: string;
    }
  | {
      type: "setFont";
      sheet: string;
      range: string;
      bold?: boolean;
      color?: string;
    }
  | { type: "autofit"; sheet: string; range: string }
  | { type: "activateWorksheet"; sheet: string }
  | { type: "deleteWorksheet"; sheet: string }
  | {
      type: "clearRange";
      sheet: string;
      range: string;
      applyTo: "all" | "contents" | "formats" | "hyperlinks";
    }
  | {
      type: "insertRange";
      sheet: string;
      range: string;
      shift: "down" | "right";
    }
  | {
      type: "deleteRange";
      sheet: string;
      range: string;
      shift: "up" | "left";
    }
  | {
      type: "copyRange";
      sheet: string;
      sourceSheet: string;
      sourceRange: string;
      targetRange: string;
      copyType: "all" | "values" | "formulas" | "formats" | "link";
      skipBlanks: boolean;
      transpose: boolean;
    }
  | { type: "writeFormulas"; sheet: string; range: string; formulas: string[][] }
  | {
      type: "sortRange";
      sheet: string;
      range: string;
      keys: Array<{ column: number; ascending: boolean }>;
      hasHeaders: boolean;
    }
  | {
      type: "filterRange";
      sheet: string;
      range: string;
      column: number;
      values: CellValue[];
    }
  | { type: "clearFilter"; sheet: string }
  | {
      type: "setDataValidation";
      sheet: string;
      range: string;
      validationType:
        | "list"
        | "wholeNumber"
        | "decimal"
        | "date"
        | "textLength"
        | "custom";
      values: CellValue[];
      formula1?: string | number | null;
      formula2?: string | number | null;
      operator: string;
      allowBlank: boolean;
      prompt?: string | null;
      errorMessage?: string | null;
    }
  | {
      type: "setConditionalFormat";
      sheet: string;
      range: string;
      ruleType: "cellValue" | "custom" | "colorScale";
      operator?: string | null;
      formula1?: string | number | null;
      formula2?: string | number | null;
      color?: string | null;
      minColor?: string | null;
      midColor?: string | null;
      maxColor?: string | null;
    }
  | { type: "setNumberFormat"; sheet: string; range: string; formatCode: string }
  | {
      type: "setBorders";
      sheet: string;
      range: string;
      sides: Array<
        "top" | "bottom" | "left" | "right" | "insideHorizontal" | "insideVertical"
      >;
      style: string;
      color: string;
      weight: string;
    }
  | {
      type: "setAlignment";
      sheet: string;
      range: string;
      horizontal?: string | null;
      vertical?: string | null;
      wrapText?: boolean | null;
    }
  | { type: "mergeCells"; sheet: string; range: string; across: boolean }
  | { type: "unmergeCells"; sheet: string; range: string }
  | {
      type: "resizeRange";
      sheet: string;
      range: string;
      rowHeight?: number | null;
      columnWidth?: number | null;
    }
  | { type: "freezePanes"; sheet: string; rows: number; columns: number }
  | {
      type: "setHyperlink";
      sheet: string;
      range: string;
      address: string;
      text?: string | null;
      screenTip?: string | null;
    }
  | { type: "addComment"; sheet: string; cell: string; text: string }
  | { type: "addNote"; sheet: string; cell: string; text: string }
  | {
      type: "createTable";
      sheet: string;
      range: string;
      name?: string | null;
      hasHeaders: boolean;
      style?: string | null;
    }
  | {
      type: "createChart";
      sheet: string;
      sourceRange: string;
      chartType: string;
      title?: string | null;
      targetRange?: string | null;
    }
  | {
      type: "createPivotTable";
      sheet: string;
      sourceSheet: string;
      sourceRange: string;
      name: string;
      destinationCell: string;
      rowFields: string[];
      columnFields: string[];
      valueFields: Array<{
        field: string;
        aggregation: "sum" | "count" | "average" | "max" | "min";
      }>;
    }
  | {
      type: "splitGroupAggregate";
      sheet: string;
      sourceRange?: string | null;
      splitBy: string;
      groupBy: string[];
      metrics: Array<{
        operation: "countRows" | "countNonBlank" | "sum";
        field?: string | null;
        outputName: string;
        ratioOutputName?: string | null;
      }>;
      includeBlankSplitValues: boolean;
      existingSheetPolicy: "rename" | "replace" | "skip";
      maxOutputSheets: number;
    }
  | {
      type: "addNamedRange";
      sheet: string;
      name: string;
      range: string;
      comment?: string | null;
    }
  | {
      type: "addImage";
      sheet: string;
      base64: string;
      targetRange: string;
      name?: string | null;
    }
  | {
      type: "addShape";
      sheet: string;
      shapeType: string;
      targetRange: string;
      text?: string | null;
      fillColor?: string | null;
    };

export type VerificationCriterion =
  | { type: "worksheetExists"; sheet: string }
  | { type: "worksheetMissing"; sheet: string }
  | {
      type: "rangeEquals";
      sheet: string;
      range: string;
      expected: CellValue[][];
    }
  | { type: "rangeEmpty"; sheet: string; range: string }
  | {
      type: "formulasEqual";
      sheet: string;
      range: string;
      expected: string[][];
    }
  | {
      type: "rangeSorted";
      sheet: string;
      range: string;
      keys: Array<{ column: number; ascending: boolean }>;
      hasHeaders: boolean;
    }
  | {
      type: "filterApplied";
      sheet: string;
      range: string;
      column: number;
      values: CellValue[];
    }
  | { type: "filterCleared"; sheet: string }
  | {
      type: "tableExists";
      sheet: string;
      range: string;
      name?: string | null;
      hasHeaders: boolean;
    }
  | {
      type: "rangeFormatMatches";
      sheet: string;
      range: string;
      fillColor?: string | null;
      bold?: boolean | null;
      fontColor?: string | null;
      numberFormat?: string | null;
      horizontal?: string | null;
      vertical?: string | null;
      wrapText?: boolean | null;
      rowHeight?: number | null;
      columnWidth?: number | null;
    }
  | {
      type: "bordersMatch";
      sheet: string;
      range: string;
      sides: Array<
        "top" | "bottom" | "left" | "right" | "insideHorizontal" | "insideVertical"
      >;
      style: string;
      color: string;
      weight: string;
    }
  | {
      type: "dataValidationMatches";
      sheet: string;
      range: string;
      validationType:
        | "list"
        | "wholeNumber"
        | "decimal"
        | "date"
        | "textLength"
        | "custom";
      values: CellValue[];
      formula1?: string | number | null;
      formula2?: string | number | null;
      operator: string;
      allowBlank: boolean;
      prompt?: string | null;
      errorMessage?: string | null;
    }
  | {
      type: "freezePanesMatches";
      sheet: string;
      rows: number;
      columns: number;
    }
  | {
      type: "chartExists";
      sheet: string;
      name?: string | null;
      chartType: string;
      sourceRange: string;
      title?: string | null;
      targetRange?: string | null;
    }
  | {
      type: "pivotTableExists";
      sheet: string;
      sourceSheet: string;
      sourceRange: string;
      name: string;
      destinationCell: string;
      rowFields: string[];
      columnFields: string[];
      valueFields: Array<{
        field: string;
        aggregation: "sum" | "count" | "average" | "max" | "min";
      }>;
    };

export interface AnalysisPlan {
  id: string;
  title: string;
  summary: string;
  sourceFingerprint?: string | null;
  sourceFingerprintSheets?: string[];
  assumptions: string[];
  warnings: string[];
  actions: ExcelAction[];
  acceptanceCriteria?: VerificationCriterion[];
}

export interface ResultContext {
  kind: "table";
  title: string;
  headers: string[];
  rows: CellValue[][];
  primaryValueColumn?: number | null;
  sourceSheets: string[];
  warnings: string[];
}

export interface ImageAttachment {
  name: string;
  mediaType: "image/png" | "image/jpeg" | "image/webp";
  data: string;
}

export interface IntentSheetContext {
  name: string;
  usedRange: string | null;
  rowCount: number;
  columnCount: number;
  headers: CellValue[];
}

export interface IntentScopeContext {
  workbookName: string;
  sourceMode: "workbook" | "folder";
  selectionMode: "auto" | "manual" | "folder";
  activeWorksheet: string;
  selectedRange?: string | null;
  totalWorksheetCount: number;
  sheets: IntentSheetContext[];
}

export interface IntentCheckRequest {
  turnId?: string | null;
  prompt: string;
  scope: IntentScopeContext;
  imageCount?: number;
  intentConfirmed?: boolean;
  clarificationRound?: number;
  conversation?: Array<{
    role: "user" | "assistant";
    text: string;
  }>;
  priorIntent?: IntentMemory | null;
  priorResult?: ResultContext | null;
  toolFailure?: {
    code: string;
    message: string;
    retryable: boolean;
    availableFields?: string[];
    request: DataToolRequest;
  } | null;
  modelId?: string | null;
}

export interface IntentOption {
  id: string;
  label: string;
  description: string;
  resolution: string;
  action?: "resolve" | "editScope";
}

export type DataFilterOperator =
  | "equals"
  | "notEquals"
  | "contains"
  | "greaterThan"
  | "greaterThanOrEqual"
  | "lessThan"
  | "lessThanOrEqual"
  | "isBlank"
  | "isNotBlank";

export interface DataFilter {
  field: string;
  operator: DataFilterOperator;
  value?: CellValue;
}

export interface DataMetric {
  operation:
    | "countRows"
    | "countDistinct"
    | "sum"
    | "average"
    | "min"
    | "max";
  field?: string | null;
  outputName: string;
  ratioOutputName?: string | null;
}

export interface QueryTableArguments {
  mode: "rows" | "aggregate" | "profile";
  scope?: "selected" | "active";
  fields?: string[];
  filters?: DataFilter[];
  groupBy?: string[];
  metrics?: DataMetric[];
  combine?: {
    mode: "union" | "deduplicate" | "join";
    deduplicateBy?: string[];
    leftSourceSheetId?: string | null;
    rightSourceSheetId?: string | null;
    leftKey?: string | null;
    rightKey?: string | null;
    joinHow?: "inner" | "left" | "right" | "outer";
  } | null;
  profileField?: string | null;
  sortBy?: string | null;
  sortDirection?: "asc" | "desc";
  limit?: number;
}

export interface DataToolRequest {
  id: string;
  tool: "query_table";
  arguments: QueryTableArguments;
}

export interface DeterministicQueryTemplate {
  id: string;
  name: string;
  description: string;
  sourceMode: "workbook" | "folder";
  request: DataToolRequest;
  sourceSheetNames: string[];
  sourceSheetIds: string[];
  expectedHeaders: string[];
}

export interface IntentMemory {
  confirmedPrompt: string;
  toolRequest?: DataToolRequest | null;
}

export interface ModelSettings {
  baseUrl: string | null;
  defaultModel: string | null;
  apiKeyConfigured: boolean;
  apiKeyHint: string | null;
  connections: ManagedModelConnection[];
}

export interface ManagedModelConnection {
  id: string;
  catalogModelId: string;
  label: string;
  baseUrl: string;
  modelId: string;
  supportsVision: boolean;
  apiKeyConfigured: boolean;
  apiKeyHint: string | null;
}

export interface UpdateModelSettingsRequest {
  apiKey: string;
}

export interface UpsertModelConnectionRequest {
  id?: string | null;
  label: string;
  baseUrl: string;
  modelId: string;
  apiKey?: string | null;
  clearApiKey?: boolean;
  supportsVision: boolean;
}

export interface TestModelConnectionResponse {
  ok: boolean;
  message: string;
}

export function assertTestModelConnectionResponse(
  value: unknown
): asserts value is TestModelConnectionResponse {
  if (!value || typeof value !== "object") {
    throw new Error("服务返回的连接测试结果不是有效对象");
  }
  const result = value as Partial<TestModelConnectionResponse>;
  if (result.ok !== true || typeof result.message !== "string") {
    throw new Error("服务返回的连接测试结果缺少必要字段");
  }
}

export function assertModelSettings(
  value: unknown
): asserts value is ModelSettings {
  if (!value || typeof value !== "object") {
    throw new Error("服务返回的模型设置不是有效对象");
  }
  const settings = value as Partial<ModelSettings>;
  if (
    (settings.baseUrl !== null && typeof settings.baseUrl !== "string") ||
    (settings.defaultModel !== null &&
      typeof settings.defaultModel !== "string") ||
    typeof settings.apiKeyConfigured !== "boolean" ||
    (settings.apiKeyHint !== null &&
      typeof settings.apiKeyHint !== "string") ||
    !Array.isArray(settings.connections) ||
    settings.connections.some(
      (connection) =>
        !connection ||
        typeof connection.id !== "string" ||
        typeof connection.catalogModelId !== "string" ||
        typeof connection.label !== "string" ||
        typeof connection.baseUrl !== "string" ||
        typeof connection.modelId !== "string" ||
        typeof connection.supportsVision !== "boolean" ||
        typeof connection.apiKeyConfigured !== "boolean" ||
        (connection.apiKeyHint !== null &&
          typeof connection.apiKeyHint !== "string")
    )
  ) {
    throw new Error("服务返回的模型设置无效");
  }
}

export interface DataToolResult {
  requestId: string;
  tool: "query_table";
  title: string;
  headers: CellValue[];
  rows: CellValue[][];
  sourceSheets: string[];
  scannedRows: number;
  complete: boolean;
  calculation: string;
  warnings: string[];
}

export function assertDataToolResult(
  value: unknown
): asserts value is DataToolResult {
  const item = value as Partial<DataToolResult>;
  if (
    !item ||
    typeof item !== "object" ||
    typeof item.requestId !== "string" ||
    item.tool !== "query_table" ||
    typeof item.title !== "string" ||
    !Array.isArray(item.headers) ||
    !item.headers.every((header) => typeof header === "string") ||
    !Array.isArray(item.rows) ||
    !item.rows.every((row) => Array.isArray(row)) ||
    !Array.isArray(item.sourceSheets) ||
    !item.sourceSheets.every((sheet) => typeof sheet === "string") ||
    typeof item.scannedRows !== "number" ||
    typeof item.complete !== "boolean" ||
    typeof item.calculation !== "string" ||
    !Array.isArray(item.warnings) ||
    !item.warnings.every((warning) => typeof warning === "string")
  ) {
    throw new Error("本地服务返回了无效的数据工具结果");
  }
}

export interface IntentClarification {
  id: string;
  summary: string;
  question: string;
  reason: string;
  scopeLabel: string;
  options: IntentOption[];
}

export type IntentCheckResponse =
  | {
      kind: "proceed";
      summary: string;
      confirmedPrompt: string;
      provider: "model" | "local";
      turnId?: string | null;
    }
  | {
      kind: "clarification";
      clarification: IntentClarification;
      provider: "model" | "local";
      turnId?: string | null;
    }
  | {
      kind: "tool_request";
      summary: string;
      confirmedPrompt: string;
      request: DataToolRequest;
      provider: "model" | "local";
      turnId?: string | null;
    };

export interface PlanRequest {
  turnId?: string | null;
  prompt: string;
  workbook: WorkbookSnapshot;
  lastResult?: ResultContext | null;
  images?: ImageAttachment[];
  dataResults?: DataToolResult[];
  modelId?: string | null;
}

export function assertIntentCheckResponse(
  value: unknown
): asserts value is IntentCheckResponse {
  if (!value || typeof value !== "object") {
    throw new Error("服务返回的意图判断不是有效对象");
  }
  const response = value as Partial<IntentCheckResponse>;
  if (response.kind === "proceed") {
    if (
      !("confirmedPrompt" in response) ||
      typeof response.confirmedPrompt !== "string"
    ) {
      throw new Error("服务返回的明确需求无效");
    }
    return;
  }
  if (response.kind === "tool_request") {
    const request = response.request;
    if (
      !request ||
      request.tool !== "query_table" ||
      typeof request.id !== "string" ||
      !request.arguments ||
      !["rows", "aggregate", "profile"].includes(request.arguments.mode)
    ) {
      throw new Error("服务返回的数据工具请求无效");
    }
    return;
  }
  if (response.kind !== "clarification") {
    throw new Error("服务返回了未知的意图判断类型");
  }
  const clarification = response.clarification;
  if (
    !clarification ||
    typeof clarification.question !== "string" ||
    !Array.isArray(clarification.options) ||
    clarification.options.length < 2 ||
    clarification.options.some(
      (option) =>
        !option ||
        typeof option.id !== "string" ||
        typeof option.label !== "string" ||
        typeof option.description !== "string" ||
        typeof option.resolution !== "string"
    )
  ) {
    throw new Error("服务返回的确认问题无效");
  }
}

export interface FolderWorksheetInfo {
  name: string;
  rowCount: number;
  columnCount: number;
}

export interface FolderFileInfo {
  id: string;
  name: string;
  relativePath: string;
  worksheets: FolderWorksheetInfo[];
  error?: string | null;
}

export interface FolderCatalog {
  sessionId: string;
  folderName: string;
  folderPath: string;
  files: FolderFileInfo[];
  totalFiles: number;
  truncated: boolean;
  expiresAt: string;
}

export interface FolderSelection {
  fileId: string;
  sheets: string[];
}

export interface FolderExecuteResult {
  filesModified: string[];
  backups: string[];
  actionResults: ActionExecutionResult[];
  verification: VerificationReport;
  executionMs: number;
  verificationMs: number;
}

export interface ActionExecutionResult {
  index: number;
  type: ExcelAction["type"];
  sheet: string;
  status: "succeeded" | "failed" | "not_run";
  message?: string | null;
}

export interface VerificationCheck {
  criterion: VerificationCriterion;
  passed: boolean;
  message: string;
  actual?: CellValue[][] | null;
}

export interface UnverifiedAction {
  index: number;
  type: ExcelAction["type"];
  sheet: string;
  message: string;
}

export interface VerificationReport {
  status: "verified" | "executed_unverified" | "failed";
  passed: boolean;
  checks: VerificationCheck[];
  unverifiedActions: UnverifiedAction[];
}

export interface PlanExecutionResult {
  actionResults: ActionExecutionResult[];
  verification: VerificationReport;
  undoSnapshot?: ExecutionUndoSnapshot | null;
  executionMs: number;
  verificationMs: number;
}

export interface UndoRangeSnapshot {
  actionIndex: number;
  sheet: string;
  range: string;
  formulas: CellValue[][];
  numberFormat: string[][];
  fillColor: string;
  fontBold: boolean;
  fontColor: string;
}

export interface ExecutionUndoSnapshot {
  planId: string;
  capturedAt: string;
  ranges: UndoRangeSnapshot[];
}

export type AssistantResponse =
  | {
      kind: "answer";
      message: string;
      provider: "model" | "local";
      resultContext?: ResultContext | null;
      turnId?: string | null;
    }
  | {
      kind: "plan";
      plan: AnalysisPlan;
      provider: "model" | "local";
      turnId?: string | null;
    };

export interface TurnStepEvent {
  phase: "planning";
  title: string;
  detail?: string | null;
  completedStep?: string | null;
}

const allowedActionTypes = new Set<ExcelAction["type"]>([
  "createWorksheet",
  "writeTable",
  "writeValues",
  "setFill",
  "setFont",
  "autofit",
  "activateWorksheet",
  "deleteWorksheet",
  "clearRange",
  "insertRange",
  "deleteRange",
  "copyRange",
  "writeFormulas",
  "sortRange",
  "filterRange",
  "clearFilter",
  "setDataValidation",
  "setConditionalFormat",
  "setNumberFormat",
  "setBorders",
  "setAlignment",
  "mergeCells",
  "unmergeCells",
  "resizeRange",
  "freezePanes",
  "setHyperlink",
  "addComment",
  "addNote",
  "createTable",
  "createChart",
  "createPivotTable",
  "splitGroupAggregate",
  "addNamedRange",
  "addImage",
  "addShape"
]);
const allowedCriterionTypes = new Set<VerificationCriterion["type"]>([
  "worksheetExists",
  "worksheetMissing",
  "rangeEquals",
  "rangeEmpty",
  "formulasEqual",
  "rangeSorted",
  "filterApplied",
  "filterCleared",
  "tableExists",
  "rangeFormatMatches",
  "bordersMatch",
  "dataValidationMatches",
  "freezePanesMatches",
  "chartExists",
  "pivotTableExists"
]);

function rangeDimensions(address: string): [number, number] | null {
  const match =
    /^\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$/.exec(
      address.trim()
    );
  if (!match) return null;
  const columnNumber = (name: string) =>
    [...name.toUpperCase()].reduce(
      (result, character) => result * 26 + character.charCodeAt(0) - 64,
      0
    );
  const startColumn = columnNumber(match[1]);
  const startRow = Number(match[2]);
  const endColumn = columnNumber(match[3] ?? match[1]);
  const endRow = Number(match[4] ?? match[2]);
  if (endColumn < startColumn || endRow < startRow) return null;
  return [endRow - startRow + 1, endColumn - startColumn + 1];
}

export function assertAssistantResponse(
  value: unknown
): asserts value is AssistantResponse {
  if (!value || typeof value !== "object") {
    throw new Error("服务返回的内容不是有效对象");
  }

  const response = value as Partial<AssistantResponse>;
  if (response.kind === "answer") {
    if (!("message" in response) || typeof response.message !== "string") {
      throw new Error("服务返回的回答缺少 message");
    }
    if (
      response.resultContext !== undefined &&
      response.resultContext !== null &&
      (response.resultContext.kind !== "table" ||
        !Array.isArray(response.resultContext.headers) ||
        !Array.isArray(response.resultContext.rows))
    ) {
      throw new Error("服务返回的 resultContext 无效");
    }
    return;
  }

  if (response.kind !== "plan") {
    throw new Error("服务返回了未知响应类型");
  }
  if (!response.plan || !Array.isArray(response.plan.actions)) {
    throw new Error("服务返回的计划缺少 actions");
  }

  for (const action of response.plan.actions as Array<{ type?: string }>) {
    if (!action.type || !allowedActionTypes.has(action.type as ExcelAction["type"])) {
      throw new Error(`计划包含未授权动作：${action.type ?? "unknown"}`);
    }
  }

  const criteria = response.plan.acceptanceCriteria;
  if (criteria !== undefined && !Array.isArray(criteria)) {
    throw new Error("计划的验收条件不是有效数组");
  }
  const actionSheets = new Set(response.plan.actions.map((action) => action.sheet));
  for (const criterion of criteria ?? []) {
    if (
      !criterion.type ||
      !allowedCriterionTypes.has(criterion.type as VerificationCriterion["type"])
    ) {
      throw new Error(`计划包含未知验收条件：${criterion.type ?? "unknown"}`);
    }
    if (!actionSheets.has(criterion.sheet)) {
      throw new Error("验收条件只能检查计划实际操作的工作表");
    }
    if (
      [
        "rangeSorted",
        "filterApplied",
        "tableExists",
        "rangeFormatMatches",
        "bordersMatch",
        "dataValidationMatches"
      ].includes(
        criterion.type
      ) &&
      typeof (criterion as { range?: unknown }).range !== "string"
    ) {
      throw new Error(`${criterion.type} 缺少有效范围`);
    }
    if (
      criterion.type === "rangeSorted" &&
      (!Array.isArray(criterion.keys) ||
        criterion.keys.length === 0 ||
        typeof criterion.hasHeaders !== "boolean" ||
        criterion.keys.some(
          (key) =>
            !Number.isInteger(key?.column) ||
            key.column < 0 ||
            typeof key.ascending !== "boolean"
        ))
    ) {
      throw new Error("rangeSorted 的排序键无效");
    }
    if (
      criterion.type === "filterApplied" &&
      (!Number.isInteger(criterion.column) ||
        criterion.column < 0 ||
        !Array.isArray(criterion.values) ||
        criterion.values.length === 0)
    ) {
      throw new Error("filterApplied 的筛选条件无效");
    }
    if (
      criterion.type === "tableExists" &&
      (typeof criterion.hasHeaders !== "boolean" ||
        (criterion.name !== undefined &&
          criterion.name !== null &&
          typeof criterion.name !== "string"))
    ) {
      throw new Error("tableExists 的表格条件无效");
    }
    if (criterion.type === "rangeFormatMatches") {
      const expectedProperties = [
        criterion.fillColor,
        criterion.bold,
        criterion.fontColor,
        criterion.numberFormat,
        criterion.horizontal,
        criterion.vertical,
        criterion.wrapText,
        criterion.rowHeight,
        criterion.columnWidth
      ];
      if (expectedProperties.every((property) => property == null)) {
        throw new Error("rangeFormatMatches 缺少格式属性");
      }
    }
    if (
      criterion.type === "bordersMatch" &&
      (!Array.isArray(criterion.sides) ||
        criterion.sides.length === 0 ||
        typeof criterion.style !== "string" ||
        typeof criterion.color !== "string" ||
        typeof criterion.weight !== "string")
    ) {
      throw new Error("bordersMatch 的边框条件无效");
    }
    if (
      criterion.type === "dataValidationMatches" &&
      (typeof criterion.validationType !== "string" ||
        !Array.isArray(criterion.values) ||
        typeof criterion.operator !== "string" ||
        typeof criterion.allowBlank !== "boolean")
    ) {
      throw new Error("dataValidationMatches 的验证条件无效");
    }
    if (
      criterion.type === "freezePanesMatches" &&
      (!Number.isInteger(criterion.rows) ||
        criterion.rows < 0 ||
        !Number.isInteger(criterion.columns) ||
        criterion.columns < 0)
    ) {
      throw new Error("freezePanesMatches 的冻结位置无效");
    }
    if (
      criterion.type === "chartExists" &&
      (typeof criterion.chartType !== "string" ||
        typeof criterion.sourceRange !== "string" ||
        (criterion.name != null && typeof criterion.name !== "string") ||
        (criterion.targetRange != null &&
          typeof criterion.targetRange !== "string"))
    ) {
      throw new Error("chartExists 的图表条件无效");
    }
    if (
      criterion.type === "pivotTableExists" &&
      (typeof criterion.name !== "string" ||
        typeof criterion.sourceSheet !== "string" ||
        typeof criterion.sourceRange !== "string" ||
        typeof criterion.destinationCell !== "string" ||
        !Array.isArray(criterion.rowFields) ||
        !Array.isArray(criterion.columnFields) ||
        !Array.isArray(criterion.valueFields))
    ) {
      throw new Error("pivotTableExists 的数据透视表条件无效");
    }
    if (
      criterion.type === "rangeEquals" ||
      criterion.type === "formulasEqual"
    ) {
      const expected = criterion.expected;
      const width = Array.isArray(expected?.[0]) ? expected[0].length : 0;
      const rectangular =
        Array.isArray(expected) &&
        expected.length > 0 &&
        expected.length <= 500 &&
        width > 0 &&
        width <= 50 &&
        expected.every((row) => Array.isArray(row) && row.length === width);
      const dimensions =
        typeof criterion.range === "string"
          ? rangeDimensions(criterion.range)
          : null;
      if (
        !rectangular ||
        !dimensions ||
        dimensions[0] !== expected.length ||
        dimensions[1] !== width
      ) {
        throw new Error(`${criterion.type} 的范围与 expected 尺寸不一致`);
      }
    }
  }
}

export function parseTurnStepEvent(value: unknown): TurnStepEvent | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const event = value as Partial<TurnStepEvent>;
  if (event.phase !== "planning" || typeof event.title !== "string" || !event.title) {
    return null;
  }
  return {
    phase: "planning",
    title: event.title,
    detail: typeof event.detail === "string" ? event.detail : null,
    completedStep:
      typeof event.completedStep === "string" ? event.completedStep : null
  };
}
