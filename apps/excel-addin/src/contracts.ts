// CellValue 的权威定义在 packages/core/src/contracts.ts（L0）。
// 本文件内部大量引用该类型，故 import 后再导出，避免两端各持一份定义后语义漂移。
import type { CellValue } from "@excel-bro/core/contracts";

export type { CellValue };

// /function 短链：AI 生成原生 Excel 公式（主路径）
export interface FormulaDictionarySheet {
  name: string;
  rows: string[][];
}

// /function 多工作簿：用户勾选的外部工作簿工作表上下文（供生成跨文件公式）。
export interface FormulaExtraSheet {
  sourceFile: string;
  sourcePath: string;
  sheetName: string;
  headers: string[];
  columns: string[];
  sampleRows: string[][];
  /** 外部表总行数（含表头行），本地拼跨表公式时用于确定引用区域末行。 */
  rowCount: number;
}

export interface GenerateFormulaRequest {
  description: string;
  activeCell: string;
  headers?: string[];
  columns?: string[];
  sampleRows?: string[][];
  dictionary?: FormulaDictionarySheet | null;
  extraSheets?: FormulaExtraSheet[];
  modelId?: string | null;
}

export interface GenerateFormulaResponse {
  modernFormula: string;
  modernExplanation: string;
  compatFormula: string;
  compatExplanation: string;
}

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
      filters?: DataFilter[];
      hasHeaders?: boolean;
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
      filters?: DataFilter[];
      hasHeaders?: boolean;
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
      filters?: DataFilter[];
      hasHeaders?: boolean;
    }
  | {
      type: "writeFormulas";
      sheet: string;
      range: string;
      formulas: string[][];
      // 锚点无关的 R1C1 形式；存在时按区域形状铺满，相对引用逐格自动平移
      formulaR1C1?: string;
    }
  | {
      type: "sortRange";
      sheet: string;
      range: string;
      keys: Array<{ column: number; ascending: boolean }>;
      hasHeaders: boolean;
    }
  | {
      type: "removeDuplicates";
      sheet: string;
      range: string;
      columns: number[];
      hasHeaders: boolean;
      filters?: DataFilter[];
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
      /**
       * 执行时求值：动作体只带查询配置，不带算好的结果。执行那一刻才读
       * sourceSheet 的实时数据算出输出，所以源数据改了计划依然成立。
       */
      type: "computeValues";
      sheet: string;
      sourceSheet: string;
      sourceRange?: string | null;
      hasHeaders: boolean;
      query: QueryTableArguments;
      destinationCell: string;
      includeHeaders: boolean;
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
      type: "formulasR1C1Equal";
      sheet: string;
      range: string;
      expected: string;
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

export interface PlanValidateRequest {
  plan: AnalysisPlan;
  authorized_sheets: string[];
  existing_sheets: string[];
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
  // 全部数据工作表名（不含 #EB_* 系统表）。sheets 只含已选表的数据，结构操作
  // （删除/重命名等用户明确要求的跨表改动）需引用 worksheetNames 才能规划。
  worksheetNames: string[];
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
  resolutionMode?: "append" | "replace";
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
  // /function 公式专用模型的 catalog id；空串=跟随全局选择。
  formulaModelId: string;
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

export interface SetFormulaModelRequest {
  modelId: string;
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
    (settings.formulaModelId !== undefined &&
      typeof settings.formulaModelId !== "string") ||
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
      kind: "conversation";
      message: string;
      provider: "model" | "local";
      turnId?: string | null;
    }
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
  planningReadPolicy?: "ifNeeded" | "none";
  modelId?: string | null;
}

export function assertIntentCheckResponse(
  value: unknown
): asserts value is IntentCheckResponse {
  if (!value || typeof value !== "object") {
    throw new Error("服务返回的意图判断不是有效对象");
  }
  const response = value as Partial<IntentCheckResponse>;
  if (response.kind === "conversation") {
    if (typeof response.message !== "string" || !response.message.trim()) {
      throw new Error("服务返回的对话回复无效");
    }
    return;
  }
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
        typeof option.resolution !== "string" ||
        (option.resolutionMode !== undefined &&
          option.resolutionMode !== "append" &&
          option.resolutionMode !== "replace")
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
  executionId?: string | null;
  filesModified: string[];
  backups: string[];
  backupDirectory?: string | null;
  backupWarnings?: string[];
  actionResults: ActionExecutionResult[];
  verification: VerificationReport;
  undo?: FolderUndoReference | null;
  executionMs: number;
  verificationMs: number;
}

export interface FolderUndoReference {
  undoId: string;
  sessionId: string;
}

export interface FolderUndoResult {
  filesRestored: string[];
  executionsUndone: number;
  remainingUndoCount: number;
}

export interface FolderExecutionStatusResult {
  operationId: string;
  status: "running" | "completed" | "failed" | "needs_reconciliation";
  lastSequence: number;
  result?: FolderExecuteResult | null;
  error?: {
    status?: number;
    code?: string;
    message?: string;
    retryable?: boolean;
  } | null;
  revision?: number | null;
  canResume?: boolean;
  checkpointStage?: "publish_prepared" | "publishing" | "published" | null;
  resumeReason?: string | null;
  resumeFilesApplied?: number | null;
  resumeFilesPending?: number | null;
}

export interface ManagedFolderRecoveryDirectory {
  directoryId: string;
  folderName: string;
  folderPath: string;
  status: "active" | "unavailable" | "unsafe";
  recoveryCount: number;
  attentionCount: number;
  operationCount: number;
  lastScannedAt?: string | null;
  statusReason?: string | null;
}

export interface ManagedFolderRecoveryDirectoriesResult {
  directories: ManagedFolderRecoveryDirectory[];
}

export interface FolderManagedRecoveryResult {
  operationId: string;
  undoId: string;
  status: "completed";
  filesRestored: string[];
  alreadyRestored: boolean;
}

export interface FolderBackupFileInfo {
  originalRelativePath: string;
  backupRelativePath?: string | null;
  recoveryAction: "restore" | "remove";
  sizeBytes: number;
}

export interface FolderBackupEntry {
  revision?: number;
  undoId: string;
  undoSessionId?: string | null;
  planId: string;
  planTitle: string;
  createdAt: string;
  status:
    | "creating"
    | "preparing"
    | "prepared"
    | "applying"
    | "ready"
    | "recovering"
    | "restored"
    | "rolled_back"
    | "needs_attention";
  statusReason?: string | null;
  undoAvailability: "eligible" | "blocked" | "view_only" | "busy";
  undoUnavailableReason?: string | null;
  deleteInvalidatesUndoIds: string[];
  pinned: boolean;
  protectedReasons: string[];
  canUndo: boolean;
  fileCount: number;
  totalBytes: number;
  expiresAt?: string | null;
  files: FolderBackupFileInfo[];
}

export interface FolderLegacyBackupInfo {
  relativePath: string;
  sizeBytes: number;
  modifiedAt: string;
}

export interface FolderBackupStorageSummary {
  automaticBytes: number;
  pinnedBytes: number;
  abnormalBytes: number;
  exportedBytes: number;
  legacyBytes: number;
  pendingCleanupBytes: number;
  totalBytes: number;
  softLimitBytes: number;
  complete: boolean;
  calculatedAt: string;
}

export interface FolderBackupsResult {
  folderPath: string;
  backupDirectory: string;
  openDirectoryToken?: string | null;
  backupManagementToken?: string | null;
  backups: FolderBackupEntry[];
  legacyCandidates: FolderLegacyBackupInfo[];
  storageSummary?: FolderBackupStorageSummary | null;
  warnings: string[];
}

export interface FolderBackupOpenResult {
  opened: boolean;
  directory: string;
}

export interface FolderBackupExportResult {
  operationId: string;
  exportDirectory: string;
  files: string[];
  complete?: boolean;
  failures?: string[];
}

export interface FolderBackupDeleteResult {
  deletedUndoId: string;
  discardedUndo: boolean;
  filesDeleted: number;
  bytesFreed: number;
  remainingUndoCount: number;
  invalidatedUndoIds: string[];
  remainingUndoIds: string[];
  cleanupPending: boolean;
  cleanupWarnings: string[];
}

export interface FolderBackupPinResult {
  undoId: string;
  pinned: boolean;
}

export interface FolderBackupCleanupResult {
  deletedExecutions: number;
  deletedFiles: number;
  bytesFreed: number;
  skippedActive: number;
  protectedExecutions?: number;
  retainedBytes?: number;
  overLimitBytes?: number;
  failures: string[];
  legacyCandidateCount: number;
}

export interface FolderLegacyCollectResult {
  collectedFiles: string[];
  failures: string[];
}

export interface ActionExecutionResult {
  index: number;
  type: ExcelAction["type"];
  sheet: string;
  status: "succeeded" | "failed" | "not_run";
  message?: string | null;
  changed?: boolean | null;
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

/**
 * 单个区域的撤销快照。
 *
 * 格式分两层：value/formula 与常用格式总是记录；`format` 里的完整格式（边框、
 * 对齐、合并）只在区域规模不超过 undo.formatCaptureMaxCells 时捕获 —— 整表操作
 * 时逐格读边框会让执行明显变慢，超限则降级为只恢复常用格式。
 */
export interface UndoRangeSnapshot {
  actionIndex: number;
  sheet: string;
  range: string;
  formulas: CellValue[][];
  numberFormat: string[][];
  fillColor: string;
  fontBold: boolean;
  fontColor: string;
  /** 本次执行完成后该范围的确定性指纹；撤销前用于阻止覆盖后续人工修改。 */
  postFingerprint?: string;
  /** 完整格式快照；区域过大而降级捕获时为 null。 */
  format?: UndoRangeFormat | null;
}

/** 超出常用格式之外、逐格恢复的格式属性。 */
export interface UndoRangeFormat {
  horizontalAlignment: string;
  verticalAlignment: string;
  wrapText: boolean | null;
  rowHeight: number;
  columnWidth: number;
  indentLevel: number;
  fontName: string;
  fontSize: number;
  fontItalic: boolean;
  fontUnderline: string;
  borders: UndoBorderState[];
}

export interface UndoBorderState {
  sideIndex: Excel.BorderIndex;
  style: string;
  color: string;
  weight: string;
}

// 被删除工作表的撤销快照：删前把整张表内容拍下来，撤销时新建并原样恢复。
export interface UndoWorksheetSnapshot {
  sheet: string;
  position: number;
  usedRange: string;
  formulas: CellValue[][];
  numberFormat: string[][];
}

export interface UndoCreatedWorksheetState {
  sheet: string;
  /** 新建表执行完成后的内容指纹；缺失时不能安全自动删除该表。 */
  postFingerprint?: string;
}

export interface ExecutionUndoSnapshot {
  planId: string;
  /** 同一计划重复执行时仍保持唯一的本次执行标识。 */
  executionId?: string;
  /** 工作簿本地标识；可取得时用于阻止把快照应用到另一个工作簿。 */
  workbookIdentity?: string;
  capturedAt: string;
  /** partial 表示失败前只捕获并执行了部分动作。 */
  coverage?: "full" | "partial";
  completedActionCount?: number;
  uncertainActionIndexes?: number[];
  /** complete 才允许直接撤销；partial 表示写后状态未能全部核对。 */
  postStateCoverage?: "complete" | "partial";
  ranges: UndoRangeSnapshot[];
  // 本次执行新建的工作表名，撤销时移除。
  createdWorksheets: string[];
  createdWorksheetStates?: UndoCreatedWorksheetState[];
  // 本次执行删除的工作表，撤销时恢复。
  deletedWorksheets: UndoWorksheetSnapshot[];
  // 删除但因过大未记录撤销快照的工作表名（无法恢复）。
  skippedWorksheets: string[];
}

/**
 * 服务端持久化的 Office.js 写入现场。只有受控区域动作且工作簿具有稳定 URL
 * 身份时才会创建；ownerToken 只在登记/领取写入权后返回。
 */
export interface OfficeExecutionState {
  operationId: string;
  status: "running" | "needs_reconciliation" | "completed" | "failed";
  revision: number;
  workbookIdentity: string;
  plan: AnalysisPlan;
  authorizedSheetNames: string[];
  allowedExternalFiles: string[];
  nextActionIndex: number;
  inFlightActionIndex?: number | null;
  beforeSnapshot?: UndoRangeSnapshot | null;
  checkpoint?: ExecutionUndoSnapshot | null;
  result?: PlanExecutionResult | null;
  error?: { code?: string; message?: string } | null;
  canResume: boolean;
  requiresReconciliation: boolean;
  ownerToken?: string | null;
}

export function assertOfficeExecutionState(
  value: unknown
): asserts value is OfficeExecutionState {
  if (!value || typeof value !== "object") {
    throw new Error("服务返回的 Office 恢复状态不是有效对象");
  }
  const state = value as Partial<OfficeExecutionState>;
  if (
    typeof state.operationId !== "string" ||
    !["running", "needs_reconciliation", "completed", "failed"].includes(
      String(state.status)
    ) ||
    typeof state.revision !== "number" ||
    typeof state.workbookIdentity !== "string" ||
    !state.plan ||
    !Array.isArray(state.plan.actions) ||
    !Array.isArray(state.authorizedSheetNames) ||
    !Array.isArray(state.allowedExternalFiles) ||
    typeof state.nextActionIndex !== "number" ||
    typeof state.canResume !== "boolean" ||
    typeof state.requiresReconciliation !== "boolean"
  ) {
    throw new Error("服务返回的 Office 恢复状态缺少必要字段");
  }
  if (
    state.inFlightActionIndex !== undefined &&
    state.inFlightActionIndex !== null &&
    typeof state.inFlightActionIndex !== "number"
  ) {
    throw new Error("服务返回的 Office 在途动作位置无效");
  }
}

export type AssistantResponse =
  | {
      kind: "answer";
      message: string;
      provider: "model" | "local" | "system";
      resultContext?: ResultContext | null;
      turnId?: string | null;
    }
  | {
      kind: "plan";
      plan: AnalysisPlan;
      provider: "model" | "local";
      turnId?: string | null;
    };

export interface PlanningPauseResult {
  kind: "planning_pause";
  turnId: string;
  status: "waiting_user" | "interrupted";
  revision: number;
  lastSequence: number;
  reason: string;
  activeElapsedSeconds: number;
  approvedWallTimeSeconds: number;
  extensionSeconds: number;
  extensionCount: number;
  maximumExtensions: number;
  canContinue: boolean;
  requiresExtension: boolean;
  pauseType: "budget" | "model_error" | "interrupted";
  currentModelId?: string | null;
  modelSwitchCount: number;
  canSwitchModel: boolean;
  availableModelConnections: Array<{
    catalogModelId: string;
    label: string;
    modelId: string;
    supportsVision: boolean;
  }>;
  sourceFingerprint?: string | null;
  sourceFingerprintSheets?: string[];
}

export type PlanningStreamResult = AssistantResponse | PlanningPauseResult;

export type PlanningTaskStatusResult =
  | (Omit<PlanningPauseResult, "status"> & {
      status: "running" | "waiting_user" | "interrupted";
    })
  | {
      turnId: string;
      status: "completed";
      revision: number;
      lastSequence: number;
      result: AssistantResponse;
      canContinue: false;
    }
  | {
      turnId: string;
      status: "failed" | "cancelled";
      revision: number;
      lastSequence: number;
      error?: {
        status?: number;
        code?: string;
        message?: string;
        retryable?: boolean;
      } | null;
      canContinue: false;
    };

export interface TurnStepEvent {
  phase: "planning";
  title: string;
  detail?: string | null;
  completedStep?: string | null;
}

export type TaskProgressPhase =
  | "intent"
  | "planning"
  | "preflight"
  | "backup"
  | "executing"
  | "saving"
  | "verifying"
  | "refreshing";

export type TaskProgressStatus =
  | "started"
  | "progress"
  | "completed"
  | "warning"
  | "failed"
  | "cancelled";

export type DecisionTraceKind =
  | "understanding"
  | "evidence"
  | "assumption"
  | "decision"
  | "conflict"
  | "correction";

export interface TaskProgressTarget {
  file?: string | null;
  sheet?: string | null;
  range?: string | null;
}

export interface TaskProgressFact {
  label: string;
  value: string;
  style?: "text" | "code";
}

export interface TaskProgressTrace {
  kind: DecisionTraceKind;
  summary: string;
  basis?: string | null;
  changedFrom?: string | null;
  requiresConfirmation?: boolean;
}

export interface TaskProgressEvent {
  schemaVersion: 1;
  operationId: string;
  executionId?: string | null;
  sequence: number;
  phase: TaskProgressPhase;
  stepId: string;
  status: TaskProgressStatus;
  title: string;
  summary?: string | null;
  detail?: string | null;
  current?: number | null;
  total?: number | null;
  unit?: "rows" | "fields" | "actions" | "sheets" | "files" | "checks" | null;
  actionIndex?: number | null;
  actionCount?: number | null;
  actionType?: string | null;
  toolName?: string | null;
  facts?: TaskProgressFact[] | null;
  target?: TaskProgressTarget | null;
  trace?: TaskProgressTrace | null;
  timestamp: string;
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
  "removeDuplicates",
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
  "computeValues",
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
  "formulasR1C1Equal",
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
  assertEditablePlan(response.plan);
}

export function assertPlanningPauseResult(
  value: unknown
): asserts value is PlanningPauseResult {
  if (!value || typeof value !== "object") {
    throw new Error("服务返回的暂停状态不是有效对象");
  }
  const result = value as Partial<PlanningPauseResult>;
  if (
    result.kind !== "planning_pause" ||
    typeof result.turnId !== "string" ||
    (result.status !== "waiting_user" && result.status !== "interrupted") ||
    typeof result.revision !== "number" ||
    typeof result.lastSequence !== "number" ||
    typeof result.reason !== "string" ||
    typeof result.activeElapsedSeconds !== "number" ||
    typeof result.approvedWallTimeSeconds !== "number" ||
    typeof result.extensionSeconds !== "number" ||
    typeof result.extensionCount !== "number" ||
    typeof result.maximumExtensions !== "number" ||
    typeof result.canContinue !== "boolean" ||
    typeof result.requiresExtension !== "boolean" ||
    (result.pauseType !== "budget" &&
      result.pauseType !== "model_error" &&
      result.pauseType !== "interrupted") ||
    typeof result.modelSwitchCount !== "number" ||
    typeof result.canSwitchModel !== "boolean" ||
    !Array.isArray(result.availableModelConnections)
  ) {
    throw new Error("服务返回的规划暂停状态缺少必要字段");
  }
  if (
    result.sourceFingerprintSheets !== undefined &&
    !Array.isArray(result.sourceFingerprintSheets)
  ) {
    throw new Error("服务返回的规划来源工作表无效");
  }
  if (
    result.availableModelConnections.some(
      (connection) =>
        !connection ||
        typeof connection !== "object" ||
        typeof connection.catalogModelId !== "string" ||
        typeof connection.label !== "string" ||
        typeof connection.modelId !== "string" ||
        typeof connection.supportsVision !== "boolean"
    )
  ) {
    throw new Error("服务返回的备用模型连接无效");
  }
}

/**
 * 校验一份计划的动作白名单与验收条件结构。
 *
 * 模型返回的计划和用户手改的 JSON 走同一道门 —— 手改的不该比模型生成的宽松，
 * 否则「能编辑」就变成了绕过白名单的后门。
 *
 * 这里只做结构与白名单校验，不查 title/summary 之类的必填字段：那些交给服务端
 * Pydantic 规范化，前端多加一层只会和既有调用方打架。
 */
export function assertEditablePlan(
  value: unknown
): asserts value is AnalysisPlan {
  if (!value || typeof value !== "object") {
    throw new Error("计划不是有效对象");
  }
  const response = { plan: value as AnalysisPlan };
  if (!Array.isArray(response.plan.actions)) {
    throw new Error("计划缺少 actions");
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
    if (criterion.type === "formulasR1C1Equal") {
      const dimensions =
        typeof criterion.range === "string"
          ? rangeDimensions(criterion.range)
          : null;
      if (
        typeof criterion.expected !== "string" ||
        criterion.expected.length === 0 ||
        !dimensions
      ) {
        throw new Error("formulasR1C1Equal 的范围或 expected 无效");
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

const taskProgressPhases = new Set<TaskProgressPhase>([
  "intent",
  "planning",
  "preflight",
  "backup",
  "executing",
  "saving",
  "verifying",
  "refreshing"
]);

const taskProgressStatuses = new Set<TaskProgressStatus>([
  "started",
  "progress",
  "completed",
  "warning",
  "failed",
  "cancelled"
]);

const decisionTraceKinds = new Set<DecisionTraceKind>([
  "understanding",
  "evidence",
  "assumption",
  "decision",
  "conflict",
  "correction"
]);

function optionalNonNegativeNumber(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) && value >= 0
    ? value
    : null;
}

const TASK_PROGRESS_FACT_LIMIT = 12;
const TASK_PROGRESS_FACT_VALUE_LIMIT = 500;
const TASK_PROGRESS_TRUNCATION_MARKER = "…（内容过长，已截断）";
const SENSITIVE_PROGRESS_LABEL =
  /token|authorization|cookie|api[ _-]?key|secret|password|credential|密钥|口令/i;
const SAFE_PROGRESS_FACT_LABELS = new Set([
  "工具",
  "动作",
  "动作数",
  "参数",
  "目标",
  "结果",
  "依据",
  "查询",
  "范围",
  "工作表",
  "工作表数",
  "文件",
  "字段数",
  "匹配数",
  "进度",
  "授权范围",
  "读取方式",
  "快照"
]);

function truncateProgressText(value: string, limit: number): string {
  if (value.length <= limit) return value;
  const keep = Math.max(0, limit - TASK_PROGRESS_TRUNCATION_MARKER.length);
  return value.slice(0, keep) + TASK_PROGRESS_TRUNCATION_MARKER;
}

function redactProgressValue(value: string): string {
  return value
    .replace(/\bBearer\s+[A-Za-z0-9._~+/=-]{8,}/gi, "Bearer 已隐藏")
    .replace(
      /([A-Za-z][A-Za-z0-9+.-]*:\/\/)([^/\s:@]+):([^@\s/]+)@/g,
      "$1已隐藏@"
    )
    .replace(/\bsk-[A-Za-z0-9_-]{12,}\b/g, "已隐藏")
    .replace(
      /\b(api[ _-]?key|token|secret|password)\s*[:=]\s*[^\s,;]+/gi,
      "$1=已隐藏"
    );
}

function parseProgressFacts(value: unknown): TaskProgressFact[] | null {
  if (!Array.isArray(value)) return null;
  const parsed: TaskProgressFact[] = [];
  for (const item of value.slice(0, TASK_PROGRESS_FACT_LIMIT)) {
    if (!item || typeof item !== "object") continue;
    const candidate = item as Record<string, unknown>;
    if (
      typeof candidate.label !== "string" ||
      !candidate.label.trim() ||
      typeof candidate.value !== "string" ||
      !candidate.value.trim()
    ) {
      continue;
    }
    const label = truncateProgressText(candidate.label.trim(), 40);
    if (
      !SENSITIVE_PROGRESS_LABEL.test(label) &&
      !SAFE_PROGRESS_FACT_LABELS.has(label)
    ) {
      continue;
    }
    const rawValue = candidate.value.trim();
    const safeValue = SENSITIVE_PROGRESS_LABEL.test(label)
      ? "已隐藏"
      : redactProgressValue(rawValue);
    parsed.push({
      label,
      value: truncateProgressText(
        safeValue,
        TASK_PROGRESS_FACT_VALUE_LIMIT
      ),
      style: candidate.style === "code" ? "code" : "text"
    });
  }
  return parsed.length > 0 ? parsed : null;
}

function parseProgressTarget(value: unknown): TaskProgressTarget | null {
  if (!value || typeof value !== "object") return null;
  const candidate = value as Record<string, unknown>;
  const rawFile =
    typeof candidate.file === "string" ? candidate.file.trim() : "";
  const file = rawFile
    ? /^[A-Za-z]:[\\/]|^\\\\|^\//.test(rawFile)
      ? (() => {
          const parts = rawFile.split(/[\\/]/).filter(Boolean);
          return parts.length > 0 ? parts[parts.length - 1] : null;
        })()
      : truncateProgressText(rawFile, 1000)
    : null;
  const sheet =
    typeof candidate.sheet === "string" && candidate.sheet.trim()
      ? truncateProgressText(candidate.sheet.trim(), 500)
      : null;
  const range =
    typeof candidate.range === "string" && candidate.range.trim()
      ? truncateProgressText(candidate.range.trim(), 80)
      : null;
  return file || sheet || range ? { file, sheet, range } : null;
}

export function parseTaskProgressEvent(
  value: unknown,
  legacyOperationId = "legacy-turn",
  legacySequence = 0
): TaskProgressEvent | null {
  if (!value || typeof value !== "object") return null;
  const candidate = value as Partial<TaskProgressEvent>;
  if (candidate.schemaVersion !== 1) {
    const legacy = parseTurnStepEvent(value);
    if (!legacy) return null;
    return {
      schemaVersion: 1,
      operationId: legacyOperationId,
      sequence: legacySequence,
      phase: "planning",
      stepId: `legacy-${legacySequence}`,
      status: "progress",
      title: legacy.title,
      detail: legacy.detail ?? null,
      summary: null,
      timestamp: new Date().toISOString()
    };
  }
  if (
    typeof candidate.operationId !== "string" ||
    !candidate.operationId.trim() ||
    typeof candidate.sequence !== "number" ||
    !Number.isInteger(candidate.sequence) ||
    candidate.sequence < 0 ||
    !taskProgressPhases.has(candidate.phase as TaskProgressPhase) ||
    !taskProgressStatuses.has(candidate.status as TaskProgressStatus) ||
    typeof candidate.stepId !== "string" ||
    !candidate.stepId.trim() ||
    typeof candidate.title !== "string" ||
    !candidate.title.trim() ||
    typeof candidate.timestamp !== "string" ||
    !candidate.timestamp.trim()
  ) {
    return null;
  }
  const current = optionalNonNegativeNumber(candidate.current);
  const total = optionalNonNegativeNumber(candidate.total);
  if (current !== null && total !== null && current > total) return null;
  const traceValue = candidate.trace;
  const trace =
    traceValue &&
    typeof traceValue === "object" &&
    decisionTraceKinds.has(traceValue.kind as DecisionTraceKind) &&
    typeof traceValue.summary === "string" &&
    traceValue.summary.trim()
      ? {
          kind: traceValue.kind as DecisionTraceKind,
          summary: traceValue.summary,
          basis: typeof traceValue.basis === "string" ? traceValue.basis : null,
          changedFrom:
            typeof traceValue.changedFrom === "string"
              ? traceValue.changedFrom
              : null,
          requiresConfirmation: Boolean(traceValue.requiresConfirmation)
        }
      : null;
  return {
    schemaVersion: 1,
    operationId: candidate.operationId,
    executionId:
      typeof candidate.executionId === "string" ? candidate.executionId : null,
    sequence: candidate.sequence,
    phase: candidate.phase as TaskProgressPhase,
    stepId: candidate.stepId,
    status: candidate.status as TaskProgressStatus,
    title: candidate.title,
    summary:
      typeof candidate.summary === "string"
        ? truncateProgressText(candidate.summary, 1000)
        : null,
    detail:
      typeof candidate.detail === "string"
        ? truncateProgressText(candidate.detail, 2000)
        : null,
    current,
    total,
    unit: candidate.unit ?? null,
    actionIndex: optionalNonNegativeNumber(candidate.actionIndex),
    actionCount: optionalNonNegativeNumber(candidate.actionCount),
    actionType:
      typeof candidate.actionType === "string"
        ? truncateProgressText(candidate.actionType, 100)
        : null,
    toolName:
      typeof candidate.toolName === "string"
        ? truncateProgressText(candidate.toolName, 100)
        : null,
    facts: parseProgressFacts(
      (candidate as TaskProgressEvent & { facts?: unknown }).facts
    ),
    target: parseProgressTarget(candidate.target),
    trace,
    timestamp: candidate.timestamp
  };
}
