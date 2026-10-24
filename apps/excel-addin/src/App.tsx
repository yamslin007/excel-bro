import { RecoveryProbeControl } from "@excel-bro/core/RecoveryProbeControl";
import { bindRecoveryProbe, interruptRecoveryProbe } from "@excel-bro/core/recoveryProbe";
import { prepareExcelValueCopies } from './excelValueCopyPlan';
import { TaskExecutionControls } from "./components/TaskExecutionControls";
import { acceptsPlanningReplay, planningReceivedSequence, waitForPlanning } from "@excel-bro/core/planningReplay";
import {
  PAUSE_UNAVAILABLE_REASON,
  PauseTimeoutError,
  controllerOwnerOf,
  executionPauseState,
  newExecutionId,
  pauseButtonState,
  planExecutionTeardown,
  recordPauseDiagnostic,
  tagControllerOwner,
  tailOf,
  withPauseTimeout
} from "@excel-bro/core/pauseControl";
import { pausePlanningTask } from "./api";
import { planningRecoveryCandidates, isOfferedPlanningRecovery } from "@excel-bro/core/planningRecoveryRouting";
import { usePrePlanningPause } from "./hooks/usePrePlanningPause";
import { PrePlanningPauseCard } from "./components/PrePlanningPauseCard";
import { PauseDetails } from "./components/PauseDetails";
import {
  isWorksheetDeletion,
  reconcileWorksheetDeletionScope,
  worksheetDeletionIssue
} from "@excel-bro/core/worksheetDeletion";
import { reconcileWorksheetRestoreScope } from "@excel-bro/core/worksheetRestore";
import { useWorksheetCheckpoint } from "./hooks/useWorksheetCheckpoint";
import {
  WorksheetCheckpointCard,
  WorksheetCheckpointRecoverySection
} from "./components/WorksheetCheckpointCard";
import { ModelConnectionProtocolFields } from "./components/ModelConnectionProtocolFields";
import {
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
} from "react";
import { ResultReview } from '@excel-bro/core/resultReview';
import { ResultReviewBar } from './components/ResultReviewBar';
import { excelRecoveryBeforeResult, captureExcelBeforeResult, showExcelBeforeResult, closeExcelNativeComparison, type ExcelBeforeResult } from './excelNativeComparison';
import { markExcelResult, clearExcelResultMarks, recoverExcelResultMarks, type ExcelResultMarks } from './resultReviewMarkers';
import type { PlanExecutionResult, ExecutionUndoSnapshot } from './contracts';
import {
  checkIntent,
  understandMessage,
  claimOfficeExecutionResume,
  completeOfficeAction,
  continuePlanningTask,
  finishOfficeExecution,
  getOfficeExecutionStatus,
  getPlanningTaskEvents,
  getPlanningTaskStatus,
  streamAssistantResponse,
  createFolderSnapshot,
  checkFolderAdmission,
  executeFolderPlan,
  executeFolderQuery,
  generateFormula,
  getFolderExecutionStatus,
  isLocalServiceConnectionError,
  reconcileOfficeActionReceipt,
  registerOfficeExecution,
  refreshFolder,
  resumeFolderExecution,
  scanFolderByPath,
  selectFolder,
  startOfficeAction,
  suspendOfficeExecution,
} from "./api";
import { ClarificationFields } from "./components/ClarificationFields";
import { brandLogoUrl } from "./brandLogo";
import { createAgentTaskContext, confirmAgentTaskDecision, type AgentTaskContext } from "@excel-bro/core/agentTaskContext";
import { renderConfirmedPrompt } from "@excel-bro/core/clarificationAnswers";
import { ModelPickerIcon } from "./components/ModelPickerIcon";
import { TargetScopeNotice } from "./components/TargetScopeNotice";
import { ScopeAdmissionSummary } from "./components/ScopeAdmissionStatus";
import { useImmediateScopeAuthorization } from "./hooks/useImmediateScopeAuthorization";
import { useScopeAdmission } from "./hooks/useScopeAdmission";
import { useAuthorizedTarget } from "./hooks/useAuthorizedTarget";
import { confirmOfficeAuthorizedScope, inspectOfficeAuthorizedTarget, inspectOfficeScopeAdmission, officeAnalysisTargetEvidence } from "./officeAuthorizedTarget";
import { buildPlanningContinuationEvidence, PLANNING_EVIDENCE_FAILURE_MESSAGES } from "@excel-bro/core/planningContinuationEvidence";
import { admissionComplete, admissionItem } from "@excel-bro/core/scopeAdmission";
import { checkAuthorizedTarget, TargetConfirmationRequiredError, type AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { folderRootIdentity, intentScopeFingerprint as sharedIntentScopeFingerprint } from "@excel-bro/core/scopeFingerprint";
import { retireTargetMessage } from "./authorizedTargetArtifacts";
import type {
  AnalysisPlan,
  AssistantResponse,
  ColumnMappingRule,
  FolderCatalog,
  FolderExecuteResult,
  FolderSelection,
  DataToolResult,
  FormulaDictionarySheet,
  FormulaExtraSheet,
  IntentCheckResponse,
  IntentClarification,
  IntentMemory,
  IntentOption,
  IntentScopeContext,
  PlanningStreamResult,
  QueryTableArguments,
  ResultContext,
  VerificationReport,
  WorkbookSnapshot
} from "./contracts";
import type {
  ChatConversation,
  ChatHistoryState,
  ChatMessage,
  FunctionPreview,
  Status
} from "./types/chat";
import type { SourceMode, WorkbookScopeMode } from "./types/workbook";
// 计划卡/结果卡的口径与 WPS 端共用同一份字面量，避免两端各写各的中文后悄悄漂移。
import {
  BlankPolicyNeededError,
  NeedsClarificationError,
  PLAN_FORMAT_IMPACT_LABEL,
  PLAN_FORMAT_SECTION,
  PLAN_NOTE_SECTIONS,
  UnmappedCategoryError,
  VERIFICATION_PASSED_LABEL,
  bridgeWorkbookIdentity,
  formatPlanPreview,
  redactCredentials,
  selectFormatActions
} from "@excel-bro/core";
import type { CellValue } from "@excel-bro/core";
import type { RangeChangeReview } from "@excel-bro/core/changeReview";
import { demoWorkbook } from "./demo";
import {
  captureSelectionContext,
  captureWorkbook,
  captureWorkbookSourceFingerprint,
  captureWorkbookStructure,
  captureWorkbookCatalog,
  dataEpochsChanged,
  currentOfficeWorkbookIdentity,
  assertUndoPostState,
  executePlan,
  isExplicitA1Address,
  isEBSystemSheet,
  isRunningInExcel,
  readSelectedRange,
  readSheetUsedRange,
  previewFormulaFirstCell,
  PlanExecutionError,
  officeRecoverySupport,
  reconcileOfficeRangeAction,
  StaleSourceError,
  snapshotDataEpochs,
  toolSchemaFingerprintForSnapshot,
  watchWorkbookStructureChanges
} from "./excel";
import {
  DataToolExecutionError,
  executeQueryTableTool
} from "./dataTools";
import {
  createQueryTool,
  createTool,
  instantiateTool,
  type SavedQueryTool,
  type SavedTool,
  type ToolParameter
} from "./storage";
import {
  executeSavedQueryTool,
  SavedQueryToolFallbackError
} from "./deterministicTools";
import { highlightToolDsl, renderToolDsl } from "./toolDsl";
import type { PlanEditorTarget } from "./hooks/usePlanEditor";
import {
  MAX_IMAGE_ATTACHMENTS,
  type PendingImage
} from "./imageAttachments";
import { useImageAttachments } from "./hooks/useImageAttachments";
import { workbookFolderFromDocumentUrl } from "./workbookIdentity";
import {
  currentOfficeHostVersion,
  officeHostRequestFields
} from "./officeHostAdapter";
import {
  buildColumnMappingPlan,
  stripDataRows as stripDataRowsForLocalMapping
} from "./columnMappingPlan";
import { columnMappingClarificationMessage } from "./columnMappingClarification";
import {
  shouldFinishActivityAsCompleted,
  type TurnPlanningOutcome
} from "./turnPlanningOutcome";
import {
  evaluateBulkWorksheetDeletion,
  parseBulkWorksheetDeletion
} from "./structuralIntent";
import {
  folderFileFullPath,
  normalizeWorkbookPath
} from "./workbookPathMatch";
import capabilities from "../../../config/capabilities.json";
import {
  buildCrossTableFormulas,
  buildCrossTableProposal,
  columnsFromRange,
  type CrossTableMatchProposal,
  type GeneratedFormulaPair
} from "./crossTableFormula";
import { formulaExternalFileNames, fuzzyLookupMatch } from "./formulaSafety";
import {
  currentModelCallCount,
  exportDiagnosticReport,
  recordDiagnosticEvent
} from "./diagnostics";
import {
  FOCUS_PAYLOAD_STORAGE_KEY,
  type FocusPayload
} from "./focusState";
import {
  describeIntentDecision
} from "./intent";
import PetCompanion from "./PetCompanion";
import { RuleManager } from "./RuleManager";
import { SlashCommandAutocomplete, type SlashCommand } from "./SlashCommandAutocomplete";
import ThemePanel from "./components/ThemePanel";
import TaskProgressTimeline from "./components/TaskProgressTimeline";
import ChangeReviewSummaryCard from "./components/ChangeReviewSummaryCard";
import { ExcelNativeReviewActions } from "./components/ExcelNativeReviewActions";
import { SaveToolDialog } from "./components/SaveToolDialog";
import { useExcelNativeComparison, excelReviewBlocksNewCopy } from "./hooks/useExcelNativeComparison";
import { assertExcelNativeComparisonUnchanged, compareExcelNativeCommittedResult, resultReviewPlanIssue, type ExcelNativeComparison } from "./excelNativeComparison";
import {
  attestResultSourceDependency,
  attestSourceDependency,
  trustedToolReads
} from "./planSourceDependency";
import { LocalMappingResolution } from "./components/LocalMappingResolution";
import MessageCopyButton from "./components/MessageCopyButton";
import { ExecutionResultCard } from "./components/ExecutionResultCard";
import ScrollToLatestButton from "./components/ScrollToLatestButton";
import { FolderBackupManager } from "./components/FolderBackupManager";
import { ScopeSelectionGuidance } from "./components/ScopeSelectionGuidance";
import HistoryStorageNotice from "./components/HistoryStorageNotice";
import {
  shouldShowWorkDirectoryGuide,
  WorkDirectoryGuide
} from "./components/WorkDirectoryGuide";
import type { FolderBackupTarget } from "./hooks/useFolderBackups";
import { MODEL_REQUIRED_HELP_TEXT } from "./helpCommand";
import { agentModelAvailable, agentModelUnavailableReason, formulaModelAvailable, chooseAvailableModel, connectionModelSuffix, isRealModelId, isSdkPlanningModel, planningReadPolicyForModel } from "./modelSelection";
import { useActivityProgress, type ActivityLog, type ActivityProgress } from "./hooks/useActivityProgress";
import {
  continuationLabel,
  freezeSegmentLog,
  openReplySegment,
  planSegmentContinuation
} from "@excel-bro/core/replySegments";
import { useSdkToolBridge } from "./hooks/useSdkToolBridge";
import { loadBridgeResult, type SdkBridgeLiveContext } from "./sdkBridgeHost";
import LocalResultCard from "./components/LocalResultCard";
import { useMessageAutoFollow } from "./hooks/useMessageAutoFollow";
import { useComposerResize } from "./hooks/useComposerResize";
import { useConversation } from "./hooks/useConversation";
import {
  canAcknowledgeStaleSource,
  redactHistoryCredentials,
  remainingLegacyCredentialConversationIds
} from "./conversation";
import {
  markFolderExecutionUndone,
  markWorkbookExecutionUndone
} from "./executionUndone";
import {
  insertChatMessage,
  persistActivityMessage
} from "./activityMessages";
import { useCopyFeedback } from "./hooks/useCopyFeedback";
import { useExecutionApproval } from "./hooks/useExecutionApproval";
import { useModelManagement } from "./hooks/useModelManagement";
import { useServiceHealth } from "./hooks/useServiceHealth";
import { useTheme } from "./hooks/useTheme";
import { useToolManagement } from "./hooks/useToolManagement";
import { useUIState } from "./hooks/useUIState";
import {
  addedFolderSheetKeys,
  resolveFolderViewAfterFileClear,
  useScopeSelection
} from "./hooks/useScopeSelection";
import { useFolderUndo } from "./hooks/useFolderUndo";
import { useUndoSnapshot } from "./hooks/useUndoSnapshot";
import { useLongPress } from "./hooks/useLongPress";
import { folderSheetKey, formulaTrialFailed } from "./utils";
import {
  clearPendingFolderOperation,
  loadPendingFolderOperation,
  savePendingFolderOperation,
  type PendingFolderOperation
} from "./folderOperationRecovery";
import {
  clearPendingOfficeExecution,
  loadPendingOfficeExecution,
  savePendingOfficeExecution,
  type PendingOfficeExecution
} from "./officeOperationRecovery";

export type { ActivityLog, ActivityProgress } from "./hooks/useActivityProgress";

// 生成耗时格式化：<1s 显示毫秒，否则显示秒（保留一位小数）。
function formatGenerateMs(ms: number): string {
  if (ms < 1000) return `${ms} 毫秒`;
  return `${(ms / 1000).toFixed(1)} 秒`;
}

const PET_ANIMATIONS = [
  "turtle-spin",
  "turtle-blink",
  "turtle-nod",
  "turtle-wave"
];

function animatePet(className: string, duration = 800): void {
  const pet = document.querySelector<HTMLElement>(".pet-avatar");
  console.log("[Pet] animatePet called:", className, pet);
  if (!pet) return;
  pet.classList.remove(...PET_ANIMATIONS, "turtle-sleepy", "turtle-encourage");
  pet.classList.add(className);
  console.log("[Pet] Animation class added:", pet.className);
  window.setTimeout(() => {
    pet.classList.remove(className);
    console.log("[Pet] Animation class removed:", className);
  }, duration);
}

const CHAT_STORAGE_KEY = "excel-bro.chat.v4";

/**
 * 接回"仍在运行的原任务"时的轮询节奏（与 WPS 端同一口径）。
 *
 * 有上限而不是死等：服务端跑一次规划本身有硬时限，超过这个窗口还没结束就不是
 * "等一等就能好"的事，如实说"仍在运行、稍后核对"比无限转圈更诚实。
 */
const RUNNING_TAIL_INTERVAL_MS = 1200;
const RUNNING_TAIL_MAX_MS = 90_000;

/**
 * 暂停请求的确认上限（与 WPS 端同一数值）。
 *
 * 本端 abort 是同步的，但"服务端收到没有"只能靠 HTTP 回执。没有上限时，一次卡住
 * 的请求会让"正在请求暂停"永远挂着 —— 用户以为按住了，其实什么都没发生。
 * 超时不是失败结论：如实说"未确认"。
 */
const PAUSE_REQUEST_TIMEOUT_MS = 8_000;

const MODEL_CONFIG_HELP_SHOWN_KEY = "ebModelConfigHelpShown";
const MAX_MESSAGES_PER_CONVERSATION =
  capabilities.conversation.maxMessagesPerConversation;
const PERSISTED_MESSAGES_PER_CONVERSATION =
  capabilities.conversation.persistedMessagesPerConversation;
const INTENT_HISTORY_MESSAGES =
  capabilities.conversation.intentHistoryMessages;
const INTENT_MESSAGE_CHARACTERS =
  capabilities.conversation.intentMessageCharacters;
const MAX_CLARIFICATION_ROUNDS =
  capabilities.conversation.maxClarificationRounds;
const INTENT_MAX_FIELDS = capabilities.intentContext.maxFieldsPerSheet;
const INTENT_MAX_PRIOR_RESULT_ROWS =
  capabilities.intentContext.maxPriorResultRows;

function verificationSummary(report: VerificationReport): string {
  if (report.status === "verified") return "并通过独立验证";
  if (report.status === "executed_unverified") {
    return "；写入已完成，但部分操作暂时无法独立验证";
  }
  return "，但结果验证未通过";
}

function folderExecutionOutcomeText(
  plan: AnalysisPlan,
  result: FolderExecuteResult,
  recovered: boolean
): string {
  const writeStatus =
    result.filesModified.length > 0
      ? `已写入 ${result.filesModified.join("、")}`
      : "目标状态原本已符合要求，没有写入文件，也没有创建恢复点";
  const verificationStatus =
    result.verification.status === "verified"
      ? "已通过独立验证"
      : result.verification.status === "executed_unverified"
        ? "写入已完成，部分操作暂时无法独立验证"
        : "写入已经结束，但结果验证未通过";
  return `${recovered ? "已找回上次中断后的执行结果。" : ""}「${plan.title}」任务状态：执行已结束，共 ${result.actionResults.length} 步。写入状态：${writeStatus}。验收状态：${verificationStatus}。${
    result.backupDirectory
      ? "执行前副本已集中保存，可在「备份与恢复」中查看。"
      : ""
  }${result.undo ? "已记录本次恢复点。" : ""}${
    result.backupWarnings?.length
      ? `注意：${result.backupWarnings.join("；")}。`
      : ""
  }`;
}

// 历史归一化（normalizeStoredVerification 等）已收敛到 ./conversation.ts，
// 由两端加载项共用；这里不再保留副本。

// isAbortError：区分「用户主动打断」和真正的错误。fetch 被 abort 时抛 AbortError，
// 本地工具取消抛 DataToolExecutionError(code=CANCELLED)。这两种都不该弹错误提示。
function isAbortError(reason: unknown): boolean {
  if (reason instanceof DOMException && reason.name === "AbortError") {
    return true;
  }
  if (
    reason instanceof DataToolExecutionError &&
    reason.code === "CANCELLED"
  ) {
    return true;
  }
  return false;
}

function latestResultContext(items: ChatMessage[]): ResultContext | null {
  const result =
    [...items].reverse().find((message) => message.resultContext)
      ?.resultContext ?? null;
  return result
    ? {
        ...result,
        rows: result.rows.slice(0, INTENT_MAX_PRIOR_RESULT_ROWS)
      }
    : null;
}

function messageId(): string {
  return globalThis.crypto?.randomUUID?.() ??
    `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

// 从选区地址取首格：兼容 "Sheet1!E2:E50" / "E2:E50" / "E2"。
function firstCellOfRange(address: string): string | null {
  const bare = address.includes("!") ? address.split("!").pop()! : address;
  const first = bare.split(":")[0]?.trim();
  if (!first || !/^\$?[A-Za-z]{1,3}\$?\d+$/.test(first)) return null;
  return first.replace(/\$/g, "");
}

// 列号 → 列字母（0→A, 25→Z, 26→AA）。
function columnLetterFromIndex(index: number): string {
  let n = index;
  let letters = "";
  do {
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return letters;
}

// 列字母 → 列号（A→0, Z→25, AA→26）。
function columnIndexFromLetter(letter: string): number {
  let index = 0;
  for (const ch of letter.toUpperCase()) {
    index = index * 26 + (ch.charCodeAt(0) - 64);
  }
  return index - 1;
}

// 从 usedRange 地址取起始列字母（"A1:E20" → "A"，"Sheet1!C2:E9" → "C"）。
function startColumnOfRange(address: string | null): string | null {
  if (!address) return null;
  const bare = address.includes("!") ? address.split("!").pop()! : address;
  const match = bare.match(/^\$?([A-Za-z]{1,3})/);
  return match ? match[1].toUpperCase() : null;
}

// 智能建议写入目标：数据区右侧第一空列 + 数据首行（表头下一行）。
// 解析 usedRange（如 "A1:D20"）取末列右移一列、起始行 +1。解析失败退回 "A2"。
function suggestWriteTarget(usedRange: string | null): string {
  if (!usedRange) return "A2";
  const bare = usedRange.includes("!")
    ? usedRange.split("!").pop()!
    : usedRange;
  const match = bare.match(
    /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/
  );
  if (!match) return "A2";
  const startRow = Number(match[2]);
  const endColLetter = (match[3] ?? match[1]).toUpperCase();
  const nextCol = columnLetterFromIndex(columnIndexFromLetter(endColLetter) + 1);
  return `${nextCol}${startRow + 1}`;
}

interface ExternalWorkbookRef {
  file: string;
  sheet: string;
}

// 解析公式中的外部工作簿引用（如 '[B.xlsx]Sheet2'!A1），去重后返回。
function externalWorkbookRefs(formula: string): ExternalWorkbookRef[] {
  const refs: ExternalWorkbookRef[] = [];
  const pattern =
    /\[([^\]\s]+\.(?:xlsx|xlsm|xlsb|xls|csv|xlw)\])'?([^'!]+)'?!/gi;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(formula)) !== null) {
    const file = match[1].slice(1, -1);
    const sheet = match[2].trim();
    if (!file || !sheet) continue;
    const key = `${file}\u0000${sheet}`;
    if (!refs.some((ref) => `${ref.file}\u0000${ref.sheet}` === key)) {
      refs.push({ file, sheet });
    }
  }
  return refs;
}

// 验算拦截文案：返回 null 表示可写入；否则返回原因（不能算就不算，不硬写）。
function functionWriteBlockReason(preview: FunctionPreview): string | null {
  const noFormula = !preview.modernFormula && !preview.compatFormula;
  if (noFormula) {
    return "无法生成可靠公式：未能得到可用公式。为避免硬写错误结果，已拦截写入；请手动填写，或调整描述后重试。";
  }
  const formula =
    preview.version === "modern"
      ? preview.modernFormula
      : preview.compatFormula;
  const trial =
    preview.version === "modern"
      ? preview.modernResult
      : preview.compatResult;
  if (
    formulaTrialFailed(trial) &&
    formulaExternalFileNames(formula).length === 0
  ) {
    return `试算未通过（${trial}）：为避免写入错误结果，已拦截写入；请检查公式或手动填写。`;
  }
  const fuzzy = fuzzyLookupMatch(formula);
  if (fuzzy !== null) {
    return `公式包含模糊匹配（${fuzzy}）：近似匹配会静默返回错误数据，已拦截写入；请手动改成精确匹配（VLOOKUP 用 FALSE、MATCH 用 0），或调整描述重试。`;
  }
  return null;
}

// /function 上下文：扫描名字含"字典"/"映射"的表，读其内容注入生成提示。
async function loadDictionaryForFormula(
  activeSheetName: string
): Promise<FormulaDictionarySheet | null> {
  try {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const match = sheets.items.find(
        (sheet) =>
          sheet.name !== activeSheetName &&
          (sheet.name.includes("字典") || sheet.name.includes("映射"))
      );
      if (!match) return null;

      const usedRange = match.getUsedRangeOrNullObject(true);
      usedRange.load("values,rowCount,isNullObject");
      await context.sync();
      if (usedRange.isNullObject || usedRange.rowCount === 0) return null;

      const rows = (usedRange.values as unknown[][])
        .slice(0, 200)
        .map((row) => row.map((cell) => String(cell ?? "")));
      return { name: match.name, rows };
    });
  } catch {
    return null;
  }
}

// ── 结论复用缓存 ──────────────────────────────────────────────────────
// 同一只读需求（结构化意图相同）且数据未变（dataEpoch 未变）时，直接复用上次
// 结论，跳过 checkIntent / 本地全表扫描 / 模型生成。缓存只放内存、不进
// localStorage，跨会话不自动复用（关闭期间文件可能被外部改动，监听器无从得知）。
const RESULT_CACHE_LIMIT = 24;
const PROMPT_KEY_CACHE_LIMIT = 48;
// 模块加载时生成一次的会话 id；换会话（刷新页面）即失效缓存。
const CACHE_SESSION_ID =
  globalThis.crypto?.randomUUID?.() ?? `session-${Date.now()}`;

interface CachedConclusion {
  intentKey: string;
  resultContext: ResultContext;
  answerText: string;
  sourceSheets: string[];
  dataEpochSnapshot: Record<string, number>;
  completeness: "complete" | "truncated";
  sourceMode: SourceMode;
  sessionId: string;
  querySourceSheetNames?: string[];
  querySourceSheetIds?: string[];
  createdAt: number;
}

// Map 迭代顺序即插入顺序，超限时删最早的键 => 朴素 LRU。
function lruSet<K, V>(cache: Map<K, V>, key: K, value: V, limit: number): void {
  if (cache.has(key)) cache.delete(key);
  cache.set(key, value);
  while (cache.size > limit) {
    const oldest = cache.keys().next().value as K | undefined;
    if (oldest === undefined) break;
    cache.delete(oldest);
  }
}

function lruGet<K, V>(cache: Map<K, V>, key: K): V | undefined {
  const value = cache.get(key);
  if (value !== undefined) {
    cache.delete(key);
    cache.set(key, value);
  }
  return value;
}

// 原样文本命中的一级 key：去首尾空白、内部空白折叠、小写。
export function normalizePrompt(text: string): string {
  return text.trim().replace(/\s+/g, " ").toLowerCase();
}

// 结构化意图 key：归一化 QueryTableArguments + scope 后排序序列化。
// filters 是 AND 关系、顺序无关 => 排序；groupBy/metrics/fields 顺序影响输出 => 保序。
// 归一化从严：误 miss（多算一次）可接受，误命中（拿错结论）危险。
export function normalizeIntentKey(
  args: QueryTableArguments,
  scope: IntentScopeContext
): string {
  const field = (value?: string | null): string =>
    (value ?? "").trim().toLowerCase();
  const filters = (args.filters ?? [])
    .map((filter) => ({
      f: field(filter.field),
      o: filter.operator,
      v: filter.value ?? null
    }))
    .sort((a, b) =>
      `${a.f}|${a.o}|${String(a.v)}`.localeCompare(
        `${b.f}|${b.o}|${String(b.v)}`
      )
    );
  const metrics = (args.metrics ?? []).map((metric) => ({
    op: metric.operation,
    f: field(metric.field),
    out: field(metric.outputName),
    ratio: field(metric.ratioOutputName)
  }));
  return JSON.stringify({
    mode: args.mode,
    scope: args.scope ?? null,
    fields: (args.fields ?? []).map(field),
    filters,
    groupBy: (args.groupBy ?? []).map(field),
    metrics,
    combine: args.combine ?? null,
    profileField: field(args.profileField),
    sortBy: field(args.sortBy),
    sortDirection: args.sortDirection ?? null,
    limit: typeof args.limit === "number" ? args.limit : null,
    workbookName: scope.workbookName,
    sourceMode: scope.sourceMode,
    sheets: [...scope.sheets.map((sheet) => sheet.name)].sort()
  });
}

function conversationTitle(messages: ChatMessage[]): string {
  const firstUserMessage = messages.find(
    (message) => message.role === "user" && message.text?.trim()
  )?.text;
  if (!firstUserMessage) return "新对话";
  const compact = firstUserMessage.replace(/\s+/g, " ").trim();
  return compact.length > 34 ? `${compact.slice(0, 34)}...` : compact;
}

// 对话历史的纯函数（loadChatHistory / deleteConversationFromHistory / createConversation /
// welcomeMessage 等）已统一收敛到 ./conversation.ts，两端加载项共用同一份实现。
// 这里不再保留副本 —— 两份实现会让"读取时是否脱敏"这种安全语义在不同入口上分叉。


/**
 * 计划卡里的「格式改动」预览。
 *
 * 行高/列宽是**行级/列级**属性：在 `B3:D5` 上设行高改的是第 3–5 整行。
 * 只写动作标签（"调整 B3:D5 的行高列宽"）会让用户以为只动这一小块，
 * 所以这里必须把真实影响范围单独念出来。
 *
 * 归组、属性文案、影响范围判定都在共用层（`@excel-bro/core`），
 * 两端同一份实现 —— 各写一份就会出现"Excel 说了、WPS 没说"。
 */
function planFormatPreview(plan: AnalysisPlan) {
  return formatPlanPreview(selectFormatActions(plan.actions));
}

function actionLabel(action: AnalysisPlan["actions"][number]): string {
  switch (action.type) {
    case "createWorksheet":
      return `新建或复用工作表「${action.sheet}」`;
    case "writeTable":
      return `向「${action.sheet}」${action.startCell} 写入 ${
        action.rows.length + 1
      } 行表格`;
    case "writeValues":
      return `写入「${action.sheet}」${action.range}`;
    case "setFill":
      return `设置「${action.sheet}」${action.range} 的填充色`;
    case "setFont":
      return `设置「${action.sheet}」${action.range} 的字体`;
    case "autofit":
      return `自动调整「${action.sheet}」${action.range}`;
    case "activateWorksheet":
      return `切换到工作表「${action.sheet}」`;
    case "deleteWorksheet":
      return `删除工作表「${action.sheet}」`;
    case "clearRange":
      return `清除「${action.sheet}」${action.range}（${action.applyTo}）`;
    case "insertRange":
      return `在「${action.sheet}」${action.range} 插入单元格`;
    case "deleteRange":
      return `删除「${action.sheet}」${action.range} 的单元格`;
    case "copyRange":
      return `复制「${action.sourceSheet}」${action.sourceRange} 到「${action.sheet}」${action.targetRange}`;
    case "writeFormulas":
      return `向「${action.sheet}」${action.range} 写入公式`;
    case "sortRange":
      return `排序「${action.sheet}」${action.range}`;
    case "filterRange":
      return `筛选「${action.sheet}」${action.range}`;
    case "clearFilter":
      return `清除「${action.sheet}」的筛选条件`;
    case "setDataValidation":
      return `设置「${action.sheet}」${action.range} 的数据验证`;
    case "setConditionalFormat":
      return `设置「${action.sheet}」${action.range} 的条件格式`;
    case "setNumberFormat":
      return `设置「${action.sheet}」${action.range} 的数字格式`;
    case "setBorders":
      return `设置「${action.sheet}」${action.range} 的边框`;
    case "setAlignment":
      return `设置「${action.sheet}」${action.range} 的对齐方式`;
    case "mergeCells":
      return `合并「${action.sheet}」${action.range}`;
    case "unmergeCells":
      return `取消合并「${action.sheet}」${action.range}`;
    case "resizeRange":
      return `调整「${action.sheet}」${action.range} 的行高列宽`;
    case "freezePanes":
      return `冻结「${action.sheet}」的窗格`;
    case "setHyperlink":
      return `为「${action.sheet}」${action.range} 添加超链接`;
    case "addComment":
      return `为「${action.sheet}」${action.cell} 添加批注`;
    case "addNote":
      return `为「${action.sheet}」${action.cell} 添加备注`;
    case "createTable":
      return `将「${action.sheet}」${action.range} 转换为表格`;
    case "createChart":
      return `基于「${action.sheet}」${action.sourceRange} 创建图表`;
    case "createPivotTable":
      return `在「${action.sheet}」创建数据透视表「${action.name}」`;
    case "splitGroupAggregate":
      return `按「${action.splitBy}」拆分「${action.sheet}」，并按 ${action.groupBy.join(
        "、"
      )} 汇总`;
    case "computeValues":
      return `从「${action.sourceSheet}」计算并写入「${action.sheet}」${action.destinationCell}`;
    case "addNamedRange":
      return `为「${action.sheet}」${action.range} 创建名称「${action.name}」`;
    case "addImage":
      return `在「${action.sheet}」${action.targetRange} 添加图片`;
    case "addShape":
      return `在「${action.sheet}」${action.targetRange} 添加形状`;
    default:
      return "未知操作";
  }
}

/**
 * "读不到宿主现场"时的桥现场。
 *
 * 刻意是**空范围 + 空身份 + 无指纹**：复核拿它去跟绑定比必然不通过，于是
 * 一次宿主读取都不会发生（fail-closed）。任何"读不到就当没变化"的降级都不许写在这里。
 */
function emptyBridgeLiveContext(): SdkBridgeLiveContext {
  return {
    workbookIdentity: "",
    sourceFingerprint: null,
    authorizedSheetNames: [],
    activeWorksheet: ""
  };
}

export default function App() {
  const [workbook, setWorkbook] = useState<WorkbookSnapshot | null>(null);
  const [prompt, setPrompt] = useState("");
  /**
   * 发送失败时把原需求放回输入框，但**不覆盖用户后来打的字**。
   *
   * 一次请求失败可能发生在几秒之后；这段时间里用户完全可能已经开始写下一句。
   * 无条件写回会把刚写的草稿冲掉 —— 那不是"恢复原样"，是丢草稿。
   */
  const restorePromptUnlessEdited = (value: string) =>
    setPrompt((current) => (current.trim() === "" ? value : current));
  const [newConversationNotice, setNewConversationNotice] = useState<{
    id: number;
    previousSaved: boolean;
  } | null>(null);
  const {
    pendingImages, setPendingImages, imageError, setImageError,
    draggingImage, setDraggingImage, preparingImages, addImageFiles,
    handleImagePaste, handleDrop: handleImageDrop
  } = useImageAttachments();
  const [clarificationDrafts, setClarificationDrafts] = useState<
    Record<string, string>
  >({});
  const [clarificationFieldDrafts, setClarificationFieldDrafts] = useState<
    Record<string, Record<string, string>>
  >({});

  // 对话管理（使用自定义 Hook）
  const conversationHook = useConversation();
  const {
    chatHistory,
    activeConversation,
    messages,
    staleCredentials,
    historyStorageStatus,
    pendingDeleteConversationId: pendingDeleteConvId,
    setMessages,
    newChat,
    openConversation,
    deleteConversation,
    confirmDeleteConversation,
    cancelDeleteConversation,
    setChatHistory,
    setPendingDeleteConversationId
  } = conversationHook;

  const [status, setStatus] = useState<Status>("idle");
  const [activePlanId, setActivePlanId] = useState<string | null>(null);
  const nativeComparison = useExcelNativeComparison();
  type ResultReceipt = { executionId: string; verified: boolean; undoable: boolean; undo: ExecutionUndoSnapshot | null };
  const [resumedReviewTitle, setResumedReviewTitle] = useState('');
  const [resultReview] = useState(() => new ResultReview<ExcelBeforeResult, ResultReceipt, ExcelResultMarks, ExcelNativeComparison>());
  /*
   * 刷新/重载后认领**上一次**遗留的审阅标记。
   *
   * 标记只存在于工作簿的条件格式里，内存里的会话对象随刷新消失；没有这一步，
   * 那些边框会一直留在表上，而加载项既不认识它们也不敢按区域批量删。
   * 认领只依据持久记录里的"规则 id + 本次随机公式"，对不上的规则一律不碰。
   */
  const recoveredMarks = useRef(false);
  useEffect(() => {
    if (recoveredMarks.current) return;
    recoveredMarks.current = true;
    void recoverExcelResultMarks().then(note => {
      if (note) appendMessage({ role: "system", text: note });
    }).catch(() => { /* 认领失败不能影响面板启动 */ });
  }, []);
  const resultReviewState = useSyncExternalStore(resultReview.subscribe, resultReview.getSnapshot, resultReview.getSnapshot);
  /*
   * 收尾状态**留在消息里**，不靠固定条保留。
   *
   * `ResultReviewBar`（共享组件）在确认/放弃的**成功终态**收起 —— 用户不再被一块
   * 重复"已确认变更"的板子挤掉输入区。但"收起"不等于"抹掉结论"：结果卡、保存
   * 为工具与撤销入口各自还在，只是少了一句"这件事已经定了"。这一条消息补的就是
   * 那句话；少了它，共享组件一改就会让 Excel 端连状态说明一起消失。
   *
   * 与 WPS 面板同一套语义（`settledLoggedRef`）：同一 `(key, phase)` 只记一次。
   */
  const settledLoggedRef = useRef<string | null>(null);
  useEffect(() => {
    const { phase, key } = resultReviewState;
    if (!key || (phase !== 'confirmed' && phase !== 'discarded')) return;
    const stamp = `${key}:${phase}`;
    if (settledLoggedRef.current === stamp) return;
    settledLoggedRef.current = stamp;
    appendMessage({
      role: 'assistant',
      text: phase === 'confirmed'
        ? '已确认变更：已接受当前结果并移除本次审阅标记；保存为工具与撤销入口可用。'
        : '已放弃并恢复：主表已按本次恢复点还原。'
    });
  }, [resultReviewState.phase, resultReviewState.key]);
  const worksheetCheckpoint = useWorksheetCheckpoint("excel");
  const [deletionCandidate, setDeletionCandidate] = useState<{ id: string; plan: AnalysisPlan; identity: string; authorization: string[]; conversationId: string } | null>(null);
  const issuedDeletions = useRef(new Set<string>());
  const [resultCandidate, setResultCandidate] = useState<{ id: string; plan: AnalysisPlan; authorization: string[]; identity: string; conversationId: string } | null>(null);
  const resultMessages = useRef(new Set<string>());
  const resultLive = useRef({ messages, workbookIdentity: '', authorization: '', conversationId: '', sourceMode: '' as string });
  const resultUndo = useRef<(id: string) => Promise<boolean>>(async () => false);
  // 只有用户主动点「重新排列」才会自增；AI 窗格据此重建自适应状态机并重套默认比例。
  // 消息更新、流式输出、回读、保存工具都不改它 —— 用户摆好的布局不会被反复抢回。
  const reviewPaneRearm = 0;
  const activeActivityMessageIdRef = useRef<string | null>(null);
  /**
   * 这一次「继续」已经开出的展示段（判据全在 ref 自己身上）。
   *
   * 一次继续动作可能分几段代码走：先接回仍在运行的原任务，任务结束后再取结果，
   * 两条分支都会要求"把进度落到当前回复"。它们属于**同一次继续**，只能有一段
   * 新回复；在 `stopTurn` 里放掉，所以下一次继续自然另开一段。
   */
  const resumeSegmentRef = useRef<{
    /**
     * 这一次继续的键。规划续跑用 `task:<turnId>`，早期分析（还没有服务端任务）
     * 用 `early:<原消息 id>`。**不伪造 turnId**：没有服务端任务就没有这个键值。
     */
    key: string;
    messageId: string;
    startedAt: number;
  } | null>(null);
  // 双栏变更对照的完整前后矩阵只活在这里（按 executionId 索引的会话内内存）。
  // 刻意不挂到 ChatMessage 上：消息会被持久化，矩阵不许进历史、也不许出站。
  const changeReviewsRef = useRef(new Map<string, RangeChangeReview[]>());
  // 活动进度管理（自定义 Hook）
  const activityHook = useActivityProgress({
    onPersistLog: (log, persistenceKey) => {
      const targetId = persistenceKey ?? activeActivityMessageIdRef.current;
      const binding = sdkBridge.binding();
      setMessages((messages) =>
        persistActivityMessage(messages, log, targetId, {
          id: targetId ?? messageId(),
          role: "assistant",
          activityLog: log,
          createdAt: new Date().toISOString()
        }, binding && binding.taskId === log.operationId ? {
          taskId: binding.taskId,
          hostId: binding.hostId as "excel-office" | "wps-et",
          workbookName: binding.workbookName,
          workbookIdentity: binding.workbookIdentity,
          sourceFingerprint: binding.sourceFingerprint
        } : undefined)
      );
      if (activeActivityMessageIdRef.current === targetId) {
        activeActivityMessageIdRef.current = null;
      }
    }
  });
  const {
    activity,
    activitySeconds,
    startActivity: startActivityProgress,
    resumeActivity: resumeActivityProgress,
    bindActivityOperation,
    getActivitySnapshot,
    advanceActivity,
    advanceActivityStage,
    completeCurrentStep,
    applyProgressEvent,
    updateActivityDetail,
    markActivityInterrupted,
    markActivityReconnecting,
    markActivityConnected,
    completeActivity: finishActivity
  } = activityHook;

  // SDK 任务工具桥：规划流里的 bridgeTool 指针 → 本机轮询/读表/回执。
  // 只读，不碰任何写入 API；值只经本机回环进出。
  const sdkBridge = useSdkToolBridge();
  // 执行前的授权复核要用**当前**范围，不能在开始规划时快照死。
  const bridgeLiveContextRef = useRef<
    () => SdkBridgeLiveContext | Promise<SdkBridgeLiveContext>
  >(() => emptyBridgeLiveContext());

  // 服务健康状态为模型管理提供模型目录刷新能力。
  const serviceHealthHook = useServiceHealth();
  const {
    serverOnline,
    serviceHealth,
    modelOptions,
    modelCatalogStatus,
    modelCatalogError,
    defaultCatalogModelId,
    refreshServiceState,
    retryModelCatalog,
    markServerOnline,
    markServerOffline
  } = serviceHealthHook;

  const modelManagementHook = useModelManagement({
    refreshServiceHealth: refreshServiceState
  });
  const {
    selectedModelId,
    modelSettings,
    apiKeyDraft,
    showApiKey,
    connectionDraft,
    pendingDeleteConnectionId,
    settingsSaving,
    settingsTesting,
    settingsLoading,
    settingsFeedback,
    connectionTestFeedback,
    connectionVerified,
    modelGuideDismissed,
    selectModel,
    dismissModelGuide,
    openSettings,
    openConnectionCreator,
    saveApiKey,
    editModelConnection,
    verifyConnection,
    saveConnection,
    removeConnection,
    saveFormulaModel,
    migrateModelSelection,
    setApiKeyDraft,
    setShowApiKey,
    setConnectionDraft,
    setPendingDeleteConnectionId,
    setSettingsLoading,
    setSettingsTesting,
    setSettingsFeedback
  } = modelManagementHook;

  const toolManagementHook = useToolManagement({
    workbook,
    onToolDslCopyError: () => {
      appendMessage({
        role: "system",
        text: "复制专家计划失败，请手动选择计划内容后复制。"
      });
    }
  });
  const {
    tools,
    queryTools,
    selectedToolId,
    selectedQueryToolId,
    toolDrawerView,
    toolDetailMode,
    pendingToolDeletion,
    copiedToolDslId,
    toolParameterValues,
    saveTool,
    saveQueryTool,
    requestToolDeletion,
    confirmToolDeletion,
    resetToolDrawer,
    openWorkflowToolDetail,
    openQueryToolDetail,
    selectTool,
    fieldOptions,
    updateToolParameter,
    copyToolDsl,
    setToolDrawerView,
    setToolDetailMode,
    setPendingToolDeletion,
    setTools
  } = toolManagementHook;

  const [toolActionBusy, setToolActionBusy] = useState<
    "prepare" | "query" | null
  >(null);

  const [contextOpen, setContextOpen] = useState(false);
  const [sheetSearch, setSheetSearch] = useState("");
  const [folderSheetSearch, setFolderSheetSearch] = useState("");
  const [activeFolderFileId, setActiveFolderFileId] = useState<string | null>(
    null
  );
  const [folderPickerView, setFolderPickerView] = useState<"files" | "sheets">(
    "files"
  );
  const [showSelectedFolderSheets, setShowSelectedFolderSheets] =
    useState(false);
  const [recentlyAddedFolderSheetKeys, setRecentlyAddedFolderSheetKeys] =
    useState<string[]>([]);
  const [folderRefreshNotice, setFolderRefreshNotice] = useState("");
  const [themePanelOpen, setThemePanelOpen] = useState(false);
  const [folderBackupOpen, setFolderBackupOpen] = useState(false);

  // 斜杠命令自动补全状态
  const [showSlashAutocomplete, setShowSlashAutocomplete] = useState(false);
  const [slashFilter, setSlashFilter] = useState("");
  // slashMode：command=选一级命令；model=已进入 /model 的模型选择二级菜单。
  const [slashMode, setSlashMode] = useState<"command" | "model">("command");

  // UI 状态管理
  const {
    toolsOpen, setToolsOpen,
    historyOpen, setHistoryOpen,
    settingsOpen, setSettingsOpen,
    closeAllDrawers,
    modelMenuOpen, setModelMenuOpen,
    moreMenuOpen, setMoreMenuOpen,
    isRuleManagerOpen, setIsRuleManagerOpen,
    closeAllMenus,
    petVisible, setPetVisible,
    focusOpening, setFocusOpening,
    isNarrowPane, setIsNarrowPane,
    widenStepDone, setWidenStepDone,
    workDirectoryGuideDismissed, dismissWorkDirectoryGuide,
    composerHeight, setComposerHeight,
    comparisonPaneNeedsResize,
    togglePetVisibility
  } = useUIState(excelReviewBlocksNewCopy(nativeComparison.review) || nativeComparison.pending,
    excelReviewBlocksNewCopy(nativeComparison.review) ? nativeComparison.review?.session : undefined, reviewPaneRearm);

  const themeApi = useTheme();
  const logoLongPress = useLongPress(() => setThemePanelOpen(true));

  // 执行验收与工具固化批准状态（自定义 Hook）
  const {
    saveCandidate,
    approveFixedContent,
    setApproveFixedContent,
    approveDestructive,
    setApproveDestructive,
    verifiedPlanIds,
    markPlanVerified,
    unmarkPlanVerified,
    saveEligibility,
    beginSaveCandidate,
    closeSaveCandidate
  } = useExecutionApproval();

  // 复制反馈状态与定时器（自定义 Hook）
  const {
    copiedMessageId,
    copiedFunctionPreviewId,
    copyMessageText,
    copyFunctionFormula
  } = useCopyFeedback({
    onMessageCopyError: () => {
      appendMessage({
        role: "system",
        text: "复制失败，请选中文字后手动复制。"
      });
    },
    onFormulaCopyError: (messageId) => {
      markFunctionPreview(messageId, {
        targetError: "复制失败，请手动选择公式文本复制。"
      });
    }
  });
  // 数据范围与工具保存字段状态（自定义 Hook）
  const {
    toolName,
    setToolName,
    toolDescription,
    setToolDescription,
    selectedSheetNames,
    committedSheetNames,
    committedScopeMode,
    scopeDraftDirty,
    commitAuthorizedScope,
    setSelectedSheetNames,
    selectionConfirmed,
    setSelectionConfirmed,
    sourceMode,
    workbookScopeMode,
    setWorkbookScopeMode,
    folderCatalog,
    setFolderCatalog,
    folderSheetKeys,
    setFolderSheetKeys,
    applyWorkbookSnapshotSelection,
    toggleSheet,
    toggleFolderSheet,
    selectAllSheetsInFile,
    clearSheetsInFile,
    applyFolderCatalog,
    chooseAutomaticScope,
    chooseManualScope,
    adoptCurrentWorkbookScope,
    chooseFolderScope,
    folderSelections,
    selectedNamesFor,
    selectAllSheets,
    clearSelectedSheets
  } = useScopeSelection();
  /**
   * 当前**已确认**的授权表名，随渲染刷新。
   *
   * 桥的执行前复核要拿"此刻的授权"去比，而这个"此刻"指的是用户当前已确认的
   * 范围，不是规划启动的那一刻。用 ref 而不是闭包，正是为了让复核读到最新值：
   * 等待期间用户收窄/撤销授权后，这里必须已经变窄。
   */
  const committedSheetNamesRef = useRef<readonly string[]>(committedSheetNames);
  committedSheetNamesRef.current = committedSheetNames;
  // Restoring a server session may replay explicit consent, never picker drafts
  // or localStorage selections left over from another panel session.
  const confirmedFolderScopeRef = useRef<{ root: string; keys: string[] } | null>(null);

  /**
   * 工作目录的**草稿 vs 已提交**。
   *
   * 草稿是 `folderSheetKeys`（勾选 / 全选 / 清空都只改它，localStorage 里存的也是它）；
   * 已提交只在上面那个 ref 里，**只有「确认」才写**。判据用共享层的
   * `folderScopeDraftDirty` —— 与 WPS 同一份，两端不许各写一套。
   *
   * 现读 ref、不缓存：它的每一次写入都与状态更新同批发生（`setWorkbook` /
   * `setSelectedSheetNames` / `setSelectionConfirmed` / `setContextOpen`），
   * 所以本次渲染读到的就是最新值，不存在"读到出生那一刻的事实"的问题。
   *
   * 根不一致一律按"没有已提交范围"处理：换了目录，旧根上的授权不许再当新目录的判据。
   */
  function intentScopeFingerprint(scope: IntentScopeContext): string {
    // This identity stays local. No absolute folder paths are added to intent
    // requests; the existing selected-field/data-row privacy contract is unchanged.
    return sharedIntentScopeFingerprint(scope, folderCatalog?.folderPath);
  }
  const retrySendRef = useRef<(text: string, understanding?: Awaited<ReturnType<typeof understandMessage>>) => void>(() => {});
  const [scopeConfirmationError, setScopeConfirmationError] = useState("");
  const targetGuard = useAuthorizedTarget({
    enabled: sourceMode === "workbook" && isRunningInExcel() && workbook !== null,
    reviewActive: excelReviewBlocksNewCopy(nativeComparison.review),
    monitorPaused: status !== "idle",
    selectionPending: () => immediateScope.isPending(),
    /*
     * 守卫读的是**已提交**范围，不是弹层里的草稿。
     *
     * 口径（跟随当前 / 固定多表）也是授权的一部分：`auto` 的授权是"这个文件的
     * 活动表"，`manual` 的是"这几张表"。旧实现读 `workbookScopeMode` ——
     * 那是**草稿**口径，用户只是在弹层里点一下「跟随当前」再点「取消」，守卫
     * 立刻改用活动表去比已提交的多表范围，于是刚确认完就被判成范围变化。
     * 与 WPS 的 `scopeMode`（已提交）/ `draftMode`（草稿）是同一份纪律。
     */
    inspect: () => inspectOfficeAuthorizedTarget(
      committedScopeMode === "auto" ? null : committedSheetNames
    ),
    analysisEvidence: officeAnalysisTargetEvidence,
    adopt: (target) => {
      const names = target.sheets.filter((sheet) => target.selected.includes(sheet.id)).map((sheet) => sheet.name);
      // 采用最新勾选的口径；异步提交前由 isCurrent 核实本次选择仍有效。
      if (workbook) adoptCurrentWorkbookScope(workbook, names, workbookScopeMode);
      if (targetGuard.binding && checkAuthorizedTarget(targetGuard.binding, target) === "same") return;
      localMappingDraftsRef.current.clear();
      clearUndoSnapshot();
      planEditorDialogRef.current?.close();
      planEditorDialogRef.current = null;
      setChatHistory((current) => ({ ...current, conversations: current.conversations.map((conversation) =>
        ({ ...conversation, messages: conversation.messages.map(retireTargetMessage) })) }));
    }
  });
  const [scopeCheckingCommit, setScopeCheckingCommit] = useState(false);
  const scopeCommitLock = useRef(false);
  // Private identity of the catalog visibly opened in this picker, not its filename.
  const scopePickerFile = useRef<Pick<AuthorizedTarget, "fileKey" | "fileRef"> | null>(null);
  const [scopeCatalogChecking, setScopeCatalogChecking] = useState(false);
  const [scopeCatalogRevision, setScopeCatalogRevision] = useState(0);
  const scopeAdmissionIds = sourceMode === "folder" ? folderSheetKeys : selectedSheetNames;
  const scopeAdmission = useScopeAdmission({
    enabled: workbook !== null && !scopeCatalogChecking,
    scopeKey: JSON.stringify([sourceMode, workbookScopeMode, workbook?.name, workbook?.capturedAt,
      folderCatalog?.sessionId, folderCatalog?.folderPath, scopeCatalogRevision]),
    ids: scopeAdmissionIds,
    inspect: async () => sourceMode === "folder"
      ? folderCatalog ? checkFolderAdmission(folderCatalog.sessionId, folderSelections())
        : { items: folderSheetKeys.map((id) => admissionItem(id, "missing")) }
      : isRunningInExcel() ? scopePickerFile.current
        ? inspectOfficeScopeAdmission(selectedSheetNames, scopePickerFile.current)
        : { items: selectedSheetNames.map((id) => admissionItem(id, "identity_unverified")) }
        : { items: selectedSheetNames.map((id) => admissionItem(id, "host_write_blocked")) }
  });
  const immediateScope = useImmediateScopeAuthorization({
    ready: !scopeCatalogChecking && (scopeAdmissionIds.length === 0 || scopeAdmission.canCommit),
    apply: confirmSheetSelection
  });
  async function openScopePicker() {
    if (immediateScope.isPending()) { setContextOpen(true); return; }
    setScopeConfirmationError("");
    setContextOpen(true);
    setFolderBackupOpen(false);
    setSheetSearch("");
    setFolderSheetSearch("");
    if (sourceMode === "folder") {
      if (folderCatalog && !busy) void refreshCurrentFolder({ quiet: true });
      return;
    }
    if (isRunningInExcel()) {
      setScopeCatalogChecking(true);
      try {
        const fileKey = currentOfficeWorkbookIdentity() ?? null;
        const fileRef = Office.context.document;
        const catalog = await captureWorkbookCatalog();
        if (fileKey !== (currentOfficeWorkbookIdentity() ?? null) || fileRef !== Office.context.document) {
          throw new Error("catalog target changed");
        }
        scopePickerFile.current = { fileKey, fileRef };
        setScopeCatalogRevision((value) => value + 1);
        setWorkbook(catalog);
        if (workbookScopeMode === "auto" && (selectedSheetNames.length !== 1 || selectedSheetNames[0] !== catalog.activeWorksheet)) {
          setSelectedSheetNames([catalog.activeWorksheet]);
        }
        else if (workbookScopeMode === "manual" && targetGuard.notice?.reason === "file_changed") setSelectedSheetNames([]);
      } catch {
        scopePickerFile.current = null;
        setScopeConfirmationError("暂时无法读取目标，请刷新选择器。");
      } finally { setScopeCatalogChecking(false); }
    }
  }

  // 备份管理目标：文件夹模式用 session，工作簿模式从文件路径推导
  const folderBackupTarget = useMemo((): FolderBackupTarget => {
    if (sourceMode === "folder" && folderCatalog) {
      return {
        kind: "session",
        sessionId: folderCatalog.sessionId,
        folderPath: folderCatalog.folderPath
      };
    }
    if (isRunningInExcel()) {
      const location = workbookFolderFromDocumentUrl(
        (Office as any).context?.document?.url
      );
      if (location.kind === "local") {
        return { kind: "path", folderPath: location.folderPath };
      }
      return { kind: "unavailable", reason: location.kind };
    }
    return { kind: "unavailable", reason: "unsaved" };
  }, [folderCatalog, sourceMode]);

  const composerInputRef = useRef<HTMLTextAreaElement>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);
  const clarificationImagesRef = useRef<Map<string, PendingImage[]>>(
    new Map()
  );
  const localMappingDraftsRef = useRef(new Map<string, {
    rule: ColumnMappingRule;
    context: { summary: string; warnings: string[]; assumptions: string[]; provider: "model" | "local" };
    values: readonly CellValue[];
    allowExactMapping: boolean;
    /** 停在哪种缺口上。`"blank"` 是**读到真实空白之后**才出现的。 */
    reason?: "unmapped" | "separator_or_mapping" | "blank";
    /** `reason === "blank"` 时真实读到的空白格数量。 */
    blankCount?: number;
    scopeFingerprint: string;
    authorizedSheets: string[];
    workbookIdentity: string | undefined;
    pendingRule?: ColumnMappingRule;
    revision: number;
  }>());
  const activityScrollSignal = [
    messages.length,
    messages.length > 0 ? messages[messages.length - 1]?.id ?? "" : "",
    activity?.operationId ?? "",
    activity?.sequence ?? -1,
    activity?.steps.length ?? 0,
    activity?.current ?? "",
    activity?.status ?? "",
    activity?.connectionState ?? ""
  ].join(":");
  const {
    containerRef: messageStreamRef,
    awayFromBottom,
    handleScroll: handleMessageStreamScroll,
    scrollToLatest
  } = useMessageAutoFollow(activityScrollSignal, chatHistory.activeConversationId);
  const focusDialogRef = useRef<Office.Dialog | null>(null);
  const planEditorDialogRef = useRef<Office.Dialog | null>(null);
  const queryAbortRef = useRef<AbortController | null>(null);
  // turnAbortRef：整个 turn（需求确认→本地取数→生成方案）共用的打断句柄。
  // 每个 turn 开始时新建；用户中途插话/停止时 abort，串在链上的模型请求会一起 reject。
  const turnAbortRef = useRef<AbortController | null>(null);
  const pauseRequestedControllerRef = useRef<AbortController | null>(null);
  /**
   * 暂停请求到了哪一步 —— 与 WPS 端同一套语义（三件事分开，不许互相冒充）：
   * `requesting`（已发出、没有答案）/ `local-stopped`（本端已停止接收）/
   * `acknowledged`（服务端接受请求，是否已停下仍看任务状态）/ `unconfirmed`
   * （失败或超时，服务端状态未知）。
   */
  const pauseRequestGenerationRef = useRef(0);
  const [pauseRequestState, setPauseRequestState] = useState<{
    phase: "requesting" | "local-stopped" | "acknowledged" | "unconfirmed";
    detail: string;
  } | null>(null);
  const prePlanningPause = usePrePlanningPause<{ prompt: string; images: PendingImage[];
    activityMessageId?: string; messageUnderstanding?: Awaited<ReturnType<typeof understandMessage>> }>(
    JSON.stringify([chatHistory.activeConversationId, sourceMode, committedScopeMode, committedSheetNames]));
  /*
   * 续跑并发闸：**记的是"哪一次执行持有它"**，不是"闸正在用"。
   *
   * 旧写法是布尔 `true/false`：被取代的那次执行在 `finally` 里无条件写 `false`，
   * 于是它一旦晚于新一轮收尾，就把**新一轮**的闸也放掉了（下一轮可以再并发进来）。
   * 换成持有者之后，只有自己仍持有才放。
   */
  const planningResumeInFlightRef = useRef<AbortController | null>(null);
  // pendingSteerRef：转向时先掐掉当前 turn，把新话暂存这里；等 busy 回落到 idle，
  // 由 effect 带着这句新话重新发起一个 turn（硬转向 = abort + 带新话重跑）。
  const pendingSteerRef = useRef<string | null>(null);
  // steerBasePromptRef：转向时被打断那一句的原始意图。被打断的 turn 没跑完，
  // 不会沉淀出 priorIntent/priorResult，导致重跑时新话（多为承接式追问）失去锚点。
  // 这里保留原句，重跑时并入新话，让 checkIntent 收到完整意图。
  const steerBasePromptRef = useRef<string | null>(null);
  // 结论复用缓存（见 CachedConclusion）。intentKey → 结论；normalizedPrompt → intentKey。
  const resultCacheRef = useRef<Map<string, CachedConclusion>>(new Map());
  const promptKeyCacheRef = useRef<Map<string, string>>(new Map());
  // "仍要重新计算"按钮设为 true，绕过命中判定并在重算后覆写缓存。
  const forceRecomputeRef = useRef(false);
  // sendMessage 起点记录的原始用户文本，写缓存时作为一级 prompt key。
  const rawPromptRef = useRef("");
  const folderRecoveryAttachStartedRef = useRef(false);
  const officeRecoveryAttachStartedRef = useRef(false);
  /**
   * `plan.id` → 产出这份计划的规划任务。
   *
   * 登记 Office 执行时必须告诉服务端"这是哪个任务产生的写入"，否则
   * `has_execution_for_task` 在 Excel 侧永远是假，「继续分析」就发现不了
   * "这个任务其实已经写过原表"。计划卡可从历史重放、`/function` 预览也会
   * 造计划 —— 那些来源没有映射，取不到就不声称关联（**不猜**）。
   */
  const planningTurnByPlanIdRef = useRef<Map<string, string>>(new Map());

  // 界面展示的模型必须与请求实际发送的 modelId 完全一致：这里只认目录中
  // 仍然有效的真实连接，解析不出有效选择时进入“未配置”状态，绝不回退到
  // 目录首项或已移除的 local 占位项。
  //
  // 这个派生值定义在斜杠命令之前：`/model` 二级菜单要靠它标注当前生效项，
  // 而 useMemo 的调用顺序必须与渲染顺序一致，不能挪到使用点之后再声明。
  const effectiveModelId = useMemo(
    () =>
      chooseAvailableModel(
        modelOptions,
        selectedModelId,
        serviceHealth?.model ?? defaultCatalogModelId
      ),
    [modelOptions, selectedModelId, serviceHealth?.model, defaultCatalogModelId]
  );
  const selectedModel = useMemo(
    () => modelOptions.find((option) => option.id === effectiveModelId),
    [modelOptions, effectiveModelId]
  );

  useEffect(() => {
    if (!newConversationNotice) return;
    const timer = window.setTimeout(() => setNewConversationNotice(null), 1500);
    return () => window.clearTimeout(timer);
  }, [newConversationNotice]);

  // 一级斜杠命令定义
  const slashCommands: SlashCommand[] = useMemo(
    () => [
      {
        value: "help",
        description: "查看使用说明与模型配置步骤"
      },
      {
        value: "function",
        description: "AI 生成原生 Excel 公式（快捷输入，智能补全）"
      },
      {
        value: "model",
        description: "切换模型（使用你已配置的 API 连接）"
      }
    ],
    []
  );

  // /model 二级菜单：把已配置的模型连接映射成候选项，标注当前生效项。
  const slashModelCommands: SlashCommand[] = useMemo(
    () =>
      modelOptions.map((option) => ({
        value: option.id,
        label: option.label,
        showSlashPrefix: false,
        active: option.id === effectiveModelId,
        disabled: !agentModelAvailable(option),
        description: agentModelAvailable(option)
          ? option.supportsVision
            ? "支持图片输入"
            : "文本模型"
          : agentModelUnavailableReason(option)
      })),
    [modelOptions, effectiveModelId]
  );

  // `busy` = 交互锁定：它多含 `phase === 'reviewing'`（结果已就绪、等用户确认/放弃），
  // 用来防重复执行与操作互斥。**不能**拿它当"EB 正在跑" —— 那会让悬浮标在只等用户
  // 查看结果时也显示三点动画。
  const operationBusy = worksheetCheckpoint.pending || status !== "idle" || nativeComparison.pending || resultReviewState.pending || resultReviewState.phase === 'reviewing';
  const busy = immediateScope.pending || operationBusy;
  /*
   * 「实际运行状态」与「交互锁定状态」分开（用户口径）：
   *  - 三点动画：模型请求（status）、生成对比（nativeComparison.pending）、
   *    确认收尾 / 放弃恢复 / 清理标记（resultReviewState.pending）等**真的在途**；
   *  - 锚图标：只是结果已生成、在等用户查看/确认/放弃。
   *
   * `resultReviewState.pending` 已覆盖审阅通道自己的每个操作（start / confirm /
   * discard / compare / retryMarkerCleanup / settleExternally，见 `resultReview.ts`
   * 的 `finally { pending:false }`），所以这里**不减** `reviewing`、也不加新条件：
   * 减掉 `reviewing` 就是答案 —— `reviewing` 是"结果在那儿、等人"，不是"还在跑"。
   * `busy` 本身原样保留，防重复执行与互斥边界不受影响。
   */
  const running = worksheetCheckpoint.pending || status !== "idle" || nativeComparison.pending || resultReviewState.pending;
  /*
   * 新一轮一开始就收掉上一次暂停请求的结论：那条小字说的是「上一轮停到哪一步」，
   * 留着它会让人以为**这一轮**也已经停了。
   */
  useEffect(() => {
    if (running) {
      pauseRequestGenerationRef.current += 1;
      setPauseRequestState(null);
    }
  }, [running]);

  // 目录成功加载后迁移旧的 local 记录和已删除连接；目录读取失败时必须保留
  // 原有选择，不能依据空列表删掉仍然有效的记录。
  useEffect(() => {
    if (modelCatalogStatus !== "ready") return;
    migrateModelSelection({ modelOptions });
  }, [migrateModelSelection, modelCatalogStatus, modelOptions]);

  useEffect(() => {
    if (!folderCatalog || sourceMode !== "folder") {
      setFolderBackupOpen(false);
    }
  }, [folderCatalog, sourceMode]);

  useEffect(() => {
    if (settingsOpen || toolsOpen || historyOpen || isRuleManagerOpen) {
      setFolderBackupOpen(false);
    }
  }, [historyOpen, isRuleManagerOpen, settingsOpen, toolsOpen]);

  useEffect(() => {
    if (!petVisible) return;

    const handlePetClick = (event: MouseEvent) => {
      const target = event.target as HTMLElement | null;
      console.log("[Pet] Click detected:", target);
      if (!target?.closest(".pet-avatar")) return;
      const animation =
        PET_ANIMATIONS[Math.floor(Math.random() * PET_ANIMATIONS.length)];
      console.log("[Pet] Animating with:", animation);
      animatePet(animation);
    };

    let idleTimer: ReturnType<typeof setTimeout> | null = null;
    const resetIdleTimer = () => {
      if (idleTimer) clearTimeout(idleTimer);
      const pet = document.querySelector<HTMLElement>(".pet-avatar");
      pet?.classList.remove("turtle-sleepy");
      idleTimer = setTimeout(() => {
        pet?.classList.add("turtle-sleepy");
      }, 60000);
    };

    document.addEventListener("click", handlePetClick);
    document.addEventListener("keydown", resetIdleTimer);
    document.addEventListener("scroll", resetIdleTimer, true);
    resetIdleTimer();

    return () => {
      document.removeEventListener("click", handlePetClick);
      document.removeEventListener("keydown", resetIdleTimer);
      document.removeEventListener("scroll", resetIdleTimer, true);
      if (idleTimer) clearTimeout(idleTimer);
    };
  }, [petVisible]);

  // 撤销栈管理（自定义 Hook）
  const {
    undoStack,
    lastUndoSnapshot,
    pushUndoSnapshot,
    clearUndoSnapshot,
    undoLastExecution,
    undoExecutionTo
  } = useUndoSnapshot({
    isBusy: status !== 'idle' || nativeComparison.pending,
    beforeOperation: targetGuard.check,
    onStatusChange: setStatus,
    onMessage: (text) => {
      appendMessage({
        role: "system",
        text
      });
    },
    onAfterUndo: () => { void scan(); },
    onRestored: (snapshot) => {
      /*
       * 从**结果卡**撤销之后，固定操作条必须同步收尾：否则它会继续显示
       * "请查看主表中的执行结果"并提供「确认变更」，而那份结果已经被还原。
       * `settleExternally` 只同步状态、不再次恢复；在 `discard()` 自己的
       * 恢复流程里因为 `pending` 在位会被它自己挡掉，不会重复处理副本。
       */
      if (snapshot.executionId) void resultReview.settleExternally(snapshot.executionId);
      setMessages((current) => {
        const marking = markWorkbookExecutionUndone(current, snapshot);
        if (marking.ambiguous) {
          appendMessage({
            role: "system",
            text: "本次执行已恢复，但无法确认它对应哪一条历史记录，未自动标记「已撤销」。"
          });
        }
        return marking.messages;
      });
    }
  });
  resultUndo.current = async executionId => {
    if (undoStack.at(-1)?.executionId !== executionId) throw new Error('本次执行不是当前恢复栈顶，未撤销其他任务。');
    return (await undoExecutionTo(executionId)) === true;
  };
  useEffect(() => {
    if (!deletionCandidate || busy || deletionCandidate.conversationId !== chatHistory.activeConversationId || issuedDeletions.current.has(deletionCandidate.id)) return;
    const candidate = deletionCandidate;
    issuedDeletions.current.add(candidate.id);
    void runPlan(candidate.plan, undefined, false, false, undefined, {
      plan: candidate.plan, workbookIdentity: candidate.identity, authorization: candidate.authorization,
      requestMessageId: candidate.id
    });
  }, [deletionCandidate, busy, chatHistory.activeConversationId]);
  let resultWorkbookIdentity = '';
  try { resultWorkbookIdentity = currentOfficeWorkbookIdentity() ?? ''; } catch { /* Unavailable identity must block writes, not crash rendering. */ }
  resultLive.current = { messages, workbookIdentity: resultWorkbookIdentity,
    authorization: JSON.stringify(workbook ? [...selectedNamesFor(workbook)].sort() : []),
    conversationId: chatHistory.activeConversationId, sourceMode };
  /*
   * 刷新后只查**当前会话 + 当前工作簿**里已有操作的状态。
   * 别的会话/工作簿的记录不会在这里被查询，也就不会冒出一张常驻的删除卡；
   * 它们仍然可以从「备份与恢复」按工作簿身份找到。
   */
  useEffect(() => {
    if (!resultWorkbookIdentity) return;
    void worksheetCheckpoint.hydrate({
      conversationId: chatHistory.activeConversationId,
      workbookIdentity: resultWorkbookIdentity
    });
  }, [chatHistory.activeConversationId, resultWorkbookIdentity]);
  useEffect(() => {
    if (!resultCandidate || status !== 'idle' || nativeComparison.pending || resultReviewState.pending || resultReviewState.phase === 'reviewing') return;
    const candidate = resultCandidate;
    setResumedReviewTitle('');
    const planKey = JSON.stringify(candidate.plan), scope = JSON.stringify([...candidate.authorization].sort());
    const isCurrent = () => {
      const now = resultLive.current;
      return now.workbookIdentity === candidate.identity && currentOfficeWorkbookIdentity() === candidate.identity &&
        now.authorization === scope && now.conversationId === candidate.conversationId &&
        now.messages.some(message => message.id === candidate.id && JSON.stringify(message.plan) === planKey);
    };
    let executionPlan = candidate.plan;
    void resultReview.start(candidate.id, {
      isCurrent,
      prepare: async () => {
        if (sdkBridge.failure()) throw new Error('本轮本地读取未通过现场核对，请重新读取后生成结果；原表未写入。');
        if (!(await targetGuard.check()) || !isCurrent()) throw new Error('当前计划或授权已变化，未执行。');
        unmarkPlanVerified(candidate.plan.id);
        executionPlan = await prepareExcelValueCopies(candidate.plan, candidate.authorization, candidate.identity, isCurrent);
        const issue = resultReviewPlanIssue(executionPlan, candidate.authorization);
        if (issue) throw new Error(issue);
        if (!isCurrent()) throw new Error('当前计划或授权已变化，未执行。');
        return captureExcelBeforeResult(executionPlan, candidate.authorization);
      },
      execute: async () => {
        let observed: PlanExecutionResult | null = null;
        const succeeded = await runPlan(executionPlan, undefined, false, false, undefined,
          { plan: executionPlan, authorization: candidate.authorization, workbookIdentity: candidate.identity,
            requestMessageId: candidate.id },
          result => { observed = result; });
        const result = observed as PlanExecutionResult | null;
        if (!succeeded || !result) throw new Error('执行未完成，请查看执行记录和恢复入口；不会自动重试。');
        return { executionId: result.undoSnapshot?.executionId ?? '', verified: result.verification.status === 'verified',
          undoable: result.undoSnapshot?.postStateCoverage === 'complete', undo: result.undoSnapshot ?? null };
      },
      mark: () => markExcelResult(executionPlan, candidate.identity),
      clearMarker: clearExcelResultMarks,
      showBefore: before => showExcelBeforeResult(before, candidate.identity, executionPlan, candidate.authorization),
      comparisonNotice: copy => copy.comparisonNotice ?? null,
      closeBefore: async copy => {
        const outcome = await closeExcelNativeComparison(copy);
        return outcome === 'closed' || outcome === 'already_closed' ? null : '执行前对比副本仍保留，请手动核对；原表不会被关闭。';
      },
      restore: async receipt => {
        if (!(await resultUndo.current(receipt.executionId))) throw new Error('恢复未成功，当前结果仍保留。');
      },
      validateResult: async receipt => {
        if (!receipt.undo) throw new Error('恢复记录不可用，未确认变更。');
        await Excel.run(context => assertUndoPostState(context, receipt.undo!));
      },
      accepted: () => markPlanVerified(candidate.plan.id)
    });
  }, [resultCandidate, status, nativeComparison.pending, resultReviewState.pending, resultReviewState.phase]);
  const {
    folderUndoStack,
    pushFolderUndo,
    reconcileFolderUndoIds,
    undoFolderExecutionTo
  } = useFolderUndo({
    isBusy: busy,
    onRestored: (entries) =>
      setMessages((current) => {
        const marking = markFolderExecutionUndone(current, entries);
        if (marking.ambiguous) {
          appendMessage({
            role: "system",
            text: "本次文件夹执行已恢复，但无法确认它对应哪一条历史记录，未自动标记「已撤销」。"
          });
        }
        return marking.messages;
      }),
    onStatusChange: setStatus,
    onMessage: (text) => {
      appendMessage({
        role: "system",
        text
      });
    },
    onAfterUndo: async (entry) => {
      const refreshed = await refreshFolder(entry.sessionId);
      applyFolderRefreshResult(refreshed, true);
    }
  });
  // 目录成功加载且确实没有真实连接 → 未配置；与“加载中/加载失败”分开表达。
  const supportsVision = selectedModel?.supportsVision === true;
  const hasConfiguredModel = modelOptions.length > 0;
  const modelCatalogFailed = modelCatalogStatus === "failed";
  const modelCatalogLoading = modelCatalogStatus === "loading";
  const modelNotConfigured =
    modelCatalogStatus === "ready" && !hasConfiguredModel;
  const headerStatusClassName = !serverOnline
    ? ""
    : modelCatalogFailed
      ? "offline"
      : effectiveModelId && serviceHealth?.configured
        ? "online model-online"
        : "online";
  const headerStatusText = !serverOnline
    ? "Excel Bro 后台未连接"
    : modelCatalogFailed
      ? "无法读取模型配置"
      : modelCatalogLoading
        ? "正在读取模型配置"
        : modelNotConfigured
          ? "尚未配置模型"
          : "";
  const showFirstModelGuide =
    serverOnline &&
    modelCatalogStatus === "ready" &&
    !hasConfiguredModel &&
    !modelGuideDismissed &&
    !settingsOpen;
  // 两步引导互斥、永不同屏：窄窗格且拉宽步未完成 → 只弹「拉宽」；
  // 拉过阈值(≥PANE_WIDEN_THRESHOLD)或点跳过、以及本就宽窗格 → 承接「模型」引导。
  const showWidenGuide =
    showFirstModelGuide && isNarrowPane && !widenStepDone;
  const showModelGuide =
    showFirstModelGuide && (!isNarrowPane || widenStepDone);
  const showWorkDirectoryGuide = shouldShowWorkDirectoryGuide({
    ready: serverOnline && modelCatalogStatus === "ready" && workbook !== null,
    previousStepComplete: hasConfiguredModel || modelGuideDismissed,
    folderConfigured: folderCatalog !== null,
    dismissed: workDirectoryGuideDismissed,
    blocked:
      settingsOpen ||
      folderBackupOpen ||
      contextOpen ||
      busy ||
      showWidenGuide ||
      showModelGuide
  });
  const hasEnvironmentModel =
    Boolean(modelSettings?.baseUrl) && Boolean(modelSettings?.defaultModel);
  const hasManagedModels = (modelSettings?.connections.length ?? 0) > 0;
  const editingConnection = connectionDraft?.id
    ? modelSettings?.connections.find(
        (connection) => connection.id === connectionDraft.id
      )
    : null;
  const filteredWorksheets = useMemo(
    () => {
      const query = sheetSearch.trim().toLocaleLowerCase();
      const sheets = workbook?.worksheets ?? [];
      return sheets.filter((sheet) =>
        sheet.name.toLocaleLowerCase().includes(query)
      );
    },
    [sheetSearch, workbook]
  );
  const folderSelectionCounts = useMemo(() => {
    const counts = new Map<string, number>();
    for (const key of folderSheetKeys) {
      const separator = key.indexOf("\0");
      if (separator <= 0) continue;
      const fileId = key.slice(0, separator);
      counts.set(fileId, (counts.get(fileId) ?? 0) + 1);
    }
    return counts;
  }, [folderSheetKeys]);
  const visibleFolderFiles = useMemo(() => {
    const files = folderCatalog?.files ?? [];
    return showSelectedFolderSheets
      ? files.filter((file) => (folderSelectionCounts.get(file.id) ?? 0) > 0)
      : files;
  }, [folderCatalog, folderSelectionCounts, showSelectedFolderSheets]);
  const activeFolderFile = useMemo(
    () =>
      visibleFolderFiles.find((file) => file.id === activeFolderFileId) ??
      visibleFolderFiles[0] ??
      null,
    [activeFolderFileId, visibleFolderFiles]
  );
  const activeFolderSheets = useMemo(
    () =>
      activeFolderFile?.worksheets.filter(
        (sheet) => !isEBSystemSheet(sheet.name)
      ) ?? [],
    [activeFolderFile]
  );
  const filteredFolderSheets = useMemo(() => {
    if (!activeFolderFile) return [];
    const query = folderSheetSearch.trim().toLocaleLowerCase();
    return activeFolderSheets.filter((sheet) => {
      const key = folderSheetKey(activeFolderFile.id, sheet.name);
      return (
        sheet.name.toLocaleLowerCase().includes(query) &&
        (!showSelectedFolderSheets || folderSheetKeys.includes(key))
      );
    });
  }, [
    activeFolderFile,
    activeFolderSheets,
    folderSheetKeys,
    folderSheetSearch,
    showSelectedFolderSheets
  ]);
  const selectedFolderFileCount = useMemo(
    () =>
      (folderCatalog?.files ?? []).filter(
        (file) => (folderSelectionCounts.get(file.id) ?? 0) > 0
      ).length,
    [folderCatalog, folderSelectionCounts]
  );
  const totalSelectableFolderSheets = useMemo(
    () =>
      (folderCatalog?.files ?? []).reduce(
        (total, file) =>
          total +
          file.worksheets.filter((sheet) => !isEBSystemSheet(sheet.name)).length,
        0
      ),
    [folderCatalog]
  );

  function clearActiveFolderFileSelection(): void {
    immediateScope.request();
    if (!activeFolderFile || !folderCatalog) return;
    const transition = resolveFolderViewAfterFileClear(
      folderSheetKeys,
      activeFolderFile.id,
      folderCatalog.files.map((file) => file.id),
      showSelectedFolderSheets
    );
    clearSheetsInFile(activeFolderFile.id);
    setShowSelectedFolderSheets(transition.showSelected);
    setActiveFolderFileId(transition.activeFileId);
    setFolderSheetSearch("");
    if (showSelectedFolderSheets && !transition.showSelected) {
      setFolderRefreshNotice(
        "已清空本文件，当前没有已选工作表，已显示全部。"
      );
    }
  }

  /** 清空整个目录的勾选并自动撤销授权。 */
  function clearFolderSelection(): void {
    immediateScope.request();
    if (folderSheetKeys.length === 0) return;
    setFolderSheetKeys([]);
    setRecentlyAddedFolderSheetKeys([]);
    if (showSelectedFolderSheets) {
      setShowSelectedFolderSheets(false);
    }
  }

  useEffect(() => {
    const files = folderCatalog?.files ?? [];
    if (files.length === 0) {
      setActiveFolderFileId(null);
      setFolderPickerView("files");
      return;
    }
    setActiveFolderFileId((current) =>
      current && files.some((file) => file.id === current)
        ? current
        : (files.find((file) => !file.error) ?? files[0]).id
    );
  }, [folderCatalog]);

  // 记录"对话历史写不进本地存储"是否已经提示过，避免每次状态变化都重复追加。
  const storageWriteFailedRef = useRef(false);

  useEffect(() => {
    // 落盘收口。`legacyCredentialConversationIds` 必须跟着历史一起写下去：
    // 它是"这条旧记录被脱敏过、需要新开会话"的持久标记，靠它才能在重开之后
    // 继续拦住续聊。用"文本里还能不能匹配到凭据"当判据是不行的——读取路径已经
    // 脱敏，重开后永远匹配不到，发送闸会不声不响地松开。
    const persist = (history: ChatHistoryState) => {
      try {
        localStorage.setItem(
          CHAT_STORAGE_KEY,
          JSON.stringify(
            redactHistoryCredentials({
              ...history,
              legacyCredentialConversationIds:
                remainingLegacyCredentialConversationIds(history)
            })
          )
        );
        return true;
      } catch {
        return false;
      }
    };
    const compact: ChatHistoryState = {
      activeConversationId: chatHistory.activeConversationId,
      legacyCredentialConversationIds:
        remainingLegacyCredentialConversationIds(chatHistory),
      conversations: chatHistory.conversations
        .slice(0, 8)
        .map((conversation) => ({
          ...conversation,
          messages: conversation.messages.slice(
            -PERSISTED_MESSAGES_PER_CONVERSATION
          )
        }))
    };
    if (!persist(chatHistory) && !persist(compact)) {
      // 存储改写失败时，**不能**退回"把原文留在内存里照常渲染/发送"的做法。
      // 内存里的历史在读入阶段已经脱敏过了，所以这里要做的不是"再脱一遍原文"
      // （原文此刻已不存在），而是明确告知用户：当前会话没有写进本地存储，
      // 关掉任务窗格后不会被保留。
      //
      // 用 ref 去重，避免每次 chatHistory 变化都追加一条同样的系统消息。
      if (!storageWriteFailedRef.current) {
        storageWriteFailedRef.current = true;
        setMessages((current) => [
          ...current,
          {
            id: messageId(),
            role: "system" as const,
            text:
              "无法把对话历史写入本机存储（浏览器或宿主可能禁止了本地存储）。本次会话仍可继续，但关掉任务窗格后不会被保留。",
            createdAt: new Date().toISOString()
          }
        ]);
      }
    } else {
      storageWriteFailedRef.current = false;
    }
  }, [chatHistory, setMessages]);

  useEffect(
    () => () => {
      focusDialogRef.current?.close();
      focusDialogRef.current = null;
      planEditorDialogRef.current?.close();
      planEditorDialogRef.current = null;
    },
    []
  );

  /*
   * 输入框的「上拉」：拖动上方把手调高度、内容变多时自动长高，长到上限之后
   * **可以滚动查看**（而不是把多出来的行切掉）。
   *
   * 这段逻辑原先只内联在这里，WPS 侧只好另写一套原生 `resize: vertical` —— 在
   * WPS 的 WebView 里既拖不动，又因为继承来的 `overflow-y: hidden` 看不到超出
   * 高度的那几行。现在两端共用同一个 hook，行为不会再各自漂移。
   */
  const { handleProps: composerResizeHandleProps } = useComposerResize({
    inputRef: composerInputRef,
    height: composerHeight,
    setHeight: setComposerHeight,
    value: prompt
  });

  useEffect(() => {
    if (folderRecoveryAttachStartedRef.current) return;
    folderRecoveryAttachStartedRef.current = true;
    const pending = loadPendingFolderOperation();
    if (!pending) return;
    const existingOperationMessage = chatHistory.conversations
      .flatMap((conversation) => conversation.messages)
      .find((message) => message.folderOperationId === pending.operationId);
    if (existingOperationMessage && !existingOperationMessage.folderResume) {
      clearPendingFolderOperation(pending.operationId);
      return;
    }

    const finishRecoveredResult = (
      result: FolderExecuteResult,
      activeSession: boolean
    ) => {
      appendFolderOperationOutcome(pending, result, true, activeSession);
      if (activeSession && result.undo) {
        pushFolderUndo({
          planId: pending.plan.id,
          undoId: result.undo.undoId,
          sessionId: result.undo.sessionId
        });
      }
      clearPendingFolderOperation(pending.operationId);
    };

    const attach = async () => {
      let persisted;
      try {
        persisted = await getFolderExecutionStatus(pending.operationId);
        markServerOnline();
      } catch (reason) {
        if (isLocalServiceConnectionError(reason)) {
          markServerOffline();
          return;
        }
        appendFolderOperationNotice(
          pending,
          `无法找回上次文件夹任务的状态：${
            reason instanceof Error ? reason.message : String(reason)
          }。没有创建新的执行，也没有重复写入；请在「备份与恢复」中核对文件。`
        );
        clearPendingFolderOperation(pending.operationId);
        return;
      }

      if (persisted.status === "completed") {
        if (persisted.result) {
          finishRecoveredResult(persisted.result, false);
        } else {
          appendFolderOperationNotice(
            pending,
            "上次任务记录为已完成，但完整结果缺失。没有重新执行；请在「备份与恢复」中核对文件。"
          );
          clearPendingFolderOperation(pending.operationId);
        }
        return;
      }

      if (
        persisted.status === "failed" ||
        persisted.status === "needs_reconciliation"
      ) {
        if (
          persisted.canResume &&
          typeof persisted.revision === "number" &&
          persisted.revision > 0
        ) {
          appendFolderOperationNotice(
            pending,
            "上次任务在安全断点处停止。已完成的拆分结果和备份仍可校验；继续时只发布尚未完成的文件，不会重新计算，也不会覆盖后来被修改的文件。",
            true,
            {
              operationId: pending.operationId,
              expectedRevision: persisted.revision,
              status: "pending"
            }
          );
          return;
        }
        appendFolderOperationNotice(
          pending,
          `上次文件夹任务已停止。任务状态：需要核对。写入状态：不能仅凭任务记录确定。验收状态：尚未完成。${
            persisted.error?.message ? `原因：${persisted.error.message}。` : ""
          }没有重新执行；请打开「备份与恢复」查看可恢复资料。`
        );
        clearPendingFolderOperation(pending.operationId);
        return;
      }

      const now = new Date().toISOString();
      setChatHistory((current) => {
        const targetConversationId = current.conversations.some(
          (conversation) => conversation.id === pending.conversationId
        )
          ? pending.conversationId
          : current.activeConversationId;
        return {
          ...current,
          activeConversationId: targetConversationId,
          conversations: current.conversations.map((conversation) => {
            if (conversation.id !== targetConversationId) return conversation;
            const existingIndex = conversation.messages.findIndex(
              (message) => message.id === pending.activityMessageId
            );
            const activityPlaceholder: ChatMessage = {
              id: pending.activityMessageId,
              role: "assistant",
              activityPending: true,
              createdAt: now
            };
            const nextMessages =
              existingIndex >= 0
                ? conversation.messages.map((message, index) =>
                    index === existingIndex
                      ? { ...message, activityPending: true }
                      : message
                  )
                : [...conversation.messages, activityPlaceholder];
            return { ...conversation, messages: nextMessages, updatedAt: now };
          })
        };
      });
      activeActivityMessageIdRef.current = pending.activityMessageId;
      startActivityProgress(
        "正在重新连接文件夹任务",
        "正在追赶原任务的进度，不会创建新的执行。",
        {
          persistenceKey: pending.activityMessageId,
          operationId: pending.operationId,
          phase: "executing",
          stepId: "execution.reattach"
        }
      );
      setActivePlanId(pending.plan.id);
      setStatus("executing");
      try {
        const result = await executeFolderPlan(pending.sessionId, pending.plan, {
          operationId: pending.operationId,
          onProgress: applyProgressEvent,
          onConnectionState: (connectionState, message) => {
            if (connectionState === "reconnecting") {
              markActivityReconnecting(
                message ?? "连接短暂中断，正在恢复原任务进度。"
              );
            } else if (connectionState === "interrupted") {
              markActivityInterrupted(
                message ?? "连接中断，执行结果待确认。"
              );
            } else {
              markActivityConnected();
            }
          }
        });
        finishRecoveredResult(result, true);
        finishActivity(
          result.verification.status === "verified" ? "completed" : "warning",
          result.verification.status === "verified"
            ? "原任务已完成并通过验收"
            : "原任务执行已结束，验收结果需要查看"
        );
      } catch (reason) {
        let latest = null;
        try {
          latest = await getFolderExecutionStatus(pending.operationId);
        } catch {
          // 状态查询同样失败时保留本地任务标识，下一次刷新继续查询原任务。
        }
        if (latest?.status === "completed" && latest.result) {
          finishRecoveredResult(latest.result, true);
          finishActivity("completed", "已找回原任务的完成结果");
        } else if (
          latest?.status === "failed" ||
          latest?.status === "needs_reconciliation"
        ) {
          appendFolderOperationNotice(
            pending,
            `原任务无法继续挂接：${
              latest.error?.message ??
              (reason instanceof Error ? reason.message : String(reason))
            }。没有重新执行；请在「备份与恢复」中核对文件。`,
            true
          );
          clearPendingFolderOperation(pending.operationId);
          finishActivity("warning", "原任务需要核对");
        } else {
          markActivityInterrupted("连接中断，原任务标识已保留。刷新后可继续查询。");
          finishActivity("warning", "连接中断，执行结果待确认");
        }
      } finally {
        setActivePlanId(null);
        setStatus("idle");
      }
    };

    void attach();
  }, []);

  useEffect(() => {
    if (officeRecoveryAttachStartedRef.current) return;
    officeRecoveryAttachStartedRef.current = true;
    const pending = loadPendingOfficeExecution();
    if (!pending) return;

    const attach = async () => {
      let persisted;
      try {
        persisted = await getOfficeExecutionStatus(pending.operationId);
        markServerOnline();
      } catch (reason) {
        if (isLocalServiceConnectionError(reason)) {
          markServerOffline();
          return;
        }
        appendOfficeExecutionNotice(
          pending,
          `无法找回上次 Office 写入任务：${
            reason instanceof Error ? reason.message : String(reason)
          }。没有创建新执行，也没有再次写入。`
        );
        return;
      }

      if (persisted.status === "completed") {
        if (persisted.result?.undoSnapshot) {
          pushUndoSnapshot(persisted.result.undoSnapshot);
        }
        appendOfficeExecutionNotice(
          pending,
          "已找回上次 Office 写入任务的完成回执；不会重复执行已经完成的动作。"
        );
        clearPendingOfficeExecution(pending.operationId);
        return;
      }

      const currentIdentity = currentOfficeWorkbookIdentity();
      const identityMatches = currentIdentity === persisted.workbookIdentity;
      const reason = !currentIdentity
        ? "当前工作簿尚未保存，无法确认它就是原工作簿。请先保存并重新打开原文件。"
        : !identityMatches
          ? "当前不是产生该断点的工作簿。请打开原工作簿后再继续。"
          : persisted.requiresReconciliation
            ? "上次停止时有一个动作可能已经写入。继续前会先读取目标范围并核对，不会盲目重放。"
            : "已完成动作会跳过，将从服务端保存的下一步继续。";
      appendOfficeExecutionNotice(pending, reason, {
        operationId: pending.operationId,
        expectedRevision: persisted.revision,
        status: identityMatches ? "pending" : "failed",
        requiresReconciliation: persisted.requiresReconciliation,
        nextActionIndex: persisted.nextActionIndex,
        totalActions: persisted.plan.actions.length,
        reason
      });
    };

    void attach();
  }, []);

  function appendMessage(
    message: Omit<ChatMessage, "id" | "createdAt">
  ): string {
    const id = messageId();
    const nextMessage: ChatMessage = {
      ...message,
      id,
      createdAt: new Date().toISOString()
    };
    setMessages((current) =>
      insertChatMessage(
        current,
        nextMessage,
        activeActivityMessageIdRef.current
      )
    );
    if (message.plan) enqueueProducedPlan(id, message.plan, Boolean(message.staleSourceConfirm));
    return id;
  }

  // Only newly delivered plans enter execution. Loading history never calls this.
  // First delivery and continuation must share this route and its identity guards.
  function enqueueProducedPlan(id: string, plan: AnalysisPlan, staleSourceConfirm = false) {
    if (staleSourceConfirm || resultMessages.current.has(id) || sourceMode !== "workbook" || !workbook) return;
    const identity = currentOfficeWorkbookIdentity();
    if (!identity || (!isWorksheetDeletion(plan) && resultReviewPlanIssue(plan, selectedNamesFor(workbook)) !== null)) return;
    resultMessages.current.add(id);
    const candidate = { id, plan, identity, authorization: [...selectedNamesFor(workbook)],
      conversationId: chatHistory.activeConversationId };
    if (isWorksheetDeletion(plan)) setDeletionCandidate(candidate);
    else setResultCandidate(candidate);
  }

  /**
   * 把一条系统消息写进**指定会话**。
   *
   * 异步回执可能晚于用户切换会话：`appendMessage` 只会写到当前会话，于是
   * "上一轮失败"会跑到用户刚打开的新会话里。这里按发起时的会话 id 落点；
   * 那段会话已被删除时什么都不写（不新建一段假历史）。
   */
  function appendSystemMessageToConversation(conversationId: string, text: string): void {
    const createdAt = new Date().toISOString();
    const safeText = redactCredentials(text);
    setChatHistory((current) => {
      if (!current.conversations.some((conversation) => conversation.id === conversationId)) return current;
      return {
        ...current,
        conversations: current.conversations.map((conversation) =>
          conversation.id !== conversationId ? conversation : {
            ...conversation,
            messages: [
              ...conversation.messages,
              { id: messageId(), role: "system" as const, text: safeText, createdAt }
            ].slice(-MAX_MESSAGES_PER_CONVERSATION),
            updatedAt: createdAt
          })
      };
    });
  }

  function appendFolderOperationOutcome(
    pending: Pick<
      PendingFolderOperation,
      "conversationId" | "operationId" | "activityMessageId" | "plan"
    >,
    result: FolderExecuteResult,
    recovered: boolean,
    preserveActivity = false
  ): void {
    const createdAt = new Date().toISOString();
    setChatHistory((current) => {
      const requestedConversationExists = current.conversations.some(
        (conversation) => conversation.id === pending.conversationId
      );
      const targetConversationId = requestedConversationExists
        ? pending.conversationId
        : current.activeConversationId;
      return {
        ...current,
        conversations: current.conversations.map((conversation) => {
          if (conversation.id !== targetConversationId) return conversation;
          const operationMessageIndex = conversation.messages.findIndex(
            (message) => message.folderOperationId === pending.operationId
          );
          const outcomeMessage: ChatMessage = {
            id:
              operationMessageIndex >= 0
                ? conversation.messages[operationMessageIndex].id
                : messageId(),
            role: "assistant",
            text: folderExecutionOutcomeText(pending.plan, result, recovered),
            verification: result.verification,
            executedPlanId: result.undo ? pending.plan.id : undefined,
            executedUndoId: result.undo?.undoId,
            folderOperationId: pending.operationId,
            createdAt
          };
          if (operationMessageIndex >= 0) {
            if (!conversation.messages[operationMessageIndex].folderResume) {
              return conversation;
            }
            const nextMessages = conversation.messages.map((message, index) =>
              index === operationMessageIndex ? outcomeMessage : message
            );
            return {
              ...conversation,
              title: conversationTitle(nextMessages),
              messages: nextMessages,
              updatedAt: createdAt
            };
          }
          const existingMessages = preserveActivity
            ? conversation.messages
            : conversation.messages.filter(
                (message) =>
                  message.id !== pending.activityMessageId ||
                  !message.activityPending ||
                  Boolean(message.activityLog)
              );
          const nextMessages = [
            ...existingMessages,
            outcomeMessage
          ].slice(-MAX_MESSAGES_PER_CONVERSATION);
          return {
            ...conversation,
            title: conversationTitle(nextMessages),
            messages: nextMessages,
            updatedAt: createdAt
          };
        })
      };
    });
    if (result.verification.status === "verified") {
      markPlanVerified(pending.plan.id);
    }
  }

  function appendFolderOperationNotice(
    pending: Pick<
      PendingFolderOperation,
      "conversationId" | "operationId" | "activityMessageId"
    >,
    text: string,
    preserveActivity = false,
    folderResume?: ChatMessage["folderResume"]
  ): void {
    const createdAt = new Date().toISOString();
    setChatHistory((current) => {
      const requestedConversationExists = current.conversations.some(
        (conversation) => conversation.id === pending.conversationId
      );
      const targetConversationId = requestedConversationExists
        ? pending.conversationId
        : current.activeConversationId;
      return {
        ...current,
        conversations: current.conversations.map((conversation) => {
          if (conversation.id !== targetConversationId) return conversation;
          const existingIndex = conversation.messages.findIndex(
            (message) => message.folderOperationId === pending.operationId
          );
          if (existingIndex >= 0) {
            if (!folderResume) return conversation;
            const nextMessages = conversation.messages.map((message, index) =>
              index === existingIndex
                ? { ...message, text, folderResume, createdAt }
                : message
            );
            return {
              ...conversation,
              title: conversationTitle(nextMessages),
              messages: nextMessages,
              updatedAt: createdAt
            };
          }
          const existingMessages = preserveActivity
            ? conversation.messages
            : conversation.messages.filter(
                (message) =>
                  message.id !== pending.activityMessageId ||
                  !message.activityPending ||
                  Boolean(message.activityLog)
              );
          const nextMessages = [
            ...existingMessages,
            {
              id: messageId(),
              role: "system" as const,
              text,
              folderOperationId: pending.operationId,
              folderResume,
              createdAt
            }
          ].slice(-MAX_MESSAGES_PER_CONVERSATION);
          return {
            ...conversation,
            title: conversationTitle(nextMessages),
            messages: nextMessages,
            updatedAt: createdAt
          };
        })
      };
    });
  }

  function appendOfficeExecutionNotice(
    pending: Pick<
      PendingOfficeExecution,
      "conversationId" | "operationId" | "activityMessageId"
    >,
    text: string,
    officeResume?: ChatMessage["officeResume"]
  ): void {
    const createdAt = new Date().toISOString();
    setChatHistory((current) => {
      const requestedConversationExists = current.conversations.some(
        (conversation) => conversation.id === pending.conversationId
      );
      const targetConversationId = requestedConversationExists
        ? pending.conversationId
        : current.activeConversationId;
      return {
        ...current,
        conversations: current.conversations.map((conversation) => {
          if (conversation.id !== targetConversationId) return conversation;
          const existingIndex = conversation.messages.findIndex(
            (message) => message.officeExecutionId === pending.operationId
          );
          if (existingIndex >= 0) {
            const nextMessages = conversation.messages.map((message, index) =>
              index === existingIndex
                ? { ...message, text, officeResume, createdAt }
                : message
            );
            return {
              ...conversation,
              title: conversationTitle(nextMessages),
              messages: nextMessages,
              updatedAt: createdAt
            };
          }
          const existingMessages = conversation.messages.filter(
            (message) =>
              message.id !== pending.activityMessageId ||
              !message.activityPending ||
              Boolean(message.activityLog)
          );
          const nextMessages = [
            ...existingMessages,
            {
              id: messageId(),
              role: "system" as const,
              text,
              officeExecutionId: pending.operationId,
              officeResume,
              createdAt
            }
          ].slice(-MAX_MESSAGES_PER_CONVERSATION);
          return {
            ...conversation,
            title: conversationTitle(nextMessages),
            messages: nextMessages,
            updatedAt: createdAt
          };
        })
      };
    });
  }

  function updateFolderResumeMessage(
    messageIdValue: string,
    update: (message: ChatMessage) => ChatMessage
  ): void {
    const updatedAt = new Date().toISOString();
    setChatHistory((current) => ({
      ...current,
      conversations: current.conversations.map((conversation) => {
        if (!conversation.messages.some((message) => message.id === messageIdValue)) {
          return conversation;
        }
        return {
          ...conversation,
          messages: conversation.messages.map((message) =>
            message.id === messageIdValue ? update(message) : message
          ),
          updatedAt
        };
      })
    }));
  }

  async function resumeInterruptedOfficeExecution(message: ChatMessage): Promise<void> {
    if (!(await targetGuard.check())) return;
    const resume = message.officeResume;
    const pending = loadPendingOfficeExecution();
    if (!resume || !pending || pending.operationId !== resume.operationId) {
      updateFolderResumeMessage(message.id, (current) => ({
        ...current,
        text: "本机没有找到这项 Office 任务的挂接记录，不能安全续跑。",
        officeResume: current.officeResume
          ? { ...current.officeResume, status: "failed" }
          : undefined
      }));
      return;
    }
    updateFolderResumeMessage(message.id, (current) => ({
      ...current,
      officeResume: current.officeResume
        ? { ...current.officeResume, status: "resuming" }
        : undefined
    }));
    beginActivity(
      "正在核对 Office 写入断点",
      "先确认工作簿身份和在途动作，再从下一步继续。",
      {
        operationId: resume.operationId,
        phase: "preflight",
        stepId: "execution.office.resume"
      }
    );
    setStatus("executing");
    let state: Awaited<ReturnType<typeof getOfficeExecutionStatus>> | null = null;
    try {
      state = await getOfficeExecutionStatus(resume.operationId);
      const currentIdentity = currentOfficeWorkbookIdentity();
      if (!currentIdentity || currentIdentity !== state.workbookIdentity) {
        throw new Error(
          currentIdentity
            ? "当前不是产生该断点的工作簿，已停止续跑。"
            : "当前工作簿尚未保存，无法确认身份，已停止续跑。"
        );
      }
      if (state.inFlightActionIndex !== null && state.inFlightActionIndex !== undefined) {
        if (!state.beforeSnapshot) {
          throw new Error("在途动作缺少写前快照，只能人工核对，不能自动续跑。");
        }
        const reconciliation = await reconcileOfficeRangeAction({
          plan: state.plan,
          operationId: state.operationId,
          workbookIdentity: state.workbookIdentity,
          actionIndex: state.inFlightActionIndex,
          before: state.beforeSnapshot,
          previous: state.checkpoint
        });
        if (reconciliation.outcome === "unknown") {
          throw new Error(
            "目标范围既不匹配执行前状态，也不匹配预期写入结果；可能存在部分写入或后续人工修改，已保持挂起。"
          );
        }
        state = await reconcileOfficeActionReceipt(state.operationId, {
          expectedRevision: state.revision,
          actionIndex: state.inFlightActionIndex,
          outcome: reconciliation.outcome,
          ...(reconciliation.checkpoint
            ? { checkpoint: reconciliation.checkpoint }
            : {})
        });
      } else {
        state = await claimOfficeExecutionResume(
          state.operationId,
          state.revision
        );
      }
      if (!state.ownerToken) {
        throw new Error("服务端没有授予本次 Office 续跑写入权。");
      }

      const resumedFromActionIndex = state.nextActionIndex;
      setActivePlanId(state.plan.id);
      const result = await executePlan(state.plan, {
        authorizedSheetNames: state.authorizedSheetNames,
        allowedExternalFiles: new Set(state.allowedExternalFiles),
        operationId: state.operationId,
        onProgress: applyProgressEvent,
        recovery: {
          startActionIndex: state.nextActionIndex,
          undoSnapshot: state.checkpoint,
          onBeforeAction: async ({ actionIndex, beforeSnapshot }) => {
            const currentState = state;
            if (!currentState?.ownerToken) throw new Error("Office 写入权已经失效");
            state = await startOfficeAction(currentState.operationId, {
              expectedRevision: currentState.revision,
              ownerToken: currentState.ownerToken,
              actionIndex,
              beforeSnapshot
            });
          },
          onActionCompleted: async ({ actionIndex, checkpoint }) => {
            const currentState = state;
            if (!currentState?.ownerToken) throw new Error("Office 写入权已经失效");
            state = await completeOfficeAction(currentState.operationId, {
              expectedRevision: currentState.revision,
              ownerToken: currentState.ownerToken,
              actionIndex,
              checkpoint
            });
          }
        }
      });
      if (!state.ownerToken) throw new Error("Office 写入权已经失效");
      state = await finishOfficeExecution(state.operationId, {
        expectedRevision: state.revision,
        ownerToken: state.ownerToken,
        result
      });
      pushUndoSnapshot(result.undoSnapshot ?? null);
      if (result.verification.status === "verified") {
        markPlanVerified(state.plan.id);
      }
      const completedPlanId = state.plan.id;
      clearPendingOfficeExecution(state.operationId);
      updateFolderResumeMessage(message.id, (current) => ({
        ...current,
        text: `Office 任务已从第 ${resumedFromActionIndex + 1} 步继续并完成；已提交每一步的写入回执。`,
        verification: result.verification,
        executedPlanId: completedPlanId,
        executedExecutionId: result.undoSnapshot?.executionId,
        officeResume: undefined
      }));
      await scan();
      const completed = state;
      const undo = result.undoSnapshot ?? null;
      const facts = {...resultLive.current};
      setResumedReviewTitle(completed.plan.title);
      await resultReview.adoptCompleted(`recovered:${completed.operationId}`,
        undo ? excelRecoveryBeforeResult(completed.plan, undo, completed.authorizedSheetNames) : {comparisonUnavailable:'恢复记录不可用。'},
        {executionId:completed.operationId,verified:result.verification.status === 'verified',
          undoable:undo?.postStateCoverage === 'complete',undo}, {
          isCurrent: () => currentOfficeWorkbookIdentity() === completed.workbookIdentity &&
            resultLive.current.workbookIdentity === completed.workbookIdentity &&
            resultLive.current.authorization === facts.authorization &&
            resultLive.current.conversationId === facts.conversationId,
          mark: () => markExcelResult(completed.plan, completed.workbookIdentity),
          clearMarker: clearExcelResultMarks,
          showBefore: before => showExcelBeforeResult(before, completed.workbookIdentity, completed.plan, completed.authorizedSheetNames),
          comparisonNotice: copy => copy.comparisonNotice ?? '恢复对比展示原始单元格值及公式文本；不重算公式，不作为原始格式对比。',
          closeBefore: async copy => {
            const outcome = await closeExcelNativeComparison(copy);
            return outcome === 'closed' || outcome === 'already_closed' ? null : '执行前副本仍保留，请手动核对。';
          },
          restore: async receipt => {
            if (!(await resultUndo.current(receipt.executionId))) throw new Error('恢复未成功，当前结果仍保留。');
          },
          validateResult: async receipt => {
            if (!receipt.undo) throw new Error('恢复记录不可用。');
            await Excel.run(context => assertUndoPostState(context, receipt.undo!));
          },
          accepted: () => markPlanVerified(completed.plan.id)
        });
      finishActivity(
        result.verification.status === "verified" ? "completed" : "warning",
        "Office 断点任务已完成"
      );
    } catch (reason) {
      const detail = reason instanceof Error ? reason.message : String(reason);
      if (state?.ownerToken && state.status === "running") {
        try {
          state = await suspendOfficeExecution(state.operationId, {
            expectedRevision: state.revision,
            ownerToken: state.ownerToken,
            code: "OFFICE_RESUME_INTERRUPTED",
            message: detail
          });
        } catch {
          // 若回执响应丢失，旧 revision 可能已推进；保留 operationId，下一次先查权威状态。
        }
      }
      updateFolderResumeMessage(message.id, (current) => ({
        ...current,
        text: detail,
        officeResume: current.officeResume
          ? {
              ...current.officeResume,
              expectedRevision: state?.revision ?? current.officeResume.expectedRevision,
              requiresReconciliation:
                state?.inFlightActionIndex !== null &&
                state?.inFlightActionIndex !== undefined,
              nextActionIndex:
                state?.nextActionIndex ?? current.officeResume.nextActionIndex,
              reason: detail,
              status: "failed"
            }
          : undefined
      }));
      finishActivity("warning", "Office 任务仍需核对");
    } finally {
      setActivePlanId(null);
      setStatus("idle");
    }
  }

  async function resumeInterruptedFolderOperation(message: ChatMessage): Promise<void> {
    const resume = message.folderResume;
    const pending = loadPendingFolderOperation();
    if (!resume || !pending || pending.operationId !== resume.operationId) {
      updateFolderResumeMessage(message.id, (current) => ({
        ...current,
        text: "本机没有找到这项任务的完整请求记录，不能安全续跑。请打开「备份与恢复」核对文件。",
        folderResume: current.folderResume
          ? { ...current.folderResume, status: "failed" }
          : undefined
      }));
      return;
    }
    updateFolderResumeMessage(message.id, (current) => ({
      ...current,
      folderResume: current.folderResume
        ? { ...current.folderResume, status: "resuming" }
        : undefined
    }));
    beginActivity(
      "正在从安全断点继续",
      "先核对实际文件，再跳过已完成发布，只继续剩余文件。",
      {
        operationId: resume.operationId,
        phase: "saving",
        stepId: "execution.resume"
      }
    );
    setStatus("executing");
    try {
      const latest = await getFolderExecutionStatus(resume.operationId);
      if (latest.status === "completed" && latest.result) {
        appendFolderOperationOutcome(pending, latest.result, true, true);
        clearPendingFolderOperation(resume.operationId);
        finishActivity("completed", "已找回原任务的完成结果");
        return;
      }
      if (!latest.canResume || typeof latest.revision !== "number") {
        throw new Error(
          latest.error?.message ?? "服务端没有确认可用的安全断点，不能继续。"
        );
      }
      const result = await resumeFolderExecution(
        resume.operationId,
        latest.revision,
        {
          onProgress: applyProgressEvent,
          onConnectionState: (connectionState, detail) => {
            if (connectionState === "interrupted") {
              markActivityInterrupted(detail ?? "连接中断，续跑结果待确认。");
            } else {
              markActivityConnected();
            }
          }
        }
      );
      appendFolderOperationOutcome(pending, result, true, true);
      if (result.undo) {
        pushFolderUndo({
          planId: pending.plan.id,
          undoId: result.undo.undoId,
          sessionId: result.undo.sessionId
        });
      }
      clearPendingFolderOperation(resume.operationId);
      finishActivity(
        result.verification.status === "verified" ? "completed" : "warning",
        result.verification.status === "verified"
          ? "断点续跑完成并通过验收"
          : "断点续跑完成，验收结果需要查看"
      );
    } catch (reason) {
      const detail = reason instanceof Error ? reason.message : String(reason);
      updateFolderResumeMessage(message.id, (current) => ({
        ...current,
        text: `续跑已安全停止：${detail} 没有覆盖无法确认状态的文件；你可以核对后重试，或在「备份与恢复」中恢复执行前版本。`,
        folderResume: current.folderResume
          ? { ...current.folderResume, status: "failed" }
          : undefined
      }));
      finishActivity("warning", "续跑已安全停止");
    } finally {
      setStatus("idle");
    }
  }

  async function applyPlanningContinuationResult(
    message: ChatMessage,
    response: PlanningStreamResult,
    baseline: NonNullable<ChatMessage["planningResume"]>
  ): Promise<"paused" | "completed" | "awaiting_input" | "failed"> {
    const signal = turnAbortRef.current?.signal;
    // A resumed result belongs beside this continuation's progress, not to the
    // old pause card. Capture ownership before awaiting the target check.
    const segment = getActivitySnapshot();
    const resultMessageId = segment?.operationId === baseline.turnId && segment.persistenceKey
      ? segment.persistenceKey : message.id;
    if (!(await waitForPlanning(targetGuard.check(), signal))) return "awaiting_input";
    if (response.kind === "planning_pause") {
      updateFolderResumeMessage(message.id, (current) => ({
        ...current,
        text: response.reason,
        provider: "system",
        planningResume: {
          ...baseline,
          expectedRevision: response.revision,
          clientLastSequence: Math.min(Math.max(baseline.clientLastSequence ?? 0, getActivitySnapshot()?.sequence ?? 0), response.lastSequence),
          awaitingStatus: false,
          lastSequence: response.lastSequence,
          status: "pending",
          reason: response.reason,
          extensionSeconds: response.extensionSeconds,
          activeElapsedSeconds: response.activeElapsedSeconds,
          pauseType: response.pauseType,
          currentModelId: response.currentModelId,
          modelSwitchCount: response.modelSwitchCount,
          canSwitchModel: response.canSwitchModel,
          availableModelConnections: response.availableModelConnections,
          selectedFallbackModelId:
            baseline.selectedFallbackModelId ??
            response.availableModelConnections[0]?.catalogModelId,
          requestId: undefined,
          sourceFingerprint:
            response.sourceFingerprint ?? baseline.sourceFingerprint,
          sourceFingerprintSheets:
            response.sourceFingerprintSheets?.length
              ? response.sourceFingerprintSheets
              : baseline.sourceFingerprintSheets
        }
      }));
      return "paused";
    }
    setMessages((items) => items.map((item) => item.planningResume?.turnId === baseline.turnId
      ? { ...item, planningResume: undefined } : item));
    if (response.kind === "clarification") {
      if (!workbook) throw new Error("工作簿范围不可用，请重新发起需求。");
      const originalPrompt = baseline.originalPrompt ?? message.intentMemory?.confirmedPrompt;
      if (!originalPrompt) throw new Error("旧规划记录缺少原需求，请重新发送需求后补充信息。");
      await continueIntentDecision({ ...response, turnId: undefined }, originalPrompt, [],
        intentScopeFingerprint(buildIntentScope(workbook,
          baseline.sourceFingerprintSheets?.length ? baseline.sourceFingerprintSheets : selectedNamesFor(workbook))),
        undefined, 0, 0, sourceMode === "workbook" ? "none" : "ifNeeded", baseline.taskContext);
      updateFolderResumeMessage(message.id, (current) => ({ ...current, planningResume: undefined }));
      return "awaiting_input";
    }
    if (response.kind === "answer") {
      updateFolderResumeMessage(resultMessageId, (current) => ({
        ...current,
        text: response.message,
        resultContext: response.resultContext ?? undefined,
        provider: response.provider,
        planningResume: undefined
      }));
      return response.outcome === "failed" ? "failed" : "completed";
    }
    if (response.kind === "local_plan_rule") {
      // 续跑通道同样要支持本地映射规则：与主路径共用同一个解析函数，
      // 免得"重新规划能算、续跑却报不支持"这种同功能两态。
      // 源表范围取续跑基线里已记录的工作表，不重新扩大授权。
      const plan = await waitForPlanning(localMappingPlanFromRule(
        response,
        baseline.sourceFingerprintSheets?.length
          ? baseline.sourceFingerprintSheets
          : selectedNamesFor(workbook!)
      ), signal);
      planningTurnByPlanIdRef.current.set(plan.id, baseline.turnId);
      updateFolderResumeMessage(resultMessageId, (current) => ({
        ...current,
        text: plan.summary,
        plan,
        provider: response.provider,
        planningResume: undefined
      }));
      enqueueProducedPlan(resultMessageId, plan);
      return "completed";
    }
    const plan: AnalysisPlan = attestSourceDependency(
      {
        ...response.plan,
        sourceFingerprint:
          baseline.sourceFingerprint ?? response.plan.sourceFingerprint,
        sourceFingerprintSheets:
          baseline.sourceFingerprintSheets?.length
            ? baseline.sourceFingerprintSheets
            : response.plan.sourceFingerprintSheets
      },
      // 续跑通道这一轮没有本机读取记录：**不采信计划自报的来源依赖**，删掉该字段。
      null
    );
    planningTurnByPlanIdRef.current.set(plan.id, baseline.turnId);
    updateFolderResumeMessage(resultMessageId, (current) => ({
      ...current,
      text: plan.summary,
      plan,
      provider: response.provider,
      planningResume: undefined
    }));
    enqueueProducedPlan(resultMessageId, plan);
    return "completed";
  }

  /**
   * 开一个**新的回复段**。
   *
   * 旧口径是把进度接回原来那条消息（"同一任务只有一张过程卡"），代价是用户按了
   * 继续之后，最新消息下面什么都没有 —— 他看到的那句话没有换来任何反馈，进度
   * 出现在上方一条更早的消息里。现在改成：**同一任务、多段回复**。
   *
   * 这一段做三件事，都不改执行身份：
   * 1. 追加一条新的助手消息作为新的展示段；
   * 2. 把上一段标成 `superseded` 并冻结它（没有收到结束事件的步骤如实交代），
   *    它保留**当时**的状态与耗时，不再被后续事件改写；
   * 3. 由上一段提供"查看后续进展"的轻量定位入口，不复制整份旧日志。
   *
   * 段链判据**只**走 `packages/core/replySegments.ts`：父段是谁、谁降级成
   * `superseded`、`nextSegmentId` 指向谁，全部由 `openReplySegment` 决定。两端
   * 各写一遍"看起来一样"的链式改写，正是"第二次继续把第一段的后继覆盖掉"那类
   * 缺陷的来源。
   *
   * 返回新段的消息 id（= 活动进度的 persistenceKey）。
   */
  function openReplySegmentMessage(params: {
    /** 同一次继续里重复调用必须用**同一个**键，否则会多开一段。 */
    key: string;
    /** 这一段在展示哪个执行任务；早期暂停没有服务端任务身份，这里给 null。 */
    taskId?: string | null;
    /** 上一段承载的消息。 */
    previous?: ChatMessage;
    /** 上一段的消息 id 已经确定时直接用（优先于 `previous`）。 */
    previousSegmentId?: string | null;
    title: string;
    detail: string;
    initialSequence: number;
    stepId: string;
  }): string {
    const {
      key,
      taskId = null,
      previous,
      title,
      detail,
      initialSequence,
      stepId
    } = params;
    // 同一次继续里的第二次调用：复用已经开出的那一段，不另开新段。
    const already = resumeSegmentRef.current;
    if (already && already.key === key) {
      activeActivityMessageIdRef.current = already.messageId;
      // 段还在跑就原样接上（不重开步骤）；已经被中间收尾关掉了才重新挂回去。
      if (getActivitySnapshot()?.persistenceKey !== already.messageId) {
        resumeActivityProgress(title, detail, {
          persistenceKey: already.messageId,
          operationId: taskId,
          phase: "planning",
          stepId,
          initialSequence,
          previousLog: null
        });
      }
      return already.messageId;
    }
    const previousLog = previous?.activityLog;
    const previousStepCount = previousLog
      ? previousLog.steps.filter((step) => step.status === "completed").length
      : 0;
    const parentId = params.previousSegmentId ?? previous?.id ?? null;
    const persistenceKey = appendMessage({
      role: "assistant",
      activityPending: true
    });
    const now = Date.now();
    setMessages((current) => {
      /*
       * 段身份 = 承载它的消息 id：段链（previous/next）与"跳到那一段"的定位都按
       * 这个键对齐，避免出现"段 id 一套、消息 id 另一套"的第二种身份。
       *
       * 早期暂停的旧消息可能**还没有**段身份（它比段机制更早落盘）。这里只给
       * "这一轮要被接管的上一段"补一条 `live`：给**所有**带日志的老消息补
       * `live` 会把它们集体降级成 `superseded`，历史卡片上的「继续」入口就全没了。
       */
      const seeded = current.map((item) =>
        parentId && item.id === parentId && !item.replySegment
          ? { ...item, replySegment: { id: item.id, status: "live" as const } }
          : item
      );
      const planned = openReplySegment(
        seeded.flatMap((item) => (item.replySegment ? [item.replySegment] : [])),
        {
          id: persistenceKey,
          taskId,
          continuesTask: true,
          previousSegmentId: parentId
        }
      );
      const byId = new Map(planned.map((segment) => [segment.id, segment]));
      return seeded.map((item) => {
        const segment = byId.get(item.replySegment?.id ?? item.id);
        if (!segment) return item;
        const supersededNow =
          segment.status === "superseded" &&
          item.replySegment?.status !== "superseded";
        return {
          ...item,
          ...(supersededNow ? { activityPending: false } : {}),
          replySegment: segment,
          ...(supersededNow &&
          item.activityLog &&
          item.activityLog.outcome === undefined
            ? {
                // 只冻结**还在进行中**的那一段：已经有结论（已暂停/已完成）的段
                // 不能在这里重算耗时，否则暂停时长会被并进执行时间。
                activityLog: freezeSegmentLog(
                  item.activityLog,
                  now
                ) as unknown as ActivityLog
              }
            : {})
        };
      });
    });
    activeActivityMessageIdRef.current = persistenceKey;
    resumeActivityProgress(title, detail, {
      persistenceKey,
      operationId: taskId ?? previousLog?.operationId ?? null,
      phase: "planning",
      stepId,
      initialSequence,
      /*
       * **不**把上一段的时间线带过来。带过来就等于把两段合成一段：新段会从
       * "第 1 步"开始重复展示旧步骤，或者反过来把旧段改写成"进行中"。
       * 上一段已完成多少步只用于文案（`previousStepCount`），不复用其时间线。
       */
      previousLog: null
    });
    resumeSegmentRef.current = {
      key,
      messageId: persistenceKey,
      startedAt: Date.now()
    };
    if (previousStepCount > 0) {
      updateActivityDetail(
        `${continuationLabel(previousStepCount)}；此前那些步骤没有重做。`
      );
    }
    return persistenceKey;
  }

  /**
   * 规划任务的「继续」：新段接在该任务**最新**那一段之后。
   *
   * 父段由共享的 `planSegmentContinuation` 判（内部是 `latestReplySegmentForTask`），
   * 不是"列表里第一条匹配这个 turnId 的消息"。用后者的话第二次继续仍然改写第一段：
   * A 的 `nextSegmentId` 被反复覆盖，B 再也没有"查看后续进展"，链就断在中间。
   */
  function resumePlanningActivity(
    resume: NonNullable<ChatMessage["planningResume"]>,
    title: string,
    detail: string,
    initialSequence: number,
    stepId: string
  ): string {
    const continuation = planSegmentContinuation(
      messages.flatMap((item) => (item.replySegment ? [item.replySegment] : [])),
      resume.turnId
    );
    const storedActivityMessage =
      (continuation.previousSegmentId
        ? messages.find((item) => item.id === continuation.previousSegmentId)
        : undefined) ??
      (resume.activityMessageId
        ? messages.find((item) => item.id === resume.activityMessageId)
        : undefined) ??
      [...messages]
        .reverse()
        .find((item) => item.activityLog?.operationId === resume.turnId);
    return openReplySegmentMessage({
      key: `task:${resume.turnId}`,
      taskId: resume.turnId,
      previous: storedActivityMessage,
      ...(continuation.previousSegmentId
        ? { previousSegmentId: continuation.previousSegmentId }
        : {}),
      title,
      detail,
      initialSequence,
      stepId
    });
  }

  /**
   * 早期分析（理解消息 / 读表结构 / 字段意图）的「继续」：**另开一个新回复段**。
   *
   * 这一步还没有服务端任务身份（没形成规划检查点），所以段上**不写** `taskId`、
   * 不伪造 turnId、也不声称 SDK 内部断点续算 —— 它是本机按原需求重跑，归属就是
   * 这条本机消息。把进度接回旧卡正是"用户又发了一句话，下方什么都没有"的那个缺陷。
   *
   * 三条入口（输入"继续"、空输入点发送、点卡片上的继续）都经 `restartEarlyAnalysis`
   * 落到这里，所以行为一致；`key` 用原消息 id，同一次继续只开一段。
   */
  function resumeEarlyAnalysisActivity(retained: ChatMessage): string {
    return openReplySegmentMessage({
      key: `early:${retained.id}`,
      previous: retained,
      title: "继续分析",
      detail: "正在核验当前现场并继续原需求",
      initialSequence: retained.activityLog?.sequence ?? -1,
      stepId: "planning.resume"
    });
  }

  /**
   * 原任务仍在服务端运行时，把**已落盘的事件**接回原来那张过程卡。
   *
   * 「继续」在服务端仍跑时会被 409 挡住（再发一次就是重复规划、重复读表），
   * 但用户按下继续之后什么都不发生同样是错的：任务确实在跑，进度也确实已经写进
   * 账本，本端只是断开了。这里只读 `GET /api/turn/{id}/events`（不创建任务、不调
   * 模型），事件照常走 `applyProgressEvent` —— 与实时流**同一个**归并入口，
   * 步骤身份与序号口径因此完全一致，既不新增卡片，也不重置编号。
   *
   * 返回 `true` 表示任务已经不在跑了（调用方应再查一次状态去取结果/判续跑）；
   * `false` 表示等到上限它还在跑 —— 如实说"仍在运行"，不重发需求。
   */
  function planningCursor(resume: NonNullable<ChatMessage["planningResume"]>) {
    const current = getActivitySnapshot();
    return planningReceivedSequence(resume.turnId, [
      { taskId: resume.turnId, sequence: resume.clientLastSequence },
      { taskId: current?.operationId, sequence: current?.sequence },
      ...messages.flatMap((item) => [
        { taskId: item.activityLog?.operationId, sequence: item.activityLog?.sequence },
        { taskId: item.planningResume?.turnId, sequence: item.planningResume?.clientLastSequence }
      ])
    ]);
  }

  function completePlanningAttachment() {
    const step = getActivitySnapshot()?.stepId;
    if (step === "planning.attach" || step?.startsWith("planning.resume.")) {
      completeCurrentStep("已接回原任务进度");
    }
  }

  async function tailRunningPlanningTask(
    resume: NonNullable<ChatMessage["planningResume"]>,
    signal: AbortSignal
  ): Promise<boolean> {
    const startCursor = planningCursor(resume);
    resumePlanningActivity(
      resume,
      "原任务仍在服务端处理中",
      "正在接回进度…",
      startCursor,
      "planning.attach"
    );
    // A pointer may have arrived before pause while its receipt was still in flight.
    // Ask the original task for pending calls; never replay a cached call ourselves.
    await waitForPlanning(sdkBridge.pollPending(), signal);
    let cursor = startCursor;
    const deadline = Date.now() + RUNNING_TAIL_MAX_MS;
    for (;;) {
      if (signal.aborted) throw new DOMException("Aborted", "AbortError");
      let page: Awaited<ReturnType<typeof getPlanningTaskEvents>>;
      try {
        page = await waitForPlanning(getPlanningTaskEvents(resume.turnId, cursor, signal), signal);
      } catch (error) {
        if (signal.aborted) throw error;
        // 读不到就停下：断线不等于可以重发需求，也不等于任务已经结束。
        return false;
      }
      if (signal.aborted) throw new DOMException("Aborted", "AbortError");
      completePlanningAttachment();
      for (const raw of page.events) {
        const event = raw as Parameters<typeof applyProgressEvent>[0];
        if (!acceptsPlanningReplay(resume.turnId, cursor, event)) continue;
        const sequence = Number(event.sequence);
        cursor = Math.max(cursor, sequence);
        applyProgressEvent(event);
        sdkBridge.handleProgressEvent(event);
      }
      if (!page.active) return true;
      if (Date.now() >= deadline) return false;
      await waitForPlanning(new Promise((resolve) => setTimeout(resolve, RUNNING_TAIL_INTERVAL_MS)), signal);
    }
  }

  async function resumePlanningTaskFromMessage(
    message: ChatMessage,
    fallbackModelId?: string,
    inspectOnly = false,
    continueRequested = false
  ): Promise<void> {
    const resume = message.planningResume;
    if (!resume || planningResumeInFlightRef.current || (busy && !inspectOnly)) return;
    let inspectAfterPause: ChatMessage | null = null;
    const requestId = resume.requestId ?? messageId();
    const turnController = new AbortController();
    // A superseded continuation must not settle the current execution's card.
    const finishResumeActivity = (...args: Parameters<typeof finishActivity>) => {
      if (turnAbortRef.current === turnController) finishActivity(...args);
    };
    planningResumeInFlightRef.current = turnController;
    turnAbortRef.current = turnController;
    setStatus("planning");
    updateFolderResumeMessage(message.id, (current) => ({
      ...current,
      planningResume: current.planningResume
        ? { ...current.planningResume, status: "resuming", requestId }
        : undefined
    }));
    try {
      if (!(await targetGuard.check())) {
        updateFolderResumeMessage(message.id, (current) => ({ ...current,
          planningResume: current.planningResume ? { ...current.planningResume, status: "pending", awaitingStatus: true,
            reason: "原任务的操作目标需要确认；尚未查询或继续原任务。" } : undefined }));
        finishResumeActivity("warning", "原任务范围需要确认", "awaiting_input");
        return;
      }
      if (turnController.signal.aborted) throw new DOMException("Aborted", "AbortError");
      let taskStatus = await waitForPlanning(getPlanningTaskStatus(
        resume.turnId,
        turnController.signal
      ), turnController.signal);
      if (turnController.signal.aborted) throw new DOMException("Aborted", "AbortError");
      const prepareBridge = async () => {
        // The user may switch files while the status request is in flight.
        if (!(await targetGuard.check())) {
          updateFolderResumeMessage(message.id, (current) => ({ ...current,
            planningResume: current.planningResume ? { ...current.planningResume, status: "pending", awaitingStatus: true,
              reason: "原任务的操作目标需要确认；尚未接回原任务。" } : undefined }));
          finishResumeActivity("warning", "原任务范围需要确认", "awaiting_input");
          return null;
        }
        if (turnController.signal.aborted) throw new DOMException("Aborted", "AbortError");
        const sourceFingerprint =
          ("sourceFingerprint" in taskStatus ? taskStatus.sourceFingerprint : null) ?? resume.sourceFingerprint ??
          (sdkBridge.binding()?.taskId === resume.turnId ? sdkBridge.binding()?.sourceFingerprint : null);
        const sourceFingerprintSheets =
          ("sourceFingerprintSheets" in taskStatus && taskStatus.sourceFingerprintSheets?.length)
            ? taskStatus.sourceFingerprintSheets
            : resume.sourceFingerprintSheets;
        if (workbook && sourceFingerprintSheets?.length &&
            JSON.stringify([...sourceFingerprintSheets].sort()) !== JSON.stringify([...selectedNamesFor(workbook)].sort())) {
          updateFolderResumeMessage(message.id, (current) => ({ ...current,
            text: "原任务的授权工作表范围已变化；不会继续旧任务或读取已撤销授权的工作表。请按当前范围重新生成预览。",
            planningResume: current.planningResume ? { ...current.planningResume, status: "failed" } : undefined }));
          finishResumeActivity("warning", "授权范围已变化，未继续旧任务");
          return null;
        }

        /*
         * 现场证据：必须**此刻**重读宿主。
         *
         * 上面几道闸只是"先停下"，它们读的仍是本地状态。真正交给服务端、也是唯一
         * 能证明"现在的还是一样"的东西，是这里现读出来的三项：
         *
         * - 工作簿身份：`currentOfficeWorkbookIdentity()` 现读文档 URL；读不到就停，
         *   **不**退回文件名或记录里的值（`name:` 身份在检查点契约里本来就不合法）；
         * - 授权成员：已提交的授权集合（与桥的授权复核同一个来源）；
         * - 来源指纹：服务端判定原任务依赖来源时才重读，并且**读取失败 != 尚未采集**。
         *
         * 判据本身在共享层（`buildPlanningContinuationEvidence`），与 WPS 同一份实现。
         */
        const authorizedNames = committedSheetNamesRef.current.length
          ? [...committedSheetNamesRef.current]
          : workbook
            ? [...selectedNamesFor(workbook)]
            : [];
        const evidenceResult = await buildPlanningContinuationEvidence({
          workbookIdentity: currentOfficeWorkbookIdentity(),
          authorizedSheetNames: authorizedNames,
          // 服务端从检查点判定；缺省按"需要核实"处理，宁可按当前数据重读一次，
          // 也不要把"数据已经变了"的现场放过去。
          sourceRequired:
            !("continuation" in taskStatus && taskStatus.continuation?.sourceVerificationRequired === false),
          readSourceFingerprint: () =>
            captureWorkbookSourceFingerprint(
              sourceFingerprintSheets?.length
                ? sourceFingerprintSheets
                : authorizedNames
            )
        });
        if (!evidenceResult.ok) {
          updateFolderResumeMessage(message.id, (current) => ({
            ...current,
            text: PLANNING_EVIDENCE_FAILURE_MESSAGES[evidenceResult.failure],
            planningResume: current.planningResume
              ? { ...current.planningResume, status: "failed" }
              : undefined
          }));
          finishResumeActivity(
            "warning",
            evidenceResult.failure === "source_unreadable"
              ? "没有读到当前数据，未继续旧任务"
              : "没有核实到当前现场，未继续旧任务"
          );
          return null;
        }
        const evidence = evidenceResult.evidence;
        if (
          evidence.sourceVerification === "verified" &&
          sourceFingerprint &&
          evidence.sourceFingerprint !== sourceFingerprint
        ) {
          // 读到了、但与原任务记录的那份不同：说明数据变过，不能沿用旧现场。
          updateFolderResumeMessage(message.id, (current) => ({
            ...current,
            text: "原任务使用的数据已经变化，不能继续旧的规划现场。请按当前数据重新发送需求。",
            planningResume: current.planningResume
              ? { ...current.planningResume, status: "failed" }
              : undefined
          }));
          finishResumeActivity("warning", "数据已经变化，未继续旧任务");
          return null;
        }

        bridgeLiveContextRef.current = async () => {
          const scopeNames = [...committedSheetNamesRef.current];
          let liveFingerprint: string | null = null;
          let activeWorksheet = "";
          try {
            activeWorksheet = (await captureSelectionContext()).activeWorksheet;
          } catch {
            activeWorksheet = "";
          }
          try {
            liveFingerprint = await captureWorkbookSourceFingerprint(scopeNames);
          } catch {
            liveFingerprint = null;
          }
          return {
            workbookIdentity: bridgeWorkbookIdentity(
              workbook?.name ?? "",
              currentOfficeWorkbookIdentity() ?? null
            ),
            sourceFingerprint: liveFingerprint,
            authorizedSheetNames: scopeNames,
            activeWorksheet
          };
      };
      const bridgeReady = await sdkBridge.ensureLive({
        // **原任务的 turnId**：不生成新 id 去绕过原账本与防重放。
        taskId: resume.turnId,
        workbookName: workbook?.name ?? "",
        workbookIdentity: evidence.workbookIdentity,
        sourceFingerprint: evidence.sourceFingerprint ?? null,
        signal: turnController.signal,
        liveContext: () => bridgeLiveContextRef.current(),
        onProgress: ({ scannedRows, totalRows, sheet }) => {
          updateActivityDetail(
            `正在读取「${sheet}」：${scannedRows.toLocaleString()} / ${totalRows.toLocaleString()} 行`
          );
        },
        onError: () => {
          updateActivityDetail("本地查询未完成，已把固定错误码交回模型继续。");
        }
      });
      if (!bridgeReady) {
        /*
         * **桥没立起来就不发继续请求。**
         *
         * `ensureLive` 返回 false 只有一种含义：这一次建立**没有成立**（初始化
         * 期间被取消、被收尾，或被另一轮取代）。以前这里不看返回值，照样把
         * 「继续」发出去 —— 服务端一发出 `bridgeTool` 指针，本端交给一条不存在的
         * 桥，模型干等到超时，用户看到的是"点继续没反应"。不发送才是如实：
         * 没有哪一次读表会被悄悄丢掉。
         */
        if (turnAbortRef.current !== turnController) {
          // 已被新一轮接管：这一轮连提示都不该写，界面归新轮次。
          return null;
        }
        const cancelled = turnController.signal.aborted;
        updateFolderResumeMessage(message.id, (current) => ({
          ...current,
          text: cancelled
            ? "已取消本次继续；没有发送继续请求，原任务现场保持不变。"
            : "本机读表桥没有为本次继续建立起来，没有发送继续请求；任务现场已保留，可以稍后再继续。",
          planningResume: current.planningResume
            ? {
                ...current.planningResume,
                status: "pending",
                awaitingStatus: false,
                requestId: undefined
              }
            : undefined
        }));
        finishResumeActivity(
          cancelled ? "warning" : "failed",
          cancelled ? "已取消本次继续，未发送请求" : "未能建立本机读表桥，未发送继续请求",
          cancelled ? "awaiting_input" : "failed"
        );
        return null;
      }
      return { evidence, sourceFingerprint, sourceFingerprintSheets };
      };
      if (taskStatus.status === "running" && !inspectOnly) {
        if (!(await prepareBridge())) return;

        /*
         * 原任务仍在服务端跑：先把**已落盘**的进度接回原来那张过程卡（不重发需求、
         * 不重新规划），等它结束后再查一次状态，走下面"取回结果 / 从检查点续跑"
         * 这两条既有分支 —— 而不是只回一句"稍后核对"。
         *
         * **自动核对（`inspectOnly`）不走这一步。** 暂停之后那一次自动查询的语义
         * 只是"读一眼真实状态"：用户刚按下暂停，服务端任务还在跑，这时去轮询补发
         * 等于把他的暂停变成"其实还在进行中"，而且会把这最后一次核对卡住最多
         * `RUNNING_TAIL_MAX_MS`。接回进度只属于**用户明确要求继续/核对**的那一次。
         */
        if (await tailRunningPlanningTask(resume, turnController.signal)) {
          taskStatus = await waitForPlanning(getPlanningTaskStatus(resume.turnId, turnController.signal), turnController.signal);
          if (turnController.signal.aborted) throw new DOMException("Aborted", "AbortError");
        }
      }
      if (taskStatus.status === "completed" && inspectOnly) {
        updateFolderResumeMessage(message.id, (current) => ({ ...current,
          planningResume: current.planningResume ? { ...current.planningResume, status: "pending", awaitingStatus: true,
            reason: "原任务已完成，结果已保留；继续时取回结果。" } : undefined }));
        finishResumeActivity("warning", "原任务已完成，结果已保留；继续时取回结果。", "awaiting_input");
        return;
      }
      if (taskStatus.status === "completed") {
        resumePlanningActivity(resume, "正在取回已有结果", "复用原任务的完成结果，不重新查询或分析。",
          planningCursor(resume), "planning.attach");
        completePlanningAttachment();
        /*
         * 完成态有两种：
         * - 账本记录里**没有值被省略** → 服务端给了完整 `result`，按原路回放；
         * - 有值被省略 → 只有 `recovery` 描述。**绝不**拿它当完整计划渲染，
         *   如实告诉用户值没有留在账本里，需要重新运行或从结果卡取值。
         */
        if (taskStatus.result === undefined) {
          const recovery = taskStatus.recovery;
          updateFolderResumeMessage(message.id, (current) => ({
            ...current,
            text: recovery?.resultRef
              ? "原任务已完成，但结果含数据值 —— 账本只保留结构与引用，未保存值。可用结果卡按引用查看，或重新运行以取得完整结果。"
              : "原任务已完成；该结果含数据值，账本只保留了恢复描述，未保存值。请重新运行以取得完整结果。",
            planningResume: undefined
          }));
          markServerOnline();
          finishResumeActivity("warning", "已完成，但账本未保存数据值");
          return;
        }
        const outcome = await applyPlanningContinuationResult(
          message,
          taskStatus.result,
          resume
        );
        markServerOnline();
        finishResumeActivity(
          outcome === "failed" ? "failed" : outcome === "completed" ? "completed" : "warning",
          outcome === "failed" ? "规划已停止" : outcome === "completed" ? "已找回任务的完成结果" : "进度已保存，等待你继续",
          outcome === "failed" ? "failed" : undefined
        );
        return;
      }
      if (taskStatus.status === "failed" || taskStatus.status === "cancelled") {
        const detail = taskStatus.error?.message ?? "该规划任务已经停止，不能继续。";
        updateFolderResumeMessage(message.id, (current) => ({
          ...current,
          text: detail,
          planningResume: current.planningResume
            ? { ...current.planningResume, status: "failed" }
            : undefined
        }));
        finishResumeActivity("warning", "任务已经停止");
        return;
      }
      if (taskStatus.status === "running") {
        updateFolderResumeMessage(message.id, (current) => ({
          ...current,
          text: "任务仍在处理中，无需重复继续。稍后再查看状态。",
          planningResume: current.planningResume
            ? { ...current.planningResume, status: "pending", awaitingStatus: true,
                reason: "原任务仍在服务端处理中；稍后核对，不重复规划。" }
            : undefined
        }));
        finishResumeActivity("warning", "任务仍在处理中");
        return;
      }
      if (!taskStatus.canContinue) {
        // 如实区分"总时长已用完"与"没有可续跑检查点"：前者不是找不到现场，
        // 而是**不应**再承诺继续，不能糊成同一句"没有已核实的可续跑检查点"。
        // 联合类型里 `pauseType` 只在实际查到任务的分支上存在，用 in 守卫收窄。
        const pauseType = "pauseType" in taskStatus ? taskStatus.pauseType : undefined;
        const exhausted = pauseType === "budget_exhausted";
        /*
         * 「为什么不能继续」只认服务端给的 `continuation`：它由共享契约 + 执行账本
         * 算出，包含"这个任务其实已经写过原表"这种本地看不出来的事实。客户端不许
         * 再自己拼一句"没有可续跑检查点"把真实原因抹平。
         */
        const blocked =
          "continuation" in taskStatus &&
          taskStatus.continuation &&
          taskStatus.continuation.available === false
            ? taskStatus.continuation
            : undefined;
        const blockedReason =
          blocked?.message?.trim() ||
          ("reason" in taskStatus ? taskStatus.reason : undefined) ||
          "没有已核实的可继续现场。";
        // 服务端文案自带句末标点；本地兜底不带。统一补一个分隔，避免"没有检查点原需求"。
        const blockedSentence = /[。！？]$/.test(blockedReason.trim())
          ? blockedReason.trim()
          : `${blockedReason.trim()}。`;
        updateFolderResumeMessage(message.id, (current) => ({
          ...current,
          text: exhausted
            ? `本轮有效运行总时长已用完，不再启动新的模型或工具调用：${("reason" in taskStatus && taskStatus.reason) || "已完成的步骤与现场已保存。"} 原需求仍保留，可重新发起任务。`
            : `该任务当前不能继续分析：${blockedSentence}原需求仍保留。`,
          planningResume: current.planningResume
            ? {
                ...current.planningResume,
                status: "failed",
                pauseType,
                reason: exhausted
                  ? "本轮有效运行总时长已用完，不能再按这里的引用直接续跑；请重新发起任务。"
                  : blockedReason
              }
            : undefined
        }));
        finishResumeActivity("warning", exhausted ? "本轮有效运行总时长已用完" : "当前任务不能继续分析");
        return;
      }

      // 自动核对永远只查状态；检查点可用也不能自动 POST 继续。
      if (inspectOnly || (resume.awaitingStatus && !continueRequested)) {
        updateFolderResumeMessage(message.id, (current) => ({ ...current, text: "已核实当前任务的现场；点击「继续分析」会按当前授权重新读取并分析，未写入工作表。",
          planningResume: current.planningResume ? { ...current.planningResume, awaitingStatus: false, status: "pending",
            expectedRevision: taskStatus.revision, lastSequence: taskStatus.lastSequence,
            pauseType: taskStatus.pauseType, reason: taskStatus.reason, extensionSeconds: taskStatus.extensionSeconds,
            currentModelId: taskStatus.currentModelId, canSwitchModel: taskStatus.canSwitchModel,
            availableModelConnections: taskStatus.availableModelConnections, requestId: undefined } : undefined }));
        finishResumeActivity("warning", "已核实检查点，等待你继续", "awaiting_input");
        return;
      }

      const maySwitchModel =
        taskStatus.pauseType === "model_error" ||
        taskStatus.pauseType === "interrupted";
      // 用解析后的有效选择（而不是原始存储值）：它已经过滤掉不存在的连接和
      // 已移除的 local，因此这里不会再拿一个失效 id 去接管断点任务。
      const globallySelectedFallback =
        maySwitchModel &&
        effectiveModelId &&
        effectiveModelId !== taskStatus.currentModelId
          ? effectiveModelId
          : undefined;
      const effectiveFallbackModelId =
        fallbackModelId ?? globallySelectedFallback;
      if (effectiveFallbackModelId) {
        const allowed = taskStatus.availableModelConnections.some(
          (connection) => connection.catalogModelId === effectiveFallbackModelId
        );
        if (!taskStatus.canSwitchModel || !allowed) {
          const label =
            modelOptions.find((option) => option.id === effectiveFallbackModelId)
              ?.label ?? effectiveFallbackModelId;
          updateFolderResumeMessage(message.id, (current) => ({
            ...current,
            text: `不能使用「${label}」接管这个断点。请选择恢复卡列出的可用模型，或切回原模型继续。`,
            planningResume: current.planningResume
              ? {
                  ...current.planningResume,
                  status: "pending",
                  pauseType: taskStatus.pauseType,
                  currentModelId: taskStatus.currentModelId,
                  canSwitchModel: taskStatus.canSwitchModel,
                  availableModelConnections:
                    taskStatus.availableModelConnections,
                  selectedFallbackModelId:
                    taskStatus.availableModelConnections[0]?.catalogModelId,
                  requestId: undefined
                }
              : undefined
          }));
          return;
        }
      }

      const prepared = await prepareBridge();
      if (!prepared) return;
      const { evidence, sourceFingerprint, sourceFingerprintSheets } = prepared;

      const clientLastSequence = Math.min(
        Math.max(
          0,
          planningCursor(resume)
        ),
        taskStatus.lastSequence
      );
      const activityMessageId = resumePlanningActivity(
        resume,
        effectiveFallbackModelId
          ? "正在使用新模型从断点继续"
          : "正在从保存的进度继续",
        effectiveFallbackModelId
          ? "已保留原任务范围、工具结果和累计预算，正在接续未完成步骤。"
          : "正在补收缺失进度，并从下一个安全边界继续。",
        clientLastSequence,
        `planning.resume.${requestId}`
      );
      const baseline = {
        ...resume,
        activityMessageId,
        clientLastSequence,
        expectedRevision: taskStatus.revision,
        lastSequence: taskStatus.lastSequence,
        requestId,
        pauseType: taskStatus.pauseType,
        currentModelId: taskStatus.currentModelId,
        modelSwitchCount: taskStatus.modelSwitchCount,
        canSwitchModel: taskStatus.canSwitchModel,
        availableModelConnections: taskStatus.availableModelConnections,
        selectedFallbackModelId:
          resume.selectedFallbackModelId ??
          taskStatus.availableModelConnections[0]?.catalogModelId,
        sourceFingerprint,
        sourceFingerprintSheets
      };
      updateFolderResumeMessage(message.id, (current) => ({
        ...current,
        planningResume: baseline
      }));
      const result = await waitForPlanning(continuePlanningTask(
        {
          turnId: resume.turnId,
          expectedRevision: taskStatus.revision,
          requestId,
          afterSequence: clientLastSequence,
          // 本轮**现读**到的证据，与上面交给桥的是同一份。
          workbookIdentity: evidence.workbookIdentity,
          authorizedSheetNames: [...evidence.authorizedSheetNames],
          sourceVerification: evidence.sourceVerification,
          ...(evidence.sourceFingerprint
            ? { sourceFingerprint: evidence.sourceFingerprint }
            : {}),
          ...(effectiveFallbackModelId
            ? {
                modelId: effectiveFallbackModelId,
                confirmModelSwitch: true as const
              }
            : {})
        },
        {
          signal: turnController.signal,
          /*
           * 续跑同样要消费桥指针。之前这里只有 `applyProgressEvent`（更新文案），
           * 指针被丢在地上 —— 模型在等本机读表，本端却在等模型出结果。
           */
          onProgress: (event) => {
            if (turnController.signal.aborted) return;
            completePlanningAttachment();
            applyProgressEvent(event);
            sdkBridge.handleProgressEvent(event);
          }
        }
      ), turnController.signal);
      if (turnController.signal.aborted) throw new DOMException("Aborted", "AbortError");
      completePlanningAttachment();
      markServerOnline();
      const outcome = await applyPlanningContinuationResult(
        message,
        result,
        baseline
      );
      finishResumeActivity(
        outcome === "failed" ? "failed" : outcome === "completed" ? "completed" : "warning",
        outcome === "completed"
          ? "已从保存的进度完成任务"
          : outcome === "failed" ? "规划已停止，未生成可执行预览" : outcome === "awaiting_input" ? "等待你补充信息；本次没有写入" : "进度已再次保存，等待你继续",
        outcome === "failed" ? "failed" : outcome !== "completed" ? "awaiting_input" : undefined
      );
    } catch (reason) {
      if (isAbortError(reason)) {
        const clientLastSequence = Math.max(resume.clientLastSequence ?? 0, getActivitySnapshot()?.sequence ?? 0);
        updateFolderResumeMessage(message.id, (current) => ({
          ...current,
          planningResume: current.planningResume
            ? { ...current.planningResume, status: "pending", requestId, awaitingStatus: true, clientLastSequence }
            : undefined
        }));
        if (pauseRequestedControllerRef.current === turnController) {
          pauseRequestedControllerRef.current = null;
          // 暂停续跑流后自动查一次；暂停查询本身则不递归查询。
          if (!inspectOnly && !resume.awaitingStatus) inspectAfterPause = { ...message,
            planningResume: { ...resume, status: "pending", requestId, awaitingStatus: true, clientLastSequence } };
        }
        finishResumeActivity("warning", "已暂停接收进度，需核对原任务现场", "awaiting_input");
      } else {
        const detail = reason instanceof Error ? reason.message : String(reason);
        if (isLocalServiceConnectionError(reason)) markServerOffline();
        updateFolderResumeMessage(message.id, (current) => ({
          ...current,
          text: `查询或接续原任务失败：${detail} 需求和任务身份仍保留；先核对状态，不重复发送需求。`,
          planningResume: current.planningResume
            ? { ...current.planningResume, status: "pending", requestId, awaitingStatus: true,
              clientLastSequence: Math.max(current.planningResume.clientLastSequence ?? 0, getActivitySnapshot()?.sequence ?? 0) }
            : undefined
        }));
        finishResumeActivity("warning", "继续分析时连接中断");
      }
    } finally {
      /*
       * **只收自己这一代。**
       *
       * 续跑可能在任何一步被新一轮取代（用户在等待期间又发了需求、或点了另一个
       * 恢复卡）。那一轮已经把桥重新立起来、正在用它读表；这里再无条件 `end()`
       * 一次，关掉的是**新轮次**的桥 —— 表现就是"新一轮规划到一半又不动了"。
       * 收尾必须核对归属，和 `turnAbortRef` 清空用的是同一个判据。
       *
       * `setStatus("idle")` 与并发闸同样要按归属收：这两处原先是无条件的，
       * 于是一次被取代的执行**晚于**新一轮收尾时，会把新一轮的运行态按回 idle
       * （暂停按钮消失、可中止对象还在）并把闸也放掉。归属不符就一个都不碰。
       */
      const ownsTurn = turnAbortRef.current === turnController;
      if (ownsTurn) {
        turnAbortRef.current = null;
        setStatus("idle");
      }
      if (planningResumeInFlightRef.current === turnController) {
        planningResumeInFlightRef.current = null;
      }
      // 本次继续到此为止：桥收尾（保留绑定给结果卡按引用取值），并让在途的
      // drain 作废 —— 迟到的事件不能落进下一轮。
      if (ownsTurn) sdkBridge.end();
    }
    if (inspectAfterPause) await resumePlanningTaskFromMessage(inspectAfterPause, undefined, true);
  }

  function beginActivity(
    ...args: Parameters<typeof startActivityProgress>
  ): void {
    const [title, detail, options] = args;
    const persistenceKey = appendMessage({
      role: "assistant",
      activityPending: true
    });
    activeActivityMessageIdRef.current = persistenceKey;
    startActivityProgress(title, detail, {
      ...(options ?? {}),
      persistenceKey
    });
  }

  /**
   * 「查看后续进展」：把视口带到接续这一段的那条回复。
   *
   * 这只是**定位**，不是"把状态搬回去"：目标任务与事件游标都不动，历史段仍然
   * 保留它停止时的内容。找不到那条消息时静默不动（不编造跳转结果）。
   */
  function jumpToMessage(id: string): void {
    const row = document.querySelector<HTMLElement>(
      `[data-message-id="${id}"]`
    );
    row?.scrollIntoView?.({ block: "start", behavior: "smooth" });
  }

  function discardRecoveryRoutingActivity(): void {    const id = getActivitySnapshot()?.persistenceKey;
    finishActivity("completed", "已识别原任务");
    if (id) setMessages((items) => items.filter((item) => item.id !== id));
  }

  // 命中判定：所有安全闸通过才返回缓存结论，否则 undefined（miss，静默重算）。
  function lookupCachedConclusion(
    intentKey: string
  ): CachedConclusion | undefined {
    const hit = lruGet(resultCacheRef.current, intentKey);
    if (!hit) return undefined;
    if (hit.sourceMode !== sourceMode) return undefined; // 闸4: 来源不同
    if (hit.sessionId !== CACHE_SESSION_ID) return undefined; // 闸5: 跨会话
    if (hit.completeness !== "complete") return undefined; // 闸6: 截断结果不复用
    if (dataEpochsChanged(hit.dataEpochSnapshot)) return undefined; // 闸1: 数据变更
    return hit;
  }

  // 命中后直接产出复用消息，跳过意图确认 / 本地扫描 / 模型生成。
  function reuseCachedConclusion(hit: CachedConclusion): void {
    appendMessage({
      role: "assistant",
      text: hit.answerText,
      resultContext: hit.resultContext,
      reused: true,
      provider: "local",
      querySourceMode: hit.sourceMode,
      querySourceSheetNames: hit.querySourceSheetNames,
      querySourceSheetIds: hit.querySourceSheetIds
    });
  }

  // 写缓存：仅只读 answer 结论且数据完整时写入；写入操作永不缓存。
  function cacheConclusion(
    intentKey: string,
    conclusion: Omit<
      CachedConclusion,
      "intentKey" | "sourceMode" | "sessionId" | "createdAt"
    >
  ): void {
    const entry: CachedConclusion = {
      ...conclusion,
      intentKey,
      sourceMode,
      sessionId: CACHE_SESSION_ID,
      createdAt: Date.now()
    };
    lruSet(resultCacheRef.current, intentKey, entry, RESULT_CACHE_LIMIT);
    const normPrompt = normalizePrompt(rawPromptRef.current);
    if (normPrompt) {
      lruSet(
        promptKeyCacheRef.current,
        normPrompt,
        intentKey,
        PROMPT_KEY_CACHE_LIMIT
      );
    }
  }

  async function scan(options?: { announce?: boolean }): Promise<WorkbookSnapshot | null> {
    const diagnosticStartedAt = performance.now();
    setStatus("scanning");
    try {
      const snapshot = isRunningInExcel()
        ? await captureWorkbookCatalog()
        : demoWorkbook;
      setWorkbook(snapshot);
      applyWorkbookSnapshotSelection(snapshot);
      setContextOpen(false);
      if (options?.announce) {
        appendMessage({
          role: "system",
          text: `已重新读取「${snapshot.name}」：${snapshot.worksheets.length} 个工作表。`
        });
      }
      recordDiagnosticEvent({
        timestamp: new Date().toISOString(),
        phase: "scan",
        durationMs: performance.now() - diagnosticStartedAt,
        modelCalls: 0,
        status: "succeeded"
      });
      // 快照返回给调用方：删表后要用它核对"剩余表集合是否就是回执里的那一份"。
      return snapshot;
    } catch (reason) {
      recordDiagnosticEvent({
        timestamp: new Date().toISOString(),
        phase: "scan",
        durationMs: performance.now() - diagnosticStartedAt,
        modelCalls: 0,
        status: "failed",
        errorCategory: "data"
      });
      appendMessage({
        role: "system",
        text: reason instanceof Error ? reason.message : "读取工作簿失败"
      });
      return null;
    } finally {
      setStatus("idle");
    }
  }

  /**
   * 删表成功后的**预期收窄**：把授权范围与目标基线同步到"剩下的表"。
   *
   * 删表本来就会让工作簿少一张表。如果什么都不做，下一轮目标守卫会把它读成
   * "用户改了授权范围"，弹一句"工作表范围发生变化 / 尚未写入" —— 与刚完成的
   * 执行结果直接冲突。
   *
   * 放过它的前提必须被证明，而不是"刚删完就放过"：核对在共享层
   * `reconcileWorksheetDeletionScope`（同一工作簿身份、回执剩余表和真实目录逐名
   * 相同、删除清单恰好等于授权表差集、收窄后仍有授权表）。这里只做两件事：
   * 1. 用真实目录读一次目标（拿当前表 id），交给 `bindVerifiedShrink` 再卡一遍
   *    "只减不增"，通过后才推进基线；
   * 2. 把已提交范围对齐到同一份集合，别让"草稿未确认"再提示一次。
   *
   * 任何一步不成立就原样返回：守卫仍会按范围变化拦下，用户照旧可以重新确认。
   */
  async function adoptVerifiedDeletionShrink(input: {
    identity: string;
    authorized: readonly string[];
    snapshot: WorkbookSnapshot | null;
    requestMessageId: string | null;
    conversationId: string;
  }): Promise<boolean> {
    if (!input.snapshot) return false;
    const receipt = worksheetCheckpoint.receiptForMessage(input.requestMessageId, input.conversationId);
    const shrink = reconcileWorksheetDeletionScope({
      beforeIdentity: input.identity,
      afterIdentity: currentOfficeWorkbookIdentity() ?? null,
      authorizedSheets: input.authorized,
      receipt,
      actualSheets: input.snapshot.worksheets.map((sheet) => sheet.name),
      isSystemSheet: isEBSystemSheet
    });
    if (!shrink.ok) return false;
    try {
      const target = await inspectOfficeAuthorizedTarget(shrink.scope);
      if (!shrink.scope.every((name) => target.sheets.some((sheet) => sheet.name === name))) return false;
      if (!targetGuard.bindVerifiedShrink(target)) return false;
    } catch {
      return false;
    }
    setSelectedSheetNames(shrink.scope);
    // 与 `bindVerifiedShrink` 同批：已提交范围就在这一帧里对齐到新基线，
    // 不留"基线已推进、已提交还是旧的"的中间态。
    commitAuthorizedScope(shrink.scope);
    return true;
  }

  /**
   * 原地撤销成功后的**预期目录增加**：把授权与目标基线就地认下来。
   *
   * 撤销把表放回了工作簿，目录因此比授权时多几张。什么都不做的话，下一轮守卫会
   * 把它读成"范围变了"，弹「尚未写入」——与刚撤销的事实冲突。
   *
   * 放过它同样要证据（共享层 `reconcileWorksheetRestoreScope`）：同簿身份、授权过的表
   * 全在、回执里放回的表真的出现、目录新增**只**来自这次撤销。任一条不成立就原样返回，
   * 守卫照常拦截外部改动与换簿。授权集合本身**不变**——恢复出来的表不会被自动授权。
   */
  async function adoptVerifiedRestore(input: {
    identity: string | null;
    authorized: readonly string[];
    restoredSheets: readonly string[];
    snapshot: WorkbookSnapshot | null;
  }): Promise<boolean> {
    if (!input.snapshot || !input.identity) return false;
    const check = reconcileWorksheetRestoreScope({
      beforeIdentity: input.identity,
      afterIdentity: currentOfficeWorkbookIdentity() ?? null,
      authorizedSheets: input.authorized,
      restoredSheets: input.restoredSheets,
      actualSheets: input.snapshot.worksheets.map((sheet) => sheet.name),
      isSystemSheet: isEBSystemSheet
    });
    if (!check.ok) return false;
    try {
      const target = await inspectOfficeAuthorizedTarget(check.authorized);
      if (!check.authorized.every((name) => target.sheets.some((sheet) => sheet.name === name))) return false;
      if (!targetGuard.bindVerifiedRestore(target, input.restoredSheets)) return false;
    } catch {
      return false;
    }
    setSelectedSheetNames(check.authorized);
    // 撤销不扩大授权：这里提交的仍是原来那几张表，只是把基线一起对齐。
    commitAuthorizedScope(check.authorized);
    return true;
  }

  useEffect(() => {
    if (typeof Office === "undefined") {
      setWorkbook(demoWorkbook);
      setSelectedSheetNames([demoWorkbook.activeWorksheet]);
      return;
    }
    let dispose: (() => void) | undefined;
    Office.onReady(async () => {
      await scan();
      if (isRunningInExcel()) {
        dispose = await watchWorkbookStructureChanges(() => {
          setSelectionConfirmed(false);
          clearUndoSnapshot();
        });
      }
    });
    return () => dispose?.();
  }, []);

  // UI 协调包装函数（调用 Hook 的方法并处理 UI 状态）
  function handleNewChat() {
    if (busy) return;
    localMappingDraftsRef.current.clear();
    setModelMenuOpen(false);
    setMoreMenuOpen(false);
    const previousSaved = Boolean(
      activeConversation?.messages.some((message) => message.role === "user")
    );
    if (previousSaved) conversationHook.newChat();
    setPrompt("");
    setPendingImages([]);
    setImageError("");
    setContextOpen(false);
    setHistoryOpen(false);
    setToolsOpen(false);
    closeSettings();
    closeSaveCandidate();
    setNewConversationNotice({ id: Date.now(), previousSaved });
    window.requestAnimationFrame(() => composerInputRef.current?.focus());
  }

  function handleOpenConversation(conversationId: string) {
    if (busy) return;
    localMappingDraftsRef.current.clear();
    conversationHook.openConversation(conversationId);
    setPrompt("");
    setPendingImages([]);
    setImageError("");
    setContextOpen(false);
    setHistoryOpen(false);
    closeSettings();
    closeSaveCandidate();
  }

  function handleDeleteConversation(conversationId: string) {
    if (busy) return;
    localMappingDraftsRef.current.clear();
    conversationHook.deleteConversation(conversationId);
  }

  function formatConversationTime(value: string): string {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "";
    return new Intl.DateTimeFormat("zh-CN", {
      month: "numeric",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit"
    }).format(date);
  }


  async function browseFolder(): Promise<FolderCatalog | null> {
    setStatus("scanning");
    try {
      const catalog = await selectFolder();
      if (!catalog) return null;
      if (
        !folderCatalog ||
        normalizeWorkbookPath(folderCatalog.folderPath) !==
          normalizeWorkbookPath(catalog.folderPath)
      ) {
        // 初次设置或切换目录时不能沿用旧授权勾选。
        setFolderSheetKeys([]);
      }
      applyFolderCatalog(catalog);
      setActiveFolderFileId(
        (catalog.files.find((file) => !file.error) ?? catalog.files[0])?.id ??
          null
      );
      setFolderPickerView("files");
      setFolderSheetSearch("");
      setShowSelectedFolderSheets(false);
      setRecentlyAddedFolderSheetKeys([]);
      setFolderRefreshNotice(
        `已读取 ${catalog.files.length} 个文件、${catalog.files.reduce(
          (total, file) =>
            total +
            file.worksheets.filter((sheet) => !isEBSystemSheet(sheet.name))
              .length,
          0
        )} 张工作表`
      );
      markServerOnline();
      return catalog;
    } catch (reason) {
      appendMessage({
        role: "system",
        text: reason instanceof Error ? reason.message : "读取文件夹失败"
      });
      return null;
    } finally {
      setStatus("idle");
    }
  }

  async function chooseWorkDirectoryFromGuide() {
    const catalog = await browseFolder();
    if (!catalog) return;
    // 这里只登记候选目录；工作表仍需用户逐项勾选后才进入授权范围。
    setFolderSheetKeys([]);
    chooseFolderScope();
    setSelectionConfirmed(false);
    dismissWorkDirectoryGuide();
    setContextOpen(true);
  }

  /** 会话失效的错误特征。后端只会说「已失效」，「已过期」是历史前端文案。 */
  function isFolderSessionLost(message: string): boolean {
    return (
      message.includes("会话已失效") ||
      message.includes("不存在或已过期") ||
      message.includes("会话已过期")
    );
  }

  /**
   * 会话丢失后按本地记录重建候选目录，只重放本面板会话明确确认过的授权。
   *
   * 后端会话存在纯内存字典里，开发期 --reload 一重启就清空。前端本地已经存了
   * localStorage 的勾选只是草稿，不能当授权。create_folder_snapshot 只恢复
   * confirmedFolderScopeRef；新范围仍须由用户显式勾选。
   *
   * 重建**不掩盖数据变更**：新快照按当前文件状态重算 source_fingerprint，旧计划带的
   * 是生成时的指纹，执行时比对不一致仍会被拒，照旧走「数据已变」二次确认。
   */
  async function rebuildFolderSession(): Promise<FolderCatalog | null> {
    const folderPath = folderCatalog?.folderPath;
    if (!folderPath) return null;
    const confirmedScope = confirmedFolderScopeRef.current;
    const wantedKeys = confirmedScope?.root === folderRootIdentity(folderPath)
      ? confirmedScope.keys : [];
    const keepConfirmation = selectionConfirmed && wantedKeys.length > 0 &&
      wantedKeys.length === folderSheetKeys.length &&
      wantedKeys.every((key) => folderSheetKeys.includes(key));
    try {
      const rebuilt = await scanFolderByPath(folderPath);
      if (folderRootIdentity(rebuilt.folderPath) !== folderRootIdentity(folderPath)) {
        throw new Error("工作目录身份发生变化");
      }
      // fileId 是相对路径的 uuid5，跨会话稳定，所以旧 key 能直接在新 catalog 里匹配。
      const selections = rebuilt.files
        .map((file) => ({
          fileId: file.id,
          sheets: file.worksheets
            .filter((sheet) => wantedKeys.includes(folderSheetKey(file.id, sheet.name)))
            .map((sheet) => sheet.name)
        }))
        .filter((selection) => selection.sheets.length > 0);
      const wantedSheetCount = wantedKeys.length;
      const foundSheetCount = selections.reduce(
        (total, selection) => total + selection.sheets.length,
        0
      );
      if (wantedSheetCount > 0 && foundSheetCount < wantedSheetCount) {
        // 有源文件或工作表在这期间被删/改名了，不静默丢弃勾选。
        throw new Error("原先勾选的工作表已不存在");
      }
      applyFolderCatalog(rebuilt);
      if (selections.length > 0) {
        const snapshot = await createFolderSnapshot(rebuilt.sessionId, selections);
        setWorkbook(snapshot);
        setSelectedSheetNames(snapshot.worksheets.map((sheet) => sheet.name));
        if (keepConfirmation) commitAuthorizedScope(snapshot.worksheets.map((sheet) => sheet.name));
        else setSelectionConfirmed(false);
      }
      appendMessage({
        role: "system",
        text: `文件夹会话已重建：${rebuilt.folderName}`
      });
      return rebuilt;
    } catch (reason) {
      setFolderCatalog(null);
      setFolderSheetKeys([]);
      setSelectionConfirmed(false);
      appendMessage({
        role: "system",
        text: `文件夹会话已失效且无法自动重建（${
          reason instanceof Error ? reason.message : String(reason)
        }），请重新选择文件夹`
      });
      return null;
    }
  }

  /**
   * 包装文件夹操作：遇到会话失效就重建后重试一次。
   *
   * operation 接收当前 sessionId 而非从闭包里读 folderCatalog，因为重建后的新
   * sessionId 要到下一次渲染才进 state，闭包里拿到的还是旧值。
   * 只重试一次，避免死循环。
   */
  async function withFolderSessionRecovery<T>(
    sessionId: string,
    operation: (sessionId: string) => Promise<T>
  ): Promise<T> {
    try {
      return await operation(sessionId);
    } catch (reason) {
      const message = reason instanceof Error ? reason.message : String(reason);
      if (!isFolderSessionLost(message)) throw reason;
      const rebuilt = await rebuildFolderSession();
      if (!rebuilt) throw reason;
      return operation(rebuilt.sessionId);
    }
  }

  function applyFolderRefreshResult(
    refreshed: FolderCatalog,
    quiet: boolean
  ) {
    setScopeConfirmationError("");
    const addedKeys = addedFolderSheetKeys(folderCatalog, refreshed);
    applyFolderCatalog(refreshed, { preserveAuthorization: true });
    setRecentlyAddedFolderSheetKeys(addedKeys);
    if (addedKeys.length > 0) {
      const separator = addedKeys[0].indexOf("\0");
      if (separator > 0) {
        setActiveFolderFileId(addedKeys[0].slice(0, separator));
        setFolderPickerView("sheets");
      }
      setFolderRefreshNotice(`发现 ${addedKeys.length} 张新增工作表`);
      return;
    }
    setFolderRefreshNotice(
      quiet
        ? ""
        : `已刷新，共 ${refreshed.files.length} 个文件、${refreshed.files.reduce(
            (total, file) =>
              total +
              file.worksheets.filter(
                (sheet) => !isEBSystemSheet(sheet.name)
              ).length,
            0
          )} 张工作表`
    );
  }

  async function refreshCurrentFolder(options: { quiet?: boolean } = {}) {
    if (!folderCatalog || status !== "idle" || scopeCommitLock.current) return;
    setStatus("scanning");
    setScopeCatalogChecking(true);
    setScopeConfirmationError("");
    try {
      const refreshed = await refreshFolder(folderCatalog.sessionId);
      applyFolderRefreshResult(refreshed, options.quiet ?? false);
      markServerOnline();
    } catch (reason) {
      const message = reason instanceof Error ? reason.message : String(reason);
      if (isFolderSessionLost(message) || message.includes("不存在")) {
        // 会话没了但本地还记着 folderPath，直接重建等价于一次刷新。
        const rebuilt = await rebuildFolderSession();
        if (rebuilt) {
          // rebuildFolderSession already rechecked the selected files and restored
          // this session's authorization. Do not revoke it using the expired id.
          setFolderRefreshNotice(options.quiet ? "" : `已刷新，共 ${rebuilt.files.length} 个文件`);
        }
        markServerOnline();
      } else {
        appendMessage({
          role: "system",
          text: `刷新文件夹失败：${message}`
        });
      }
    } finally {
      // Retry access checks even if the catalog is unchanged or refresh failed.
      setScopeCatalogRevision((value) => value + 1);
      setScopeCatalogChecking(false);
      setStatus("idle");
    }
  }

  /** 每次点「刷新」自增：作为图标的 `key`，换 key 即重挂载 → 旋转动画重放一次。 */
  const [refreshSpin, setRefreshSpin] = useState(0);

  /**
   * 重新读取当前工作簿，把选择器拉回真实宿主。
   *
   * 返回读到的快照（失败/被挡时 `null`）——「从工作目录切回工作簿来源」要拿它重置名字草稿，
   * 因为那时 `workbook` 与草稿都还是文件夹快照的「文件.xlsx › Sheet」复合名。
   */
  async function refreshWorkbook(
    options: { spin?: boolean; ignoreBusy?: boolean } = {}
  ): Promise<WorkbookSnapshot | null> {
    // 先记一次点击反馈（放在最前面，`busy` 早退时同样要转）。
    if (options.spin !== false) setRefreshSpin((value) => value + 1);
    // `ignoreBusy` 只给「从工作目录切回工作簿来源」用：那一刻可能还有一次目录刷新在飞行中
    // （打开面板就会静默刷一次），等它结束再读会把切回变成静默失效 —— 用户实测踩到的正是这个。
    // 重读本身有 fileKey/fileRef 的一致性复核，并发重入最坏也只是把同一份真实宿主多写一遍。
    if (!options.ignoreBusy && (busy || scopeCommitLock.current)) return null;
    setScopeConfirmationError("");
    setStatus("scanning");
    setScopeCatalogChecking(true);
    try {
      const fileKey = isRunningInExcel() ? currentOfficeWorkbookIdentity() ?? null : null;
      const fileRef = isRunningInExcel() ? Office.context.document : null;
      const refreshed = isRunningInExcel()
        ? await captureWorkbookCatalog()
        : demoWorkbook;
      if (isRunningInExcel()) {
        if (fileKey !== (currentOfficeWorkbookIdentity() ?? null) || fileRef !== Office.context.document) {
          throw new Error("文件在刷新期间发生变化，请重新刷新。");
        }
        scopePickerFile.current = { fileKey, fileRef };
        setScopeCatalogRevision((value) => value + 1);
      }
      setWorkbook(refreshed);
      if (sourceMode === "workbook" && workbookScopeMode === "auto" &&
          (selectedSheetNames.length !== 1 || selectedSheetNames[0] !== refreshed.activeWorksheet)) {
        // Refresh updates the automatic draft, never grants a new scope. Manual
        // selections (including missing items) remain for the user to resolve.
        setSelectedSheetNames([refreshed.activeWorksheet]);
      }
      return refreshed;
    } catch (reason) {
      scopePickerFile.current = null;
      setScopeCatalogRevision((value) => value + 1);
      setScopeConfirmationError("暂时无法读取目标，请刷新选择器；未变更授权。");
      appendMessage({
        role: "system",
        text: `刷新工作簿失败：${
          reason instanceof Error ? reason.message : String(reason)
        }`
      });
      return null;
    } finally {
      setScopeCatalogChecking(false);
      setStatus("idle");
    }
  }

  /** 选择范围模式即请求授权。由工作目录切回时先读取真实宿主目录，再核验本次选择。 */
  async function switchScopeMode(mode: "auto" | "manual"): Promise<void> {
    if (sourceMode !== "folder") {
      if (mode === "auto") chooseAutomaticScope(workbook);
      else chooseManualScope(workbook);
      immediateScope.request();
      return;
    }
    if (mode === "auto") chooseAutomaticScope(null);
    else chooseManualScope(null);
    if (scopeCommitLock.current) return;
    // 重读真实宿主不受 `busy` 闸拦：切回工作簿来源的那一刻，打开面板触发的目录刷新
    // 可能仍在飞行中（`busy=true`），等它结束再读会把切回变成静默失效。`refreshWorkbook`
    // 用 `ignoreBusy` 放行这一次，重读内部仍有 fileKey/fileRef 的一致性复核兜底。
    const refreshed = await refreshWorkbook({ spin: false, ignoreBusy: true });
    if (refreshed && mode === "manual") {
      setSelectedSheetNames([refreshed.activeWorksheet]);
    }
    if (refreshed) immediateScope.request();
  }

  /** 收起保留勾选，不撤回授权，也不取消在途核验。 */
  function discardWorkbookScopeEdit(): void { setContextOpen(false); }

  async function confirmSheetSelection(isCurrent: () => boolean): Promise<boolean> {
    scopeCommitLock.current = true;
    setScopeCheckingCommit(true);
    setScopeConfirmationError("");
    try {
      if (scopeAdmissionIds.length === 0) {
        if (sourceMode === "folder") confirmedFolderScopeRef.current = null;
        else if (isRunningInExcel()) {
          const empty = await inspectOfficeAuthorizedTarget([]);
          if (!isCurrent()) return false;
          if (!(await targetGuard.confirmScope(empty, { isCurrent, allowEmpty: true }))) return false;
        }
        if (!isCurrent()) return false;
        commitAuthorizedScope([], "manual");
        return true;
      }
      const checked = await scopeAdmission.verify();
      if (!isCurrent() || !checked) return false;
    if (sourceMode === "folder") {
      if (!folderCatalog || folderSheetKeys.length === 0) return false;
      setStatus("scanning");
      try {
        const selections = folderSelections();
        const selectedFile =
          selections.length === 1
            ? folderCatalog.files.find(
                (file) => file.id === selections[0].fileId
              )
            : null;
        const currentWorkbookPath = isRunningInExcel()
          ? normalizeWorkbookPath(Office.context.document.url)
          : null;
        const selectedFilePath = selectedFile
          ? folderFileFullPath(
              folderCatalog.folderPath,
              selectedFile.relativePath
            )
          : null;
        if (
          selectedFile &&
          currentWorkbookPath &&
          currentWorkbookPath === selectedFilePath
        ) {
          try {
            const nameSnapshot = await captureWorkbookCatalog();
            if (!isCurrent()) return false;
            const actualNames = new Map(
              nameSnapshot.worksheets.map((sheet) => [
                sheet.name.toLocaleLowerCase(),
                sheet.name
              ])
            );
            const selectedNames = selections[0].sheets
              .map((name) => actualNames.get(name.toLocaleLowerCase()))
              .filter((name): name is string => name !== undefined);
            if (selectedNames.length === selections[0].sheets.length) {
              const officeCheck = await inspectOfficeScopeAdmission(selectedNames);
              if (!isCurrent()) return false;
              if (!admissionComplete(selectedNames, officeCheck) || !officeCheck.target) {
                setScopeConfirmationError("当前打开的工作簿暂时无法核实，请刷新后重试；未变更授权。");
                return false;
              }
              const snapshot = await captureWorkbook(selectedNames);
              const after = await inspectOfficeAuthorizedTarget(selectedNames);
              if (!isCurrent()) return false;
              if (checkAuthorizedTarget(officeCheck.target, after) !== "same" || officeCheck.target.fileRef !== after.fileRef) {
                setScopeConfirmationError("目标在确认期间发生变化，请刷新后重新选择；未变更授权。");
                return false;
              }
              const snapshotNames = new Set(
                snapshot.worksheets.map((sheet) =>
                  sheet.name.toLocaleLowerCase()
                )
              );
              if (
                selectedNames.every((name) =>
                  snapshotNames.has(name.toLocaleLowerCase())
                )
              ) {
                setWorkbook(snapshot);
                targetGuard.bind(after);
                adoptCurrentWorkbookScope(snapshot, selectedNames);
                appendMessage({
                  role: "system",
                  text: "所选文件就是当前工作簿，已改为直接在 Excel 中处理。"
                });
                return true;
              }
            }
          } catch {
            // 无法从宿主实时确认时保留文件夹语义，由后端锁探测继续兜底。
          }
        }
        const snapshot = await withFolderSessionRecovery(
          folderCatalog.sessionId,
          (sessionId) => createFolderSnapshot(sessionId, selections)
        );
        if (!isCurrent()) return false;
        const root = folderRootIdentity(folderCatalog.folderPath);
        if (!root) throw new Error("无法核实工作目录身份，请重新选择目录。");
        confirmedFolderScopeRef.current = {
          root,
          keys: selections.flatMap((selection) => selection.sheets.map((name) => folderSheetKey(selection.fileId, name)))
        };
        setWorkbook(snapshot);
        setSelectedSheetNames(snapshot.worksheets.map((sheet) => sheet.name));
        // 文件夹的已提交范围住在 `confirmedFolderScopeRef`（上面刚写），
        // 工作簿口径的已提交名单同步到同一批表；守卫此刻不启用（来源是文件夹）。
        commitAuthorizedScope(snapshot.worksheets.map((sheet) => sheet.name));
        return true;
      } catch (reason) {
        const message = reason instanceof Error ? reason.message : "读取所选工作表失败";
        // 会话失效的分支已由 withFolderSessionRecovery 处理（重建失败时它已提示重选）。
        if (!isFolderSessionLost(message)) {
          if (isCurrent()) setScopeConfirmationError(message);
        }
      } finally {
        setStatus("idle");
      }
      return false;
    }

    if (!workbook || selectedSheetNames.length === 0) return false;
    if (isRunningInExcel()) {
      setScopeConfirmationError("");
      try {
        if (!(await confirmOfficeAuthorizedScope(
          workbookScopeMode === "auto" ? null : selectedSheetNames,
          (target) => targetGuard.confirmScope(target, { isCurrent }), checked.target
        ))) {
          setScopeConfirmationError("暂时无法核实目标，请刷新后重试；未变更授权。");
          return false;
        }
        // Target-dialog adoption defaults to manual; the range picker must retain
        // the automatic/manual mode explicitly chosen for this confirmation.
        if (!isCurrent()) return false;
        setWorkbookScopeMode(workbookScopeMode);
      } catch {
        setScopeConfirmationError("暂时无法读取目标身份，请刷新后重试；未变更授权。");
        return false;
      }
      /*
       * 有守卫的那条路：`targetGuard.confirmScope` 里的 `adopt` 已经把
       * "范围 + 口径"同批写进已提交（守卫要立刻拿到新基线）。这里**不再**用草稿
       * 覆写一次 —— 草稿里若有被目标过滤掉的表名，覆写会让已提交与基线差一个名字。
       */
    } else {
      // 没有守卫的那条路（未在 Excel 中运行）：确认动作自己提交。
      commitAuthorizedScope(selectedSheetNames, workbookScopeMode);
    }
    return true;
    } finally {
      scopeCommitLock.current = false;
      setScopeCheckingCommit(false);
    }
  }

  function buildIntentScope(
    snapshot: WorkbookSnapshot,
    sheetNames: string[]
  ): IntentScopeContext {
    const selected = new Set(sheetNames);
    const sheets = snapshot.worksheets
      .filter((sheet) => selected.has(sheet.name))
      .map((sheet) => ({
        name: sheet.name,
        usedRange: sheet.usedRange,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        headers: sheet.headers.slice(0, INTENT_MAX_FIELDS)
      }));
    const fallback = snapshot.worksheets.find(
      (sheet) => sheet.name === snapshot.activeWorksheet
    );
    if (sheets.length === 0 && fallback) {
      sheets.push({
        name: fallback.name,
        usedRange: fallback.usedRange,
        rowCount: fallback.rowCount,
        columnCount: fallback.columnCount,
        headers: fallback.headers.slice(0, INTENT_MAX_FIELDS)
      });
    }
    const folderSheetCount =
      folderCatalog?.files.reduce(
        (total, file) => total + file.worksheets.length,
        0
      ) ?? snapshot.worksheets.length;
    return {
      workbookName: snapshot.name,
      sourceMode,
      selectionMode:
        sourceMode === "folder" ? "folder" : workbookScopeMode,
      activeWorksheet: snapshot.activeWorksheet,
      selectedRange: snapshot.selectedRange ?? null,
      worksheetNames: snapshot.worksheets.map((ws) => ws.name),
      totalWorksheetCount:
        sourceMode === "folder"
          ? Math.max(folderSheetCount, sheets.length)
          : snapshot.worksheets.length,
      sheets
    };
  }

  /**
   * 服务端下发列映射**规则**时，在本端读出授权源列、算成受白名单约束的计划。
   *
   * 为什么 Excel 也走这条：服务端只在 `planningReadPolicy: "none"` 时下发规则
   * （它手上没有源值）。此前的 Excel 分发直接抛"本端不支持该模式"，于是同一个
   * `map_column_values` 需求在 WPS 能用、在 Excel 报错。现在两端复用共享 core
   * 的同一套计算（`buildColumnMappingPlan` → `buildColumnMappingActions`）。
   *
   * 三条不变量：
   * - **只读授权工作表**：源表不在本次已勾选范围内就拒绝，不自动扩大授权；
   * - **源值不外发**：读出的值只在本端计算，算好的矩阵不回传服务端；
   * - **不放宽写入闸门**：产出的仍是普通 `AnalysisPlan`，照常进预览卡、
   *   通过授权与恢复准备后自动写入，再回读审阅；确认不重复执行。
   */
  async function localMappingPlanFromRule(
    response: Extract<AssistantResponse, { kind: "local_plan_rule" }>,
    effectiveSheetNames: readonly string[]
  ): Promise<AnalysisPlan> {
    const rule = response.rule;
    if (!isRunningInExcel()) {
      throw new Error(
        "本地列映射需要读取真实工作簿，浏览器演示模式不支持。请在 Excel 中打开任务窗格。"
      );
    }
    if (
      !effectiveSheetNames.some(
        (name) => name.toLocaleLowerCase() === rule.sourceSheet.toLocaleLowerCase()
      )
    ) {
      throw new Error(
        `映射规则要读取「${rule.sourceSheet}」，但它不在本次已授权的工作表里。`
      );
    }
    advanceActivity(
      "正在读取授权源列",
      "只在已勾选的工作表内读取映射源值，源值留在本机，不上传服务端。"
    );
    const sheet = await readSheetUsedRange(rule.sourceSheet);
    if (!sheet) {
      throw new Error(`来源工作表「${rule.sourceSheet}」当前没有可用区域。`);
    }
    advanceActivity(
      "正在按规则计算目标列",
      "使用与 WPS、服务端同一套确定性映射实现，在本机逐格计算。"
    );
    /*
     * 实际来源依赖 = **刚才真的读过的那张表**（`readSheetUsedRange(rule.sourceSheet)`），
     * 这是确定性本地读取的记账，不是模型声明的参数，也不是"授权范围减目标"。
     * 映射规则的目标列通常就落在同一张表上 → 读写重叠 → 照旧不登记断点恢复。
     */
    return attestSourceDependency(
      buildColumnMappingPlan(rule, sheet, {
        summary: response.summary,
        warnings: response.warnings ?? [],
        assumptions: response.assumptions ?? []
      }),
      { origin: "deterministic-local-read", sheets: [rule.sourceSheet] }
    );
  }

  /** 空白策略未定且**真读到空白**：结构化澄清，不是报错，也不替用户挑默认值。 */
  function recordExcelBlankPolicyGap(
    input: {
      response: Extract<AssistantResponse, { kind: "local_plan_rule" }>;
      originalPrompt: string;
      authorizedSheets: string[];
      scopeFingerprint: string;
      messageId?: string;
    },
    blankCount: number
  ): boolean {
    const workbookIdentity = currentOfficeWorkbookIdentity();
    if (!workbookIdentity) {
      if (input.messageId) {
        localMappingDraftsRef.current.delete(input.messageId);
        setMessages((current) => current.map((message) =>
          message.id === input.messageId && message.clarification
            ? { ...message, clarification: { ...message.clarification, status: "invalidated" } }
            : message
        ));
      }
      appendMessage({
        role: "system",
        text: "当前工作簿没有可验证的身份，无法安全保留本地映射草稿。请先保存工作簿、确认授权范围，再重新发送需求；原表未写入。"
      });
      setPrompt(input.originalPrompt);
      return true;
    }
    const question =
      `源区域读出来有 ${blankCount} 个空白格，但这些位置的目标格应当怎么处理还没定。` +
      "请选择：保留目标原值 / 清空目标内容 / 写入指定值 / 写入空字符串。";
    const id = input.messageId ?? appendMessage({
      role: "assistant",
      text: "",
      clarification: {
        id: `local-mapping-${Date.now().toString(36)}`,
        localMapping: true,
        summary: "源区域有空白，要先定空白怎么处理",
        question,
        reason: "空白策略只在本机决定，不会送到模型；选定后会重新读取并生成预览，原表尚未写入。",
        scopeLabel: `${input.response.rule.sourceSheet}!${input.response.rule.sourceRange}`,
        options: [],
        originalPrompt: input.originalPrompt,
        scopeFingerprint: input.scopeFingerprint,
        round: 0,
        status: "pending"
      },
      provider: "local"
    });
    const revision = (localMappingDraftsRef.current.get(id)?.revision ?? 0) + 1;
    localMappingDraftsRef.current.set(id, {
      rule: input.response.rule,
      context: {
        summary: input.response.summary,
        warnings: input.response.warnings,
        assumptions: input.response.assumptions,
        provider: input.response.provider
      },
      // 空白格没有值可带；带值只会违反"源值只留本端"。
      values: [],
      allowExactMapping: false,
      reason: "blank",
      blankCount,
      scopeFingerprint: input.scopeFingerprint,
      authorizedSheets: [...input.authorizedSheets],
      workbookIdentity,
      revision
    });
    if (input.messageId) {
      setMessages((current) => current.map((message) =>
        message.id === id && message.clarification
          ? { ...message, clarification: {
              ...message.clarification, question, summary: "源区域有空白，要先定空白怎么处理", status: "pending"
            } }
          : message
      ));
    }
    return true;
  }

  function recordExcelLocalMappingGap(input: {
    error: unknown;
    response: Extract<AssistantResponse, { kind: "local_plan_rule" }>;
    originalPrompt: string;
    authorizedSheets: string[];
    scopeFingerprint: string;
    messageId?: string;
  }): boolean {
    if (input.error instanceof BlankPolicyNeededError) {
      return recordExcelBlankPolicyGap(input, input.error.blankCount);
    }
    if (!(input.error instanceof UnmappedCategoryError) &&
        !(input.error instanceof NeedsClarificationError)) return false;
    const values = input.error instanceof UnmappedCategoryError
      ? input.error.values : input.error.cells;
    const allowExactMapping = input.error instanceof NeedsClarificationError ||
      !input.response.rule.delimiters?.length;
    const workbookIdentity = currentOfficeWorkbookIdentity();
    if (!workbookIdentity) {
      // 没有可验证的工作簿身份时，展示可填写却永远过不了续答复核的卡片是死路。
      // 明确请用户先保存并重发；不把旧源值或补充映射带到下一轮。
      if (input.messageId) {
        localMappingDraftsRef.current.delete(input.messageId);
        setMessages((current) => current.map((message) =>
          message.id === input.messageId && message.clarification
            ? { ...message, clarification: { ...message.clarification, status: "invalidated" } }
            : message
        ));
      }
      appendMessage({
        role: "system",
        text: "当前工作簿没有可验证的身份，无法安全保留本地映射草稿。请先保存工作簿、确认授权范围，再重新发送需求；原表未写入。"
      });
      setPrompt(input.originalPrompt);
      return true;
    }
    const question = allowExactMapping
      ? `源区域有 ${values.length} 个不同来源值未能按现有规则归类。请选择整格映射，或确认拆分符与优先顺序。`
      : `源区域有 ${values.length} 个拆分项未被规则覆盖。请补齐映射并确认多个类别的优先顺序。`;
    const id = input.messageId ?? appendMessage({
      role: "assistant",
      text: "",
      clarification: {
        id: `local-mapping-${Date.now().toString(36)}`,
        localMapping: true,
        summary: "本地映射需要你决定",
        question,
        reason: "来源值只在本机展示，不会送到模型；确认后会重新读取并生成预览，原表尚未写入。",
        scopeLabel: `${input.response.rule.sourceSheet}!${input.response.rule.sourceRange}`,
        options: [],
        originalPrompt: input.originalPrompt,
        scopeFingerprint: input.scopeFingerprint,
        round: 0,
        status: "pending"
      },
      provider: "local"
    });
    const revision = (localMappingDraftsRef.current.get(id)?.revision ?? 0) + 1;
    localMappingDraftsRef.current.set(id, {
      rule: input.response.rule,
      context: {
        summary: input.response.summary,
        warnings: input.response.warnings,
        assumptions: input.response.assumptions,
        provider: input.response.provider
      },
      values,
      allowExactMapping,
      // 归因按**错误类型**定，不按 allowExactMapping 反推 —— 后者在
      // "未命中且没有分隔符"时也是 true，会把它错认成分隔符问题。
      reason: input.error instanceof UnmappedCategoryError
        ? "unmapped"
        : "separator_or_mapping",
      scopeFingerprint: input.scopeFingerprint,
      authorizedSheets: [...input.authorizedSheets],
      workbookIdentity,
      revision
    });
    if (input.messageId) {
      setMessages((current) => current.map((message) =>
        message.id === id && message.clarification
          ? { ...message, clarification: {
              ...message.clarification, question, status: "pending"
            } }
          : message
      ));
    }
    return true;
  }

  async function analyzeConfirmedIntent(
    confirmedPrompt: string,
    sentImages: PendingImage[],
    expectedScopeFingerprint?: string,
    clarificationMessageId?: string,
    dataResults: DataToolResult[] = [],
    preparedWorkbook?: WorkbookSnapshot,
    intentMemory?: IntentMemory,
    turnId?: string | null,
    planningReadPolicy: "ifNeeded" | "none" = sourceMode === "workbook" ? "none" : "ifNeeded",
    /** 这一轮的澄清轮次。列映射的澄清卡要沿用同一个计数，不能各自从 0 起。 */
    clarificationRound = 0,
    taskContext?: AgentTaskContext,
    planningGoal?: "read_only" | "write_preview"
  ): Promise<TurnPlanningOutcome> {
    if (!workbook) return "skipped";
    if (!(await targetGuard.checkForAnalysis())) return "skipped";
    let activityOutcome: TurnPlanningOutcome = "planned";
    const planningController = turnAbortRef.current;
    planningReadPolicy = planningReadPolicyForModel(effectiveModelId, planningReadPolicy);
    setStatus("planning");
    if (turnId) bindActivityOperation(turnId);
    if (dataResults.length > 0) {
      const scannedRows = dataResults.reduce(
        (total, result) => total + result.scannedRows,
        0
      );
      advanceActivity(
        "正在准备规划请求",
        "整理本地紧凑结果与授权范围，不接收完整工作表数据。",
        `本地计算完成，共扫描 ${scannedRows.toLocaleString()} 行`
      );
    } else {
      advanceActivity(
        "正在准备规划请求",
        "正在重新核对授权范围并准备受约束的规划请求。"
      );
    }
    try {
      let latestWorkbook = preparedWorkbook ?? workbook;
      let effectiveSheetNames = selectedNamesFor(latestWorkbook);
      if (
        dataResults.length === 0 &&
        sourceMode === "workbook" &&
        isRunningInExcel()
      ) {
        latestWorkbook = await captureWorkbook(
          workbookScopeMode === "manual" ? selectedSheetNames : undefined
        );
        effectiveSheetNames = selectedNamesFor(latestWorkbook);
        setWorkbook(latestWorkbook);
        setSelectedSheetNames(effectiveSheetNames);
      }

      const latestScope = buildIntentScope(
        latestWorkbook,
        effectiveSheetNames
      );
      if (
        expectedScopeFingerprint &&
        intentScopeFingerprint(latestScope) !== expectedScopeFingerprint
      ) {
        if (clarificationMessageId) {
          setMessages((current) =>
            current.map((message) =>
              message.id === clarificationMessageId &&
              message.clarification
                ? {
                    ...message,
                    clarification: {
                      ...message.clarification,
                      status: "invalidated"
                    }
                  }
                : message
            )
          );
        }
        setPrompt(confirmedPrompt);
        appendMessage({
          role: "system",
          text: "确认期间数据范围发生了变化。本次没有继续分析，请检查上方数据范围后重新发送。"
        });
        return "skipped";
      }

      const selected = new Set(effectiveSheetNames);
      const currentWorksheets = latestWorkbook.worksheets
        .filter((sheet) => selected.has(sheet.name))
        .map((sheet) =>
          dataResults.length > 0
            ? { ...sheet, dataRows: [], truncated: false }
            : sheet
        );
      const activeWorksheetIsSelected = effectiveSheetNames.includes(
        latestWorkbook.activeWorksheet
      );
      /*
       * 桥宿主现场：**执行那一刻**重读，不是规划启动时的快照。
       * 旧实现把 `latestWorkbook` / `effectiveSheetNames` 闭包进来，等价于用一
       * 个静止的现场去"证明"等待期间什么都没变 —— 用户收窄/撤销授权、换簿、
       * 改源值全都漏过去。这里改成真的重读：
       *   - 身份：`currentOfficeWorkbookIdentity()`（读当前 Office 上下文）
       *   - 来源指纹：`captureWorkbookSourceFingerprint`（真读单元格内容）
       *   - 当前授权：`committedSheetNamesRef`（用户此刻已确认的范围）
       *   - 活动表：`captureSelectionContext()`
       * 读不到就如实降级（指纹置 null）—— 复核绝不对"读不到"下"没变化"的结论。
       */
      bridgeLiveContextRef.current = async () => {
        const scopeNames = [...committedSheetNamesRef.current];
        let sourceFingerprint: string | null = null;
        let activeWorksheet = "";
        try {
          activeWorksheet = (await captureSelectionContext()).activeWorksheet;
        } catch {
          activeWorksheet = "";
        }
        try {
          sourceFingerprint = await captureWorkbookSourceFingerprint(scopeNames);
        } catch {
          sourceFingerprint = null;
        }
        return {
          workbookIdentity: bridgeWorkbookIdentity(
            latestWorkbook.name,
            currentOfficeWorkbookIdentity() ?? null
          ),
          sourceFingerprint,
          authorizedSheetNames: scopeNames,
          activeWorksheet
        };
      };
      /*
       * 本轮的**来源基线**（身份 + 来源指纹）取自上面这个读取器本身，不能另起一份。
       *
       * WPS 侧早就是这么接的：先读一次现场，把 name / 身份 / 指纹一起交给桥。
       * Excel 侧原来交的是**快照**上的 `workbookIdentity` / `sourceFingerprint`，
       * 而这两者与复核读到的值根本不在一个口径上：
       *   - 身份：快照上的身份过去没有任何捕获点写入（恒为空），只配到"文件名兜底"，
       *     复核读到的却是 `url:`/路径身份 —— 同一个工作簿自比不自等；
       *   - 指纹：快照指纹由 `sourceFingerprintForSnapshot` 产出（`fnv1a32:`），
       *     复核用的是 `captureWorkbookSourceFingerprint`（`fnv1a32x2:`），
       *     连算法都不同，永远不可能相等。
       * 于是复核必然判"换了一个工作簿/来源已变化"，工具桥在真实工作簿上永远读不到
       * 数据；读不到就没有"本机实际读过哪些表"的可信记账，写侧的断点恢复登记
       * 也就永远不可达。这里把基线固定成"复核将要读的那一份"。
       *
       * 同一个值还要交给规划请求：服务端按 `workbook.workbookIdentity`/
       * `sourceFingerprint` 建立桥绑定与授权归属，也会拿它跟回执里的指纹比
       * （与本地复核是同一判据）。三处必须是**同一个值**，否则链上总有一处
       * 会判来源变化。
       */
      let turnSourceBaseline: SdkBridgeLiveContext | null = null;
      try {
        turnSourceBaseline = await bridgeLiveContextRef.current();
      } catch {
        /*
         * 读不到现场（例如 Office 上下文已失效）：不把整轮规划打掉，也**绝不**
         * 假装"什么都没变"。退回原有的保守路径 —— 规划请求与桥绑定照旧，而
         * 执行前的复核会因为读不到现场而一律拒绝（fail-closed）。
         */
        turnSourceBaseline = null;
      }
      const requestWorkbook: WorkbookSnapshot = {
        ...latestWorkbook,
        // 身份/指纹/覆盖范围一律采信"复核将读到的那一份"，不采信快照（理由见上）。
        ...(turnSourceBaseline
          ? {
              workbookIdentity: turnSourceBaseline.workbookIdentity,
              sourceFingerprint: turnSourceBaseline.sourceFingerprint,
              sourceFingerprintSheets: [...turnSourceBaseline.authorizedSheetNames]
            }
          : {}),
        activeWorksheet: activeWorksheetIsSelected
          ? latestWorkbook.activeWorksheet
          : effectiveSheetNames[0],
        selectedRange: activeWorksheetIsSelected
          ? latestWorkbook.selectedRange
          : null,
        worksheets: currentWorksheets
      };
      const lastResult =
        [...messages]
          .reverse()
          .find((message) => message.resultContext)?.resultContext ?? null;
      // 本地列映射通道（`planningReadPolicy: "none"`）：本端声明"源值不外发"，
      // 规划请求也必须**剥掉数据行** —— 否则服务端手里其实有源值，就轮不到
      // 下发规则这条通道，"本地算"会退化成一个说辞。表头/区域/行列数保留，
      // 服务端据此判越界并产出规则；源值只在 `localMappingPlanFromRule` 里读。
      const plannedWorkbook: WorkbookSnapshot =
        planningReadPolicy === "none"
          ? stripDataRowsForLocalMapping(requestWorkbook)
          : requestWorkbook;
      // This local preparation ends here; remote generation has its own steps.
      // Do not leave an umbrella "generating preview" row running beside them.
      completeCurrentStep("已准备规划请求", {
        detail: "已核对本轮授权范围并整理请求；尚未生成预览或写入工作表。"
      });
      // 任务 ID 必须由本端确定：缺省时服务端会自己生成，客户端就再也对不上桥。
      const planningTurnId = turnId?.trim() || `turn-${messageId()}`;
      const bridgeReady = await sdkBridge.begin({
        taskId: planningTurnId,
        workbookName: plannedWorkbook.name,
        // 与规划请求、与执行前复核读到的**同一份**基线（见上面 `turnSourceBaseline`）；
        // 读不到现场时退回快照值，复核会因为读不到而 fail-closed。
        workbookIdentity:
          turnSourceBaseline?.workbookIdentity ??
          plannedWorkbook.workbookIdentity ??
          null,
        sourceFingerprint:
          turnSourceBaseline?.sourceFingerprint ??
          plannedWorkbook.sourceFingerprint ??
          null,
        signal: turnAbortRef.current?.signal,
        liveContext: () => bridgeLiveContextRef.current(),
        onProgress: ({ scannedRows, totalRows, sheet }) => {
          updateActivityDetail(
            `正在读取「${sheet}」：${scannedRows.toLocaleString()} / ${totalRows.toLocaleString()} 行`
          );
        },
        onError: () => {
          updateActivityDetail("本地查询未完成，已把固定错误码交回模型继续。");
        }
      });
      if (!bridgeReady) {
        /*
         * **桥没建起来就不发规划请求。**
         *
         * `begin` 返回 false = 初始化（读宿主现场）期间这一轮已经被取消/收尾，
         * 或被另一轮取代。以前这里不看返回值：请求照样发出去，模型一要读表就
         * 拿到一条不存在的桥，只能等超时。发了就是一次"注定没人服务"的调用，
         * 不如不发，并且如实告诉用户走到哪一步停的。
         */
        if (planningController && turnAbortRef.current !== planningController) {
          // 已被新一轮接管：界面归新轮次，这里一个字的提示都不写。
          return "skipped";
        }
        if (planningController?.signal.aborted) {
          setPrompt(confirmedPrompt);
          appendMessage({ role: "system", text: "已取消本次规划；没有发送规划请求，原表未改动。" });
          finishActivity("warning", "已取消本次规划，未发送请求", "awaiting_input");
          return "skipped";
        }
        appendMessage({
          role: "system",
          text: "本机读表桥没有建立起来，本次没有发送规划请求；原表未改动，可以重新发送需求。"
        });
        finishActivity("failed", "未能建立本机读表桥，未发送规划请求", "failed");
        return "failed";
      }
      const response = await waitForPlanning(streamAssistantResponse(
        {
          turnId: planningTurnId,
          prompt: confirmedPrompt,
          planningGoal: planningGoal ?? (intentMemory?.toolRequest || dataResults.length > 0 ? "read_only" : undefined),
          taskContext: taskContext ?? createAgentTaskContext(confirmedPrompt),
          workbook: plannedWorkbook,
          lastResult: planningReadPolicy === "none" && dataResults.length === 0 ? null : lastResult,
          images: sentImages.map(({ name, mediaType, data }) => ({
            name,
            mediaType,
            data
          })),
          dataResults,
          planningReadPolicy,
          modelId: effectiveModelId || null,
          // 规划前能力收窄（方案 §4.3）：服务端按可信配置算上界，与本端自报求交。
          ...officeHostRequestFields(isRunningInExcel(), currentOfficeHostVersion())
        },
        {
          signal: turnAbortRef.current?.signal,
          // 规划流里的 bridgeTool 指针交给本机桥：轮询→复核授权→读表→回执。
          onProgress: (event) => {
            if (planningController?.signal.aborted) return;
            applyProgressEvent(event);
            sdkBridge.handleProgressEvent(event);
          }
        }
      ), planningController?.signal);
      planningController?.signal.throwIfAborted();
      markServerOnline();
      if (!(await waitForPlanning(targetGuard.check(), planningController?.signal))) {
        setPrompt(confirmedPrompt);
        finishActivity("completed", "等待确认操作目标", "awaiting_input");
        return "skipped";
      }
      if (response.kind === "planning_pause") {
        const planningActivity = getActivitySnapshot();
        appendMessage({
          role: "assistant",
          text: response.reason,
          intentMemory,
          provider: "system",
          querySourceMode: sourceMode,
          querySourceSheetNames: effectiveSheetNames,
          planningResume: {
            turnId: response.turnId,
            originalPrompt: confirmedPrompt,
            taskContext,
            activityMessageId:
              planningActivity?.persistenceKey ??
              activeActivityMessageIdRef.current ??
              undefined,
            clientLastSequence: Math.min(planningActivity?.sequence ?? 0, response.lastSequence),
            expectedRevision: response.revision,
            lastSequence: response.lastSequence,
            status: "pending",
            reason: response.reason,
            extensionSeconds: response.extensionSeconds,
            activeElapsedSeconds: response.activeElapsedSeconds,
            pauseType: response.pauseType,
            currentModelId: response.currentModelId,
            modelSwitchCount: response.modelSwitchCount,
            canSwitchModel: response.canSwitchModel,
            availableModelConnections: response.availableModelConnections,
            selectedFallbackModelId:
              response.availableModelConnections[0]?.catalogModelId,
            sourceFingerprint:
              response.sourceFingerprint ?? requestWorkbook.sourceFingerprint,
            sourceFingerprintSheets:
              response.sourceFingerprintSheets?.length
                ? response.sourceFingerprintSheets
                : effectiveSheetNames
          }
        });
        finishActivity("warning", "进度已保存，等待你继续", "awaiting_input");
        activityOutcome = "skipped";
        return "skipped";
      }
      if (response.kind === "clarification") {
        // A completed planning invocation has an immutable request fingerprint.
        // The user's answers start a fresh invocation with the full original request.
        await continueIntentDecision({ ...response, turnId: undefined }, confirmedPrompt, sentImages,
          intentScopeFingerprint(latestScope), clarificationMessageId, clarificationRound + 1, 0, planningReadPolicy, taskContext);
        finishActivity("completed", "等待你补充信息；本次没有写入", "awaiting_input");
        activityOutcome = "skipped";
        return activityOutcome;
      }
      if (response.kind === "answer") {
        if (response.outcome === "failed") {
          appendMessage({ role: "system", text: response.message, provider: response.provider });
          finishActivity("failed", "规划已停止，未生成可执行预览", "failed");
          return "failed";
        }
        const querySourceSheetIds = currentWorksheets
          .map((sheet) => sheet.sourceSheetId)
          .filter((value): value is string => Boolean(value));
        const localResultBinding = sdkBridge.binding();
        appendMessage({
          role: "assistant",
          text: response.message,
          resultContext: response.resultContext ?? undefined,
          // SDK 本地结果卡：只带引用与结构，值由本机回环按引用取回后渲染。
          localResult: response.localResult ?? undefined,
          // 会话关闭后服务端按授权归属校验，所以要带上**这一轮**的绑定。
          localResultBinding:
            response.localResult && localResultBinding
              ? {
                  taskId: localResultBinding.taskId,
                  hostId: localResultBinding.hostId,
                  workbookName: localResultBinding.workbookName,
                  workbookIdentity: localResultBinding.workbookIdentity,
                  sourceFingerprint: localResultBinding.sourceFingerprint
                }
              : undefined,
          intentMemory,
          ...(dataResults.length > 0
            ? {
                querySourceMode: sourceMode,
                querySourceSheetNames: effectiveSheetNames,
                querySourceSheetIds
              }
            : {}),
          provider: response.provider
        });
        // 写缓存：只缓存 workbook 只读结论。folder 外部文件不可控（闸4）；
        // 写入分支（else）永不缓存（闸3）。completeness 由本地工具 complete 派生，
        // 无法可靠判定时按 truncated（不复用），安全优先（闸6）。
        const toolRequest = intentMemory?.toolRequest;
        if (
          sourceMode === "workbook" &&
          dataResults.length > 0 &&
          toolRequest &&
          response.resultContext
        ) {
          const intentKey = normalizeIntentKey(
            toolRequest.arguments,
            latestScope
          );
          const completeness = dataResults.every(
            (result) => result.complete === true
          )
            ? "complete"
            : "truncated";
          cacheConclusion(intentKey, {
            resultContext: response.resultContext,
            answerText: response.message,
            sourceSheets: response.resultContext.sourceSheets,
            dataEpochSnapshot: snapshotDataEpochs(effectiveSheetNames),
            completeness,
            querySourceSheetNames: effectiveSheetNames,
            querySourceSheetIds
          });
        }
      } else if (response.kind === "local_plan_rule") {
        if (!(await targetGuard.check())) return "skipped";
        // 服务端已经判定"这条规则还缺参数/缺类别，算不出正确答案"。
        // 这时**不要**读源值、更不要算：算出来的矩阵必然是猜的，
        // 而猜错在单元格层面看不出来。把归因接进澄清卡，让用户补完再跑。
        if (response.rule.clarification) {
          appendMessage({
            role: "assistant",
            text: response.summary,
            intentMemory,
            clarification: { ...columnMappingClarificationMessage({
              clarification: response.rule.clarification,
              scopeLabel: `${response.rule.sourceSheet}!${response.rule.sourceRange}`,
              originalPrompt: confirmedPrompt,
              scopeFingerprint: intentScopeFingerprint(latestScope),
              round: clarificationRound + 1
            }), taskContext },
            provider: response.provider
          });
          activityOutcome = "skipped";
          return activityOutcome;
        }
        // 服务端手上没有源值，只下发了规则。在本端读授权源列、算出计划：
        // 与 WPS 复用同一套确定性实现，源值不外发，写入仍走预览确认。
        let plan: AnalysisPlan;
        try {
          plan = await localMappingPlanFromRule(response, effectiveSheetNames);
        } catch (error) {
          if (recordExcelLocalMappingGap({
            error,
            response,
            originalPrompt: confirmedPrompt,
            authorizedSheets: effectiveSheetNames,
            scopeFingerprint: intentScopeFingerprint(latestScope)
          })) {
            finishActivity("completed", undefined, "awaiting_input");
            activityOutcome = "skipped";
            return activityOutcome;
          }
          throw error;
        }
        plan.sourceFingerprintSheets = [...effectiveSheetNames];
        plan.sourceFingerprint = await captureWorkbookSourceFingerprint(effectiveSheetNames);
        planningTurnByPlanIdRef.current.set(plan.id, planningTurnId);
        appendMessage({
          role: "assistant",
          text: plan.summary,
          plan,
          intentMemory,
          provider: response.provider
        });
      } else {
        const sourceFingerprint =
          sourceMode === "workbook" && isRunningInExcel()
            ? await captureWorkbookSourceFingerprint(effectiveSheetNames)
            : latestWorkbook.sourceFingerprint;
        /*
         * 计划来自服务端（SDK 编译结果），来源依赖**不能采信它自报的**：只认本端
         * 这一轮桥上的确定性读取记录（`recordTrustedToolReads` 在执行器读完那一刻
         * 记账）。没有记录 = 依赖不确定 = 删掉该字段，恢复登记退回既有保守口径。
         *
         * 例外是**由本地结果编译出的计划**（带 `localResultRef`）：它的来源由服务端
         * 按"实际交给执行器的输入句柄"认定，比本端"本轮读过谁"的并集更精确
         * （本轮可能读过 A/B 却只把 A 交给执行器）。两侧取**交集**：既要有服务端
         * 的结果依赖声明，又要有本端真的读过它的凭据。
         */
        const trustedReads = trustedToolReads(planningTurnId);
        const basePlan: AnalysisPlan = {
          ...response.plan,
          sourceFingerprint,
          sourceFingerprintSheets: effectiveSheetNames
        };
        const plan: AnalysisPlan = basePlan.localResultRef
          ? attestResultSourceDependency(
              basePlan,
              trustedReads ? [trustedReads] : []
            )
          : attestSourceDependency(basePlan, trustedReads);
        planningTurnByPlanIdRef.current.set(plan.id, planningTurnId);
        appendMessage({
          role: "assistant",
          text: plan.summary,
          plan,
          intentMemory,
          provider: response.provider
        });
      }
    } catch (reason) {
      const manualPause = isAbortError(reason) && pauseRequestedControllerRef.current === planningController;
      if (isAbortError(reason) && !manualPause) {
        activityOutcome = "aborted";
        setStatus("idle");
        return "aborted";
      }
      if (manualPause) {
        pauseRequestedControllerRef.current = null;
      } else setPendingImages(sentImages);
      if (!manualPause && isLocalServiceConnectionError(reason)) {
        markServerOffline();
        markActivityInterrupted("连接中断，执行结果待确认。");
      } else {
        markServerOnline();
      }
      if ((manualPause || isLocalServiceConnectionError(reason)) && turnId) {
        const planningActivity = getActivitySnapshot();
        const pausedMessage: Omit<ChatMessage, "id" | "createdAt"> = {
          role: "assistant",
          provider: "system",
          text: manualPause ? "已暂停接收进度；需求已保留。先核对原任务，确认检查点后再继续。未写入工作表。"
            : "连接中断，但服务端可能已经保存了规划现场。恢复连接后可在这里核对并继续。",
          intentMemory,
          planningResume: {
            turnId,
            activityMessageId:
              planningActivity?.persistenceKey ??
              activeActivityMessageIdRef.current ??
              undefined,
            clientLastSequence: Math.max(
              0,
              planningActivity?.sequence ?? 0
            ),
            expectedRevision: 1,
            lastSequence: Math.max(0, planningActivity?.sequence ?? 0),
            status: "pending",
            awaitingStatus: true,
            reason: manualPause ? "主动暂停，等待核对任务状态。" : "连接中断，等待恢复后核对任务状态。",
            originalPrompt: confirmedPrompt,
            taskContext,
            extensionSeconds: 0,
            activeElapsedSeconds: 0,
            pauseType: "interrupted",
            currentModelId: effectiveModelId || null,
            sourceFingerprint: workbook.sourceFingerprint,
            sourceFingerprintSheets: selectedNamesFor(workbook)
          }
        };
        const pausedMessageId = appendMessage(pausedMessage);
        finishActivity(
          "warning",
          manualPause ? "已暂停，等待核对检查点" : "连接中断，执行结果待确认。",
          manualPause ? "awaiting_input" : "interrupted"
        );
        activityOutcome = manualPause ? "skipped" : "failed";
        if (manualPause) await resumePlanningTaskFromMessage({ ...pausedMessage, id: pausedMessageId,
          createdAt: new Date().toISOString() }, undefined, true);
      } else if (manualPause) {
        appendMessage({ role: "system", text: "已暂停；尚未形成可续跑检查点，不能从断点继续。原需求保留在对话中，如需处理请重新开始；未修改工作表。" });
        finishActivity("warning", "尚未形成可续跑检查点", "awaiting_input");
        activityOutcome = "skipped";
      } else {
        appendMessage({
          role: "system",
          text:
            reason instanceof Error
              ? reason.message
              : "没有连接到本地 AI 服务，请确认服务已经启动。"
        });
        // 规划没能产出计划：确保活动以失败收尾，绝不落到默认的「已完成」。
        finishActivity("failed", "规划未完成，本次没有写入", "failed");
        activityOutcome = "failed";
      }
    } finally {
      /*
       * 只收自己这一代的桥：这一轮已经被新一轮取代时，桥是**新轮次**的资产，
       * 在这里 `end()` 一次就把它关了（新一轮随后一条读表都发不出去）。
       * 判据与上面清 `turnAbortRef` 是同一个。
       */
      const ownsTurn = turnAbortRef.current === planningController;
      if (ownsTurn) {
        // 本轮结束：停止桥轮询。服务端的结果登记有独立保留期，结果卡仍可取回值。
        sdkBridge.end();
        setStatus("idle");
      }
    }
    return activityOutcome;
  }

  async function executeRequestedDataTool(
    intent: Extract<IntentCheckResponse, { kind: "tool_request" }>,
    sentImages: PendingImage[],
    expectedScopeFingerprint: string,
    clarificationMessageId?: string,
    correctionAttempt = 0,
    taskContext?: AgentTaskContext
  ): Promise<TurnPlanningOutcome> {
    if (!workbook) return "skipped";
    if (!(await targetGuard.checkForAnalysis())) return "skipped";
    if (sourceMode === "folder") {
      if (!folderCatalog) return "skipped";      setStatus("tooling");
      advanceActivity(
        "正在读取文件夹完整数据",
        "pandas 只会读取本次已选择的文件和工作表。"
      );
      try {
        const result = await executeFolderQuery(
          folderCatalog.sessionId,
          intent.request
        );
        return await analyzeConfirmedIntent(
          intent.confirmedPrompt,
          sentImages,
          undefined,
          clarificationMessageId,
          [result],
          workbook,
          {
            confirmedPrompt: intent.confirmedPrompt,
            toolRequest: intent.request
          },
          intent.turnId, "none", 0, taskContext, intent.planningGoal ?? "read_only"
        );
      } catch (reason) {
        appendMessage({
          role: "system",
          text:
            reason instanceof Error
              ? `文件夹数据工具未完成：${reason.message}`
              : "文件夹数据工具未完成"
        });
        setStatus("idle");
        finishActivity("failed", "规划未完成，本次没有写入", "failed");
        return "failed";
      }
    }
    if (sourceMode !== "workbook" || !isRunningInExcel()) {
      return await analyzeConfirmedIntent(
        intent.confirmedPrompt,
        sentImages,
        expectedScopeFingerprint,
        clarificationMessageId,
        [],
        undefined,
        {
          confirmedPrompt: intent.confirmedPrompt,
          toolRequest: intent.request
        },
        intent.turnId,
        sourceMode === "workbook" ? "none" : "ifNeeded",
        0,
        taskContext, intent.planningGoal ?? "read_only"
      );
    }
    setStatus("tooling");
    const diagnosticStartedAt = performance.now();
    const modelCallsBefore = currentModelCallCount();
    const queryArgs = intent.request.arguments;
    const planParts: string[] = [];
    if (queryArgs.groupBy && queryArgs.groupBy.length > 0) {
      planParts.push(`按「${queryArgs.groupBy.join("、")}」分组`);
    }
    if (queryArgs.metrics && queryArgs.metrics.length > 0) {
      planParts.push(
        `聚合 ${queryArgs.metrics.map((metric) => metric.field).join("、")}`
      );
    }
    if (queryArgs.sortBy) {
      planParts.push(
        `按 ${queryArgs.sortBy} ${queryArgs.sortDirection === "asc" ? "升序" : "降序"}`
      );
    }
    if (typeof queryArgs.limit === "number") {
      planParts.push(`取前 ${queryArgs.limit} 条`);
    }
    advanceActivity(
      "正在本地读取并计算",
      `将扫描 ${selectedNamesFor(workbook).length} 张已选工作表；完整数据只在 Excel 本地处理。`,
      "需求已确认，已选择本地数据工具"
    );
    if (planParts.length > 0) {
      advanceActivity(
        "已确定计算方式",
        `本地执行：${planParts.join("，")}。`,
        `计算方式：${planParts.join("，")}`
      );
    }
    let correctionScope: IntentScopeContext | null = null;
    try {
      const selection = await captureSelectionContext();
      const liveWorkbook: WorkbookSnapshot = {
        ...workbook,
        activeWorksheet: selection.activeWorksheet,
        selectedRange: selection.selectedRange
      };
      const effectiveSheetNames = selectedNamesFor(liveWorkbook);
      const liveScope = buildIntentScope(
        liveWorkbook,
        effectiveSheetNames
      );
      correctionScope = liveScope;
      if (
        intentScopeFingerprint(liveScope) !== expectedScopeFingerprint
      ) {
        if (clarificationMessageId) {
          setMessages((current) =>
            current.map((message) =>
              message.id === clarificationMessageId &&
              message.clarification
                ? {
                    ...message,
                    clarification: {
                      ...message.clarification,
                      status: "invalidated"
                    }
                  }
                : message
            )
          );
        }
        setPrompt(intent.confirmedPrompt);
        appendMessage({
          role: "system",
          text: "数据范围已经变化，本地工具没有运行。请确认上方范围后重新发送。"
        });
        finishActivity("warning", "数据范围已变化，本次没有继续", "warning");
        return "skipped";
      }
      /*
       * SDK 适配器：本地**不预置**计算结果。
       *
       * SDK 只拿得到表结构，模型要自己调 query_table 经本机桥在真实宿主上执行
       * 确定性查询。若这里仍把算好的结果塞进 dataResults，服务端的"不上传数据行"
       * 守卫会直接拒绝整条请求，新工具桥就永远拿不到执行权。读表与计算推迟到
       * 桥指针回来那一刻（`sdkBridge.handleProgressEvent`）。
       */
      if (isSdkPlanningModel(effectiveModelId)) {
        setWorkbook(liveWorkbook);
        setSelectedSheetNames(effectiveSheetNames);
        setStatus("idle");
        return await analyzeConfirmedIntent(
          intent.confirmedPrompt,
          sentImages,
          undefined,
          clarificationMessageId,
          [],
          liveWorkbook,
          {
            confirmedPrompt: intent.confirmedPrompt,
            toolRequest: intent.request
          },
          intent.turnId, "none", 0, taskContext, intent.planningGoal ?? "read_only"
        );
      }
      // 二级命中：结构化意图相同且数据未变时，跳过本地全表扫描直接复用结论。
      if (!forceRecomputeRef.current) {
        const intentKey = normalizeIntentKey(
          intent.request.arguments,
          liveScope
        );
        const hit = lookupCachedConclusion(intentKey);
        if (hit) {
          reuseCachedConclusion(hit);
          setStatus("idle");
          return "skipped";
        }
      }
      // 本地取数复用 turn 级 controller：打断整个 turn 时它也会一起停。
      // turnAbortRef 为空（如从缓存直接进入的边缘路径）时兜底自建一个。
      const controller = turnAbortRef.current ?? new AbortController();
      queryAbortRef.current = controller;
      const result = await executeQueryTableTool(
        intent.request,
        effectiveSheetNames,
        selection.activeWorksheet,
        {
          signal: controller.signal,
          onProgress: ({ scannedRows, totalRows, sheet }) => {
            updateActivityDetail(
              `正在读取「${sheet}」：${scannedRows.toLocaleString()} / ${totalRows.toLocaleString()} 行`
            );
          }
        }
      );
      recordDiagnosticEvent({
        timestamp: new Date().toISOString(),
        phase: "local_query",
        durationMs: performance.now() - diagnosticStartedAt,
        scannedRows: result.scannedRows,
        modelCalls: currentModelCallCount() - modelCallsBefore,
        status: "succeeded"
      });
      queryAbortRef.current = null;
      advanceActivity(
        "本地计算完成",
        `已扫描 ${result.scannedRows.toLocaleString()} 行，正在准备紧凑结果。`,
        `扫描 ${result.scannedRows.toLocaleString()} 行，得到 ${result.rows.length.toLocaleString()} 条结果`
      );
      setWorkbook(liveWorkbook);
      setSelectedSheetNames(effectiveSheetNames);
      setStatus("idle");
      return await analyzeConfirmedIntent(
        intent.confirmedPrompt,
        sentImages,
        undefined,
        clarificationMessageId,
        [result],
        liveWorkbook,
        {
          confirmedPrompt: intent.confirmedPrompt,
          toolRequest: intent.request
        },
        intent.turnId, "none", 0, taskContext, intent.planningGoal ?? "read_only"
      );
    } catch (reason) {
      recordDiagnosticEvent({
        timestamp: new Date().toISOString(),
        phase: "local_query",
        durationMs: performance.now() - diagnosticStartedAt,
        modelCalls: currentModelCallCount() - modelCallsBefore,
        status:
          reason instanceof DataToolExecutionError &&
          reason.code === "CANCELLED"
            ? "cancelled"
            : "failed",
        errorCategory: "data_tool"
      });
      queryAbortRef.current = null;
      // 用户主动打断：不做自纠重试，直接静默收尾。
      if (isAbortError(reason)) {
        setStatus("idle");
        return "skipped";
      }
      if (correctionAttempt < 1 && correctionScope) {
        try {
          const failureMessage =
            reason instanceof Error ? reason.message : "本地工具参数无效";
          const structuredFailure =
            reason instanceof DataToolExecutionError ? reason : null;
          if (structuredFailure && !structuredFailure.retryable) {
            throw structuredFailure;
          }
          const corrected = await checkIntent({
            turnId: intent.turnId,
            prompt: intent.confirmedPrompt,
            taskContext,
            scope: correctionScope,
            imageCount: sentImages.length,
            intentConfirmed: true,
            clarificationRound: 1,

            conversation: messages
              .filter(
                (message) =>
                  (message.role === "user" ||
                    message.role === "assistant") &&
                  message.text?.trim()
              )
              .slice(-INTENT_HISTORY_MESSAGES)
              .map((message) => ({
                role: message.role as "user" | "assistant",
                text: message.text!.slice(0, INTENT_MESSAGE_CHARACTERS)
              })),
            priorIntent: {
              confirmedPrompt: intent.confirmedPrompt,
              toolRequest: intent.request
            },
            priorResult: latestResultContext(messages),
            toolFailure: {
              code: structuredFailure?.code ?? "TOOL_EXECUTION_FAILED",
              message: failureMessage,
              retryable: structuredFailure?.retryable ?? true,
              availableFields:
                structuredFailure?.availableFields ?? [],
              request: intent.request
            },
            modelId: effectiveModelId || null
          }, turnAbortRef.current?.signal);
          return await continueIntentDecision(
            corrected,
            intent.confirmedPrompt,
            sentImages,
            expectedScopeFingerprint,
            clarificationMessageId,
            1,
            correctionAttempt + 1,
            sourceMode === "workbook" ? "none" : "ifNeeded",
            taskContext
          );
        } catch {
          // Fall through to the user-facing tool error below.
        }
      }
      setPendingImages(sentImages);
      appendMessage({
        role: "system",
        text:
          reason instanceof Error
            ? `本地数据工具未完成：${reason.message}`
            : "本地数据工具未完成"
      });
      finishActivity("failed", "本地计算未完成，本次没有写入", "failed");
      return "failed";
    } finally {
      setStatus("idle");
    }
  }

  async function continueIntentDecision(
    intent: IntentCheckResponse,
    originalPrompt: string,
    sentImages: PendingImage[],
    scopeFingerprint: string,
    clarificationMessageId?: string,
    clarificationRound = 0,
    correctionAttempt = 0,
    planningReadPolicy: "ifNeeded" | "none" = sourceMode === "workbook" ? "none" : "ifNeeded",
    taskContext: AgentTaskContext = createAgentTaskContext(originalPrompt)
  ): Promise<TurnPlanningOutcome> {
    if (intent.kind === "conversation") {
      appendMessage({ role: "assistant", provider: intent.provider, text: intent.message });
      return "skipped";
    }
    if (intent.kind === "clarification") {
      clarificationImagesRef.current.set(
        intent.clarification.id,
        sentImages
      );
      appendMessage({
        role: "assistant",
        provider: intent.provider,
        clarification: {
          ...intent.clarification,
          turnId: intent.turnId ?? undefined,
          planningReadPolicy: planningReadPolicyForModel(effectiveModelId, planningReadPolicy),
          originalPrompt,
          taskContext,
          scopeFingerprint,
          hadImages: sentImages.length > 0,
          round: clarificationRound,
          status: "pending"
        }
      });
      return "skipped";
    }
    if (intent.kind === "tool_request") {
      return executeRequestedDataTool(
        intent,
        sentImages,
        scopeFingerprint,
        clarificationMessageId,
        correctionAttempt,
        taskContext
      );
    }
    return analyzeConfirmedIntent(
      intent.confirmedPrompt,
      sentImages,
      scopeFingerprint,
      clarificationMessageId,
      [],
      undefined,
      {
        confirmedPrompt: intent.confirmedPrompt
      },
      intent.turnId,
      planningReadPolicy,
      clarificationRound,
      taskContext,
      intent.planningGoal ?? "write_preview"
    );
  }

  async function checkExcelLocalMappingScope(message: ChatMessage) {
    const draft = localMappingDraftsRef.current.get(message.id);
    if (!draft || !workbook || !(await targetGuard.check())) throw new TargetConfirmationRequiredError();
    const fresh = isRunningInExcel() ? await captureWorkbookStructure(draft.authorizedSheets) : workbook;
    const names = selectedNamesFor(fresh);
    const scope = intentScopeFingerprint(buildIntentScope(fresh, names));
    const identity = currentOfficeWorkbookIdentity();
    if (!identity || scope !== draft.scopeFingerprint ||
        JSON.stringify(names) !== JSON.stringify(draft.authorizedSheets) ||
        (draft.workbookIdentity && identity !== draft.workbookIdentity)) {
      await targetGuard.requestConfirmation();
      throw new TargetConfirmationRequiredError();
    }
    if (!draft.workbookIdentity) {
      await targetGuard.requestConfirmation();
      throw new TargetConfirmationRequiredError();
    }
  }

  async function resolveExcelLocalMapping(
    message: ChatMessage,
    rule: ColumnMappingRule,
    reauthorized = false
  ): Promise<void> {
    const draft = localMappingDraftsRef.current.get(message.id);
    if (!draft || busy || !workbook || message.clarification?.status !== "pending") return;
    if (!(await targetGuard.check())) return;
    if (!reauthorized) await checkExcelLocalMappingScope(message);
    let freshWorkbook = workbook;
    if (isRunningInExcel()) {
      try {
        freshWorkbook = await captureWorkbookStructure(draft.authorizedSheets);
      } catch {
        appendMessage({ role: "system", text: "无法读取当前授权范围。已保留补充规则，原表未写入；请检查工作簿后重试。" });
        return;
      }
    }
    const currentSheets = selectedNamesFor(freshWorkbook);
    const currentScope = intentScopeFingerprint(buildIntentScope(freshWorkbook, currentSheets));
    const currentIdentity = currentOfficeWorkbookIdentity();
    if (
      currentScope !== draft.scopeFingerprint ||
      JSON.stringify(currentSheets) !== JSON.stringify(draft.authorizedSheets) ||
      (draft.workbookIdentity && currentIdentity && currentIdentity !== draft.workbookIdentity)
    ) {
      setMessages((current) => current.map((item) =>
        item.id === message.id && item.clarification
          ? { ...item, clarification: { ...item.clarification, status: "invalidated" } }
          : item
      ));
      localMappingDraftsRef.current.delete(message.id);
      setPrompt(message.clarification.originalPrompt);
      appendMessage({
        role: "system",
        text: "无法确认当前仍是原工作簿与授权范围。请重新确认数据范围并发送需求；不会沿用旧源值或旧计划。"
      });
      return;
    }
    if (!currentIdentity || (!draft.workbookIdentity && !reauthorized)) {
      draft.pendingRule = rule;
      // 只缺身份正面证据：留住已提交规则，不销毁草稿，也不假称范围发生变化。
      setMessages((current) => current.map((item) => item.id === message.id ? { ...item } : item));
      appendMessage({ role: "system", text: "补充规则已保留。请确认当前范围后重新读取并生成预览；若仍无法读取工作簿身份，请先保存文件并刷新范围。" });
      return;
    }
    draft.workbookIdentity = currentIdentity;
    draft.pendingRule = undefined;
    setStatus("planning");
    beginActivity("重新读取授权源列", "按你确认的本地规则重新计算；不向模型发送源值。");
    const response: Extract<AssistantResponse, { kind: "local_plan_rule" }> = {
      kind: "local_plan_rule",
      rule,
      ...draft.context
    };
    try {
      let plan: AnalysisPlan;
      try {
        plan = await localMappingPlanFromRule(response, draft.authorizedSheets);
      } catch (error) {
        // A gap is also a result of the async read: do not attach another target's values
        // to this clarification when the workbook changed during that read.
        await checkExcelLocalMappingScope(message);
        if (recordExcelLocalMappingGap({
          error,
          response,
          originalPrompt: message.clarification.originalPrompt,
          authorizedSheets: draft.authorizedSheets,
          scopeFingerprint: draft.scopeFingerprint,
          messageId: message.id
        })) {
          finishActivity("completed", undefined, "awaiting_input");
          return;
        }
        throw error;
      }
      plan.sourceFingerprintSheets = [...draft.authorizedSheets];
      plan.sourceFingerprint = await captureWorkbookSourceFingerprint(draft.authorizedSheets);
      await checkExcelLocalMappingScope(message);
      localMappingDraftsRef.current.delete(message.id);
      setMessages((current) => current.map((item) =>
        item.id === message.id && item.clarification
          ? { ...item, clarification: {
              ...item.clarification, status: "resolved", resolvedLabel: "已补全映射并重新预览"
            } }
          : item
      ));
      appendMessage({ role: "assistant", text: plan.summary, plan, provider: "local" });
      finishActivity("completed", undefined, "awaiting_input");
    } catch (failure) {
      if (failure instanceof TargetConfirmationRequiredError) {
        finishActivity("completed", undefined, "awaiting_input");
        throw failure;
      }
      // Office.js 异常可能含单元格原文；不把异常原文持久化进消息历史。
      appendMessage({ role: "system", text: "重新读取或生成预览失败，原表未写入。请检查授权范围后重试。" });
      finishActivity("failed", "本地映射预览未生成", "failed");
    } finally {
      setStatus("idle");
    }
  }

  async function resolveClarification(
    message: ChatMessage,
    resolution: string,
    label: string,
    finalSelection = false,
    resolutionMode: IntentOption["resolutionMode"] = "append"
  ) {
    const clarification = message.clarification;
    if (!clarification || clarification.status !== "pending" || busy || preparingImages) return;
    if (sourceMode === "folder" && !folderCatalog) {
      setContextOpen(true);
      appendMessage({ role: "system", text: "请先在上方「数据范围」中选择文件夹和工作表。" });
      return;
    }
    if (!(await targetGuard.check())) return;
    if (clarification.localMapping) return; // 只走本地结构化补全，绝不把源值送进模型澄清。
    if (!resolution.trim()) return;
    const rememberedImages =
      clarificationImagesRef.current.get(clarification.id) ?? [];
    const sentImages = [...rememberedImages, ...pendingImages].slice(
      0,
      MAX_IMAGE_ATTACHMENTS
    );
    if (clarification.hadImages && sentImages.length === 0) {
      setImageError("原需求包含图片，请重新添加图片后再确认。");
      return;
    }
    if (sentImages.length && !supportsVision) {
      setImageError("当前模型不支持图片，请切换到支持视觉输入的模型后继续。附件已保留。");
      setModelMenuOpen(true);
      return;
    }
    let taskContext: AgentTaskContext;
    try {
      taskContext = confirmAgentTaskDecision(
        clarification.taskContext ?? createAgentTaskContext(clarification.originalPrompt),
        { questionId: clarification.id, label: clarification.summary.slice(0, capabilities.taskContext.maxLabelChars) || "补充说明", value: resolution },
        resolutionMode
      );
    } catch (error) {
      appendMessage({ role: "system", text: error instanceof Error ? error.message : "请检查补充说明，原表未写入。" });
      return;
    }
    setMessages((current) =>
      current.map((item) =>
        item.id === message.id && item.clarification
          ? {
              ...item,
              clarification: {
                ...item.clarification,
                // Save explicit semantics before the model request: a failed
                // request must not erase the answer or duplicate it on retry.
                taskContext,
                originalPrompt: taskContext.originalPrompt,
                status: "resolving"
              }
            }
          : item
      )
    );
    setPendingImages([]);
    setImageError("");
    const confirmedPrompt = renderConfirmedPrompt(taskContext.originalPrompt, taskContext.confirmations);
    if (!workbook) return;
    const sheetNames = selectedNamesFor(workbook);
    const scope = buildIntentScope(workbook, sheetNames);
    if (
      intentScopeFingerprint(scope) !== clarification.scopeFingerprint
    ) {
      setPrompt(confirmedPrompt);
      setMessages((current) =>
        current.map((item) =>
          item.id === message.id && item.clarification
            ? {
                ...item,
                clarification: {
                  ...item.clarification,
                  status: "invalidated"
                }
              }
            : item
        )
      );
      appendMessage({
        role: "system",
        text: "数据范围已经变化，请检查范围后重新发送需求。"
      });
      setPendingImages(sentImages);
      return;
    }
    setStatus("planning");
    // 澄清后的确认也是一个新 turn，同样建 controller 供打断。
    const turnController = new AbortController();
    turnAbortRef.current = turnController;
    beginActivity(
      "正在处理确认结果",
      `正在核对 ${sheetNames.length} 张已选工作表的字段结构。`
    );
    try {
      const intent: IntentCheckResponse = finalSelection
        ? {
            kind: "proceed",
            summary: `已按用户选择确认：${label}`,
            confirmedPrompt,
            provider: "local",
            turnId: clarification.turnId
          }
        : await checkIntent(
            {
              turnId: clarification.turnId,
              prompt: confirmedPrompt,
              taskContext,
              scope,
              imageCount: sentImages.length,
              intentConfirmed: true,
              clarificationRound: Math.min(
                MAX_CLARIFICATION_ROUNDS,
                clarification.round + 1
              ),
              conversation: messages
                .filter(
                  (item) =>
                    (item.role === "user" || item.role === "assistant") &&
                    item.text?.trim()
                )
                .slice(-INTENT_HISTORY_MESSAGES)
                .map((item) => ({
                  role: item.role as "user" | "assistant",
                  text: item.text!.slice(0, INTENT_MESSAGE_CHARACTERS)
                })),
              priorIntent:
                [...messages]
                  .reverse()
                  .find((item) => item.intentMemory)?.intentMemory ?? null,
              priorResult: latestResultContext(messages),
              modelId: effectiveModelId || null
            },
            turnController.signal
          );
      advanceActivity(
        "需求确认完成",
        "正在根据确认结果选择本地工具或生成操作预览。",
        `已识别 ${scope.sheets.length} 张工作表的字段结构`
      );
      setMessages((current) =>
        current.map((item) =>
          item.id === message.id && item.clarification
            ? {
                ...item,
                clarification: {
                  ...item.clarification,
                  status: "resolved",
                  resolvedLabel: label
                }
              }
            : item
        )
      );
      appendMessage({
        role: "user",
        text: `确认需求：${label}`
      });
      setClarificationDrafts((current) => {
        const next = { ...current };
        delete next[clarification.id];
        return next;
      });
      clarificationImagesRef.current.delete(clarification.id);
      const decision = describeIntentDecision(intent);
      if (intent.kind === "clarification") {
        completeCurrentStep(decision.label, {
          note: decision.note,
          detail: decision.detail
        });
        finishActivity("completed", undefined, "awaiting_input");
        await continueIntentDecision(
          intent,
          confirmedPrompt,
          sentImages,
          clarification.scopeFingerprint,
          message.id,
          Math.min(
            MAX_CLARIFICATION_ROUNDS,
            clarification.round + 1
          ),
          0,
          clarification.planningReadPolicy ?? (finalSelection ? "none" : "ifNeeded"),
          taskContext
        );
      } else {
        advanceActivity(
          decision.label,
          decision.note ?? "正在根据确认结果继续。",
          decision.label,
          { note: decision.note, detail: decision.detail }
        );
        const outcome = await continueIntentDecision(
          intent,
          confirmedPrompt,
          sentImages,
          clarification.scopeFingerprint,
          message.id,
          Math.min(
            MAX_CLARIFICATION_ROUNDS,
            clarification.round + 1
          ),
          0,
          clarification.planningReadPolicy ?? (finalSelection ? "none" : "ifNeeded"),
          taskContext
        );
        // 规划失败时已在内部以失败收尾，这里不再覆盖成「已完成」。
        if (shouldFinishActivityAsCompleted(outcome)) {
          finishActivity();
        }
      }
      setStatus("idle");
    } catch (reason) {
      // 用户主动打断：把澄清状态复位为 pending，静默收尾。
      if (isAbortError(reason)) {
        setMessages((current) =>
          current.map((item) =>
            item.id === message.id && item.clarification
              ? {
                  ...item,
                  clarification: {
                    ...item.clarification,
                    status: "pending"
                  }
                }
              : item
          )
        );
        finishActivity("cancelled", "用户已取消");
        setStatus("idle");
        return;
      }
      setPendingImages(sentImages);
      setMessages((current) =>
        current.map((item) =>
          item.id === message.id && item.clarification
            ? {
                ...item,
                clarification: {
                  ...item.clarification,
                  status: "pending"
                }
              }
            : item
        )
      );
      appendMessage({
        role: "system",
        text:
          reason instanceof Error
            ? reason.message
            : "需求确认失败，本次没有继续处理。"
      });
      setStatus("idle");
      finishActivity();
    } finally {
      if (turnAbortRef.current === turnController) {
        turnAbortRef.current = null;
      }
    }
  }

  function cancelClarification(message: ChatMessage) {
    if (!message.clarification || busy) return;
    localMappingDraftsRef.current.delete(message.id);
    clarificationImagesRef.current.delete(message.clarification.id);
    setClarificationDrafts((current) => {
      const next = { ...current };
      delete next[message.clarification!.id];
      return next;
    });
    setMessages((current) =>
      current.map((item) =>
        item.id === message.id && item.clarification
          ? {
              ...item,
              clarification: {
                ...item.clarification,
                status: "cancelled"
              }
            }
          : item
      )
    );
    appendMessage({ role: "system", text: "已取消这次需求，没有进行分析或写入。" });
  }

  function editClarificationScope(message: ChatMessage) {
    if (!message.clarification || busy) return;
    const rememberedImages =
      clarificationImagesRef.current.get(message.clarification.id) ?? [];
    if (rememberedImages.length > 0) {
      setPendingImages(rememberedImages);
    }
    clarificationImagesRef.current.delete(message.clarification.id);
    setClarificationDrafts((current) => {
      const next = { ...current };
      delete next[message.clarification!.id];
      return next;
    });
    setMessages((current) =>
      current.map((item) =>
        item.id === message.id && item.clarification
          ? {
              ...item,
              clarification: {
                ...item.clarification,
                status: "cancelled"
              }
            }
          : item
      )
    );
    setPrompt(message.clarification.originalPrompt);
    setContextOpen(true);
  }

  /**
   * 暂停规划接收，不自动重跑；保存任务身份后仍须核对服务端检查点。
   *
   * 与 WPS 端同一套语义（三件事分开）：
   *  ① **请求取消**：`abort()` 一调用，本端立刻停止接收 —— 这是本端唯一能自己
   *     确定的事实，与 WPS 端同口径；
   *  ② **服务端暂停**：只有 `/pause` 回 `pauseRequested: true` 才算请求被接受，
   *     是否真的停下仍以任务状态为准；
   *  ③ **界面停止**：活动卡与按钮由各自的状态更新负责。
   *
   * 旧实现没有可中止对象时照样 `markActivityInterrupted` 再发一次暂停请求：
   * 前者会给一张已经收尾的卡扣上"已中断"，后者只能拿到 404。现在先判断有没有
   * 可中止对象，没有就**明确说出来**，不再假装按住了。
   */
  function stopTurn() {
    const controller = turnAbortRef.current ?? queryAbortRef.current;
    const owner = controllerOwnerOf(controller);
    const taskId = getActivitySnapshot()?.operationId;
    recordPauseDiagnostic({
      kind: "pause-click",
      source: "composer",
      executionId: owner ?? undefined,
      turnIdTail: tailOf(taskId),
      controller: !controller ? "none" : controller.signal.aborted ? "aborted" : "owned"
    });
    if (!controller || controller.signal.aborted) {
      recordPauseDiagnostic({
        kind: "pause-unavailable", source: "composer", executionId: owner ?? undefined, reasonCode: "no-controller"
      });
      setPauseRequestState({ phase: "unconfirmed", detail: PAUSE_UNAVAILABLE_REASON });
      return;
    }
    pendingSteerRef.current = null;
    steerBasePromptRef.current = null;
    /*
     * 用户真的按了暂停：这一次「继续」动作到此结束。
     *
     * 只有在这里放掉"本次继续已开的那一段"，才能同时满足：① 同一次继续内部的
     * 多个分支（接回 running → 取回结果）只落**一段**；② 用户再暂停后重新继续，
     * 另开一段而不是继续往旧段里灌。若改在收尾时放掉，① 会断掉。
     */
    resumeSegmentRef.current = null;
    pauseRequestedControllerRef.current = controller;
    const pauseGeneration = ++pauseRequestGenerationRef.current;
    setPauseRequestState({ phase: "requesting", detail: "正在请求暂停…" });
    markActivityInterrupted("已暂停接收进度，等待核对检查点。");
    controller.abort();
    if (queryAbortRef.current && queryAbortRef.current !== controller) queryAbortRef.current.abort();
    if (!taskId) {
      /*
       * 还没有服务端任务身份：不请求。编一个请求只会拿到 404，再把 404 说成
       * "已暂停"就是撒谎。
       */
      setPauseRequestState({
        phase: "local-stopped",
        detail: "已在本端停止接收进度；这次执行还没有提交到后台任务，不需要后台暂停。"
      });
      recordPauseDiagnostic({ kind: "pause-request", source: "composer", executionId: owner ?? undefined, outcome: "skipped" });
      return;
    }
    void (async () => {
      try {
        const ack = await withPauseTimeout(pausePlanningTask(taskId), PAUSE_REQUEST_TIMEOUT_MS, "暂停请求");
        if (pauseGeneration !== pauseRequestGenerationRef.current) return;
        if (ack.pauseRequested) {
          recordPauseDiagnostic({ kind: "pause-ack", source: "composer", executionId: owner ?? undefined, turnIdTail: tailOf(taskId), outcome: "confirmed" });
          setPauseRequestState({
            phase: "acknowledged",
            detail: "本端已停止接收；服务端已接受暂停请求（是否已停下以任务状态为准）。"
          });
          return;
        }
        recordPauseDiagnostic({ kind: "pause-ack", source: "composer", executionId: owner ?? undefined, turnIdTail: tailOf(taskId), outcome: "skipped" });
        setPauseRequestState({
          phase: "local-stopped",
          detail: "本端已停止接收；服务端这条任务此刻不在运行（可能已结束），稍后核对即可看到结果。"
        });
      } catch (error) {
        if (pauseGeneration !== pauseRequestGenerationRef.current) return;
        const timedOut = error instanceof PauseTimeoutError;
        recordPauseDiagnostic({
          kind: timedOut ? "pause-timeout" : "pause-failed",
          source: "composer",
          executionId: owner ?? undefined,
          turnIdTail: tailOf(taskId),
          outcome: "unconfirmed"
        });
        setPauseRequestState({
          phase: "unconfirmed",
          detail: timedOut
            ? `本端已停止接收，但暂停请求在 ${Math.round(PAUSE_REQUEST_TIMEOUT_MS / 1000)} 秒内没有确认；后台可能仍在运行。`
            : "本端已停止接收，但未能确认后台暂停；请稍后核对原任务状态。"
        });
      }
    })();
  }

  // 硬转向：把新话暂存，掐掉当前 turn。等 busy 回落，下面的 effect 会带这句话重跑。
  // 新话不清空输入框内容？——这里选择先清空，重跑走 overrideText，体验和正常发送一致。
  function steerTurn(text: string) {
    const trimmed = text.trim();
    if (!trimmed) {
      stopTurn();
      return;
    }
    pendingSteerRef.current = trimmed;
    // 记下被打断那一句在飞的意图，重跑时并入新话，避免承接式追问丢失前一句约束。
    steerBasePromptRef.current = busy ? rawPromptRef.current || null : null;
    setPrompt("");
    turnAbortRef.current?.abort();
    queryAbortRef.current?.abort();
  }

  // 当一个 turn 被打断、status 归 idle 后，若有待转向的话，就带着它重跑一个新 turn。
  useEffect(() => {
    if (busy) return;
    const steerText = pendingSteerRef.current;
    if (!steerText) return;
    pendingSteerRef.current = null;
    void sendMessage({ overrideText: steerText });
    // sendMessage / busy 是稳定引用与派生状态，这里仅在 busy 变化时检查一次。
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [busy]);

  // /function 阶段一：不立即生成，先弹「确定写入单元格」卡（预填智能建议）。
  // 光标点哪都无所谓——目标由用户在卡里确认后，才作 activeCell 生成+试算。
  async function handleFunctionCommand(description: string) {
    if (!(await targetGuard.check())) return;
    setPrompt("");
    appendMessage({ role: "user", text: `/function ${description}` });

    if (!isRunningInExcel()) {
      appendMessage({
        role: "system",
        text: "当前是浏览器演示模式，/function 需在 Excel 中打开任务窗格后使用。"
      });
      return;
    }

    setStatus("planning");
    // 同 confirmFunctionTarget：建 activity 让计时器起步，否则 planning 卡显示"0 秒"。
    beginActivity("正在准备公式生成", "正在读取当前选区与表结构，建议写入位置。");
    try {
      const selection = await captureSelectionContext();
      const sheetName = selection.activeWorksheet;
      // 智能建议：数据区右侧第一空列首行。取不到就退到当前选区首格 / A2。
      const snapshot = await captureWorkbook([sheetName]);
      const activeSheet = snapshot.worksheets.find(
        (sheet) => sheet.name === sheetName
      );
      const suggested = suggestWriteTarget(activeSheet?.usedRange ?? null);
      // 跨表匹配确定性提案：仅凭表头结构 + 描述关键词本地判断，不调模型。
      // 快照读取失败不阻断流程：退回模型生成路径。
      let match: CrossTableMatchProposal | null = null;
      try {
        const extraSheets = await resolveFormulaExtraSheets();
        match = buildCrossTableProposal(
          description,
          columnsFromRange(
            activeSheet?.headers ?? [],
            activeSheet?.usedRange ?? null
          ),
          extraSheets
        );
      } catch {
        match = null;
      }

      appendMessage({
        role: "assistant",
        functionPreview: {
          phase: "target",
          description,
          sheet: sheetName,
          writeTarget: suggested,
          mode: match ? "deterministic" : "model",
          match: match ?? undefined,
          version: "compat",
          modernFormula: "",
          modernExplanation: "",
          modernResult: "",
          compatFormula: "",
          compatExplanation: "",
          compatResult: ""
        }
      });
    } catch (reason) {
      appendMessage({
        role: "system",
        text: reason instanceof Error ? reason.message : "准备公式生成失败"
      });
    } finally {
      finishActivity();
      setStatus("idle");
    }
  }

  // /function 多工作簿：从 folder 勾选解析每张外部表的文件信息与样本上下文。
  // 优先复用 workbook 里已缓存的 folder snapshot（confirmSheetSelection 存的，
  // sheet 带 sourceFile/sourceFileId）；勾选集未被缓存完整覆盖时现取。
  async function resolveFormulaExtraSheets(): Promise<FormulaExtraSheet[]> {
    if (
      sourceMode !== "folder" ||
      !folderCatalog ||
      folderSheetKeys.length === 0
    ) {
      return [];
    }
    const selectedKeys = new Set(folderSheetKeys);
    const snapshotKeys = new Set(
      workbook?.worksheets
        .filter((sheet) => sheet.sourceFileId || sheet.sourceFile)
        .map((sheet) => folderSheetKey(sheet.sourceFileId ?? "", sheet.sourceSheet ?? sheet.name)) ??
        []
    );
    const snapshot =
      workbook !== null &&
      snapshotKeys.size > 0 &&
      [...selectedKeys].every((key) => snapshotKeys.has(key))
        ? workbook
        : await withFolderSessionRecovery(folderCatalog.sessionId, (sessionId) =>
            createFolderSnapshot(sessionId, folderSelections())
          );
    const extraSheets: FormulaExtraSheet[] = [];
    for (const sheet of snapshot.worksheets) {
      const fileId = sheet.sourceFileId ?? "";
      const originalSheetName = sheet.sourceSheet ?? sheet.name;
      if (!selectedKeys.has(folderSheetKey(fileId, originalSheetName))) continue;
      const sourcePath = sheet.sourceFile ?? "";
      const sourceFile = sourcePath.split(/[\\/]/).pop() ?? sourcePath;
      if (!sourceFile) continue;
      const startCol = startColumnOfRange(sheet.usedRange);
      const startIndex = startCol ? columnIndexFromLetter(startCol) : 0;
      const headers = sheet.headers.map((cell) => String(cell ?? ""));
      const columns = headers.map((_, index) =>
        columnLetterFromIndex(startIndex + index)
      );
      extraSheets.push({
        sourceFile,
        sourcePath,
        sheetName: originalSheetName,
        headers,
        columns,
        sampleRows: sheet.dataRows
          .slice(0, 5)
          .map((row) => row.map((cell) => String(cell ?? ""))),
        rowCount: sheet.rowCount
      });
    }
    return extraSheets;
  }

  // /function 阶段二：用户确认写入单元格后，以目标首格作 activeCell 生成两版公式，
  // 就地试算，翻到 preview 阶段。生成/试算/写入三者同锚，R1C1 错位不再发生。
  // 优先本地确定性路径（mode=deterministic），否则走模型兜底（mode=model）。
  async function confirmFunctionTarget(
    message: ChatMessage,
    options?: { forceModel?: boolean }
  ) {
    const preview = message.functionPreview;
    if (!preview || preview.phase !== "target" || busy) return;
    if (!(await targetGuard.check())) return;

    const resolved = resolveWriteTarget(preview);
    if (!resolved.address) {
      markFunctionPreview(message.id, { targetError: resolved.error });
      return;
    }
    const firstCell = firstCellOfRange(resolved.address);
    if (!firstCell) {
      markFunctionPreview(message.id, {
        targetError: `「${resolved.address}」不是有效的写入位置。`
      });
      return;
    }

    const useModel =
      options?.forceModel === true ||
      preview.mode !== "deterministic" ||
      !preview.match;

    setStatus("planning");
    // 建 activity 让实时计时器起步：不建的话计时 effect 因 activity 为空提前返回，
    // 生成公式的多秒模型调用期间会一直显示"0 秒"。
    beginActivity(
      useModel ? "正在生成公式" : "正在生成本地公式",
      useModel
        ? `目标 ${resolved.address}，正在结合表结构生成两版公式。`
        : `目标 ${resolved.address}，正在按跨表匹配参数拼公式并试算。`
    );
    try {
      const sheetName = preview.sheet;
      // 计时起点：从此刻(调模型)到两版公式试算完成，作为"生成耗时"展示。
      const generateStart = performance.now();

      let generated: GeneratedFormulaPair;
      let externalFiles: string[];
      if (useModel) {
        // 模型兜底：抓活动表列头/样本行喂模型；扫描"字典/映射"表注入其内容。
        const snapshot = await captureWorkbook([sheetName]);
        const activeSheet = snapshot.worksheets.find(
          (sheet) => sheet.name === sheetName
        );
        const headers =
          activeSheet?.headers.map((cell) => String(cell ?? "")) ?? [];
        // 列头对应的列字母：从 usedRange 起始列往右枚举，喂给模型建立"列头→列"映射，
        // 避免模型猜错列（如把输出列当输入列导致循环引用）。
        const columns = columnsFromRange(
          activeSheet?.headers ?? [],
          activeSheet?.usedRange ?? null
        ).map((column) => column.letter);
        const sampleRows = (activeSheet?.dataRows ?? [])
          .slice(0, 5)
          .map((row) => row.map((cell) => String(cell ?? "")));
        const dictionary = await loadDictionaryForFormula(sheetName);
        const extraSheets = await resolveFormulaExtraSheets();
        externalFiles = extraSheets.map((sheet) => sheet.sourceFile);

        // 后端优先用「/function 公式模型」（模型配置里单独选的那个），忽略 modelId。
        // modelId 仅作回退：若公式模型设为「跟随全局」，才用全局选择。这让 /function
        // 与聊天窗口全局选择脱钩，避免有人切推理模型时公式生成被 reasoning 卡住。
        generated = await generateFormula({
          description: preview.description,
          activeCell: firstCell,
          headers,
          columns,
          sampleRows,
          dictionary,
          extraSheets,
          modelId: effectiveModelId || null
        });
      } else if (preview.match) {
        // 本地确定性路径：不调模型、不发数据，直接拼两版公式。
        externalFiles = [preview.match.externalFile];
        generated = buildCrossTableFormulas(preview.match, firstCell);
      } else {
        markFunctionPreview(message.id, {
          targetError: "缺少跨表匹配参数，请重试或换 AI 生成。"
        });
        return;
      }

      const allowedExternal = new Set(externalFiles);

      // 两版公式各在目标首格真实试算（试算失败不阻断预览，让用户自行判断）。
      const trialCalc = async (formula: string): Promise<string> => {
        if (!formula) return "";
        try {
          const trial = await previewFormulaFirstCell(
            sheetName,
            firstCell,
            formula,
            allowedExternal
          );
          return trial.sampleResult;
        } catch (previewError) {
          const detail =
            previewError instanceof Error ? previewError.message : "";
          return detail
            ? `（试算失败：${detail}）`
            : "（试算失败，请检查公式）";
        }
      };

      const modernResult = await trialCalc(generated.modernFormula);
      const compatResult = await trialCalc(generated.compatFormula);

      const generateMs = Math.round(performance.now() - generateStart);

      markFunctionPreview(message.id, {
        phase: "preview",
        writeTarget: resolved.address,
        targetError: undefined,
        pickingTarget: false,
        version: "compat",
        modernFormula: generated.modernFormula,
        modernExplanation: generated.modernExplanation,
        modernResult,
        externalFiles,
        compatFormula: generated.compatFormula,
        compatExplanation: generated.compatExplanation,
        compatResult,
        generateMs
      });
    } catch (reason) {
      const errorMsg = reason instanceof Error ? reason.message : "生成公式失败";
      // 会话失效已在上面尝试过重建；能走到这里说明重建也失败了，状态已被清空。
      markFunctionPreview(message.id, {
        targetError: isFolderSessionLost(errorMsg)
          ? "文件夹会话已失效，请重新选择文件夹"
          : errorMsg
      });
    } finally {
      finishActivity();
      setStatus("idle");
    }
  }

  // 用户觉得预填的跨表匹配不对：切到模型兜底，用原描述直接生成。
  async function switchToModelFunction(message: ChatMessage) {
    if (!message.functionPreview || busy) return;

    // 清除之前的验证消息：避免重新生成时出现重复的"已执行 N 步并通过独立验证"
    const planId = `function-${message.id}`;
    setMessages((current) =>
      current.filter(
        (msg) => msg.executedPlanId !== planId
      )
    );

    markFunctionPreview(message.id, {
      phase: "target",
      mode: "model",
      match: undefined,
      targetError: undefined,
      applied: false,
      cancelled: false
    });
    await confirmFunctionTarget(message, { forceModel: true });
  }

  async function sendMessage(options?: {
    forceRecompute?: boolean;
    overrideText?: string;
    overrideImages?: PendingImage[];
    restartingAnalysis?: boolean;
    activityMessageId?: string;
    messageUnderstanding?: Awaited<ReturnType<typeof understandMessage>>;
    /**
     * 这次进来是「操作目标确认之后」的回调重入（`retrySendRef`），不是用户又发了一遍。
     * 那句话**已经**在对话里、输入框也**已经**清空，不能再送一遍。
     * 与 WPS 的 `targetConfirmedRetry` 同义（`WpsAgentPanel.tsx`）。
     */
    targetConfirmedRetry?: boolean;
  }) {
    const inputImages = options?.overrideImages ?? pendingImages;
    const originalTarget = targetGuard.binding;
    const originalMode = `${sourceMode}:${committedScopeMode}`;
    let completedUnderstanding = options?.messageUnderstanding;
    const rememberEarlyPause = (originalPrompt: string) => {
      const messageId = getActivitySnapshot()?.persistenceKey ?? appendMessage({ role: "assistant", text: "已暂停" });
      prePlanningPause.remember({ messageId, conversationId: chatHistory.activeConversationId,
        target: originalTarget, mode: originalMode, data: { prompt: originalPrompt, images: [...inputImages],
          activityMessageId: messageId, messageUnderstanding: completedUnderstanding } });
    };
    const enteredText = (options?.overrideText ?? prompt).trim();
    const text =
      enteredText ||
      (inputImages.length > 0
        ? "请结合附件图片分析当前工作簿，并说明发现的问题。"
        : "");
    if (!text || text.length < 2 || busy || preparingImages) return;
    if (inputImages.length && !supportsVision) {
      setImageError("当前模型不支持图片，请切换到支持视觉输入的模型后发送。附件已保留。");
      setModelMenuOpen(true);
      return;
    }
    // 续聊前的旧记录闸。这条历史在读取时就被脱敏过（真凭据不会进入界面、内存
    // 或请求），同时被登记进 `legacyCredentialConversationIds`。判据是那个
    // **持久标记**，不是"文本里还能不能匹配到凭据" —— 后者在脱敏那一刻就永远
    // 不再成立，靠它做闸门会在重开之后自动松开（这正是上一版的缺陷）。
    // 只查当前会话，别的历史不该挡住这次提问。
    if (staleCredentials) {
      appendMessage({
        role: "system",
        text:
          "这段历史里的凭据是在旧版本写进去的，已被拦下，不会随本次提问发送。" +
          "请点上方「新对话」开始一段干净的对话，或先删除这段历史。"
      });
      setPrompt(enteredText);
      return;
    }
    // /function 短链：绕开 planner（checkIntent/streamAssistantResponse），
    // 命中式单发生成原生公式 + 首格真实试算 + 预览卡。
    if (enteredText.startsWith("/function ") && inputImages.length === 0) {
      const description = enteredText.slice("/function ".length).trim();
      if (description) {
        if (!(await targetGuard.checkForAnalysis(() => retrySendRef.current(text)))) return;
        await handleFunctionCommand(description);
        return;
      }
    }
    // —— 受理即送达 ——
    //
    // 拿到这句话就把它送进对话、并清空输入框，**不等「理解消息」跑完**。
    // 「理解消息」是一次模型往返，后面还夹着操作目标核对与授权表结构读取，可能要等几秒。
    // 这期间输入框里还留着这句话会有两个后果：
    // ① 运行中的发送键停在「打断并补充」（↑）而不是「暂停」（■）——用户再按一次
    //    Enter 就会拿着**同一句**去转向重跑；
    // ② 对话里看不到自己刚发的话，看起来像"还没发出去"。
    //
    // 两处例外不能在这里提前送：
    // - **回答澄清卡**（`resolveClarification`）：那句回答由澄清卡承载，不由这里另起气泡；
    // - **目标确认后的回调重入**（`targetConfirmedRetry`）：上一次调用已经送过了。
    //
    // 下面两个 `find` 只读 `messages` 与 `enteredText`，提前到这里取值不会变
    // ——后面「取消本地映射澄清 / 按补充回答继续」两段逻辑仍留在原位，不搬家。
    const pendingLocalMapping = [...messages]
      .reverse()
      .find((message) => message.clarification?.status === "pending" && message.clarification.localMapping);
    const pendingClarification = [...messages]
      .reverse()
      .find(
        (message) =>
          message.clarification?.status === "pending" && !message.clarification.localMapping
      );
    const answeringClarification = Boolean(enteredText) && Boolean(pendingClarification);
    if (!answeringClarification && options?.targetConfirmedRetry !== true && !options?.restartingAnalysis) {
      appendMessage({
        role: "user",
        text: enteredText || "请结合附件图片分析当前工作簿。",
        attachmentNames: inputImages.map((image) => image.name)
      });
      setPrompt("");
    }
    // 转向重跑：把被打断那句的意图并进这句，让 checkIntent 收到完整承接式需求。
    // 仅影响送模型与缓存的意图文本（intentText），气泡仍只显示用户实际输入（text）。
    const steerBasePrompt = steerBasePromptRef.current;
    steerBasePromptRef.current = null;
    const standaloneStructuralIntent =
      sourceMode === "workbook"
        ? parseBulkWorksheetDeletion(text)
        : null;
    const intentText =
      !standaloneStructuralIntent && steerBasePrompt && steerBasePrompt !== text
        ? `${steerBasePrompt}；补充：${text}`
        : text;
    // Stage 1 is message-only. Do not build scope, read headers or send prior results here.
    forceRecomputeRef.current = options?.forceRecompute === true;
    rawPromptRef.current = intentText;
    const messageController = new AbortController();
    turnAbortRef.current = messageController;
    setStatus("planning");
    const retainedActivity = options?.activityMessageId && messages.find((item) => item.id === options.activityMessageId);
    if (retainedActivity && retainedActivity.activityLog) {
      /*
       * 早期分析（理解消息 / 读表结构 / 字段意图）的继续：**另开一个新回复段**，
       * 不再把进度接回旧卡。旧卡保留它暂停那一刻的内容与耗时，并给出
       * "查看后续进展"；它自己也就不再是"还能继续"的那一段。
       */
      resumeEarlyAnalysisActivity(retainedActivity);
    } else beginActivity("理解消息", "先判断是否需要工作簿上下文；尚未读取字段或数据。", { stepId: "message-intent" });
    let understanding: Awaited<ReturnType<typeof understandMessage>>;
    let needsContext = false;
    const earlyTasks = prePlanningPause.pendingEntries().filter((entry) => entry.conversationId === chatHistory.activeConversationId);
    // The original request has already been selected and verified. This dispatch
    // analyses it; it must not select a paused task for a second time.
    const recoveryTasks = options?.restartingAnalysis ? [] : planningRecoveryCandidates(messages.map((message) =>
      message.planningResume?.status === "pending" ? message.planningResume.turnId
        : earlyTasks.find((entry) => entry.messageId === message.id)?.id));
    try {
      understanding = options?.messageUnderstanding ?? await waitForPlanning(understandMessage({
        prompt: intentText,
        ...(recoveryTasks.length ? { recoveryTasks } : {}),
        imageCount: inputImages.length,
        conversation: messages
          .filter((message) => (message.role === "user" || message.role === "assistant") && message.text?.trim())
          .slice(-INTENT_HISTORY_MESSAGES)
          .map((message) => ({ role: message.role as "user" | "assistant", text: message.text!.slice(0, INTENT_MESSAGE_CHARACTERS) })),
        modelId: effectiveModelId || null
      }, messageController.signal), messageController.signal);
      if (messageController.signal.aborted) throw new DOMException("Aborted", "AbortError");
      markServerOnline();
      if (understanding.kind === "conversation") {
        // 普通发送/转向在受理时就把用户气泡送进对话了（见上面「受理即送达」），
        // 这里只补**回答澄清卡**那一种：它的气泡不由那条链路产生。
        if (answeringClarification) appendMessage({ role: "user", text });
        appendMessage({ role: "assistant", text: understanding.message, provider: understanding.provider });
        setPrompt("");
        setPendingImages([]);
        setImageError("");
        finishActivity("completed", "已理解并直接回复；未读取工作表", "completed", { defaultCollapsed: true });
        return;
      }
      if (understanding.kind === "task_recovery") {
        const turnId = understanding.turnId;
        if (!inputImages.length && isOfferedPlanningRecovery(turnId, recoveryTasks) && earlyTasks.some((entry) => entry.id === turnId)) {
          discardRecoveryRoutingActivity();
          if (understanding.action === "continue") await restartEarlyAnalysis(turnId);
          else appendMessage({ role: "assistant", text: "原需求已暂停，尚未写入。可以继续分析。" });
          return;
        }
        const original = messages.find((message) =>
          message.planningResume?.turnId === turnId && message.planningResume.status === "pending");
        if (inputImages.length || !isOfferedPlanningRecovery(turnId, recoveryTasks) || !original) {
          throw new Error("无法确定要恢复的原任务；请在原任务卡片上核对状态。");
        }
        discardRecoveryRoutingActivity();
        await resumePlanningTaskFromMessage(original, undefined, understanding.action !== "continue", understanding.action === "continue");
        return;
      }
      needsContext = true;
      completedUnderstanding = understanding;
      if (!options?.messageUnderstanding) completeCurrentStep("已理解消息：需要工作簿上下文");
    } catch (reason) {
      if (isAbortError(reason) && pauseRequestedControllerRef.current === messageController) {
        pauseRequestedControllerRef.current = null;
        setPrompt((current) => current.trim() === text ? "" : current);
        // 同上：只有回答澄清卡时气泡还没送出；普通发送的气泡已经在受理那一刻落盘。
        if (answeringClarification) {
          appendMessage({ role: "user", text, attachmentNames: inputImages.map((image) => image.name) });
        }
        rememberEarlyPause(intentText);
        finishActivity("warning", "需求分析已暂停，原需求已保留", "awaiting_input");
      } else if (isAbortError(reason)) finishActivity("cancelled", "用户已取消");
      else {
        setPrompt(text);
        if (isLocalServiceConnectionError(reason)) markServerOffline();
        appendMessage({ role: "system", text: reason instanceof Error ? reason.message : "消息理解失败，尚未读取工作表。" });
        finishActivity("failed", "消息理解未完成；未读取工作表");
      }
      return;
    } finally {
      if (!needsContext) {
        if (turnAbortRef.current === messageController) {
          turnAbortRef.current = null;
          setStatus("idle");
        }
      }
    }
    const pauseBeforeContext = () => {
      if (turnAbortRef.current === messageController) turnAbortRef.current = null;
      setStatus("idle");
    };
    const finishEarlyCancellation = () => {
      if (pauseRequestedControllerRef.current === messageController) {
        pauseRequestedControllerRef.current = null;
        rememberEarlyPause(intentText);
        finishActivity("warning", "需求分析已暂停，原需求已保留", "awaiting_input");
      } else finishActivity("cancelled", "用户已取消");
      pauseBeforeContext();
    };
    /**
     * 「停下来让用户先调整、这句话还没被消费」时把它放回输入框。
     *
     * 受理那一刻已经清空了输入框（见「受理即送达」），所以这些**在读取字段之前**
     * 就交还控制权的出口必须把它还回去 —— 否则用户白打一遍。这也解释了为什么
     * 结构化删除那几条出口本来就写着 `setPrompt(enteredText)`：口径一致。
     *
     * 只填**空的**输入框：停在这里期间用户可能已经打了新草稿，不能覆盖掉。
     * （同款判据见上面暂停出口的 `current.trim() === text ? "" : current`。）
     */
    const keepRequestInComposer = () => {
      if (!enteredText) return;
      setPrompt((current) => (current.trim() ? current : enteredText));
    };
    if (!workbook) {
      setContextOpen(true);
      keepRequestInComposer();
      finishActivity("completed", "请先选择工作簿", "awaiting_input");
      pauseBeforeContext();
      return;
    }
    /*
     * 刷新/打开弹层不能把草稿变成权限，**换口径也算改了范围**。
     *
     * 判据是"草稿 ≠ 已提交"（表名或口径任一不同），不只看手动模式：
     * 发送那一刻的**读取**是按草稿口径走的（跟随当前 ⇒ 读活动表），若放行，
     * 读到的可能是已提交授权之外的表。停在这里不读一个字段。
     */
    if (sourceMode === "workbook" && scopeDraftDirty) {
      setContextOpen(true);
      setScopeConfirmationError("正在核对勾选的工作表；需求仍保留，核对完成后可继续。");
      keepRequestInComposer();
      finishActivity("completed", "等待确认数据范围", "awaiting_input");
      pauseBeforeContext();
      return;
    }
    if (!(await targetGuard.checkForAnalysis(() => retrySendRef.current(intentText, understanding)))) {
      keepRequestInComposer();
      finishActivity("completed", "等待确认操作目标", "awaiting_input");
      pauseBeforeContext();
      return;
    }
    if (messageController.signal.aborted) {
      finishEarlyCancellation();
      return;
    }
    // "仍要重新计算"绕过整个命中判定，并在重算后覆写缓存。
    forceRecomputeRef.current = options?.forceRecompute === true;
    rawPromptRef.current = intentText;
    if (inputImages.length > 0 && !supportsVision) {
      setImageError("当前模型不支持图片，请切换模型或移除附件。");
      keepRequestInComposer();
      return;
    }

    // `pendingLocalMapping` / `pendingClarification` 已在上面「受理即送达」处算好
    // （只依赖 `messages` 与 `enteredText`，中间没有变更），这里直接用，不重算。
    if (pendingLocalMapping && enteredText) {
      // 自由文本是用户发起的新请求；不能误走通用澄清，把本地来源值带入模型。
      localMappingDraftsRef.current.delete(pendingLocalMapping.id);
      setMessages((current) => current.map((message) =>
        message.id === pendingLocalMapping.id && message.clarification
          ? { ...message, clarification: { ...message.clarification, status: "cancelled" } }
          : message
      ));
    }
    if (pendingClarification && enteredText) {
      finishActivity("completed", "按补充回答继续");
      pauseBeforeContext();
      await resolveClarification(
        pendingClarification,
        enteredText,
        enteredText
      );
      return;
    }

    if (
      sourceMode === "workbook" &&
      workbookScopeMode === "manual" &&
      selectedSheetNames.length === 0
    ) {
      setContextOpen(true);
      appendMessage({
        role: "system",
        text: "请至少选择一个工作表，或改为跟随当前工作表。"
      });
      keepRequestInComposer();
      finishActivity("completed", "等待选择工作表", "awaiting_input");
      pauseBeforeContext();
      return;
    }
    if (sourceMode === "folder" && (!folderCatalog || !selectionConfirmed || folderSheetKeys.length === 0)) {
      setContextOpen(true);
      appendMessage({
        role: "system",
        text: "请先在上方「数据范围」中选择文件夹和工作表。"
      });
      keepRequestInComposer();
      finishActivity("completed", "等待选择文件和工作表", "awaiting_input");
      pauseBeforeContext();
      return;
    }
    /*
     * 这里**没有**"发送即确认"：草稿不是授权。
     *
     * 旧实现在这里把 `selectionConfirmed` 置真，而当时"已确认 ⇒ 把草稿抄成已提交"
     * 是靠一个被动 effect 生效的 —— 于是"直接发一句、没点确认"就悄悄改了已提交范围，
     * 守卫下一轮读到新范围、基线还是旧的，弹「工作表范围发生变化」。
     * 现在只有确认动作（`commitAuthorizedScope`）能写已提交；手动口径下草稿未确认
     * 的请求更早就在上面被 `scopeDraftDirty` 拦下并保留了输入。
     */

    let intentWorkbook = workbook;
    if (sourceMode === "workbook" && isRunningInExcel()) {
      setStatus("planning");
      advanceActivityStage(
        "read-schemas",
        "读取已授权表结构",
        `消息需要工作簿上下文；正在本地检查 ${selectedNamesFor(workbook).length} 张工作表，不上传数据行。`
      );
      try {
        intentWorkbook = await waitForPlanning(captureWorkbookStructure(
          workbookScopeMode === "manual" ? selectedSheetNames : undefined
        ), messageController.signal);
        setWorkbook(intentWorkbook);
      } catch (reason) {
        if (messageController.signal.aborted) {
          finishEarlyCancellation();
          return;
        }
        appendMessage({
          role: "system",
          text:
            reason instanceof Error
              ? `读取工作表结构失败：${reason.message}`
              : "读取工作表结构失败"
        });
        // 本地读表失败，需求一个字都没消费掉，放回输入框让用户可以直接重发。
        keepRequestInComposer();
        setStatus("idle");
        pauseBeforeContext();
        finishActivity();
        return;
      }
    }
    const effectiveSheetNames = selectedNamesFor(intentWorkbook);
    if (messageController.signal.aborted) {
      finishEarlyCancellation();
      return;
    }
    setSelectedSheetNames(effectiveSheetNames);
    const structuralDeletion =
      sourceMode === "workbook"
        ? evaluateBulkWorksheetDeletion(
            intentText,
            intentWorkbook.worksheets
              .filter((sheet) => !isEBSystemSheet(sheet.name))
              .map((sheet) => sheet.name),
            effectiveSheetNames
          )
        : { kind: "notMatched" as const };
    if (structuralDeletion.kind === "keeperMissing") {
      // 用户气泡已在「受理即送达」时落盘；这里只补结果说明。
      appendMessage({
        role: "system",
        text:
          `当前工作簿中未找到要保留的工作表"${structuralDeletion.requestedKeeperName}"。` +
          "本次没有生成计划，也没有删除任何工作表。"
      });
      restorePromptUnlessEdited(enteredText);
      setStatus("idle");
      finishActivity();
      pauseBeforeContext();
      return;
    }
    if (structuralDeletion.kind === "selectionRequired") {
      const missingPreview = structuralDeletion.missingSheetNames
        .slice(0, 6)
        .map((name) => `"${name}"`)
        .join("、");
      const missingSuffix =
        structuralDeletion.missingSheetNames.length > 6
          ? `等 ${structuralDeletion.missingSheetNames.length} 张工作表`
          : "";
      appendMessage({
        role: "system",
        text:
          "当前授权范围不包含全部工作表，因此不能删除尚未授权的工作表。" +
          `请点击输入框上方的「工作表范围」，勾选当前工作簿的全部工作表，然后再次发送这条指令。请同时勾选需要保留的「${structuralDeletion.keeperName}」。` +
          `\n\n尚未授权：${missingPreview}${missingSuffix}。` +
          "\n\n本次未生成执行计划，也未删除任何工作表。",
        scopeSelectionGuidance: {
          keeperName: structuralDeletion.keeperName,
          missingSheetNames: structuralDeletion.missingSheetNames
        }
      });
      chooseManualScope(intentWorkbook);
      setSelectionConfirmed(false);
      restorePromptUnlessEdited(enteredText);
      setStatus("idle");
      finishActivity();
      pauseBeforeContext();
      return;
    }
    if (structuralDeletion.kind === "nothingToDelete") {
      appendMessage({
        role: "assistant",
        text: `当前工作簿只剩"${structuralDeletion.keeperName}"，没有其他工作表需要删除。`
      });
      setPrompt("");
      setPendingImages([]);
      setImageError("");
      setStatus("idle");
      finishActivity();
      pauseBeforeContext();
      return;
    }
    const intentScope = buildIntentScope(
      intentWorkbook,
      effectiveSheetNames
    );
    const scopeFingerprint = intentScopeFingerprint(intentScope);
    const sentImages = inputImages;
    animatePet("turtle-encourage", 1000);
    setPrompt("");
    setPendingImages([]);
    setImageError("");
    // 一级 prompt 缓存：原样再问且无图片、非强制重算时，跳过 checkIntent 直接复用。
    // folder 模式外部文件监听器无从得知变化，永不复用。
    if (
      !forceRecomputeRef.current &&
      sourceMode === "workbook" &&
      sentImages.length === 0
    ) {
      const intentKey = promptKeyCacheRef.current.get(
        normalizePrompt(intentText)
      );
      const hit = intentKey
        ? lookupCachedConclusion(intentKey)
        : undefined;
      if (hit) {
        reuseCachedConclusion(hit);
        finishActivity("completed", "复用已核对的查询结论");
        setStatus("idle");
        pauseBeforeContext();
        return;
      }
    }
    setStatus("planning");
    // 开启一个新的 turn：建 controller 存进 ref，串到本 turn 所有模型请求上。
    const turnController = messageController;
    turnAbortRef.current = turnController;
    advanceActivityStage(
      "intent",
      "结合字段理解任务",
      "正在结合字段结构、最近对话和上一轮紧凑结果判断是否需要本地工具。"
    );
    try {
      const intent = await waitForPlanning(checkIntent({
        prompt: intentText,
        taskContext: createAgentTaskContext(intentText),
        scope: intentScope,
        imageCount: sentImages.length,
        conversation: messages
          .filter(
            (message) =>
              (message.role === "user" || message.role === "assistant") &&
              message.text?.trim()
          )
          .slice(-INTENT_HISTORY_MESSAGES)
          .map((message) => ({
            role: message.role as "user" | "assistant",
            text: message.text!.slice(0, INTENT_MESSAGE_CHARACTERS)
          })),
        priorIntent:
          [...messages]
            .reverse()
            .find((message) => message.intentMemory)?.intentMemory ?? null,
        priorResult: latestResultContext(messages),
        modelId: effectiveModelId || null
      }, turnController.signal), turnController.signal);
      turnController.signal.throwIfAborted();
      markServerOnline();
      const decision = describeIntentDecision(intent);
      if (!(await targetGuard.check())) {
        restorePromptUnlessEdited(text);
        setPendingImages(sentImages);
        setStatus("idle");
        finishActivity("completed", "等待确认操作目标", "awaiting_input");
        return;
      }
      if (intent.kind === "clarification") {
        completeCurrentStep(decision.label, {
          note: decision.note,
          detail: decision.detail
        });
        finishActivity("completed", undefined, "awaiting_input");
        await continueIntentDecision(
          intent,
          intentText,
          sentImages,
          scopeFingerprint
        );
      } else {
        advanceActivity(
          decision.label,
          decision.note ?? "正在根据确认结果继续。",
          decision.label,
          { note: decision.note, detail: decision.detail }
        );
        const outcome = await continueIntentDecision(
          intent,
          intentText,
          sentImages,
          scopeFingerprint
        );
        // 规划没产出计划（含本地列映射读取/计算失败）时已经以失败收尾，
        // 这里不能再补一次默认的 finishActivity()，否则会把失败标成「已完成」。
        if (shouldFinishActivityAsCompleted(outcome)) {
          finishActivity();
        }
      }
      setStatus("idle");
    } catch (reason) {
      // 主动暂停与转向取消分开；没有任务身份就不能编造检查点。
      if (isAbortError(reason)) {
        if (pauseRequestedControllerRef.current === turnController) {
          pauseRequestedControllerRef.current = null;
          rememberEarlyPause(intentText);
          finishActivity("warning", "需求分析已暂停，原需求已保留", "awaiting_input");
        } else finishActivity("cancelled", "用户已取消");
        setStatus("idle");
        return;
      }
      setPendingImages(sentImages);
      if (sentImages.length) restorePromptUnlessEdited(text);
      if (isLocalServiceConnectionError(reason)) {
        markServerOffline();
        markActivityInterrupted("连接中断，执行结果待确认。");
      } else {
        markServerOnline();
      }
      appendMessage({
        role: "system",
        text:
          reason instanceof Error
            ? reason.message
            : "需求确认失败，本次没有继续分析或执行。"
      });
      setStatus("idle");
      finishActivity();
    } finally {
      // 只清自己这个 turn 的 controller；若已被新 turn 覆盖则不动。
      if (turnAbortRef.current === turnController) {
        turnAbortRef.current = null;
      }
    }
  }

  const composerPausedMessage = [...messages].reverse().find((message) =>
    message.planningResume?.status === "pending" ||
    prePlanningPause.entries.some((entry) => entry.messageId === message.id &&
      entry.conversationId === chatHistory.activeConversationId));
  /*
   * 三种"继续"进**同一个恢复入口**，但入口不同层：点按钮 / 空输入点发送直接
   * 走暂停卡（本函数）；在输入框里打「继续」「接着做」仍交给意图理解（模型判
   * 语义）后落到同一条续跑路径，不在本地再写一套关键词分类。
   */
  const composerCanContinue = Boolean(composerPausedMessage);
  /*
   * 输入框右侧按钮的判据来自**共享层**（`pauseButtonState`，WPS 端读同一份）。
   *
   * 三件事显式分开，不再由 `busy` 一个变量推：
   *  - `taskRunning`：真的在途（模型请求 / 取数 / 生成对比 / 确认收尾）；
   *  - `controllerAlive`：存在未 abort 的句柄 —— **只有**它能让 ■ 成立；
   *  - `hasInput`：有字时是"打断并补充"（↑），不是暂停。
   *
   * 旧写法 `busy ? ■ : ↑` 的后果实测过：只是"结果在等你看"（`phase === 'reviewing'`）
   * 也会画出一个 ■，点下去 `stopTurn` 找不到任何可中止对象 —— 点了没反应。
   */
  const composerPauseButton = pauseButtonState({
    hasInput: false,
    paused: composerCanContinue,
    taskRunning: running,
    controllerAlive: Boolean(
      (turnAbortRef.current && !turnAbortRef.current.signal.aborted) ||
      (queryAbortRef.current && !queryAbortRef.current.signal.aborted)
    )
  });
  /** 过程卡上的「停止」：同一条判据，但**不受输入影响**（打了字也要能停）。 */
  const cardPauseButton = executionPauseState({
    taskRunning: running,
    controllerAlive: Boolean(
      (turnAbortRef.current && !turnAbortRef.current.signal.aborted) ||
      (queryAbortRef.current && !queryAbortRef.current.signal.aborted)
    )
  });

  async function submitComposer() {
    if (busy || preparingImages || (!prompt.trim() && pendingImages.length === 0)) return;
    return sendMessage();
  }

  async function resumeComposerTask() {
    if (busy || preparingImages || !composerPausedMessage || pauseRequestState?.phase === "requesting") return;
    const early = prePlanningPause.pendingEntries().find((entry) =>
      entry.messageId === composerPausedMessage.id && entry.conversationId === chatHistory.activeConversationId);
    if (early) await restartEarlyAnalysis(early.id);
    else await resumePlanningTaskFromMessage(composerPausedMessage, undefined, false, true);
  }

  async function restartEarlyAnalysis(id: string) {
    const issued = turnAbortRef.current;
    const entry = prePlanningPause.pendingEntries().find((item) => item.id === id);
    if (busy || preparingImages) return;
    if (entry?.data.images.length && !supportsVision) {
      setImageError("原需求包含图片，请切换到支持视觉输入的模型后继续。");
      setModelMenuOpen(true);
      return;
    }
    await prePlanningPause.restart(id, chatHistory.activeConversationId, `${sourceMode}:${committedScopeMode}`,
      async () => {
        if (sourceMode !== "workbook" || !(await targetGuard.check())) return null;
        return inspectOfficeAuthorizedTarget(committedScopeMode === "auto" ? null : committedSheetNames);
      },
      async (saved) => {
        await sendMessage({ overrideText: saved.prompt, overrideImages: saved.images,
          messageUnderstanding: saved.messageUnderstanding, activityMessageId: saved.activityMessageId, restartingAnalysis: true });
      }, () => turnAbortRef.current === issued);
  }

  /**
   * **直接执行**路径的执行前复核凭据。
   *
   * 与 `nativeReview` 互斥：直接执行没有副本可比，所以它自己的"复核"不能是
   * `assertExcelNativeComparisonUnchanged`（那会去找一个不存在的副本）。
   * 它复核的是同样重要的三件事 —— 同一计划版本、同一授权范围、同一本工作簿。
   *
   * 存在的意义是让"直接执行没跳过执行前复核"这句话有**代码位置**：
   * 没有它，直接执行就会成为唯一一条没有 `beforeFirstWrite` 的写入路径。
   */
  type DirectExecutionRecheck = {
    plan: AnalysisPlan;
    authorization: readonly string[];
    workbookIdentity: string;
    /**
     * 发起这次执行的请求消息。删表检查点用它把结果卡画回**这一条**，
     * 而不是追加到聊天底部。
     */
    requestMessageId?: string;
  };

  async function runPlan(
    plan: AnalysisPlan,
    allowedExternalFiles?: ReadonlySet<string>,
    forceLocal?: boolean,
    /** 用户在「数据已变」二次确认里点了「仍要执行」。 */
    acknowledgeStaleSource?: boolean,
    nativeReview?: ExcelNativeComparison,
    /** 直接执行路径的执行前复核凭据；与 `nativeReview` 互斥。 */
    directRecheck?: DirectExecutionRecheck,
    onResultReview?: (result: PlanExecutionResult) => void
  ): Promise<boolean> {
    const deletionState = () => JSON.stringify({ identity: resultLive.current.workbookIdentity,
      authorization: resultLive.current.authorization, conversation: resultLive.current.conversationId,
      sourceMode: resultLive.current.sourceMode });
    const deletionBinding = isWorksheetDeletion(plan) ? deletionState() : null;
    if (nativeReview && directRecheck) {
      appendMessage({ role: "system", text: "一次执行只能有一条复核来源；原表未写入。" });
      return false;
    }
    if (sourceMode === "workbook" && !(await targetGuard.check())) return false;
    if (nativeReview && sourceMode !== "workbook") {
      appendMessage({ role: "system", text: "数据来源模式已变化，不能把原生审阅用于其他数据来源。" });
      return false;
    }
    /*
     * 直接执行的执行前复核与审阅路径**同样在首写之前**：授权勾选、计划版本与
     * 原工作簿身份三项任一变化都当场停下，不写。放在这里而不是执行器内部，
     * 是为了让两条路径的"写之前"位置一目了然。
     */
    const directRecheckGuard = directRecheck
      ? async () => {
        if (currentOfficeWorkbookIdentity() !== directRecheck.workbookIdentity) {
          throw new Error("原工作簿身份已变化，本次直接执行已停止；原表未写入。");
        }
        const currentNames = workbook ? selectedNamesFor(workbook) : [];
        if (JSON.stringify(currentNames) !== JSON.stringify([...directRecheck.authorization])) {
          throw new Error("授权工作表范围在确认后发生变化，本次直接执行已停止；原表未写入。");
        }
      }
      : null;
    if (sourceMode === "folder" && !forceLocal) {
      if (!folderCatalog) return false;
      const operationId = `execution-${messageId()}`;
      setActivePlanId(plan.id);
      beginActivity("正在执行文件夹操作计划", plan.summary, {
        operationId,
        phase: "preflight",
        stepId: "execution.preflight"
      });
      const pendingFolderOperationBase = {
        schemaVersion: 1 as const,
        operationId,
        folderPath: folderCatalog.folderPath,
        conversationId:
          activeConversation?.id ?? chatHistory.activeConversationId,
        activityMessageId:
          activeActivityMessageIdRef.current ?? `activity-${operationId}`,
        plan,
        createdAt: new Date().toISOString()
      };
      savePendingFolderOperation({
        ...pendingFolderOperationBase,
        sessionId: folderCatalog.sessionId
      });
      setStatus("executing");
      let folderActivityStatus: "completed" | "warning" | "failed" =
        "completed";
      let folderActivitySummary: string | undefined;
      try {
        const result = await withFolderSessionRecovery(
          folderCatalog.sessionId,
          (sessionId) => {
            savePendingFolderOperation({
              ...pendingFolderOperationBase,
              sessionId
            });
            return executeFolderPlan(sessionId, plan, {
              operationId,
              onProgress: applyProgressEvent,
              onConnectionState: (connectionState, message) => {
                if (connectionState === "reconnecting") {
                  markActivityReconnecting(
                    message ?? "连接短暂中断，正在恢复进度。"
                  );
                } else if (connectionState === "interrupted") {
                  markActivityInterrupted(
                    message ?? "连接中断，执行结果待确认。"
                  );
                } else {
                  markActivityConnected();
                }
              }
            });
          }
        );
        recordDiagnosticEvent({
          timestamp: new Date().toISOString(),
          phase: "execution",
          durationMs: result.executionMs,
          modelCalls: 0,
          status: "succeeded"
        });
        recordDiagnosticEvent({
          timestamp: new Date().toISOString(),
          phase: "verification",
          durationMs: result.verificationMs,
          modelCalls: 0,
          status:
            result.verification.status === "failed" ? "failed" : "succeeded"
        });
        appendMessage({
          role: "assistant",
          text: folderExecutionOutcomeText(plan, result, false),
          verification: result.verification,
          executedPlanId: result.undo ? plan.id : undefined,
          executedUndoId: result.undo?.undoId,
          folderOperationId: operationId
        });
        pushFolderUndo(
          result.undo
            ? {
                planId: plan.id,
                undoId: result.undo.undoId,
                sessionId: result.undo.sessionId
              }
            : null
        );
        if (result.verification.status === "verified") {
          markPlanVerified(plan.id);
        }
        clearPendingFolderOperation(operationId);
        return true;
      } catch (reason) {
        folderActivitySummary =
          reason instanceof Error ? reason.message : "执行文件夹计划失败";
        if (isLocalServiceConnectionError(reason)) {
          folderActivityStatus = "warning";
          markActivityInterrupted("连接中断，执行结果待确认。");
        } else {
          folderActivityStatus = "failed";
        }
        appendMessage({
          role: "system",
          text: reason instanceof Error ? reason.message : "执行文件夹计划失败"
        });
        return false;
      } finally {
        setActivePlanId(null);
        setStatus("idle");
        finishActivity(folderActivityStatus, folderActivitySummary);
      }
    }

    if (!isRunningInExcel()) {
      appendMessage({
        role: "system",
        text: "当前是浏览器演示模式，不会写入文件。请在 Excel 中打开任务窗格。"
      });
      return false;
    }

    if (plan.actions.some(action => action.type === "deleteWorksheet") && !isWorksheetDeletion(plan)) {
      appendMessage({ role: "system", text: "删除工作表需要完整检查点，请将删表与其他修改拆分执行；本次未写入。" });
      return false;
    }
    if (isWorksheetDeletion(plan)) {
      /*
       * 归属在**发起这一刻**固定：会话、工作簿与请求消息三者一起交给检查点。
       * 之后用户切走会话，回执也只更新这条请求下的结果卡，不会被写到新会话，
       * 更不会在聊天底部另起一张"最新"卡。
       */
      const deletionConversationId = chatHistory.activeConversationId;
      const requestMessageId = directRecheck?.requestMessageId ?? null;
      setStatus("executing");
      try {
        await directRecheckGuard?.();
        if (deletionBinding !== deletionState()) throw new Error("授权、计划或会话在检查期间变化，未删除。");
        const authorized = selectedNamesFor(workbook!);
        const issue = worksheetDeletionIssue(plan, authorized, workbook!.worksheets.map(sheet => sheet.name));
        if (issue) throw new Error(issue);
        const identity = currentOfficeWorkbookIdentity();
        if (!identity) throw new Error("无法确认原工作簿身份，未删除。");
        /*
         * 成功之后固定走这一条收尾：重读目录 → 按回执核对 → 只收窄授权与目标基线。
         * 直接执行与「允许仅本地备份并继续」两条路径共用它，否则补授权那条会在
         * 下一轮被守卫读成"范围被改了"，又弹一次"尚未写入"。
         */
        const finishDeletion = async () => {
          const snapshot = await scan({ announce: false });
          await adoptVerifiedDeletionShrink({
            identity,
            authorized,
            snapshot,
            requestMessageId,
            conversationId: deletionConversationId
          });
        };
        const succeeded = await worksheetCheckpoint.execute(`${deletionConversationId}:${plan.id}`, {
          workbookPath: identity, authorizedSheets: authorized, plan
        }, {
          conversationId: deletionConversationId,
          workbookIdentity: identity,
          requestMessageId,
          beforeResume: async () => {
            if (!(await targetGuard.check())) throw new Error("目标检查未通过，未继续删除。");
            await directRecheckGuard?.();
            if (deletionBinding !== deletionState() || identity !== currentOfficeWorkbookIdentity()) throw new Error("授权、工作簿或会话变化，请重新发起；未继续删除。");
          },
          onCompleted: finishDeletion
        });
        if (succeeded) await finishDeletion();
        return succeeded;
      } catch (reason) {
        const text = reason instanceof Error ? reason.message : "未开始删除工作表。";
        // 已经建卡就写在卡里（对应请求的位置）；否则写回**发起时的会话**，不是当前会话。
        if (!worksheetCheckpoint.reportIssue(requestMessageId, deletionConversationId, text)) {
          appendSystemMessageToConversation(deletionConversationId, text);
        }
        return false;
      } finally { setStatus("idle"); }
    }

    const operationId = `execution-${messageId()}`;
    const authorizedSheetNames = forceLocal
      ? [workbook!.activeWorksheet]
      : selectedNamesFor(workbook!);
    const recoverySupport = officeRecoverySupport(plan);
    const workbookIdentity = currentOfficeWorkbookIdentity();
    let officeState: Awaited<ReturnType<typeof registerOfficeExecution>> | null = null;
    let recoveryLimitation: string | null = null;
    /*
     * 登记成功时的**恢复范围**说明。不写"支持断点恢复"这种笼统话：说清是哪一级
     * 断点（动作级）、什么时候还能续（关闭或重载任务窗格后）、续跑前做什么
     * （先核对真实单元格状态），以及范围（只续跑本次登记的动作）。
     */
    let recoveryNotice: string | null = null;
    setActivePlanId(plan.id);
    beginActivity("正在执行操作计划", plan.summary, {
      operationId,
      phase: "preflight",
      stepId: "execution.preflight"
    });
    setStatus("executing");
    let workbookActivityStatus: "completed" | "warning" | "failed" =
      "completed";
    let workbookActivitySummary: string | undefined;
    try {
      if (nativeReview) await Excel.run(context => assertExcelNativeComparisonUnchanged(context, nativeReview, authorizedSheetNames, plan));
      if (recoverySupport.supported) {
        if (!workbookIdentity) {
          recoveryLimitation =
            "当前工作簿尚未保存，本次仍按普通模式执行，但不能在关闭任务窗格后自动续跑。";
        } else {
          officeState = await registerOfficeExecution({
            operationId,
            plan,
            authorizedSheetNames,
            allowedExternalFiles: [...(allowedExternalFiles ?? [])],
            workbookIdentity,
            // 只有本端确实知道这份计划由哪个规划任务产出时才声称关联：
            // 服务端据此拒绝"这个任务已经写过原表"的继续分析。
            taskId: planningTurnByPlanIdRef.current.get(plan.id)
          });
          if (!officeState.ownerToken) {
            throw new Error("服务端没有授予本次 Office 写入权，已停止执行。");
          }
          const pendingSaved = savePendingOfficeExecution({
            schemaVersion: 1,
            operationId,
            conversationId: chatHistory.activeConversationId,
            activityMessageId:
              activeActivityMessageIdRef.current ?? `activity-${operationId}`,
            workbookIdentity,
            createdAt: new Date().toISOString()
          });
          if (!pendingSaved) {
            throw new Error("无法保存 Office 任务挂接索引，已停止执行以避免丢失恢复入口。");
          }
          recoveryNotice =
            "已登记动作级断点：关闭或重载任务窗格后可从已完成动作之后继续；续跑前先核对真实单元格状态，只续跑本次登记的动作。";
        }
      } else {
        recoveryLimitation = `${recoverySupport.reason ?? "该计划包含复杂对象操作"}；本次不承诺跨重启自动续跑。`;
      }
      bindRecoveryProbe((import.meta.env.DEV || import.meta.env.VITE_ENABLE_RECOVERY_PROBE === "true"), currentOfficeWorkbookIdentity(), operationId, plan.actions, Boolean(officeState?.ownerToken));
      const result = await executePlan(plan, {
        /*
       * 直接执行**也**挂 `beforeFirstWrite`：这就是"只跳过副本比对"的落点。
       * 少了这一行，两条路径的差别就不止"有没有副本"，还会多出"有没有执行前复核"。
       */
      beforeFirstWrite: nativeReview
        ? context => assertExcelNativeComparisonUnchanged(context, nativeReview, authorizedSheetNames, plan)
        : directRecheckGuard ?? undefined,
        authorizedSheetNames,
        allowedExternalFiles,
        acknowledgeStaleSource,
        operationId,
        onProgress: applyProgressEvent,
        onChangeReview: (reviews) => {
          if (reviews.length > 0) {
            changeReviewsRef.current.set(operationId, reviews);
          }
        },
        ...(officeState
          ? {
              recovery: {
                startActionIndex: officeState.nextActionIndex,
                undoSnapshot: officeState.checkpoint,
                onBeforeAction: async ({ actionIndex, beforeSnapshot }) => {
                  if (!officeState?.ownerToken) {
                    throw new Error("Office 写入权已经失效");
                  }
                  officeState = await startOfficeAction(operationId, {
                    expectedRevision: officeState.revision,
                    ownerToken: officeState.ownerToken,
                    actionIndex,
                    beforeSnapshot
                  });
                },
                onActionCompleted: async ({ actionIndex, checkpoint }) => {
                  if (!officeState?.ownerToken) {
                    throw new Error("Office 写入权已经失效");
                  }
                  interruptRecoveryProbe((import.meta.env.DEV || import.meta.env.VITE_ENABLE_RECOVERY_PROBE === "true"), operationId, actionIndex);
                  officeState = await completeOfficeAction(operationId, {
                    expectedRevision: officeState.revision,
                    ownerToken: officeState.ownerToken,
                    actionIndex,
                    checkpoint
                  });
                }
              }
            }
          : {})
      });
      let nativeResultNotice = "";
      if (directRecheck) {
        // 直接执行没有副本可比：如实说"未与副本比对"，而不是让这句话缺席。
        // 动作回读与恢复点仍然照常，措辞不许让人以为跳过了验证。
        nativeResultNotice = "本次直接执行未生成副本，因此没有与副本比对；上述验证为执行侧逐格回读与恢复点核对。";
      }
      if (nativeReview?.mode === "manual") {
        nativeResultNotice = "兼容副本未自动回读；上述验证仅针对原表执行结果，不代表已核对原表与副本的计算值一致。";
      } else if (nativeReview) {
        const comparison = await compareExcelNativeCommittedResult(nativeReview);
        nativeResultNotice = comparison === "matched" ? "原表目标回读与审阅副本一致。" : comparison === "different" ? "原表已写入，但目标计算值或公式与审阅副本不同，请核对执行结果与恢复点。" : "原表已写入，但无法独立核对其结果与审阅副本是否一致，请核对执行结果与恢复点。";
        if (comparison !== "matched") {
          result.verification = { ...result.verification, passed: false,
            status: comparison === "different" || result.verification.status === "failed" ? "failed" : "executed_unverified",
            unverifiedActions: [...result.verification.unverifiedActions, ...plan.actions.map((action, index) => ({ index, type: action.type, sheet: action.sheet, message: nativeResultNotice }))]
          };
          const reviews = changeReviewsRef.current.get(operationId);
          if (reviews) changeReviewsRef.current.set(operationId, reviews.map(review => ({ ...review, status: comparison === "different" ? "readback_mismatch" : "readback_failed" })));
        }
      }
      if (officeState) {
        if (!officeState.ownerToken) throw new Error("Office 写入权已经失效");
        officeState = await finishOfficeExecution(operationId, {
          expectedRevision: officeState.revision,
          ownerToken: officeState.ownerToken,
          result
        });
        clearPendingOfficeExecution(operationId);
      }
      recordDiagnosticEvent({
        timestamp: new Date().toISOString(),
        phase: "execution",
        durationMs: result.executionMs,
        modelCalls: 0,
        status: "succeeded"
      });
      recordDiagnosticEvent({
        timestamp: new Date().toISOString(),
        phase: "verification",
        durationMs: result.verificationMs,
        modelCalls: 0,
        status:
          result.verification.status === "failed" ? "failed" : "succeeded"
      });
      appendMessage({
        role: "assistant",
        text: `「${plan.title}」已执行 ${result.actionResults.length} 步${verificationSummary(result.verification)}。${
          result.undoSnapshot
            ? result.undoSnapshot.postStateCoverage === "complete"
              ? `已记录 ${result.undoSnapshot.ranges.length} 项本次执行撤销数据，并已核对执行后状态。`
              : "已保留本次执行捕获的数据，但执行后状态未能完整核对，暂不能直接撤销。"
            : ""
        }${recoveryLimitation ?? recoveryNotice ?? ""}${nativeResultNotice}原始工作表没有被删除或清空。`,
        verification: result.verification,
        executedPlanId: plan.id,
        executedExecutionId: result.undoSnapshot?.executionId
      });
      pushUndoSnapshot(result.undoSnapshot ?? null);
      onResultReview?.(result);
      if (result.verification.status === "verified" && !onResultReview) {
        markPlanVerified(plan.id);
      }
      await scan();
      return true;
    } catch (reason) {
      if (officeState?.ownerToken && officeState.status === "running") {
        try {
          officeState = await suspendOfficeExecution(operationId, {
            expectedRevision: officeState.revision,
            ownerToken: officeState.ownerToken,
            code: "OFFICE_EXECUTION_INTERRUPTED",
            message: reason instanceof Error ? reason.message : String(reason)
          });
          if (!officeState.requiresReconciliation && officeState.nextActionIndex === 0) {
            clearPendingOfficeExecution(operationId);
          }
        } catch {
          // 动作完成回执可能已提交但响应丢失；保留 operationId，重载后查询权威状态。
        }
      }
      if (reason instanceof StaleSourceError) {
        workbookActivityStatus = "warning";
        workbookActivitySummary = reason.message;
        if (nativeReview) {
          appendMessage({ role: "system", text: `${reason.message}原生审阅已失效，请关闭副本并重新生成预览；不能跳过并排审阅直接写入旧结果。` });
          return false;
        }
        if (directRecheck) {
          // 直接执行同样**不豁免**来源复核：数据变了要么重新生成计划，要么走上面的二次确认，
          // 但绝不因为"没生成副本"就当这份计划可以直接写。
          appendMessage({ role: "system", text: `${reason.message}本次直接执行已停止；请重新生成计划，或确认数据已变后再次执行。原表未写入。` });
          return false;
        }
        // 数据变了但计划本身仍可执行：给出二次确认，而不是像字段缺失那样堵死。
        appendMessage({
          role: "system",
          text: reason.message,
          plan,
          staleSourceConfirm: {
            plan,
            bakedActionCount: reason.bakedActionCount,
            status: "pending"
          }
        });
        return false;
      }
      if (reason instanceof PlanExecutionError) {
        workbookActivityStatus = "failed";
        workbookActivitySummary = reason.message;
        const succeeded = reason.actionResults.filter(
          (result) => result.status === "succeeded"
        ).length;
        const failed = reason.actionResults.find(
          (result) => result.status === "failed"
        );
        const notRun = reason.actionResults.filter(
          (result) => result.status === "not_run"
        ).length;
        pushUndoSnapshot(reason.undoSnapshot);
        appendMessage({
          role: "system",
          text: `${reason.message}。${
            succeeded > 0
              ? `已有 ${succeeded} 步成功写入当前工作簿；`
              : "本次没有步骤成功写入；"
          }${
            reason.undoSnapshot
              ? `已保留本次执行捕获的数据（${reason.undoSnapshot.coverage === "partial" ? "部分覆盖" : "完整覆盖"}，${reason.undoSnapshot.postStateCoverage === "complete" ? "可在撤销前再次核对" : "执行后状态未完整核对，暂不能直接撤销"}）；`
              : "没有可用的撤销快照；"
          }${failed ? `失败步骤：${failed.type}（${failed.sheet}）；` : ""}${
            notRun > 0 ? `其余 ${notRun} 步未执行。` : ""
          }`,
          executedPlanId: reason.undoSnapshot ? plan.id : undefined,
          executedExecutionId: reason.undoSnapshot?.executionId,
          officeExecutionId: officeState ? operationId : undefined,
          officeResume:
            officeState &&
            (officeState.requiresReconciliation || officeState.nextActionIndex > 0)
              ? {
                  operationId,
                  expectedRevision: officeState.revision,
                  status: "pending",
                  requiresReconciliation: officeState.requiresReconciliation,
                  nextActionIndex: officeState.nextActionIndex,
                  totalActions: plan.actions.length,
                  reason:
                    officeState.error?.message ??
                    "执行现场已保存；继续前会先核对真实单元格状态。"
                }
              : undefined
        });
        return false;
      }
      workbookActivityStatus = isLocalServiceConnectionError(reason)
        ? "warning"
        : "failed";
      workbookActivitySummary =
        reason instanceof Error ? reason.message : "执行计划失败";
      if (workbookActivityStatus === "warning") {
        markActivityInterrupted("连接中断，执行结果待确认。");
      }
      appendMessage({
        role: "system",
        text: reason instanceof Error ? reason.message : "执行计划失败",
        officeExecutionId: officeState ? operationId : undefined,
        officeResume:
          officeState &&
          (officeState.requiresReconciliation || officeState.nextActionIndex > 0)
            ? {
                operationId,
                expectedRevision: officeState.revision,
                status: "pending",
                requiresReconciliation: officeState.requiresReconciliation,
                nextActionIndex: officeState.nextActionIndex,
                totalActions: plan.actions.length,
                reason:
                  officeState.error?.message ??
                  "写入回执尚未完整确认，继续前会先核对真实单元格状态。"
              }
            : undefined
      });
      return false;
    } finally {
      setActivePlanId(null);
      setStatus("idle");
      finishActivity(workbookActivityStatus, workbookActivitySummary);
    }
  }

  async function acknowledgeStaleSource(message: ChatMessage) {
    if (!canAcknowledgeStaleSource(message)) {
      appendMessage({
        role: "system",
        text: "无法恢复这份计划的执行预览，已阻止写入。请按当前数据重新生成预览。"
      });
      return;
    }
    const confirmation = message.staleSourceConfirm!;
    const succeeded = await runPlan(
      confirmation.plan,
      undefined,
      undefined,
      true
    );
    if (!succeeded) return;
    setMessages((current) =>
      current.map((item) =>
        item.id === message.id && item.staleSourceConfirm
          ? {
              ...item,
              staleSourceConfirm: {
                ...item.staleSourceConfirm,
                status: "completed"
              }
            }
          : item
      )
    );
  }

  // 更新某条消息的 functionPreview（确认/取消后标记状态）。
  function markFunctionPreview(
    messageId: string,
    patch: Partial<FunctionPreview>
  ) {
    setMessages((current) =>
      current.map((message) =>
        message.id === messageId && message.functionPreview
          ? {
              ...message,
              functionPreview: { ...message.functionPreview, ...patch }
            }
          : message
      )
    );
  }

  // 解析写入目标：接受裸 A1（E2 / E2:E20）或与预览同表的「表名!地址」。
  // 拒绝跨表与整列/整行（E:E、2:2），后者会写出百万格公式。
  function resolveWriteTarget(preview: FunctionPreview): {
    address?: string;
    error?: string;
  } {
    const raw = preview.writeTarget.trim();
    if (!raw) {
      return { error: "请先输入或拾取要写入的单元格/区域。" };
    }
    let address = raw;
    if (raw.includes("!")) {
      const sheetPart = raw.slice(0, raw.indexOf("!")).replace(/^'|'$/g, "");
      if (sheetPart !== preview.sheet) {
        return {
          error: `只能写入当前工作表「${preview.sheet}」，请先在该表中操作。`
        };
      }
      address = raw.slice(raw.indexOf("!") + 1);
    }
    address = address.replace(/\$/g, "").trim();
    if (/^[A-Za-z]{1,3}:[A-Za-z]{1,3}$/.test(address) || /^\d+:\d+$/.test(address)) {
      return { error: "请指定具体单元格区域（如 E2:E20），暂不支持整列/整行写入。" };
    }
    if (!isExplicitA1Address(address)) {
      return { error: `「${raw}」不是有效的 A1 地址（如 E2 或 E2:E20）。` };
    }
    return { address };
  }

  // 拾取：即刻读当前工作表选区填入写入目标。用户先在表里选好，再点此按钮。
  // 跨表选择被拒（只能写当前预览表），提示用户切回该表再选。
  async function pickFunctionTarget(message: ChatMessage) {
    const preview = message.functionPreview;
    if (!preview || preview.pickingTarget || busy) return;
    if (!(await targetGuard.check())) return;
    markFunctionPreview(message.id, {
      pickingTarget: true,
      targetError: undefined
    });
    try {
      const picked = await readSelectedRange();
      if (!picked) {
        markFunctionPreview(message.id, {
          pickingTarget: false,
          targetError: "没读到选区，请先在表里点选目标单元格再点拾取。"
        });
        return;
      }
      if (picked.sheet !== preview.sheet) {
        markFunctionPreview(message.id, {
          pickingTarget: false,
          targetError: `请在工作表「${preview.sheet}」里选，当前选区在「${picked.sheet}」。`
        });
        return;
      }
      markFunctionPreview(message.id, {
        pickingTarget: false,
        writeTarget: picked.address,
        targetError: undefined
      });
    } catch {
      markFunctionPreview(message.id, {
        pickingTarget: false,
        targetError: "拾取失败，请手动输入地址。"
      });
    }
  }

  // 确认写入：写入目标即生成时的锚点，公式落在其首格。多格区域按 R1C1 铺满、
  // 相对引用逐格平移——R1C1 在写入前从目标首格现取（此时锚点已正确）。
  // 复用 runPlan 的撤销与独立验收。
  async function applyFunctionPreview(message: ChatMessage) {
    if (!(await targetGuard.check())) return;
    const preview = message.functionPreview;
    if (
      !preview ||
      preview.phase !== "preview" ||
      preview.applied ||
      preview.cancelled ||
      busy
    ) {
      return;
    }
    const resolved = resolveWriteTarget(preview);
    if (!resolved.address) {
      markFunctionPreview(message.id, { targetError: resolved.error });
      return;
    }
    const formula =
      preview.version === "modern"
        ? preview.modernFormula
        : preview.compatFormula;
    const explanation =
      preview.version === "modern"
        ? preview.modernExplanation
        : preview.compatExplanation;

    // 验算不通过拦截：拿不到公式，或试算报错且不是"外部工作簿未打开"导致——不硬写。
    const noFormula = !preview.modernFormula && !preview.compatFormula;
    if (noFormula) {
      markFunctionPreview(message.id, {
        targetError:
          "无法生成可靠公式：未能得到可用公式。为避免硬写错误结果，已拦截写入；请手动填写，或调整描述后重试。"
      });
      return;
    }
    const activeTrial =
      preview.version === "modern"
        ? preview.modernResult
        : preview.compatResult;
    if (
      formulaTrialFailed(activeTrial) &&
      formulaExternalFileNames(formula).length === 0
    ) {
      markFunctionPreview(message.id, {
        targetError: `试算未通过（${activeTrial}）：为避免写入错误结果，已拦截写入；请检查公式或手动填写。`
      });
      return;
    }
    const fuzzy = fuzzyLookupMatch(formula);
    if (fuzzy !== null) {
      markFunctionPreview(message.id, {
        targetError: `公式包含模糊匹配（${fuzzy}）：近似匹配会静默返回错误数据，已拦截写入；请手动改成精确匹配（VLOOKUP 用 FALSE、MATCH 用 0），或调整描述重试。`
      });
      return;
    }

    // 多格区域才需 R1C1 铺满；单格直写字面公式。R1C1 从目标首格现取。
    let formulaR1C1 = "";
    if (resolved.address.includes(":")) {
      const firstCell = firstCellOfRange(resolved.address);
      if (firstCell) {
        try {
          const trial = await previewFormulaFirstCell(
            preview.sheet,
            firstCell,
            formula,
            preview.externalFiles
              ? new Set(preview.externalFiles)
              : undefined
          );
          formulaR1C1 = trial.formulaR1C1;
        } catch {
          formulaR1C1 = "";
        }
      }
    }

    const plan: AnalysisPlan = {
      id: `function-${message.id}`,
      title: "填入公式",
      summary: explanation,
      assumptions: [],
      warnings: [],
      actions: [
        {
          type: "writeFormulas",
          sheet: preview.sheet,
          range: resolved.address,
          formulas: [[formula]],
          ...(formulaR1C1 ? { formulaR1C1 } : {})
        }
      ],
      // 显式挂验收条件：verificationGaps 只看 acceptanceCriteria，不看内部推断，
      // 不挂会误报"缺少独立验收"。多格用 R1C1、单格用字面公式各自对应。
      acceptanceCriteria: [
        formulaR1C1
          ? {
              type: "formulasR1C1Equal",
              sheet: preview.sheet,
              range: resolved.address,
              expected: formulaR1C1
            }
          : {
              type: "formulasEqual",
              sheet: preview.sheet,
              range: resolved.address,
              expected: [[formula]]
            }
      ]
    };
    // /function 预览锚定在当前实时工作簿（captureWorkbook、previewFormulaFirstCell
    // 都通过 Excel.run 与打开文档交互），写入必须直接操作实时工作簿，不能路由到
    // folder 后端写磁盘文件——否则用户看不到结果。
    // 注意：/function 只写单个公式单元格，无需源数据指纹校验（与多表分析不同）。

    // 强制走本地 Excel 路径：/function 预览基于实时工作簿，必须写入实时工作簿。
    const planSucceeded = await runPlan(
      plan,
      preview.externalFiles ? new Set(preview.externalFiles) : undefined,
      true  // forceLocal: 强制本地路径
    );
    markFunctionPreview(message.id, {
      applied: planSucceeded,
      appliedTarget: planSucceeded ? resolved.address : undefined,
      targetError: undefined,
      pickingTarget: false
    });
  }

  function cancelFunctionPreview(message: ChatMessage) {
    if (!message.functionPreview) return;
    markFunctionPreview(message.id, { cancelled: true, pickingTarget: false });
  }

  function beginSaveTool(plan: AnalysisPlan) {
    beginSaveCandidate(plan);
    setToolName(plan.title);
    setToolDescription(plan.summary);
  }

  function confirmSaveTool() {
    if (!saveCandidate) return;
    try {
      const tool = createTool(
        saveCandidate,
        toolName,
        toolDescription,
        selectedSheetNames,
        {
          fixedContent: approveFixedContent,
          destructive: approveDestructive
        }
      );
      saveTool(tool);
      closeSaveCandidate();
      appendMessage({
        role: "system",
        text: `已把「${tool.name}」保存为参数化工作流。`
      });
    } catch (reason) {
      appendMessage({
        role: "system",
        text: reason instanceof Error ? reason.message : "无法保存工具"
      });
    }
  }

  function renderToolParameter(tool: SavedTool, parameter: ToolParameter) {
    return (
      <label key={parameter.id}>
        <span>{parameter.label}</span>
        {parameter.type === "outputWorksheet" ||
        parameter.type === "range" ? (
          <input
            value={
              toolParameterValues[parameter.id] ?? parameter.defaultValue
            }
            placeholder={
              parameter.type === "outputWorksheet"
                ? "输入新的工作表名称"
                : "例如 A1:E812"
            }
            onChange={(event) =>
              updateToolParameter(tool, parameter, event.target.value)
            }
          />
        ) : (
          <select
            value={
              toolParameterValues[parameter.id] ?? parameter.defaultValue
            }
            onChange={(event) =>
              updateToolParameter(tool, parameter, event.target.value)
            }
          >
            {parameter.type === "worksheet"
              ? (workbook?.worksheets ?? []).map((sheet) => (
                  <option key={sheet.name} value={sheet.name}>
                    {sheet.name}
                  </option>
                ))
              : [
                  <option key="field-placeholder" value="">
                    请选择对应字段
                  </option>,
                  ...fieldOptions(tool, parameter).map((field) => (
                    <option key={field} value={field}>
                      {field}
                    </option>
                  ))
                ]}
          </select>
        )}
      </label>
    );
  }

  function workflowToolPresentation(tool: SavedTool) {
    const actionTypes = new Set(
      tool.planTemplate.actions.map((action) => action.type)
    );
    if (actionTypes.has("splitGroupAggregate")) {
      return { label: "拆分统计", glyph: "分", tone: "sage" };
    }
    if (
      actionTypes.has("createPivotTable") ||
      actionTypes.has("createChart")
    ) {
      return { label: "分析呈现", glyph: "析", tone: "blue" };
    }
    if (
      actionTypes.has("filterRange") ||
      actionTypes.has("sortRange")
    ) {
      return { label: "整理数据", glyph: "整", tone: "amber" };
    }
    return { label: "工作流程", glyph: "工", tone: "slate" };
  }

  function formatToolDate(value: string): string {
    const parsed = new Date(value);
    if (Number.isNaN(parsed.getTime())) return "";
    return new Intl.DateTimeFormat("zh-CN", {
      month: "numeric",
      day: "numeric"
    }).format(parsed);
  }

  function buildFocusPayload(
    initialView: FocusPayload["initialView"]
  ): FocusPayload {
    return {
      type: "focus-state",
      workbookName: workbook?.name ?? "",
      initialView,
      activeConversationId: chatHistory.activeConversationId,
      conversations: chatHistory.conversations.map((conversation) => ({
        id: conversation.id,
        title: conversation.title,
        updatedAt: conversation.updatedAt,
        messages: conversation.messages.slice(-100).map((message) => ({
          id: message.id,
          role: message.role,
          text: message.text ?? "",
          createdAt: message.createdAt,
          plan: message.plan
            ? {
                title: message.plan.title,
                summary: message.plan.summary,
                steps: message.plan.actions.map(actionLabel)
              }
            : undefined,
          result: message.resultContext
            ? {
                title: message.resultContext.title,
                headers: message.resultContext.headers,
                rows: message.resultContext.rows.slice(0, 200)
              }
            : undefined
        }))
      })),
      tools: [
        ...tools.map((tool) => ({
          id: tool.id,
          kind: "workflow" as const,
          name: tool.name,
          description: tool.description,
          category: workflowToolPresentation(tool).label,
          steps: tool.planTemplate.actions.length,
          stepLabels: tool.planTemplate.actions.map(actionLabel),
          dsl: renderToolDsl(tool.planTemplate)
        })),
        ...queryTools.map((tool) => ({
          id: tool.id,
          kind: "query" as const,
          name: tool.name,
          description: tool.description,
          category: "本地查询",
          steps: 1,
          stepLabels: []
        }))
      ]
    };
  }

  function openFocusWindow() {
    if (!isRunningInExcel()) {
      appendMessage({
        role: "system",
        text: "专注窗口需要从 Excel 任务窗格打开。"
      });
      return;
    }
    const initialView: FocusPayload["initialView"] = toolsOpen
      ? "tools"
      : "conversation";
    const payload = buildFocusPayload(initialView);
    try {
      localStorage.setItem(FOCUS_PAYLOAD_STORAGE_KEY, JSON.stringify(payload));
    } catch {
      // DialogApi 1.2 messaging remains the primary synchronization path.
    }
    const canMessageChild = Office.context.requirements.isSetSupported(
      "DialogApi",
      "1.2"
    );
    if (focusDialogRef.current) {
      if (canMessageChild) {
        focusDialogRef.current.messageChild(JSON.stringify(payload));
      }
      return;
    }
    setFocusOpening(true);
    const url = new URL("focus.html", window.location.href).toString();
    Office.context.ui.displayDialogAsync(
      url,
      { height: 93, width: 100, displayInIframe: false },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          setFocusOpening(false);
          appendMessage({
            role: "system",
            text: `无法打开专注窗口：${result.error.message}`
          });
          return;
        }
        const dialog = result.value;
        focusDialogRef.current = dialog;
        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (event) => {
            if (!("message" in event)) return;
            try {
              const message = JSON.parse(event.message) as { type?: string };
              if (message.type === "focus-ready" && canMessageChild) {
                dialog.messageChild(JSON.stringify(payload));
                setFocusOpening(false);
              }
              if (message.type === "focus-ready" && !canMessageChild) {
                setFocusOpening(false);
              }
              if (message.type === "focus-close") {
                dialog.close();
                focusDialogRef.current = null;
                setFocusOpening(false);
              }
            } catch {
              // Ignore messages that don't belong to the focus workspace.
            }
          }
        );
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          focusDialogRef.current = null;
          setFocusOpening(false);
        });
      }
    );
  }

  function openPlanEditorDialog(plan: AnalysisPlan, target: PlanEditorTarget) {
    if (!isRunningInExcel()) {
      appendMessage({
        role: "system",
        text: "计划编辑器需要从 Excel 任务窗格打开。"
      });
      return;
    }

    const payload = {
      type: "plan-editor-init" as const,
      payload: {
        plan,
        target,
        workbook: workbook!,
        authorizedSheets: selectedNamesFor(workbook!)
      }
    };

    try {
      localStorage.setItem("EB_PLAN_EDITOR_PAYLOAD", JSON.stringify(payload));
    } catch {
      // DialogApi 1.2 messaging remains the primary synchronization path.
    }

    const canMessageChild = Office.context.requirements.isSetSupported(
      "DialogApi",
      "1.2"
    );

    if (planEditorDialogRef.current) {
      if (canMessageChild) {
        planEditorDialogRef.current.messageChild(JSON.stringify(payload));
      }
      return;
    }

    const url = new URL("plan-editor.html", window.location.href).toString();
    Office.context.ui.displayDialogAsync(
      url,
      { height: 93, width: 100, displayInIframe: false },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          appendMessage({
            role: "system",
            text: `无法打开编辑器：${result.error.message}`
          });
          return;
        }

        const dialog = result.value;
        planEditorDialogRef.current = dialog;

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (event) => {
            if (!("message" in event)) return;
            try {
              const message = JSON.parse(event.message) as { type?: string; payload?: any };

              if (message.type === "plan-editor-ready" && canMessageChild) {
                dialog.messageChild(JSON.stringify(payload));
              }

              if (message.type === "plan-editor-save" && message.payload) {
                if (planEditorDialogRef.current !== dialog) return;
                const {
                  plan: updatedPlan,
                  target: saveTarget,
                  updatedTools,
                  closeAfterSave
                } = message.payload;

                if (saveTarget.kind === "message") {
                  setMessages((prev) =>
                    prev.map((msg) =>
                      msg.id === saveTarget.messageId
                        ? { ...msg, plan: updatedPlan }
                        : msg
                    )
                  );
                  // 编辑后旧版本作废：撤回「已验证」资格（须重新执行并验证才能保存为工具），
                  // 并放弃该消息上仍未写回的审阅副本，避免拿旧副本去写已改的计划。
                  unmarkPlanVerified(updatedPlan.id);
                  if (nativeComparison.review?.messageId === saveTarget.messageId) {
                    void nativeComparison.discard(saveTarget.messageId);
                  }
                  appendMessage({
                    role: "system",
                    text: "已更新计划。编辑后需重新确认预览才能写入。"
                  });
                }

                if (saveTarget.kind === "savedTool" && updatedTools) {
                  // updateToolPlanTemplate 返回的是 SavedTool[]（workflow 工具）
                  setTools(updatedTools);
                  appendMessage({
                    role: "system",
                    text: "已更新工具计划。运行前仍会按当前授权重新校验并预览。"
                  });
                }

                // 「保存」按钮触发的保存由父窗口关闭对话框；Ctrl+S 快捷保存则留在窗口
                if (closeAfterSave) {
                  dialog.close();
                  planEditorDialogRef.current = null;
                }
              }

              if (message.type === "plan-editor-close") {
                dialog.close();
                planEditorDialogRef.current = null;
              }
            } catch {
              // Ignore malformed messages
            }
          }
        );

        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          planEditorDialogRef.current = null;
        });
      }
    );
  }

  function openTools() {
    setModelMenuOpen(false);
    setMoreMenuOpen(false);
    setHistoryOpen(false);
    closeSettings();
    resetToolDrawer();
    setToolsOpen(true);
  }

  async function prepareToolRun(tool: SavedTool) {
    setToolActionBusy("prepare");
    try {
      if (!isRunningInExcel() || !workbook) {
        selectTool(tool);
        setToolDrawerView("run");
        return;
      }
      setStatus("scanning");
      try {
        const snapshot = await captureWorkbookStructure(
          workbookScopeMode === "manual" ? selectedSheetNames : undefined
        );
        setWorkbook(snapshot);
        selectTool(tool, snapshot);
        setToolDrawerView("run");
      } catch (reason) {
        appendMessage({
          role: "system",
          text:
            reason instanceof Error
              ? `读取工具所需字段失败：${reason.message}`
              : "读取工具所需字段失败"
        });
        selectTool(tool);
        setToolDrawerView("run");
      } finally {
        setStatus("idle");
      }
    } finally {
      setToolActionBusy(null);
    }
  }

  function openHistory() {
    setModelMenuOpen(false);
    setMoreMenuOpen(false);
    setToolsOpen(false);
    closeSettings();
    setHistoryOpen(true);
  }


  function closeSettings() {
    setSettingsOpen(false);
    setApiKeyDraft("");
    setShowApiKey(false);
    setSettingsLoading(false);
    setSettingsTesting(false);
    setSettingsFeedback("");
    setConnectionDraft(null);
    setPendingDeleteConnectionId(null);
  }

  async function handleOpenSettings(): Promise<boolean> {
    setModelMenuOpen(false);
    setMoreMenuOpen(false);
    setToolsOpen(false);
    setHistoryOpen(false);
    setSettingsOpen(true);
    setApiKeyDraft("");
    return await openSettings();
  }

  async function handleOpenConnectionCreator() {
    setModelMenuOpen(false);
    setMoreMenuOpen(false);
    setToolsOpen(false);
    setHistoryOpen(false);
    setSettingsOpen(true);
    setApiKeyDraft("");
    await openConnectionCreator();
  }

  function handleSelectModel(modelId: string) {
    // 只接受目录里仍然有效的真实连接。旧的 "local" 占位项已从模型列表移除，
    // 万一从旧存储或旧界面残留里传进来，也不能写回选择存储——那会让后续请求
    // 携带一个不存在的 modelId。
    if (!isRealModelId(modelId)) return;
    selectModel(modelId);
    setModelMenuOpen(false);
  }

  async function previewTool(tool: SavedTool) {
    if (!(await targetGuard.check())) return;
    try {
      /*
       * 已保存工具：模板里可能残留上一次保存时带进来的 `sourceDependencySheets`，
       * 但这一步**没有真的读表**（只是按模板+参数实例化），所以依赖不确定 ——
       * 剥掉该字段，恢复登记退回既有保守口径，不拿模板当"读过"的证据。
       */
      const plan = attestSourceDependency(
        instantiateTool(tool, toolParameterValues, workbook),
        null
      );
      if (
        sourceMode === "workbook" &&
        isRunningInExcel() &&
        workbook
      ) {
        const sourceSheets = selectedNamesFor(workbook);
        plan.sourceFingerprintSheets = sourceSheets;
        plan.sourceFingerprint = toolSchemaFingerprintForSnapshot(
          plan,
          workbook
        );
      } else if (sourceMode === "folder" && workbook) {
        plan.sourceFingerprintSheets = selectedNamesFor(workbook);
        plan.sourceFingerprint = workbook.sourceFingerprint;
      }
      appendMessage({
        role: "assistant",
        text: `已加载工具「${tool.name}」。请确认下面的执行预览。`,
        plan,
        provider: "local"
      });
      setToolsOpen(false);
    } catch (reason) {
      appendMessage({
        role: "system",
        text: reason instanceof Error ? reason.message : "无法加载工具"
      });
    }
  }

  // "仍要重新计算"：取复用消息前最近的用户提问，绕过缓存重发并覆写缓存。
  function recomputeReusedMessage(message: ChatMessage) {
    if (busy) return;
    const index = messages.findIndex((item) => item.id === message.id);
    if (index < 0) return;
    let originalPrompt = "";
    for (let i = index - 1; i >= 0; i -= 1) {
      const candidate = messages[i];
      if (candidate.role === "user" && candidate.text?.trim()) {
        originalPrompt = candidate.text.trim();
        break;
      }
    }
    if (!originalPrompt) return;
    void sendMessage({ forceRecompute: true, overrideText: originalPrompt });
  }

  function saveQueryFromMessage(message: ChatMessage) {
    const request = message.intentMemory?.toolRequest;
    if (!request || !workbook) return;
    const sourceNames = message.querySourceSheetNames;
    const querySourceMode = message.querySourceMode;
    if (!sourceNames || !querySourceMode) {
      appendMessage({
        role: "system",
        text: "这条历史结果缺少原始数据范围，不能保存为固化查询。"
      });
      return;
    }
    const sourceIds = message.querySourceSheetIds ?? [];
    const tool = createQueryTool(
      message.resultContext?.title ?? "本地查询",
      message.text ?? "重复运行确定性本地查询",
      request,
      querySourceMode,
      sourceNames,
      sourceIds,
      message.resultContext?.headers ?? []
    );
    saveQueryTool(tool);
    appendMessage({
      role: "system",
      text: `已保存固化查询「${tool.name}」。重复运行将直接使用本地执行器。`
    });
  }

  async function runQueryTool(tool: SavedQueryTool) {
    if (!(await targetGuard.check())) return;
    if (!workbook) return;
    setToolActionBusy("query");
    setStatus("tooling");
    beginActivity("正在运行固化查询", "直接调用本地确定性执行器。");
    try {
      if (tool.sourceMode !== sourceMode) {
        throw new SavedQueryToolFallbackError([
          `工具来源是${tool.sourceMode === "folder" ? "文件夹" : "当前工作簿"}，与当前数据来源不一致`
        ]);
      }
      const result = await executeSavedQueryTool(tool, workbook, {
        workbook: (request) =>
          executeQueryTableTool(
            request,
            tool.sourceSheetNames,
            workbook.activeWorksheet
          ),
        folder: (request) => {
          if (!folderCatalog) {
            throw new Error("文件夹会话已失效，请重新选择文件夹");
          }
          return withFolderSessionRecovery(folderCatalog.sessionId, (sessionId) =>
            executeFolderQuery(sessionId, request)
          );
        }
      });
      appendMessage({
        role: "assistant",
        text: `固化查询「${tool.name}」已完成：扫描 ${result.scannedRows.toLocaleString()} 行，返回 ${result.rows.length} 行；模型调用 0 次。`,
        resultContext: {
          kind: "table",
          title: result.title,
          headers: result.headers.map(String),
          rows: result.rows,
          sourceSheets: result.sourceSheets,
          warnings: result.warnings
        },
        provider: "local"
      });
      setToolsOpen(false);
    } catch (reason) {
      appendMessage({
        role: "system",
        text:
          reason instanceof SavedQueryToolFallbackError
            ? `${reason.message}。请重新描述需求，由模型确认是否更新工具。`
            : reason instanceof Error
              ? reason.message
              : "固化查询执行失败"
      });
    } finally {
      setStatus("idle");
      setToolActionBusy(null);
      finishActivity();
    }
  }

  // 当输入内容变化时检测斜杠命令：仅当整段以 "/" 开头且尚未含空格时展开候选。
  function handleComposerChange(event: React.ChangeEvent<HTMLTextAreaElement>) {
    const value = event.target.value;
    setPrompt(value);
    // 仅在开头处触发：以 / 开头、还未输入空格（即仍在拼命令名）时显示补全。
    // "/model " 之后进入模型选择二级菜单，空格后的文字作为模型过滤词。
    const modelMatch = /^\/model\s+(.*)$/.exec(value);
    if (modelMatch) {
      setSlashMode("model");
      setSlashFilter(modelMatch[1]);
      setShowSlashAutocomplete(true);
      return;
    }
    // 仍在拼一级命令名：以 / 开头且未含空格。
    const commandMatch = /^\/([^\s]*)$/.exec(value);
    if (commandMatch) {
      setSlashMode("command");
      setSlashFilter(commandMatch[1]);
      setShowSlashAutocomplete(true);
    } else {
      setShowSlashAutocomplete(false);
    }
  }

  // 选中某个斜杠命令或模型：一级命令填入 "/命令 " 继续打字；模型项直接切换并清空输入。
  function handleSlashCommandSelect(value: string) {
    if (slashMode === "model") {
      // 模型选择：直接切换，清空输入框，关闭补全。
      handleSelectModel(value);
      setPrompt("");
      setShowSlashAutocomplete(false);
    } else {
      if (value === "help") {
        appendMessage({
          role: "assistant",
          text: MODEL_REQUIRED_HELP_TEXT
        });
        setPrompt("");
        setShowSlashAutocomplete(false);
        const input = composerInputRef.current;
        if (input) {
          input.focus();
        }
        return;
      }

      // 一级命令：/function 填入等继续输入描述；/model 进二级菜单。
      if (value === "model") {
        setPrompt("/model ");
        setSlashMode("model");
        setSlashFilter("");
        // 保持补全开启，接下来会切到模型列表。
      } else {
        const nextValue = `/${value} `;
        setPrompt(nextValue);
        setShowSlashAutocomplete(false);
        const input = composerInputRef.current;
        if (input) {
          input.focus();
          requestAnimationFrame(() => {
            input.setSelectionRange(nextValue.length, nextValue.length);
          });
        }
      }
    }
  }

  function handleComposerKeyDown(event: React.KeyboardEvent<HTMLTextAreaElement>) {
    // 补全菜单展开时，方向键/回车/Tab 交给菜单处理（见 SlashCommandAutocomplete）。
    if (
      showSlashAutocomplete &&
      ["ArrowDown", "ArrowUp", "Enter", "Tab"].includes(event.key)
    ) {
      return;
    }
    if (event.key === "Escape" && showSlashAutocomplete) {
      setShowSlashAutocomplete(false);
      return;
    }
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      // 运行中回车 = 转向（带这句话打断重跑）；空闲回车 = 正常发送。
      if (busy) {
        if (!prompt.trim()) return;
        steerTurn(prompt);
        return;
      }
      void submitComposer();
    }
  }

  // 「操作目标确认之后」的回调重入：这句话已经送过一次（气泡已落盘、输入框已清空），
  // 带上 `targetConfirmedRetry` 让 `sendMessage` 不再重复送达一遍。与 WPS 同款标记。
  retrySendRef.current = (text, messageUnderstanding) => {
    void sendMessage({ overrideText: text, messageUnderstanding, targetConfirmedRetry: true });
  };

  return (
    <main className="chat-shell">
      <header
        className={`chat-header${modelMenuOpen || moreMenuOpen ? " menu-open" : ""}`}
      >
        <img className="brand-mark" src={brandLogoUrl} alt="" />
        <div className="brand-copy">
          <strong className="app-logo" {...logoLongPress}>
            Excel Bro
          </strong>
          {headerStatusText && <span className={headerStatusClassName}>
            <i />
            {serverOnline && modelNotConfigured && (
              <svg
                className="model-status-icon"
                viewBox="0 0 16 16"
                aria-hidden="true"
                focusable="false"
              >
                <circle cx="8" cy="8" r="6.25" />
                <path d="M5.5 8h5" />
              </svg>
            )}
            {headerStatusText}
            {serverOnline && modelNotConfigured && (
              <button
                type="button"
                className="model-help-button"
                title="查看模型配置说明"
                aria-label="查看模型配置说明"
                onClick={() => handleSlashCommandSelect("help")}
              >
                <svg viewBox="0 0 16 16" aria-hidden="true" focusable="false">
                  <circle cx="8" cy="8" r="6.25" />
                  <path d="M6.15 6.15c.14-1.05 1.08-1.78 2.12-1.62 1.02.16 1.72 1.02 1.67 2.02-.04.83-.54 1.42-1.36 1.86-.38.2-.47.38-.43.78" />
                  <path d="M8 11.2h.01" />
                </svg>
              </button>
            )}
          </span>}
        </div>
        <div className="header-actions">
        <div
          className={`model-picker${showModelGuide ? " needs-model" : ""}`}
        >
          <button
            type="button"
            className="model-picker-trigger"
            disabled={!serverOnline || busy}
            aria-haspopup="menu"
            aria-expanded={modelMenuOpen}
            aria-label={hasConfiguredModel ? "选择或添加模型" : "添加模型"}
            title={selectedModel ? `选择或添加模型 · 当前模型：${selectedModel.label}` : "选择或添加模型"}
            onClick={() => {
              setHistoryOpen(false);
              setToolsOpen(false);
              setMoreMenuOpen(false);
              closeSettings();
              setModelMenuOpen((current) => !current);
            }}
          >
            <ModelPickerIcon />
          </button>
          {modelMenuOpen && (
            <div className="model-menu" role="menu" aria-label="模型">
              <span className="model-menu-title">选择模型</span>
              {modelOptions.map((option) => (
                <button
                  type="button"
                  role="menuitemradio"
                  aria-checked={option.id === effectiveModelId}
                  key={option.id}
                  disabled={!agentModelAvailable(option)}
                  title={!agentModelAvailable(option) ? agentModelUnavailableReason(option) : undefined}
                  onClick={() => handleSelectModel(option.id)}
                >
                  <i />
                  <span className="model-option-copy">
                    <span>{option.label}</span>
                    {!agentModelAvailable(option) && <small>{agentModelUnavailableReason(option)}</small>}
                  </span>
                  {option.id === effectiveModelId && <b>✓</b>}
                </button>
              ))}
              <div className="model-menu-actions">
                <button
                  type="button"
                  role="menuitem"
                  onClick={() => void handleOpenConnectionCreator()}
                >
                  ＋ 添加模型连接
                </button>
                <button
                  type="button"
                  role="menuitem"
                  onClick={() => void handleOpenSettings()}
                >
                  管理模型连接
                </button>
              </div>
            </div>
          )}
          {showModelGuide && (
            <section
              className="first-model-guide"
              aria-label="首次模型设置引导"
            >
              <div className="guide-spark" aria-hidden="true">
                <i />
                <i />
                <i />
              </div>
              <div>
                <strong>添加你的第一个大模型</strong>
                <span>
                  Excel Bro 不会预置模型或密钥，由你选择 Kimi、DeepSeek、OpenAI
                  或其他兼容服务。
                </span>
              </div>
              <div className="first-model-guide-actions">
                <button type="button" onClick={dismissModelGuide}>
                  稍后
                </button>
                <button
                  type="button"
                  className="primary"
                  onClick={() => void handleOpenConnectionCreator()}
                >
                  现在添加
                </button>
              </div>
            </section>
          )}
        </div>
          <button
            className="header-button new-chat-entry"
            onClick={handleNewChat}
            disabled={busy}
            title="新对话"
            aria-label="新对话"
          >
            <span aria-hidden="true">＋</span>
          </button>
          <button
            className="header-button tools-entry"
            onClick={openTools}
            disabled={busy}
            title="我的工具"
            aria-label="我的工具"
          >
            <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
              <path d="M8.5 7V5.8c0-1 .8-1.8 1.8-1.8h3.4c1 0 1.8.8 1.8 1.8V7" />
              <rect x="3" y="7" width="18" height="12.5" rx="2.5" />
              <path d="M3.5 11.5h17M10 11.5v2h4v-2" />
            </svg>
          </button>
          <button
            className="header-button folder-backup-entry"
            onClick={() => {
              closeAllDrawers();
              closeAllMenus();
              setFolderBackupOpen(true);
            }}
            disabled={busy}
            title={
              folderBackupTarget.kind === "unavailable"
                ? "恢复点与备份"
                : `恢复点与备份：${folderBackupTarget.folderPath}`
            }
            aria-label="备份与恢复"
          >
            <span aria-hidden="true">↶</span>
            {undoStack.length + folderUndoStack.length > 0 && (
              <b>{undoStack.length + folderUndoStack.length}</b>
            )}
          </button>
          <button
            className={`header-button pet-toggle${petVisible ? " active" : ""}`}
            onClick={togglePetVisibility}
            title={petVisible ? "隐藏格仔" : "显示格仔"}
            aria-label={petVisible ? "隐藏格仔" : "显示格仔"}
            aria-pressed={petVisible}
          >
            <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
              <ellipse cx="6.3" cy="10" rx="1.9" ry="2.4" />
              <ellipse cx="10.4" cy="7.1" rx="1.9" ry="2.5" />
              <ellipse cx="14.6" cy="7.1" rx="1.9" ry="2.5" />
              <ellipse cx="18.7" cy="10" rx="1.9" ry="2.4" />
              <path d="M7.4 16c0-2.7 2-4.5 4.6-4.5s4.6 1.8 4.6 4.5c0 2.4-1.9 4-4.6 4s-4.6-1.6-4.6-4Z" />
            </svg>
          </button>
          <div className="view-menu-wrap">
            <button
              className="header-button"
              onClick={() => {
                setModelMenuOpen(false);
                setHistoryOpen(false);
                setToolsOpen(false);
                closeSettings();
                setMoreMenuOpen((current) => !current);
              }}
              aria-haspopup="menu"
              aria-expanded={moreMenuOpen}
              title="更多"
              aria-label="更多"
            >
              <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
                <circle cx="5" cy="12" r="1.6" />
                <circle cx="12" cy="12" r="1.6" />
                <circle cx="19" cy="12" r="1.6" />
              </svg>
            </button>
            {moreMenuOpen && (
              <>
                <div
                  className="more-menu-backdrop"
                  onClick={() => setMoreMenuOpen(false)}
                  aria-hidden="true"
                />
                <div className="view-menu" role="menu" aria-label="更多">
                  <button
                    type="button"
                    role="menuitem"
                    onClick={openHistory}
                    disabled={busy}
                  >
                    <strong>历史对话</strong>
                    <span>查看并恢复以前的对话</span>
                  </button>
                  <button
                    type="button"
                    role="menuitem"
                    onClick={() => {
                      setMoreMenuOpen(false);
                      setIsRuleManagerOpen(true);
                    }}
                    disabled={busy}
                  >
                    <strong>EB 函数说明</strong>
                    <span>查看内置 =EB() 函数用法</span>
                  </button>
                  <button
                    type="button"
                    role="menuitem"
                    onClick={() => {
                      setMoreMenuOpen(false);
                      openFocusWindow();
                    }}
                    disabled={focusOpening}
                  >
                    <strong>{focusOpening ? "正在打开..." : "专注窗口"}</strong>
                    <span>在独立窗口中打开对话</span>
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      </header>

      {showWidenGuide && (
        <section
          className="widen-pane-guide"
          aria-label="拉宽窗格引导"
        >
          <div className="guide-spark" aria-hidden="true">
            <i />
            <i />
            <i />
          </div>
          <strong>把面板拉宽一点</strong>
          <span>
            向左拖动面板左边缘，留出更多对话空间，然后我们会引导你添加第一个大模型。
          </span>
          <div className="widen-pane-guide-actions">
            <button type="button" onClick={() => setWidenStepDone(true)}>
              跳过
            </button>
          </div>
        </section>
      )}

      {folderBackupOpen && (
        <FolderBackupManager
          target={folderBackupTarget}
          onClose={() => setFolderBackupOpen(false)}
          onUndo={async (planId, undoId, undoSessionId) => {
            await undoFolderExecutionTo(undoId, {
              sessionId: undoSessionId,
              undoId
            });
          }}
          onBackupDeleted={reconcileFolderUndoIds}
          workbookUndoCount={undoStack.length}
          folderUndoCount={folderUndoStack.length}
          checkpointSection={
            /*
              删除恢复点统一从「备份与恢复」查看：按**工作簿身份**过滤，
              旧版只存 operationId 的记录也在这里被找回来，而不是被猜进当前会话。
            */
            <WorksheetCheckpointRecoverySection
              controller={worksheetCheckpoint}
              workbookPath={resultWorkbookIdentity || null}
            />
          }
        />
      )}

      {settingsOpen && (
        <aside className="tool-drawer settings-drawer" aria-label="模型设置">
          <div className="tool-drawer-header">
            <div>
              <strong>模型设置</strong>
              <span>密钥仅保存在这台电脑的本地服务中</span>
            </div>
            <button
              type="button"
              className="drawer-text-action"
              onClick={exportDiagnosticReport}
              title="导出不含原始数据和密钥的诊断 JSON"
            >
              导出诊断
            </button>
            <button onClick={closeSettings} aria-label="关闭">
              ×
            </button>
          </div>
          <div className="settings-content">
            {settingsLoading ? (
              <div className="settings-loading">
                <i />
                正在读取本地模型连接...
              </div>
            ) : (
              <>
                {hasEnvironmentModel ? (
                  <>
                    <section className="settings-summary">
                      <div>
                        <span>服务地址</span>
                        <strong>{modelSettings?.baseUrl}</strong>
                      </div>
                      <div>
                        <span>默认模型</span>
                        <strong>{modelSettings?.defaultModel}</strong>
                      </div>
                      <div>
                        <span>API Key</span>
                        <strong>
                          {modelSettings?.apiKeyConfigured
                            ? `已配置 ${modelSettings.apiKeyHint ?? ""}`
                            : "未配置（本地服务可无需密钥）"}
                        </strong>
                      </div>
                    </section>

                    <section className="api-key-settings">
                      <div>
                        <strong>更换默认连接的 API Key</strong>
                        <span>保存后立即用于新的模型请求，无需重启后端。</span>
                      </div>
                      <label>
                        <span>新的 API Key</span>
                        <div className="api-key-input-row">
                          <input
                            type={showApiKey ? "text" : "password"}
                            value={apiKeyDraft}
                            disabled={settingsSaving || !serverOnline}
                            autoComplete="off"
                            spellCheck={false}
                            placeholder={
                              modelSettings?.apiKeyConfigured
                                ? "留空不会覆盖现有密钥"
                                : "粘贴 API Key"
                            }
                            onChange={(event) => {
                              setApiKeyDraft(event.target.value);
                              setSettingsFeedback("");
                            }}
                            onKeyDown={(event) => {
                              if (event.key === "Enter") void saveApiKey();
                            }}
                          />
                          <button
                            type="button"
                            disabled={!apiKeyDraft || settingsSaving}
                            onClick={() =>
                              setShowApiKey((current) => !current)
                            }
                          >
                            {showApiKey ? "隐藏" : "显示"}
                          </button>
                        </div>
                      </label>
                      <p>
                        完整密钥不会从后端读回，也不会保存到聊天记录或浏览器存储。
                      </p>
                      <button
                        className="tool-preview-button"
                        disabled={
                          !apiKeyDraft.trim() ||
                          settingsSaving ||
                          !serverOnline
                        }
                        onClick={() => void saveApiKey()}
                      >
                        {settingsSaving ? "正在保存..." : "保存并立即生效"}
                      </button>
                    </section>
                  </>
                ) : (
                  <section className="settings-first-run">
                    <div className="settings-first-run-icon">＋</div>
                    <div>
                      <strong>
                        {hasManagedModels
                          ? "你的模型连接"
                          : "从添加第一个模型开始"}
                      </strong>
                      <span>
                        {hasManagedModels
                          ? "模型和密钥都由下方连接独立管理。"
                          : "没有预置模型，也不需要先去后端修改配置。填写服务地址、模型 ID 和自己的 API Key 即可。"}
                      </span>
                    </div>
                  </section>
                )}

            <section className="connection-settings">
              <div className="connection-settings-header">
                <div>
                  <strong>
                    {hasEnvironmentModel ? "其他模型连接" : "模型连接"}
                  </strong>
                  <span>每个连接使用独立的服务地址、模型和 API Key。</span>
                </div>
                {!connectionDraft && (
                  <button
                    type="button"
                    disabled={settingsSaving || !serverOnline}
                    onClick={() => void handleOpenConnectionCreator()}
                  >
                    ＋ 添加
                  </button>
                )}
              </div>

              {(modelSettings?.connections.length ?? 0) === 0 &&
              !connectionDraft ? (
                <div className="connection-empty">
                  还没有额外连接。可以添加 DeepSeek、OpenAI 或其他
                  OpenAI-compatible 模型。
                </div>
              ) : (
                <div className="connection-list">
                  {(modelSettings?.connections ?? []).map((connection) => (
                    <article
                      className={`connection-card${
                        effectiveModelId === connection.catalogModelId
                          ? " current"
                          : ""
                      }`}
                      key={connection.id}
                    >
                      <div>
                        <div className="connection-card-title">
                          <strong>{connection.label}</strong>
                          {effectiveModelId === connection.catalogModelId && (
                            <b>当前使用</b>
                          )}
                        </div>
                        <span>{connection.modelId}{connectionModelSuffix(modelSettings?.connections ?? [], connection.id)}</span>
                        <small>{connection.baseUrl}</small>
                          <small>{connection.agentAvailable === false ? connection.agentUnavailableReason : "SDK Agent 可用性请以连接测试为准"}</small>
                        <small>
                          Key：
                          {connection.apiKeyConfigured
                            ? `已配置 ${connection.apiKeyHint ?? ""}`
                            : "未配置"}
                          {connection.supportsVision ? " · 支持图片" : ""}
                          {modelSettings?.formulaModelId ===
                          connection.catalogModelId
                            ? " · /function 公式模型"
                            : ""}
                        </small>
                      </div>
                      {pendingDeleteConnectionId === connection.id ? (
                        <div className="connection-card-actions confirm">
                          <button
                            type="button"
                            onClick={() => setPendingDeleteConnectionId(null)}
                          >
                            取消
                          </button>
                          <button
                            type="button"
                            className="delete"
                            disabled={settingsSaving || settingsTesting}
                            onClick={() => void removeConnection(connection.id)}
                          >
                            确认删除
                          </button>
                        </div>
                      ) : (
                        <div className="connection-card-actions">
                          <button
                            type="button"
                            disabled={settingsSaving || settingsTesting}
                            onClick={() => editModelConnection(connection.id)}
                          >
                            编辑
                          </button>
                          <button
                            type="button"
                            className="delete"
                            disabled={settingsSaving}
                            onClick={() =>
                              setPendingDeleteConnectionId(connection.id)
                            }
                          >
                            删除
                          </button>
                        </div>
                      )}
                    </article>
                  ))}
                </div>
              )}

              {connectionDraft && (
                <div className="connection-editor">
                  <div>
                    <strong>
                      {connectionDraft.id ? "编辑模型连接" : "添加模型连接"}
                    </strong>
                    <button
                      type="button"
                      onClick={() => {
                        setShowApiKey(false);
                        setConnectionDraft(null);
                      }}
                      aria-label="关闭连接编辑"
                    >
                      ×
                    </button>
                  </div>
                  <ModelConnectionProtocolFields draft={connectionDraft} onChange={setConnectionDraft} />
                  <label>
                    <span>连接名称</span>
                    <input
                      value={connectionDraft.label}
                      placeholder="例如 DeepSeek"
                      onChange={(event) =>
                        setConnectionDraft((current) =>
                          current
                            ? { ...current, label: event.target.value }
                            : current
                        )
                      }
                    />
                  </label>
                  <label>
                    <span>服务地址</span>
                    <input
                      value={connectionDraft.baseUrl}
                      placeholder={connectionDraft.protocol === "anthropic" ? "供应商提供的 Anthropic 兼容地址" : "https://api.example.com/v1"}
                      onChange={(event) =>
                        setConnectionDraft((current) =>
                          current
                            ? { ...current, baseUrl: event.target.value }
                            : current
                        )
                      }
                    />
                  </label>
                  <label>
                    <span>模型 ID</span>
                    <input
                      value={connectionDraft.modelId}
                      placeholder="供应商实际接受的模型 ID"
                      onChange={(event) =>
                        setConnectionDraft((current) =>
                          current
                            ? { ...current, modelId: event.target.value }
                            : current
                        )
                      }
                    />
                  </label>
                  <label>
                    <span>API Key</span>
                    <input
                      type={showApiKey ? "text" : "password"}
                      value={connectionDraft.apiKey}
                      disabled={connectionDraft.clearApiKey}
                      autoComplete="off"
                      spellCheck={false}
                      placeholder={
                        connectionDraft.id
                          ? "留空保留现有密钥"
                          : "本地无鉴权服务可留空"
                      }
                      onChange={(event) =>
                        setConnectionDraft((current) =>
                          current
                            ? {
                                ...current,
                                apiKey: event.target.value,
                                clearApiKey: false
                              }
                            : current
                        )
                      }
                    />
                    <button
                      type="button"
                      className="connection-key-toggle"
                      disabled={
                        !connectionDraft.apiKey ||
                        connectionDraft.clearApiKey
                      }
                      onClick={() => setShowApiKey((current) => !current)}
                    >
                      {showApiKey ? "隐藏 Key" : "显示 Key"}
                    </button>
                  </label>
                  {connectionDraft.id &&
                    editingConnection?.apiKeyConfigured && (
                      <label className="connection-clear-key">
                        <input
                          type="checkbox"
                          checked={connectionDraft.clearApiKey}
                          onChange={(event) =>
                            setConnectionDraft((current) =>
                              current
                                ? {
                                    ...current,
                                    apiKey: "",
                                    clearApiKey: event.target.checked
                                  }
                                : current
                            )
                          }
                        />
                        <span>
                          清除现有 API Key（仅适用于无需鉴权的本地服务）
                        </span>
                      </label>
                    )}
                  <label className="connection-vision-toggle">
                    <input
                      type="checkbox"
                      checked={connectionDraft.supportsVision}
                      onChange={(event) =>
                        setConnectionDraft((current) =>
                          current
                            ? {
                                ...current,
                                supportsVision: event.target.checked
                              }
                            : current
                        )
                      }
                    />
                    <span>这个模型支持图片输入</span>
                  </label>
                  <div className="connection-editor-actions">
                    <button
                      type="button"
                      disabled={
                        settingsSaving || settingsTesting || !serverOnline
                      }
                      onClick={() => void verifyConnection()}
                    >
                      {settingsTesting ? "正在测试..." : "测试连接"}
                    </button>
                    <button
                      type="button"
                      className="tool-preview-button"
                      disabled={
                        settingsSaving ||
                        settingsTesting ||
                        !serverOnline ||
                        !connectionVerified
                      }
                      title={
                        !connectionVerified
                          ? "请先测试连接并确认成功"
                          : undefined
                      }
                      onClick={() => void saveConnection()}
                    >
                      {settingsSaving
                        ? "正在保存..."
                        : connectionDraft.id
                          ? "保存修改"
                          : "添加并使用"}
                    </button>
                  </div>
                  {connectionTestFeedback && (
                    <div
                      className={`connection-test-feedback ${connectionTestFeedback.kind}`}
                      role={
                        connectionTestFeedback.kind === "error"
                          ? "alert"
                          : "status"
                      }
                      aria-live={
                        connectionTestFeedback.kind === "error"
                          ? "assertive"
                          : "polite"
                      }
                    >
                      <span
                        className="connection-test-feedback-icon"
                        aria-hidden="true"
                      >
                        {connectionTestFeedback.kind === "success" ? "✓" : "!"}
                      </span>
                      <span>{connectionTestFeedback.message}</span>
                    </div>
                  )}
                </div>
              )}
            </section>

            <section className="connection-settings">
              <div className="connection-settings-header">
                <div>
                  <strong>/function 公式模型</strong>
                  <span>
                    /function 生成公式走的是短链、小任务，可单独选一个极速（非推理）模型，与上方聊天所选模型无关。
                  </span>
                </div>
              </div>
              <label className="connection-formula-field">
                <span>用于生成公式的模型</span>
                <select
                  className="connection-formula-select"
                  value={modelSettings?.formulaModelId ?? ""}
                  disabled={settingsSaving || !serverOnline}
                  onChange={(event) =>
                    void saveFormulaModel(event.target.value)
                  }
                >
                  <option value="">跟随全局选择（默认）</option>
                  {modelOptions
                    .filter((option) => option.id !== "local")
                    .map((option) => (
                      <option key={option.id} value={option.id} disabled={!formulaModelAvailable(option)}>
                        {option.label}
                      </option>
                    ))}
                </select>
              </label>
            </section>
              </>
            )}

            {settingsFeedback && (
              <div
                className={
                  (settingsFeedback.includes("已") ||
                    settingsFeedback.includes("成功")) &&
                  !settingsFeedback.startsWith("请")
                    ? "settings-feedback success"
                    : "settings-feedback"
                }
                role="status"
              >
                {settingsFeedback}
              </div>
            )}
          </div>
        </aside>
      )}

      {toolsOpen && (
        <aside className="tool-drawer" aria-label="我的工具">
          <div className="tool-drawer-header">
            <div>
              <strong>
                {toolDrawerView === "library"
                  ? "我的工具"
                  : toolDrawerView === "detail"
                    ? "工具说明"
                    : "运行工具"}
              </strong>
              <span>
                {toolDrawerView === "library"
                  ? "先选择工具，了解清楚后再运行"
                  : toolDrawerView === "detail"
                    ? "确认用途、输入和结果"
                    : "逐步确认本次运行所需信息"}
              </span>
            </div>
            <button onClick={() => setToolsOpen(false)} aria-label="关闭">
              ×
            </button>
          </div>

          {toolDrawerView === "library" && (
            <div className="tool-library">
              {tools.length === 0 && queryTools.length === 0 ? (
                <div className="tool-empty">
                  <i>▦</i>
                  <strong>工具箱还是空的</strong>
                  <span>执行并验证一个计划后，可以把它保存到这里。</span>
                </div>
              ) : (
                <>
                  <div className="tool-library-intro">
                    <span>工具箱</span>
                    <strong>今天想用哪一个？</strong>
                    <p>
                      每个工具都保留自己的处理规则。选择后先看说明，不会立即执行。
                    </p>
                  </div>
                  <div className="tool-card-grid">
                    {tools.map((tool) => {
                      const presentation = workflowToolPresentation(tool);
                      return (
                        <article className="tool-card-wrap" key={tool.id}>
                          <button
                            className="tool-card"
                            onClick={() => openWorkflowToolDetail(tool)}
                          >
                            <span
                              className={`tool-card-icon tone-${presentation.tone}`}
                            >
                              {presentation.glyph}
                            </span>
                            <span className="tool-card-copy">
                              <small>{presentation.label}</small>
                              <strong>{tool.name}</strong>
                              <span title={tool.description}>
                                {tool.description}
                              </span>
                            </span>
                          </button>
                          <div className="tool-card-footer">
                            <span className="tool-card-meta">
                              {tool.planTemplate.actions.length} 个步骤
                              {formatToolDate(tool.updatedAt)
                                ? ` · ${formatToolDate(tool.updatedAt)} 更新`
                                : ""}
                            </span>
                            <button
                              className="tool-card-delete"
                              aria-label={`删除工具「${tool.name}」`}
                              title={`删除「${tool.name}」`}
                              onClick={() =>
                                requestToolDeletion("workflow", tool)
                              }
                            >
                              删除
                            </button>
                          </div>
                        </article>
                      );
                    })}
                    {queryTools.map((tool) => (
                      <article className="tool-card-wrap" key={tool.id}>
                        <button
                          className="tool-card"
                          onClick={() => openQueryToolDetail(tool)}
                        >
                          <span className="tool-card-icon tone-blue">查</span>
                          <span className="tool-card-copy">
                            <small>本地查询</small>
                            <strong>{tool.name}</strong>
                            <span title={tool.description}>
                              {tool.description}
                            </span>
                          </span>
                        </button>
                        <div className="tool-card-footer">
                          <span className="tool-card-meta">
                            本地确定性执行
                            {formatToolDate(tool.updatedAt)
                              ? ` · ${formatToolDate(tool.updatedAt)} 更新`
                              : ""}
                          </span>
                          <button
                            className="tool-card-delete"
                            aria-label={`删除工具「${tool.name}」`}
                            title={`删除「${tool.name}」`}
                            onClick={() => requestToolDeletion("query", tool)}
                          >
                            删除
                          </button>
                        </div>
                      </article>
                    ))}
                  </div>
                </>
              )}
            </div>
          )}

          {toolDrawerView === "detail" &&
            tools
              .filter((tool) => tool.id === selectedToolId)
              .map((tool) => {
                const presentation = workflowToolPresentation(tool);
                return (
                  <section className="tool-page tool-detail-page" key={tool.id}>
                    <button
                      className="tool-back-button"
                      onClick={() => setToolDrawerView("library")}
                    >
                      ← 返回工具箱
                    </button>
                    <div className="tool-detail-hero">
                      <span
                        className={`tool-card-icon large tone-${presentation.tone}`}
                      >
                        {presentation.glyph}
                      </span>
                      <div>
                        <small>{presentation.label}</small>
                        <strong>{tool.name}</strong>
                        <span>
                          版本 {tool.version} ·{" "}
                          {tool.planTemplate.actions.length} 个处理步骤
                        </span>
                      </div>
                    </div>
                    <div
                      className="tool-view-switch"
                      role="tablist"
                      aria-label="工具说明视图"
                    >
                      <button
                        role="tab"
                        aria-selected={toolDetailMode === "standard"}
                        className={toolDetailMode === "standard" ? "active" : ""}
                        onClick={() => setToolDetailMode("standard")}
                      >
                        普通视图
                      </button>
                      <button
                        role="tab"
                        aria-selected={toolDetailMode === "expert"}
                        className={toolDetailMode === "expert" ? "active" : ""}
                        onClick={() => setToolDetailMode("expert")}
                      >
                        专家视图
                      </button>
                    </div>
                    {toolDetailMode === "standard" ? (
                      <>
                        <div className="tool-purpose">
                          <strong>它会为你完成什么</strong>
                          <p>{tool.description}</p>
                        </div>
                        <div className="tool-facts">
                          <div>
                            <span>处理方式</span>
                            <strong>{presentation.label}</strong>
                          </div>
                          <div>
                            <span>执行步骤</span>
                            <strong>{tool.planTemplate.actions.length} 步</strong>
                          </div>
                          <div>
                            <span>安全机制</span>
                            <strong>运行前预览</strong>
                          </div>
                        </div>
                        <div className="tool-conversation-note">
                          <span>运行时怎么沟通</span>
                          <p>
                            我会先识别当前工作簿和字段。能够自动匹配的内容不会打扰你，
                            只有发现字段变化时才会请你确认。
                          </p>
                        </div>
                        <details className="tool-steps">
                          <summary>
                            查看完整处理步骤（{tool.planTemplate.actions.length}）
                          </summary>
                          <ol>
                            {tool.planTemplate.actions.map((action, index) => (
                              <li key={`${tool.id}-${index}`}>
                                {actionLabel(action)}
                              </li>
                            ))}
                          </ol>
                        </details>
                        <p className="tool-expert-hint">
                          需要调整处理规则？{" "}
                          <button
                            type="button"
                            onClick={() => setToolDetailMode("expert")}
                          >
                            切换到专家视图
                          </button>
                          后可以编辑计划。
                        </p>
                      </>
                    ) : (
                      <div className="tool-expert-view">
                        <div className="tool-expert-heading">
                          <div>
                            <small>CONTROLLED PLAN</small>
                            <strong>专家计划</strong>
                            <span>由保存的白名单计划确定性生成</span>
                          </div>
                          <div className="tool-expert-actions">
                            <button
                              className="secondary-button"
                              disabled={busy}
                              onClick={() =>
                                openPlanEditorDialog(tool.planTemplate, {
                                  kind: "savedTool",
                                  toolId: tool.id
                                })
                              }
                            >
                              编辑计划
                            </button>
                            <button
                              className="tool-copy-dsl"
                              onClick={() => void copyToolDsl(tool)}
                            >
                              {copiedToolDslId === tool.id
                                ? "已复制"
                                : "复制计划"}
                            </button>
                          </div>
                        </div>
                        <div className="tool-expert-badges">
                          <span>可编辑 JSON</span>
                          <span>运行前预览</span>
                          <span>禁止任意代码</span>
                        </div>
                        <pre className="tool-dsl" tabIndex={0}>
                          <code
                            dangerouslySetInnerHTML={{
                              __html: highlightToolDsl(
                                renderToolDsl(tool.planTemplate)
                              )
                            }}
                          />
                        </pre>
                        <p className="tool-expert-note">
                          这里和普通视图是同一份工具。上面的计划文本是只读展示；要改内容请用
                          「编辑计划」，保存时会检查 JSON 格式与白名单结构。实际运行仍使用受控
                          AnalysisPlan，并在写入前让你确认预览。
                        </p>
                      </div>
                    )}
                    <div className="tool-run-actions">
                      <button
                        className="tool-preview-button"
                        disabled={busy || !workbook || toolActionBusy !== null}
                        onClick={() => void prepareToolRun(tool)}
                      >
                        {toolActionBusy === "prepare" ? (
                          <>
                            <span
                              className="tool-button-spinner"
                              aria-hidden="true"
                            />
                            正在读取工作簿...
                          </>
                        ) : (
                          "使用这个工具"
                        )}
                      </button>
                      <span>下一步先确认本次数据来源，不会立即修改工作簿</span>
                    </div>
                  </section>
                );
              })}

          {toolDrawerView === "detail" &&
            queryTools
              .filter((tool) => tool.id === selectedQueryToolId)
              .map((tool) => (
                <section className="tool-page tool-detail-page" key={tool.id}>
                  <button
                    className="tool-back-button"
                    onClick={() => setToolDrawerView("library")}
                  >
                    ← 返回工具箱
                  </button>
                  <div className="tool-detail-hero">
                    <span className="tool-card-icon large tone-blue">查</span>
                    <div>
                      <small>本地查询</small>
                      <strong>{tool.name}</strong>
                      <span>确定性执行 · 不调用模型</span>
                    </div>
                  </div>
                  <div className="tool-purpose">
                    <strong>它会为你查什么</strong>
                    <p>{tool.description}</p>
                  </div>
                  <div className="tool-facts">
                    <div>
                      <span>运行位置</span>
                      <strong>本地</strong>
                    </div>
                    <div>
                      <span>模型调用</span>
                      <strong>0 次</strong>
                    </div>
                    <div>
                      <span>字段变化</span>
                      <strong>停止并提醒</strong>
                    </div>
                  </div>
                  <div className="tool-conversation-note">
                    <span>运行前检查</span>
                    <p>
                      如果来源或字段发生变化，工具会停止，不会用错误字段继续计算。
                    </p>
                  </div>
                  <div className="tool-run-actions">
                    <button
                      className="tool-preview-button"
                      disabled={busy || !workbook || toolActionBusy !== null}
                      onClick={() => void runQueryTool(tool)}
                    >
                      {toolActionBusy === "query" ? (
                        <>
                          <span
                            className="tool-button-spinner"
                            aria-hidden="true"
                          />
                          正在运行...
                        </>
                      ) : (
                        "运行这个查询"
                      )}
                    </button>
                    <span>查询结果会回到当前对话，不会写入工作簿</span>
                  </div>
                </section>
              ))}

          {toolDrawerView === "run" &&
            tools
              .filter((tool) => tool.id === selectedToolId)
              .map((tool) => {
                const primaryParameters = tool.parameters.filter(
                  (parameter) =>
                    parameter.type === "worksheet" ||
                    parameter.type === "outputWorksheet"
                );
                const advancedParameters = tool.parameters.filter(
                  (parameter) =>
                    parameter.type === "field" ||
                    parameter.type === "range"
                );
                const fieldParameters = advancedParameters.filter(
                  (parameter) => parameter.type === "field"
                );
                const missingFieldCount = fieldParameters.filter(
                  (parameter) =>
                    !(toolParameterValues[parameter.id] ?? "").trim()
                ).length;
                const primaryReady =
                  primaryParameters.length === 0 ||
                  primaryParameters.every(
                    (parameter) =>
                      (toolParameterValues[parameter.id] ?? "").trim()
                  );
                return (
                  <section className="tool-page tool-run-page" key={tool.id}>
                    <button
                      className="tool-back-button"
                      onClick={() => setToolDrawerView("detail")}
                    >
                      ← 返回工具说明
                    </button>
                    <div className="tool-run-heading">
                      <small>正在准备</small>
                      <strong>{tool.name}</strong>
                      <span>
                        我会逐项确认本次运行环境，已匹配的内容保持收起。
                      </span>
                    </div>
                    <div className="tool-run-guide">
                      <section
                        className={`tool-guide-step${
                          primaryParameters.length > 0 && !primaryReady
                            ? " needs-attention"
                            : ""
                        }`}
                      >
                        <span className="tool-guide-number">1</span>
                        <div className="tool-guide-content">
                          <div className="tool-guide-title">
                            <div>
                              <strong>选择数据来源</strong>
                              <span>这次要处理哪张工作表？</span>
                            </div>
                            <small>
                              {primaryParameters.length === 0
                                ? "固定来源"
                                : primaryReady
                                  ? "已完成"
                                  : "待处理"}
                            </small>
                          </div>
                          <div className="tool-guide-fields">
                            {primaryParameters.length > 0 ? (
                              primaryParameters.map((parameter) =>
                                renderToolParameter(tool, parameter)
                              )
                            ) : (
                              <span>使用保存时的固定数据来源。</span>
                            )}
                          </div>
                        </div>
                      </section>
                      <section
                        className={`tool-guide-step${
                          missingFieldCount > 0 ? " needs-attention" : ""
                        }`}
                      >
                        <span className="tool-guide-number">2</span>
                        <div className="tool-guide-content">
                          <div className="tool-guide-title">
                            <div>
                              <strong>核对字段</strong>
                              <span>
                                {fieldParameters.length === 0
                                  ? "这个工具不需要字段匹配"
                                  : missingFieldCount > 0
                                  ? `有 ${missingFieldCount} 个字段需要你指定`
                                  : `${fieldParameters.length} 个字段已按表头自动匹配`}
                              </span>
                            </div>
                            <small>
                              {missingFieldCount > 0 ? "待处理" : "已完成"}
                            </small>
                          </div>
                          {advancedParameters.length > 0 && (
                            <details
                              className="tool-field-mapping"
                              open={missingFieldCount > 0}
                            >
                              <summary>
                                <span>
                                  {missingFieldCount > 0
                                    ? "完成字段匹配"
                                    : "查看字段映射"}
                                </span>
                                <small>
                                  {fieldParameters.length > 0
                                    ? `${
                                        fieldParameters.length -
                                        missingFieldCount
                                      }/${fieldParameters.length} 已匹配`
                                    : `${advancedParameters.length} 项设置`}
                                </small>
                              </summary>
                              <div>
                                <p>
                                  只有字段名称发生变化时才需要调整；数据范围也可以在这里检查。
                                </p>
                                {advancedParameters.map((parameter) =>
                                  renderToolParameter(tool, parameter)
                                )}
                              </div>
                            </details>
                          )}
                        </div>
                      </section>
                      <section className="tool-guide-step">
                        <span className="tool-guide-number">3</span>
                        <div className="tool-guide-content">
                          <div className="tool-guide-title">
                            <div>
                              <strong>生成安全预览</strong>
                              <span>
                                下一步只展示将要执行的内容，不会立即修改工作簿。
                              </span>
                            </div>
                            <small>
                              {missingFieldCount > 0 ? "待处理" : "准备就绪"}
                            </small>
                          </div>
                        </div>
                      </section>
                    </div>
                    <div className="tool-run-actions">
                      <button
                        className="tool-preview-button"
                        disabled={
                          busy || !workbook || missingFieldCount > 0
                        }
                        onClick={() => void previewTool(tool)}
                      >
                        查看执行预览
                      </button>
                      <span>
                        预览确认后才会执行；当前页面不会直接写入 Excel
                      </span>
                    </div>
                  </section>
                );
              })}
        </aside>
      )}

      {pendingToolDeletion && (
        <div className="tool-dialog-backdrop" role="presentation">
          <section
            className="tool-dialog history-delete-dialog"
            role="dialog"
            aria-modal="true"
            aria-label="删除保存的工具"
          >
            <div className="tool-dialog-title">
              <div>
                <strong>删除这个工具？</strong>
                <span>删除后无法恢复，但不会影响工作簿中的数据</span>
              </div>
              <button
                onClick={() => setPendingToolDeletion(null)}
                aria-label="关闭"
              >
                ×
              </button>
            </div>
            <p>确定从工具箱删除「{pendingToolDeletion.name}」吗？</p>
            <div className="tool-dialog-actions">
              <button
                className="secondary-button"
                onClick={() => setPendingToolDeletion(null)}
              >
                取消
              </button>
              <button
                className="danger-button"
                onClick={confirmToolDeletion}
              >
                确认删除
              </button>
            </div>
          </section>
        </div>
      )}

      {historyOpen && (
        <aside className="tool-drawer history-drawer" aria-label="历史对话">
          <div className="tool-drawer-header">
            <div>
              <strong>历史对话</strong>
              <span>新建对话不会覆盖之前的记录</span>
            </div>
            <button onClick={() => setHistoryOpen(false)} aria-label="关闭">
              ×
            </button>
          </div>
          <div className="history-list">
            {[...chatHistory.conversations]
              .sort((left, right) =>
                right.updatedAt.localeCompare(left.updatedAt)
              )
              .map((conversation) => {
                const active =
                  conversation.id === chatHistory.activeConversationId;
                const questionCount = conversation.messages.filter(
                  (message) => message.role === "user"
                ).length;
                return (
                  <article
                    className={`history-item${active ? " active" : ""}`}
                    key={conversation.id}
                  >
                    <button
                      className="history-open-button"
                      onClick={() => handleOpenConversation(conversation.id)}
                      disabled={busy}
                    >
                      <span className="history-title-row">
                        <strong>{conversation.title}</strong>
                        {active && <em>当前</em>}
                      </span>
                      <span>
                        {formatConversationTime(conversation.updatedAt)}
                        {questionCount > 0
                          ? ` · ${questionCount} 条提问`
                          : " · 尚未提问"}
                      </span>
                    </button>
                    <button
                      className="history-delete-button"
                      onClick={() => handleDeleteConversation(conversation.id)}
                      disabled={busy}
                      title={`删除「${conversation.title}」`}
                      aria-label={`删除「${conversation.title}」`}
                    >
                      删除
                    </button>
                  </article>
                );
              })}
          </div>
          <button
            className="history-new-button"
            onClick={handleNewChat}
            disabled={busy}
          >
            ＋ 新建对话
          </button>
        </aside>
      )}

      {pendingDeleteConvId &&
        chatHistory.conversations
          .filter(
            (conversation) =>
              conversation.id === pendingDeleteConvId
          )
          .map((conversation) => (
            <div
              className="tool-dialog-backdrop"
              role="presentation"
              key={conversation.id}
            >
              <section
                className="tool-dialog history-delete-dialog"
                role="dialog"
                aria-modal="true"
                aria-label="删除历史对话"
              >
                <div className="tool-dialog-title">
                  <div>
                    <strong>删除历史对话？</strong>
                    <span>此操作无法撤销</span>
                  </div>
                  <button
                    onClick={cancelDeleteConversation}
                    aria-label="关闭"
                  >
                    ×
                  </button>
                </div>
                <p>
                  确定删除「{conversation.title}」吗？删除后不会影响已保存的工具。
                </p>
                <div className="tool-dialog-actions">
                  <button
                    className="secondary-button"
                    onClick={cancelDeleteConversation}
                  >
                    取消
                  </button>
                  <button
                    className="danger-button"
                    onClick={confirmDeleteConversation}
                  >
                    确认删除
                  </button>
                </div>
              </section>
            </div>
          ))}

      {saveCandidate && (
        <SaveToolDialog
          toolName={toolName}
          toolDescription={toolDescription}
          eligibility={saveEligibility}
          approveFixedContent={approveFixedContent}
          approveDestructive={approveDestructive}
          onToolNameChange={setToolName}
          onToolDescriptionChange={setToolDescription}
          onApproveFixedContentChange={setApproveFixedContent}
          onApproveDestructiveChange={setApproveDestructive}
          onCancel={closeSaveCandidate}
          onConfirm={confirmSaveTool}
          subtitle="验证后可从“我的工具”重复使用"
          note="只保存白名单 Excel 操作，不会保存 API Key，也不会执行系统命令。"
        />
      )}

      <div className={`message-stream-shell${newConversationNotice ? " conversation-starting" : ""}`}>
      {newConversationNotice && (
        <div className="new-conversation-toast" role="status">
          {newConversationNotice.previousSaved ? "已新建对话" : "当前已是新对话"}
        </div>
      )}
      <section
        ref={messageStreamRef}
        className="message-stream"
        aria-live="polite"
        onScroll={handleMessageStreamScroll}
      >
        {messages.map((message) => (
          <article
            className={`message-row ${message.role}${
              message.staleSourceConfirm ? " stale-source-message" : ""
            }`}
            key={message.id}
            data-message-id={message.id}
          >
            <div className="message-content">
              {message.role === "assistant" && (
                <span className="message-author">
                  Excel Bro
                  {message.provider && (
                    <em>
                      {message.clarification
                        ? "需求确认"
                        : message.provider === "model"
                          ? "Agent"
                          : message.provider === "system"
                            ? "系统"
                            : "本地工具"}
                    </em>
                  )}
                </span>
              )}
              {message.reused && (
                <div className="reuse-badge">
                  <span>♻ 复用上次结果（数据未变化）</span>
                  <button
                    type="button"
                    className="link-button"
                    disabled={busy}
                    onClick={() => recomputeReusedMessage(message)}
                  >
                    仍要重新计算
                  </button>
                </div>
              )}
              {message.staleSourceConfirm ? (
                <section
                  className="stale-source-warning"
                  aria-label="数据源变化提示"
                >
                  <span className="stale-source-warning-icon" aria-hidden="true">
                    !
                  </span>
                  <div className="stale-source-warning-copy">
                    <strong>数据源已发生变化</strong>
                    <p>
                      这份计划里有 {message.staleSourceConfirm.bakedActionCount} 处写入
                      使用了预览时计算的固定值。继续执行会写入旧结果，请先检查下方计划预览。
                    </p>
                  </div>
                </section>
              ) : message.scopeSelectionGuidance ? (
                <ScopeSelectionGuidance {...message.scopeSelectionGuidance} />
              ) : message.text &&
                message.intentMemory?.toolRequest &&
                message.resultContext ? (
                <section className="query-result-card">
                  <p className="message-text">{message.text}</p>
                  <div className="query-result-actions">
                    <button
                      type="button"
                      className="secondary-button query-result-save-button"
                      disabled={busy}
                      onClick={() => saveQueryFromMessage(message)}
                    >
                      保存为固化查询
                    </button>
                  </div>
                </section>
              ) : (
                message.text && !message.planningResume &&
                !prePlanningPause.entries.some((entry) => entry.messageId === message.id) &&
                <p className="message-text">{message.text}</p>
              )}
              {message.folderResume && (
                <section className="folder-resume-card" aria-label="继续未完成的文件夹任务">
                  <div className="folder-resume-copy">
                    <strong>可从安全断点继续</strong>
                    <span>先核对文件状态；已完成的会跳过，外部修改会立即停止。</span>
                  </div>
                  <div className="folder-resume-actions">
                    <button
                      type="button"
                      className="primary-button"
                      disabled={busy || message.folderResume.status === "resuming"}
                      onClick={() => void resumeInterruptedFolderOperation(message)}
                    >
                      {message.folderResume.status === "resuming"
                        ? "正在继续..."
                        : message.folderResume.status === "failed"
                          ? "重新核对并继续"
                          : "继续剩余任务"}
                    </button>
                    <button
                      type="button"
                      className="secondary-button"
                      disabled={busy}
                      onClick={() => {
                        setContextOpen(false);
                        closeAllDrawers();
                        setFolderBackupOpen(true);
                      }}
                    >
                      查看备份
                    </button>
                  </div>
                </section>
              )}
              {message.officeResume && (
                <section
                  className="folder-resume-card office-resume-card"
                  aria-label="继续已保存的 Office 写入任务"
                >
                  <div className="folder-resume-copy">
                    <strong>
                      {message.officeResume.requiresReconciliation
                        ? "需要先核对一个在途写入"
                        : "可从 Office 动作断点继续"}
                    </strong>
                    <span>{message.officeResume.reason}</span>
                    <span>
                      已确认到第 {message.officeResume.nextActionIndex} / {message.officeResume.totalActions} 步；
                      只支持已保存工作簿中的受控区域写入。
                    </span>
                  </div>
                  <div className="folder-resume-actions">
                    <button
                      type="button"
                      className="primary-button"
                      disabled={busy || message.officeResume.status === "resuming"}
                      onClick={() => void resumeInterruptedOfficeExecution(message)}
                    >
                      {message.officeResume.status === "resuming"
                        ? "正在核对..."
                        : message.officeResume.requiresReconciliation
                          ? "核对并继续"
                          : "继续剩余步骤"}
                    </button>
                  </div>
                </section>
              )}
              {!message.activityLog && prePlanningPause.entries.filter((entry) => entry.messageId === message.id &&
                entry.conversationId === chatHistory.activeConversationId).map((entry) =>
                <PrePlanningPauseCard key={entry.id} busy={busy} error={entry.error}
                  onContinue={() => void restartEarlyAnalysis(entry.id)} />)}
              {message.planningResume && (
                <PauseDetails blocked={message.planningResume.status === "failed"} active={message.planningResume.status === "resuming"}>
                <section
                  className="folder-resume-card planning-resume-card"
                  aria-label="继续已保存的规划任务"
                >
                  <div className="folder-resume-copy">
                    {message.text && <p>{message.text}</p>}
                    <strong>
                      {message.planningResume.status === "resuming"
                        ? message.planningResume.awaitingStatus ? "正在核对原任务" : "正在从保存的进度继续"
                        : message.planningResume.status === "failed" ? "任务不能继续，需要重新开始"
                        : message.planningResume.awaitingStatus ? "已暂停接收，待核对原任务"
                        : message.planningResume.pauseType === "model_error"
                          ? "模型连接中断，任务现场已保存"
                          : message.planningResume.pauseType === "interrupted" ? "任务已暂停，可继续"
                          : message.planningResume.pauseType === "budget_exhausted"
                            ? "本轮有效运行总时长已用完"
                          : "本轮处理已暂停"}
                    </strong>
                    <span>{message.planningResume.reason}</span>
                    <span>
                      {message.planningResume.awaitingStatus
                        ? "只查询原任务；尚未确认存在可继续的现场，不会自动重新规划或写入。"
                        : `已完成约 ${Math.round(message.planningResume.activeElapsedSeconds)} 秒有效处理；将按当前授权重新读取并分析，已确认的口径会保留。`}
                    </span>
                    {(message.planningResume.pauseType === "model_error" ||
                      message.planningResume.pauseType === "interrupted") &&
                      message.planningResume.currentModelId && (
                        <span>当前连接：{message.planningResume.currentModelId}</span>
                      )}
                  </div>
                  <div className="folder-resume-actions">
                    <button
                      type="button"
                      className="primary-button"
                      disabled={
                        busy || message.planningResume.status === "resuming" || message.planningResume.status === "failed"
                      }
                      onClick={() => void resumePlanningTaskFromMessage(message)}
                    >
                      {message.planningResume.status === "resuming"
                        ? message.planningResume.awaitingStatus ? "正在核对..." : "正在继续..."
                        : message.planningResume.awaitingStatus ? "核对任务状态"
                        : (message.planningResume.pauseType === "model_error" ||
                              message.planningResume.pauseType === "interrupted") &&
                            effectiveModelId &&
                            effectiveModelId !== message.planningResume.currentModelId
                          ? `使用「${selectedModel?.label ?? effectiveModelId}」从断点继续`
                          : message.planningResume.pauseType === "model_error"
                            ? "重试当前模型"
                          : message.planningResume.pauseType === "budget_exhausted"
                            ? "重新发起任务"
                          : message.planningResume.extensionSeconds > 0
                          ? `继续分析（约 ${Math.round(
                              message.planningResume.extensionSeconds
                            )} 秒）`
                          : "继续分析"}
                    </button>
                    {(message.planningResume.pauseType === "model_error" ||
                      message.planningResume.pauseType === "interrupted") &&
                      message.planningResume.canSwitchModel &&
                      (message.planningResume.availableModelConnections?.length ?? 0) > 0 && (
                        <>
                          <label className="planning-fallback-field">
                            <span>备用模型</span>
                            <select
                              value={
                                message.planningResume.selectedFallbackModelId ??
                                message.planningResume.availableModelConnections?.[0]
                                  ?.catalogModelId ??
                                ""
                              }
                              disabled={
                                busy || message.planningResume.status === "resuming"
                              }
                              onChange={(event) =>
                                updateFolderResumeMessage(message.id, (current) => ({
                                  ...current,
                                  planningResume: current.planningResume
                                    ? {
                                        ...current.planningResume,
                                        selectedFallbackModelId: event.target.value,
                                        requestId: undefined
                                      }
                                    : undefined
                                }))
                              }
                            >
                              {message.planningResume.availableModelConnections?.map(
                                (connection) => (
                                  <option
                                    key={connection.catalogModelId}
                                    value={connection.catalogModelId}
                                  >
                                    {modelOptions.find((model) => model.id === connection.catalogModelId)?.label ?? `${connection.label} · ${connection.modelId}`}
                                  </option>
                                )
                              )}
                            </select>
                          </label>
                          <button
                            type="button"
                            className="secondary-button"
                            disabled={
                              busy ||
                              message.planningResume.status === "resuming" ||
                              !message.planningResume.selectedFallbackModelId
                            }
                            onClick={() =>
                              void resumePlanningTaskFromMessage(
                                message,
                                message.planningResume?.selectedFallbackModelId
                              )
                            }
                          >
                            使用备用模型继续
                          </button>
                        </>
                      )}
                  </div>
                </section></PauseDetails>
              )}
              {message.attachmentNames &&
                message.attachmentNames.length > 0 && (
                  <div className="message-attachments">
                    {message.attachmentNames.map((name, index) => (
                      <span key={`${message.id}-attachment-${index}`}>
                        ▧ {name}
                      </span>
                    ))}
                  </div>
                )}
              {message.activityPending &&
                activity?.persistenceKey === message.id &&
                activity.steps.length > 0 && (
                  <TaskProgressTimeline
                    live
                    log={{
                      timelineStartedAt: activity.timelineStartedAt ?? activity.startedAt,
                      pausedMs: activity.pausedMs, waitingMs: activity.waitingMs,
                      excludedSince: activity.excludedSince, excludedKind: activity.excludedKind,
                      operationId: activity.operationId,
                      steps: activity.steps,
                      traces: activity.traces,
                      totalMs: activitySeconds * 1000,
                      connectionState: activity.connectionState,
                      connectionMessage: activity.connectionMessage
                    }}
                    /*
                     * 卡片上的「暂停」与输入框按钮读**同一份判据**：只有真的能停时才给入口。
                     * 只按 `status` 判断的话，状态已经是 planning、但本端没有任何可中止对象
                     * 的窗口里会画出一个点不动的按钮。
                     */
                    onStop={
                      cardPauseButton.canPause &&
                      (status === "planning" || status === "tooling")
                        ? stopTurn
                        : undefined
                    }
                    segment={{
                      live: true,
                      continuesTask: message.replySegment?.continuesTask
                    }}
                  />
                )}
              {message.activityLog &&
                !(
                  message.activityPending &&
                  activity?.persistenceKey === message.id
                ) &&
                message.activityLog.steps.length > 0 && (
                  <TaskProgressTimeline log={message.activityLog}
                    segment={{
                      continuesTask: message.replySegment?.continuesTask,
                      ...(message.replySegment?.nextSegmentId
                        ? {
                            onShowLater: () =>
                              jumpToMessage(message.replySegment!.nextSegmentId!),
                            laterLabel: "查看后续进展"
                          }
                        : {})
                    }}
                    onContinue={
                      // 已被后续段接管的段是历史：不再给"继续"入口。留着一个还能点的
                      // 按钮，等于让用户对一个已经继续过的任务再点一次。
                      message.replySegment?.status === "superseded"
                        ? undefined
                        : prePlanningPause.entries.some((entry) => entry.messageId === message.id && entry.conversationId === chatHistory.activeConversationId)
                          ? () => void restartEarlyAnalysis(`preplanning-${message.id}`) : undefined
                    }
                    continueBusy={busy}
                    continuationError={prePlanningPause.entries.find((entry) => entry.messageId === message.id)?.error} />
                )}
              {message.clarification && (
                <section
                  className={`clarification-card ${message.clarification.status}`}
                  aria-label="确认需求"
                >
                  <div className="clarification-heading">
                    <div>
                      <span>执行前确认</span>
                      <strong>
                        {message.clarification.status === "pending"
                          ? "需要你确认需求"
                          : message.clarification.status === "resolving"
                            ? "正在理解你的补充"
                          : message.clarification.status === "resolved"
                            ? "需求已确认"
                            : message.clarification.status === "invalidated"
                              ? "数据范围已变化"
                              : "需求已取消"}
                      </strong>
                    </div>
                    <em>{message.clarification.scopeLabel}</em>
                  </div>
                  <p className="clarification-summary">
                    {message.clarification.summary}
                  </p>
                  <strong className="clarification-question">
                    {message.clarification.question}
                  </strong>
                  <span className="clarification-reason">
                    {message.clarification.reason}
                  </span>

                  {message.clarification.status === "pending" ? (
                    message.clarification.localMapping ? (
                      localMappingDraftsRef.current.get(message.id) ? (
                        <LocalMappingResolution
                          key={`${message.id}:${localMappingDraftsRef.current.get(message.id)!.revision}`}
                          rule={localMappingDraftsRef.current.get(message.id)!.rule}
                          values={localMappingDraftsRef.current.get(message.id)!.values}
                          allowExactMapping={localMappingDraftsRef.current.get(message.id)!.allowExactMapping}
                          reason={localMappingDraftsRef.current.get(message.id)!.reason}
                          blankCount={localMappingDraftsRef.current.get(message.id)!.blankCount}
                          busy={busy}
                          checkScope={() => checkExcelLocalMappingScope(message)}
                          onReauthorize={localMappingDraftsRef.current.get(message.id)!.pendingRule
                            ? () => resolveExcelLocalMapping(message,
                                localMappingDraftsRef.current.get(message.id)!.pendingRule!, true)
                            : undefined}
                          onResolve={(rule) => resolveExcelLocalMapping(message, rule)}
                          onCancel={() => cancelClarification(message)}
                        />
                      ) : <p>本地规则草稿已失效，请重新发送需求。</p>
                    ) : <>
                      <div className="clarification-options">
                        {message.clarification.options.map(
                          (option: IntentOption) => (
                            <button
                              key={option.id}
                              disabled={busy}
                              onClick={() => {
                                if (option.action === "editScope") {
                                  editClarificationScope(message);
                                  return;
                                }
                                void resolveClarification(
                                  message,
                                   option.resolution,
                                   option.label,
                                   true,
                                   option.resolutionMode
                                 );
                              }}
                            >
                              <strong>{option.label}</strong>
                              <span>{option.description}</span>
                            </button>
                          )
                        )}
                      </div>
                      {message.clarification.fields?.length ? (
                        <ClarificationFields fields={message.clarification.fields}
                          values={clarificationFieldDrafts[message.id] ?? {}} busy={busy}
                          onChange={(values) => setClarificationFieldDrafts((current) => ({ ...current, [message.id]: values }))}
                          onSubmit={(answer) => void resolveClarification(message, answer, answer)} />
                      ) : <>
                      <span className="clarification-custom-hint">
                        都不符合？请直接写下你的想法。
                      </span>
                      <div className="clarification-custom-input">
                        <textarea
                          value={
                            clarificationDrafts[
                              message.clarification.id
                            ] ?? ""
                          }
                          disabled={busy}
                          rows={2}
                          placeholder="例如：按所有已选工作表汇总后，再比较每个分类的整体占比"
                          onChange={(event) =>
                            setClarificationDrafts((current) => ({
                              ...current,
                              [message.clarification!.id]:
                                event.target.value
                            }))
                          }
                        />
                        <button
                          disabled={
                            busy ||
                            !(
                              clarificationDrafts[
                                message.clarification.id
                              ] ?? ""
                            ).trim()
                          }
                          onClick={() => {
                            const customAnswer = (
                              clarificationDrafts[
                                message.clarification!.id
                              ] ?? ""
                            ).trim();
                            void resolveClarification(
                              message,
                              customAnswer,
                              customAnswer
                            );
                          }}
                        >
                          提交补充
                        </button>
                      </div>
                      </>}
                      <div className="clarification-actions">
                        <button
                          disabled={busy}
                          onClick={() => editClarificationScope(message)}
                        >
                          修改数据范围
                        </button>
                        {!message.clarification.fields?.length && <button
                          disabled={busy}
                          onClick={() =>
                            void resolveClarification(
                              message,
                              "按用户原始需求继续；若仍需自行判断，请在结果中明确列出采用的假设。",
                              "按原话继续",
                              true
                            )
                          }
                        >
                          按原话继续
                        </button>}
                        <button
                          disabled={busy}
                          onClick={() => cancelClarification(message)}
                        >
                          取消
                        </button>
                      </div>
                    </>
                  ) : (
                    <div className="clarification-outcome">
                      {message.clarification.status === "resolving"
                        ? "正在结合你的回答重新判断，不会执行任何写入。"
                        : message.clarification.status === "resolved"
                        ? `已选择：${message.clarification.resolvedLabel}`
                        : message.clarification.status === "invalidated"
                          ? "请检查数据范围后重新发送需求。"
                          : "没有进行后续分析或写入。"}
                    </div>
                  )}
                </section>
              )}
              {message.localResult && (
                <LocalResultCard
                  card={message.localResult}
                  load={(resultRef) =>
                    message.localResultBinding
                      ? loadBridgeResult(message.localResultBinding, resultRef)
                      : Promise.resolve(null)
                  }
                />
              )}
              {(message.verification || message.executedPlanId || message.executedExecutionId) && (
                <ExecutionResultCard
                  tone={message.executionUndone ? "undone" : message.verification?.status ?? "unverified"}
                  title={message.executionUndone ? "已撤销并恢复" : !message.verification ? "执行记录" : message.verification.status === "verified"
                      ? VERIFICATION_PASSED_LABEL
                      : message.verification.status === "executed_unverified"
                        ? "已执行，部分操作未独立验证"
                        : "执行完成，但验证未通过"}
                  summary={message.executionUndone
                    ? /* 已撤销：按模式说清恢复了什么 —— 文件夹模式恢复的是文件，
                         当前工作簿模式才是单元格的内容与格式。不混用同一句笼统描述。 */
                      message.executedUndoId
                        ? "已把本次执行影响的文件恢复到执行前状态。"
                        : "已恢复本次执行记录的内容和格式。"
                    : message.verification ? <>
                    {
                      message.verification.checks.filter((check) => check.passed)
                        .length
                    }
                    /{message.verification.checks.length} 项符合预期
                    {message.verification.unverifiedActions.length > 0
                      ? `；${message.verification.unverifiedActions.length} 步缺少独立验收`
                      : ""}
                  </> : "详细状态见执行过程"}
                >
                  {!message.executionUndone && message.verification?.status === "failed" && (
                    <ul>
                      {message.verification.checks
                        .filter((check) => !check.passed)
                        .map((check, index) => (
                          <li key={`${message.id}-verification-${index}`}>
                            {check.message}
                          </li>
                        ))}
                    </ul>
                  )}
                  {!message.executionUndone && message.verification?.status ===
                    "executed_unverified" && (
                    <ul>
                      {message.verification.unverifiedActions.map(
                        (action) => (
                          <li
                            key={`${message.id}-unverified-${action.index}`}
                          >
                            {action.message}
                          </li>
                        )
                      )}
                    </ul>
                  )}
              {!message.executionUndone && message.executedExecutionId &&
                changeReviewsRef.current.has(message.executedExecutionId) && (
                  <ChangeReviewSummaryCard
                    reviews={
                      changeReviewsRef.current.get(
                        message.executedExecutionId
                      ) ?? []
                    }
                  />
                )}
              {message.executedPlanId &&
                (() => {
                  const workbookIndex = undoStack.findIndex((entry) =>
                    message.executedExecutionId
                      ? entry.executionId === message.executedExecutionId
                      : entry.planId === message.executedPlanId
                  );
                  const folderIndex = folderUndoStack.findIndex((entry) =>
                    message.executedUndoId
                      ? entry.undoId === message.executedUndoId
                      : entry.planId === message.executedPlanId
                  );
                  if (message.executionUndone || (workbookIndex < 0 && folderIndex < 0)) {
                    return <div className="execution-recovery-footer">
                      <span>{message.executionUndone
                        ? "本次执行已撤销"
                        : "当前没有可用的撤销记录；这不表示执行已撤销。"}</span>
                    </div>;
                  }
                  const isFolderUndo = folderIndex >= 0;
                  const workbookUndoEntry =
                    workbookIndex >= 0 ? undoStack[workbookIndex] : null;
                  const workbookUndoReady =
                    isFolderUndo ||
                    workbookUndoEntry?.postStateCoverage === "complete";
                  const activeStackLength = isFolderUndo
                    ? folderUndoStack.length
                    : undoStack.length;
                  const stackIndex = isFolderUndo
                    ? folderIndex
                    : workbookIndex;
                  // 栈顶之后还有多少次执行会被连带撤销。
                  const alsoUndone = activeStackLength - 1 - stackIndex;
                  return (
                    <div className="undo-row execution-undo-row execution-recovery-footer">
                      <button
                        type="button"
                        className="secondary-button"
                        disabled={busy || !workbookUndoReady}
                        onClick={() =>
                          void (isFolderUndo
                            ? undoFolderExecutionTo(
                                message.executedUndoId ?? message.executedPlanId
                              )
                            : undoExecutionTo(
                                message.executedExecutionId ??
                                  message.executedPlanId
                              ))
                        }
                      >
                        {!workbookUndoReady
                          ? "↶ 撤销需核对"
                          : alsoUndone > 0
                          ? `↶ 撤销本次执行（含之后 ${alsoUndone} 次）`
                          : "↶ 撤销本次执行"}
                      </button>
                      {isFolderUndo && (
                        <button
                          type="button"
                          className="secondary-button"
                          disabled={busy}
                          onClick={() => {
                            setContextOpen(false);
                            closeAllDrawers();
                            setFolderBackupOpen(true);
                          }}
                        >
                          查看备份
                        </button>
                      )}
                      <span className="undo-row-hint">
                        {!workbookUndoReady
                          ? "执行后状态未能完整核对，已停止自动覆盖"
                          : alsoUndone > 0
                          ? "多次执行可能修改同一文件，只能从最近一次往前逐级撤销"
                          : isFolderUndo
                            ? "恢复到执行前状态；撤销前会检查文件是否又被修改"
                            : `保留最近 ${undoStack.length} 次执行的撤销数据`}
                      </span>
                    </div>
                  );
                })()}
                </ExecutionResultCard>
              )}
              {message.functionPreview &&
                message.functionPreview.phase === "target" && (
                  <div className="inline-plan">
                    <div className="inline-plan-title">
                      <div>
                        <span>写入单元格内</span>
                        <strong>{message.functionPreview.sheet}</strong>
                      </div>
                    </div>
                    <p className="function-preview-explain">
                      确定公式要写到哪个单元格或区域，再生成。已为你预填建议位置，可直接改，或先在表里点选目标格、再点「拾取当前选区」。
                    </p>
                    <div className="function-write-target">
                      <input
                        className="function-write-target-input"
                        type="text"
                        value={message.functionPreview.writeTarget}
                        disabled={busy || message.functionPreview.pickingTarget}
                        placeholder="如 E2 或 E2:E20"
                        spellCheck={false}
                        onChange={(event) =>
                          markFunctionPreview(message.id, {
                            writeTarget: event.target.value,
                            targetError: undefined
                          })
                        }
                      />
                      <button
                        className="secondary-button"
                        disabled={busy}
                        onClick={() => void pickFunctionTarget(message)}
                      >
                        {message.functionPreview.pickingTarget
                          ? "读取中..."
                          : "拾取当前选区"}
                      </button>
                    </div>
                    <div className="inline-notes function-pick-hint">
                      拾取方式：先在工作表里点一下（或框选）目标单元格，再点上面的「拾取当前选区」，会自动填入那一刻选中的位置。
                    </div>
                    {message.functionPreview.match && (
                      <div className="function-match-proposal">
                        <strong>
                          跨表匹配 · 本地生成公式（不调用 AI，数据不出本机）
                        </strong>
                        <div className="function-match-row">
                          <span>外部表</span>
                          <em>
                            {message.functionPreview.match.externalFile} ›{" "}
                            {message.functionPreview.match.externalSheet}
                          </em>
                        </div>
                        <div className="function-match-row">
                          <label>
                            匹配键列
                            <select
                              value={
                                message.functionPreview.match.selectedKey
                              }
                              disabled={busy}
                              onChange={(event) => {
                                const match = message.functionPreview?.match;
                                if (!match) return;
                                markFunctionPreview(message.id, {
                                  match: {
                                    ...match,
                                    selectedKey: event.target.value
                                  }
                                });
                              }}
                            >
                              {message.functionPreview.match.keyCandidates.map(
                                (candidate) => (
                                  <option
                                    key={candidate.name}
                                    value={candidate.name}
                                  >
                                    {candidate.name}
                                  </option>
                                )
                              )}
                            </select>
                          </label>
                          <label>
                            取值列
                            <select
                              value={
                                message.functionPreview.match.selectedValue
                              }
                              disabled={busy}
                              onChange={(event) => {
                                const match = message.functionPreview?.match;
                                if (!match) return;
                                markFunctionPreview(message.id, {
                                  match: {
                                    ...match,
                                    selectedValue: event.target.value
                                  }
                                });
                              }}
                            >
                              {message.functionPreview.match.valueCandidates.map(
                                (candidate) => (
                                  <option
                                    key={candidate.name}
                                    value={candidate.name}
                                  >
                                    {candidate.name}
                                  </option>
                                )
                              )}
                            </select>
                          </label>
                        </div>
                      </div>
                    )}
                    {message.functionPreview.targetError && (
                      <div className="function-write-target-error">
                        {message.functionPreview.targetError}
                      </div>
                    )}
                    <div className="inline-plan-buttons">
                      <button
                        className="secondary-button"
                        disabled={busy}
                        onClick={() => cancelFunctionPreview(message)}
                      >
                        取消
                      </button>
                      <button
                        className="primary-button"
                        disabled={busy}
                        onClick={() => void confirmFunctionTarget(message)}
                      >
                        {status === "planning" ? "正在生成..." : "确定并生成"}
                      </button>
                    </div>
                  </div>
                )}
              {message.functionPreview &&
                message.functionPreview.phase === "preview" && (
                  <div className="inline-plan">
                    <div className="inline-plan-title">
                      <div>
                        <span>公式预览</span>
                        <strong>{message.functionPreview.writeTarget}</strong>
                        {message.functionPreview.mode === "deterministic" && (
                          <em className="function-local-badge">
                            本地生成 · 未调用 AI
                          </em>
                        )}
                      </div>
                      {message.functionPreview.generateMs !== undefined && (
                        <span className="function-preview-timing">
                          生成耗时 {formatGenerateMs(message.functionPreview.generateMs)}
                        </span>
                      )}
                    </div>

                    {!message.functionPreview.applied &&
                      !message.functionPreview.cancelled && (
                        <div className="function-preview-versions">
                          <button
                            className={
                              message.functionPreview.version === "compat"
                                ? "version-tab active"
                                : "version-tab"
                            }
                            onClick={() =>
                              markFunctionPreview(message.id, {
                                version: "compat"
                              })
                            }
                          >
                            兼容版 · 2016/2019
                          </button>
                          <button
                            className={
                              message.functionPreview.version === "modern"
                                ? "version-tab active"
                                : "version-tab"
                            }
                            onClick={() =>
                              markFunctionPreview(message.id, {
                                version: "modern"
                              })
                            }
                          >
                            现代版 · 365/2021
                          </button>
                        </div>
                      )}

                    <pre className="function-preview-formula">
                      {message.functionPreview.version === "modern"
                        ? message.functionPreview.modernFormula
                        : message.functionPreview.compatFormula}
                    </pre>

                    <div className="function-preview-trial">
                      <span>首格试算</span>
                      <strong>
                        {message.functionPreview.version === "modern"
                          ? message.functionPreview.modernResult
                          : message.functionPreview.compatResult}
                      </strong>
                    </div>

                    {(() => {
                      const preview = message.functionPreview;
                      if (!preview) return null;
                      const formula =
                        preview.version === "modern"
                          ? preview.modernFormula
                          : preview.compatFormula;
                      const trial =
                        preview.version === "modern"
                          ? preview.modernResult
                          : preview.compatResult;
                      const refs = externalWorkbookRefs(formula);
                      if (refs.length === 0) return null;
                      const unavailable = /#(?:REF|NAME|VALUE)/.test(trial);
                      return (
                        <>
                          <div className="function-external-refs">
                            <span>本公式引用外部工作簿：</span>
                            {refs.map((ref) => (
                              <em
                                key={`${ref.file}\u0000${ref.sheet}`}
                              >
                                {ref.file} › {ref.sheet}
                              </em>
                            ))}
                          </div>
                          {unavailable && (
                            <p className="function-external-hint">
                              外部工作簿未在 Excel 中打开，打开后会自动重算。
                            </p>
                          )}
                        </>
                      );
                    })()}

                    {(() => {
                      const reason = functionWriteBlockReason(
                        message.functionPreview
                      );
                      return reason ? (
                        <p className="function-verification-block">
                          {reason}
                        </p>
                      ) : null;
                    })()}

                    {(message.functionPreview.version === "modern"
                      ? message.functionPreview.modernExplanation
                      : message.functionPreview.compatExplanation) && (
                      <p className="function-preview-explain">
                        {message.functionPreview.version === "modern"
                          ? message.functionPreview.modernExplanation
                          : message.functionPreview.compatExplanation}
                      </p>
                    )}

                    {message.functionPreview.applied ? (
                      <>
                        <div className="inline-notes">
                          已写入 {message.functionPreview.appliedTarget}。
                          下拉填充后若发现公式不对，可用下面的按钮重新生成。
                        </div>
                        <div className="inline-plan-buttons">
                          <button
                            className="secondary-button"
                            disabled={busy}
                            onClick={() => void switchToModelFunction(message)}
                          >
                            换 AI 重新生成
                          </button>
                        </div>
                      </>
                    ) : message.functionPreview.cancelled ? (
                      <div className="inline-notes">已取消。</div>
                    ) : (
                      <>
                        <div className="function-write-target">
                          <input
                            className="function-write-target-input"
                            type="text"
                            value={message.functionPreview.writeTarget}
                            disabled={
                              busy || message.functionPreview.pickingTarget
                            }
                            placeholder="如 E2 或 E2:E20"
                            spellCheck={false}
                            onChange={(event) =>
                              markFunctionPreview(message.id, {
                                writeTarget: event.target.value,
                                targetError: undefined
                              })
                            }
                          />
                          <button
                            className="secondary-button"
                            disabled={busy}
                            onClick={() => void pickFunctionTarget(message)}
                          >
                            {message.functionPreview.pickingTarget
                              ? "读取中..."
                              : "拾取当前选区"}
                          </button>
                        </div>
                        {message.functionPreview.targetError && (
                          <div className="function-write-target-error">
                            {message.functionPreview.targetError}
                          </div>
                        )}
                        <div className="inline-plan-buttons">
                          <button
                            className="secondary-button"
                            disabled={busy}
                            onClick={() => {
                              const preview = message.functionPreview;
                              if (!preview) return;
                              const formula =
                                preview.version === "modern"
                                  ? preview.modernFormula
                                  : preview.compatFormula;
                              if (formula) {
                                void copyFunctionFormula(message.id, formula);
                              }
                            }}
                          >
                            {copiedFunctionPreviewId === message.id
                              ? "已复制"
                              : "复制公式"}
                          </button>
                          {message.functionPreview.mode === "deterministic" && (
                            <button
                              className="secondary-button"
                              disabled={busy}
                              onClick={() => void switchToModelFunction(message)}
                            >
                              换 AI 生成
                            </button>
                          )}
                          <button
                            className="secondary-button"
                            disabled={busy}
                            onClick={() => cancelFunctionPreview(message)}
                          >
                            取消
                          </button>
                          <button
                            className="primary-button"
                            disabled={
                              busy ||
                              functionWriteBlockReason(
                                message.functionPreview
                              ) !== null
                            }
                            onClick={() => void applyFunctionPreview(message)}
                          >
                            {status === "executing" ? "正在写入..." : "确认写入"}
                          </button>
                        </div>
                      </>
                    )}
                  </div>
                )}
              {message.plan && (
                <div
                  className={`inline-plan${
                    message.staleSourceConfirm ? " stale-source-plan" : ""
                  }`}
                >
                  <div className="inline-plan-title">
                    <div>
                      {/*
                        删表成功之后，这条计划已经不是"待执行"了。再挂着「准备就绪」
                        会和下面的「工作表删除已完成」互相打脸 —— 计划内容仍留在折叠
                        明细里，只是不再冒充一个还能执行的预览。
                      */}
                      {!worksheetCheckpoint.entryForMessage(message.id, chatHistory.activeConversationId)?.result && (
                        <span>准备就绪</span>
                      )}
                      <strong>{message.plan.title}</strong>
                    </div>
                  </div>

                  {message.plan.summary && (
                    <p className="inline-plan-summary">{message.plan.summary}</p>
                  )}
                  <p className="inline-plan-scale">
                    {message.plan.actions.length} 个动作 ·{" "}
                    {message.plan.actions
                      .slice(0, 3)
                      .map((action) => actionLabel(action))
                      .join("；")}
                    {message.plan.actions.length > 3
                      ? `；等 ${message.plan.actions.length} 步`
                      : ""}
                  </p>

                  <div className="inline-plan-buttons">
                    {/*
                      未通过验证的「验证后保存」只是执行前的引导。这次执行已经有回执
                      （删表尤其如此）之后，它就是一个永远不会亮起来的假入口，收起来。
                      通过验证的计划仍保留「保存为工具」。
                    */}
                    {(verifiedPlanIds.has(message.plan.id) ||
                      !worksheetCheckpoint.entryForMessage(message.id, chatHistory.activeConversationId)?.result) && (
                      <button
                        className="secondary-button"
                        disabled={busy || !verifiedPlanIds.has(message.plan.id)}
                        title={
                          verifiedPlanIds.has(message.plan.id)
                            ? "保存为可重复使用的个人工具"
                            : "请先执行，并通过结果验证"
                        }
                        onClick={() => beginSaveTool(message.plan!)}
                      >
                        {verifiedPlanIds.has(message.plan.id)
                          ? "保存为工具"
                          : "验证后保存"}
                      </button>
                    )}
                    {!message.staleSourceConfirm && !isWorksheetDeletion(message.plan) && nativeComparison.review?.messageId !== message.id && (sourceMode !== "workbook" || !message.plan.actions.some(action => action.type === "copyRange")) && (sourceMode !== "workbook" || !workbook || resultReviewPlanIssue(message.plan, selectedNamesFor(workbook)) !== null) && (
                      <button className="primary-button" disabled={busy} onClick={() => void runPlan(message.plan!)}>
                        {status === "executing" && activePlanId === message.plan.id ? "正在执行..." : "执行到 Excel"}
                      </button>
                    )}
                  </div>
                  {sourceMode === "workbook" && workbook && !message.staleSourceConfirm && message.plan.actions.some(action => action.type === "copyRange") &&
                    resultReviewPlanIssue(message.plan, selectedNamesFor(workbook)) !== null && (
                      <p role="status">{resultReviewPlanIssue(message.plan, selectedNamesFor(workbook))}</p>
                    )}
                  {!resultMessages.current.has(message.id) && !message.staleSourceConfirm && (nativeComparison.review?.messageId === message.id ||
                    (sourceMode === "workbook" && workbook && !isWorksheetDeletion(message.plan) && resultReviewPlanIssue(message.plan, selectedNamesFor(workbook)) === null)) && (
                    <ExcelNativeReviewActions messageId={message.id} plan={message.plan}
                      paneNeedsResize={comparisonPaneNeedsResize}
                      authorization={workbook ? selectedNamesFor(workbook) : []} controller={nativeComparison} disabled={busy}
                      beforeOperation={targetGuard.check}
                      commit={session => session
                        ? runPlan(message.plan!, undefined, false, false, session)
                        : runPlan(message.plan!, undefined, false, false, undefined, {
                          plan: message.plan!, authorization: workbook ? selectedNamesFor(workbook) : [],
                          workbookIdentity: currentOfficeWorkbookIdentity() ?? ""
                        })} />
                  )}

                  <details
                    className="plan-details"
                    open={Boolean(message.staleSourceConfirm)}
                  >
                    <summary>
                      查看执行细节
                      <span>{message.plan.actions.length} 步</span>
                    </summary>

                    {message.plan.assumptions.length > 0 && (
                      <div className="inline-notes">
                        <strong>{PLAN_NOTE_SECTIONS.assumptions}</strong>
                        <ul>
                          {message.plan.assumptions.map((item) => (
                            <li key={item}>{item}</li>
                          ))}
                        </ul>
                      </div>
                    )}

                    {message.plan.warnings.length > 0 && (
                      <div className="inline-notes warning">
                        <strong>{PLAN_NOTE_SECTIONS.warnings}</strong>
                        <ul>
                          {message.plan.warnings.map((item) => (
                            <li key={item}>{item}</li>
                          ))}
                        </ul>
                      </div>
                    )}

                    <ol className="inline-actions">
                      {message.plan.actions.map((action, index) => (
                        <li key={`${message.id}-${action.type}-${index}`}>
                          <span>{index + 1}</span>
                          {actionLabel(action)}
                        </li>
                      ))}
                    </ol>

                    {planFormatPreview(message.plan).length > 0 && (
                      <div className="inline-notes format-preview">
                        <strong>{PLAN_FORMAT_SECTION}</strong>
                        <ul>
                          {planFormatPreview(message.plan).map((entry) => (
                            <li key={`${message.id}-format-${entry.key}`}>
                              <span className="format-preview-target">
                                「{entry.sheet}!{entry.range}」
                              </span>
                              <span className="format-preview-changes">
                                {entry.changes.join("、")}
                              </span>
                              {entry.impactNotice && (
                                <span className="format-preview-impact">
                                  <b>{PLAN_FORMAT_IMPACT_LABEL}</b>
                                  {entry.impactNotice}
                                </span>
                              )}
                            </li>
                          ))}
                        </ul>
                      </div>
                    )}

                    <div className="inline-plan-buttons">
                      <button
                        className="secondary-button"
                        disabled={busy}
                        onClick={() => {
                          const plan = message.plan;
                          if (!plan) return;
                          openPlanEditorDialog(plan, {
                            kind: "message",
                            messageId: message.id
                          });
                        }}
                      >
                        编辑计划
                      </button>
                    </div>
                  </details>
                </div>
              )}
              {/*
                删表结果卡只画在**发起这次删除的那条消息**下面。
                这里没有"全局最新一条"的语义：换会话、刷新、后续提问都不会让它
                漂到聊天底部；旧的失败/暂缓也留在各自的请求位置。

                判据是**检查点的归属**，不是 `message.plan`：删表会让已授权表变少，
                用户随后重新确认范围时，`retireTargetMessage` 会把**待执行计划**
                整条退役（`plan: undefined`）。那是计划预览的正确归宿，但它不能
                连带抹掉已经发生的执行回执 —— 所以这里只认 `requestMessageId`。
              */}
              {(() => {
                const checkpoint = worksheetCheckpoint.entryForMessage(
                  message.id,
                  chatHistory.activeConversationId
                );
                if (!checkpoint) return null;
                return (
                  <WorksheetCheckpointCard
                    entry={checkpoint}
                    controller={worksheetCheckpoint}
                    onRestored={async (result, entry) => {
                      // 先在**撤销前已提交的授权**上做核对；快照由真实目录重读得来。
                      const authorized = selectedNamesFor(workbook!);
                      const snapshot = await scan({ announce: false });
                      await adoptVerifiedRestore({
                        identity: entry.workbookIdentity,
                        authorized,
                        restoredSheets: result.restoredSheets,
                        snapshot
                      });
                    }}
                  />
                );
              })()}
              {message.staleSourceConfirm &&
                message.staleSourceConfirm.status !== "completed" &&
                (canAcknowledgeStaleSource(message) ? (
                  <section
                    className="stale-source-actions"
                    aria-label="选择数据变化后的处理方式"
                  >
                    <div className="stale-source-action-copy">
                      <strong>推荐：按当前数据重新生成预览</strong>
                      <span>重新提问后会使用当前数据生成一份新计划。</span>
                    </div>
                    <button
                      type="button"
                      className="secondary-button stale-source-continue-button"
                      disabled={busy}
                      onClick={() => void acknowledgeStaleSource(message)}
                    >
                      仍要执行旧结果
                    </button>
                  </section>
                ) : (
                  <section
                    className="stale-source-actions blocked"
                    aria-label="计划预览不可用"
                  >
                    <div className="stale-source-action-copy">
                      <strong>计划预览不可用，已阻止执行</strong>
                      <span>请重新提问，按当前数据生成一份新计划预览。</span>
                    </div>
                  </section>
                ))}
              {message.text && !message.staleSourceConfirm && (
                <div className="message-copy-row">
                  <MessageCopyButton
                    copied={copiedMessageId === message.id}
                    onClick={() =>
                      void copyMessageText(message.id, message.text?.trim() ?? "")
                    }
                  />
                </div>
              )}
            </div>
          </article>
        ))}

        <div />
      </section>
      {/*
        用户往上滚离底部时出现，判据是 `useMessageAutoFollow` 的 `awayFromBottom`
        —— **只看位置**，不要求"期间又来了一次新内容"。旧口径（`hasUnseenProgress`）
        要求"离底 **且** 有新内容"两个条件同时成立，于是往上翻历史、期间没有新内容时
        界面上什么都不会出现，只能手动滚回去。
        字形跟着**本轮是否在跑**走：跑着是三点动效，跑完换成锚。判据用 `busy`
        （本轮 turn 是否在途），不是"最近有没有报过进度" —— 后者在没有进度回调的
        快速回答里会是 false，于是一个正在跑的 turn 被画成"已停下"。
      */}
      {awayFromBottom && (
        <ScrollToLatestButton running={running} onClick={scrollToLatest} />
      )}
      </div>

      {staleCredentials && (
        <div className="stale-credentials-warning" role="alert">
          <strong>这段历史来自旧版本，曾包含凭据</strong>
          <span>
            凭据已在读取时抹掉，界面里不会再显示原文。为避免把它当作正常历史
            续聊，发送仍会被拦下；请点上方「新对话」开始一段干净的对话，或删掉
            这段历史。
          </span>
        </div>
      )}

      <footer className="composer">
        <ResultReviewBar state={resultReviewState} title={resumedReviewTitle || resultCandidate?.plan.title || ''}
          onCompare={() => void resultReview.compare(resultReviewState.key ?? '')}
          onConfirm={() => void resultReview.confirm(resultReviewState.key ?? '')}
          onDiscard={() => void resultReview.discard(resultReviewState.key ?? '')}
          onRetryMarker={() => void resultReview.retryMarkerCleanup(resultReviewState.key ?? '')} />
        {targetGuard.notice && !contextOpen && <TargetScopeNotice notice={targetGuard.notice} error={targetGuard.error}
          busy={targetGuard.confirming || scopeCheckingCommit} onOpenScope={() => void openScopePicker()}
          onRecheck={() => void targetGuard.recheck()}
        />}
        {showWorkDirectoryGuide && (
          <WorkDirectoryGuide
            busy={busy}
            onDismiss={dismissWorkDirectoryGuide}
            onChoose={() => void chooseWorkDirectoryFromGuide()}
          />
        )}
        {contextOpen && workbook && (
          <section className="scope-popover" aria-label="选择数据范围">
              <div className="scope-popover-header">
              <div>
                <strong>这次对话使用哪些数据？</strong>
                <span className="scope-authorization-hint" role="status">
                  勾选即授权{scopeCheckingCommit || scopeAdmission.checking ? " · 正在核对…" : "，取消勾选即撤销"}
                </span>
              </div>
              <button
                disabled={scopeCheckingCommit}
                onClick={() => setContextOpen(false)}
                aria-label="关闭数据范围选择"
              >
                ×
              </button>
            </div>

            <div className="scope-popover-scroll">
              <fieldset className="scope-selection-fields">
            {sourceMode === "workbook" && (
              <div className="workbook-identity">
                <i aria-hidden="true">XLSX</i>
                <span>
                  <small>当前文件</small>
                  <strong title={workbook.name}>{workbook.name}</strong>
                  {/*
                    当前文件卡只说**用户此刻要用的那两件事**：哪个文件、哪张表。
                    这里**不认文件名里的日期** —— 报表日期识别整条已经下线（2026-10-16 产品决定：
                    "不需要识别日期"）。下线的不是"那句否定式提示"，而是**识别这件事本身**：
                    上一版改成"有日期就显示、没日期就不说"，用户看到的仍然是"这软件在猜我的日期"。
                    所以 `extractWorkbookDataPeriod` 与它派生出的 `workbookDataPeriod` 都已删除，
                    不是留着一个没人读的派生值 —— 留着它，任何一处 JSX 都能把它请回来。
                    与 WPS 的当前文件卡同口径：那边本来就只有「当前工作表 X」。
                  */}
                  <em>当前工作表 {workbook.activeWorksheet}</em>
                </span>
              </div>
            )}

            <div className="scope-controls">
              <fieldset className="scope-selection-fields">
              <div className="scope-mode-list">
                <button
                  className={
                    sourceMode === "workbook" && workbookScopeMode === "auto"
                      ? "selected"
                      : ""
                  }
                  onClick={() => void switchScopeMode("auto")}
                >
                  <i>◎</i>
                  <span>
                    <strong>跟随当前</strong>
                    <small>发送时自动使用正在查看的工作表</small>
                  </span>
                </button>
                <button
                  className={
                    sourceMode === "workbook" && workbookScopeMode === "manual"
                      ? "selected"
                      : ""
                  }
                  onClick={() => void switchScopeMode("manual")}
                >
                  <i>☷</i>
                  <span>
                    <strong>选择多个</strong>
                    <small>用于跨表查询、比较或汇总</small>
                  </span>
                </button>
                {/* 子表选择紧跟「选择多个」，而不是落在「工作目录」下方：
                    它是该模式的展开内容，不能隔着一条无关的模式项。
                    面板常挂载、由 `.open` 类驱动收放 —— 条件渲染会在切走时
                    直接卸载，收缩无从播放。收起态靠 CSS 的 0fr + visibility:hidden。 */}
                {sourceMode === "workbook" && (
                  <div className={`scope-picker-reveal${workbookScopeMode === "manual" ? " open" : ""}`}>
                    <div className="workbook-sheet-picker">
                      <div className="sheet-search-row">
                        <input
                          value={sheetSearch}
                          onChange={(event) => setSheetSearch(event.target.value)}
                          placeholder="搜索工作表..."
                          aria-label="搜索工作表"
                        />
                        <span>
                          已选 {selectedSheetNames.length}/{workbook.worksheets.length}
                        </span>
                      </div>
                      <div className="sheet-picker-toolbar">
                        <button
                          onClick={() => {
                            immediateScope.request();
                            selectAllSheets(
                              filteredWorksheets.map((sheet) => sheet.name)
                            );
                          }}
                        >
                          全选搜索结果
                        </button>
                        <button
                          onClick={() => { clearSelectedSheets(); immediateScope.request(); }}
                        >
                          清空
                        </button>
                      </div>
                      <div className="workbook-sheet-list">
                        {filteredWorksheets.map((sheet) => {
                          const selected = selectedSheetNames.includes(sheet.name);
                          return (
                            <button
                              className={selected ? "selected" : ""}
                              key={sheet.name}
                              onClick={() => { toggleSheet(sheet.name); immediateScope.request(); }}
                              aria-pressed={selected}
                            >
                              <i
                                className="sheet-checkbox"
                                aria-hidden="true"
                              >
                                {selected ? "✓" : ""}
                              </i>
                              {/*
                                行内**不显示**授权状态文案：这一行的选择控件才是它要表达的
                                东西，而「可分析 · 修改仍需预览确认」等长句子会把单行撑成两行、
                                行高随之跳动，也在视觉上压过工作表名。授权状态的唯一去处是
                                列表下方那一条 ScopeAdmissionSummary（整次选择的汇总）。
                              */}
                              <span>
                                <strong>{sheet.name}</strong>
                                <small>
                                  {sheet.name === workbook.activeWorksheet
                                    ? "当前工作表 · "
                                    : ""}
                                  {workbook.catalogOnly ? "字段按需读取" : `${sheet.rowCount} 行 · ${sheet.columnCount} 列`}
                                </small>
                              </span>
                            </button>
                          );
                        })}
                        {filteredWorksheets.length === 0 && (
                          <p>没有匹配的工作表</p>
                        )}
                      </div>
                    </div>
                  </div>
                )}
                <button
                  className={sourceMode === "folder" ? "selected" : ""}
                  onClick={() => {
                    chooseFolderScope();
                    immediateScope.request();
                    setFolderPickerView("files");
                    setFolderSheetSearch("");
                    if (folderCatalog && !busy) {
                      void refreshCurrentFolder({ quiet: true });
                    }
                  }}
                >
                  <i>⌑</i>
                  <span>
                    <strong>工作目录</strong>
                    <small>批量处理多个本地工作簿</small>
                  </span>
                </button>
              </div>
              </fieldset>
              {sourceMode === "workbook" && <ScopeAdmissionSummary items={scopeAdmission.report?.items ?? []}
                count={scopeAdmissionIds.length} checking={scopeAdmission.checking} />}
            </div>

            {sourceMode === "folder" && (
              <div className="folder-scope-picker">
                {!folderCatalog ? (
                  <section
                    className="work-directory-empty"
                    aria-label="设置工作目录"
                  >
                    <i aria-hidden="true">⌂</i>
                    <div>
                      <strong>设置工作目录</strong>
                      <span>
                        选择存放待处理 Excel 文件的文件夹。这里仅确定候选范围，稍后仍需勾选允许处理的工作表。
                      </span>
                      <small>
                        恢复点只会在发生写入后创建，不会生成空的备份目录。
                      </small>
                    </div>
                    <button
                      type="button"
                      className="browse-folder-button"
                      disabled={operationBusy}
                      onClick={() => void browseFolder()}
                    >
                      {status === "scanning" ? "正在读取…" : "选择工作目录"}
                    </button>
                  </section>
                ) : (
                  <div className="folder-toolbar">
                    <button
                      type="button"
                      className="browse-folder-button"
                      disabled={operationBusy}
                      onClick={() => void browseFolder()}
                    >
                      {`更换工作目录 · ${folderCatalog.folderName}`}
                    </button>
                    <button
                      type="button"
                      className="folder-refresh-button"
                      disabled={operationBusy}
                      onClick={() => { void refreshCurrentFolder(); if (immediateScope.isPending()) immediateScope.request(); }}
                      title="重新扫描文件和工作表"
                    >
                      ↻ {status === "scanning" ? "刷新中" : "刷新"}
                    </button>
                  </div>
                )}

                {folderCatalog && (
                  <>
                    <p className="folder-save-hint">
                      在其他工作簿新增或改名后，请先保存文件再刷新。
                    </p>
                    {folderRefreshNotice && (
                      <div
                        className={`folder-refresh-notice${
                          recentlyAddedFolderSheetKeys.length > 0
                            ? " has-new"
                            : ""
                        }`}
                        role="status"
                      >
                        {folderRefreshNotice}
                      </div>
                    )}
                    <div className="folder-selection-filter">
                      <span>
                        {folderCatalog.files.length} 个文件 · {totalSelectableFolderSheets} 张表
                      </span>
                      <div className="folder-selection-filter-actions">
                        <button
                          type="button"
                          disabled={folderSheetKeys.length === 0}
                          onClick={clearFolderSelection}
                        >
                          清空
                        </button>
                        <button
                          type="button"
                          className={showSelectedFolderSheets ? "selected" : ""}
                          disabled={
                            !showSelectedFolderSheets &&
                            folderSheetKeys.length === 0
                          }
                          onClick={() =>
                            setShowSelectedFolderSheets((current) => !current)
                          }
                          aria-pressed={showSelectedFolderSheets}
                        >
                          {showSelectedFolderSheets ? "显示全部" : "只看已选"}
                        </button>
                      </div>
                    </div>
                    <div
                      className={`folder-master-detail folder-picker-view-${folderPickerView}`}
                    >
                      <div
                        className="folder-file-list"
                        role="list"
                        aria-label="工作簿列表"
                      >
                        {visibleFolderFiles.map((file) => {
                          const selectableCount = file.worksheets.filter(
                            (sheet) => !isEBSystemSheet(sheet.name)
                          ).length;
                          const selectedCount =
                            folderSelectionCounts.get(file.id) ?? 0;
                          const active = activeFolderFile?.id === file.id;
                          return (
                            <button
                              type="button"
                              className={`folder-file-option${
                                active ? " active" : ""
                              }${file.error ? " error" : ""}`}
                              key={file.id}
                              onClick={() => {
                                setActiveFolderFileId(file.id);
                                setFolderPickerView("sheets");
                                setFolderSheetSearch("");
                              }}
                              aria-pressed={active}
                              role="listitem"
                            >
                              <i aria-hidden="true">X</i>
                              <span>
                                <strong title={file.relativePath}>
                                  {file.relativePath}
                                </strong>
                                <small>
                                  {file.error
                                    ? "文件暂时无法读取"
                                    : `${selectableCount} 张工作表`}
                                </small>
                              </span>
                              <b>
                                {selectedCount}/{selectableCount}
                              </b>
                            </button>
                          );
                        })}
                        {visibleFolderFiles.length === 0 && (
                          <p className="folder-empty">
                            {showSelectedFolderSheets
                              ? "还没有勾选工作表"
                              : "文件夹中没有可用工作簿"}
                          </p>
                        )}
                      </div>

                      {activeFolderFile && (
                        <section
                          className="folder-sheet-detail"
                          aria-label={`选择 ${activeFolderFile.relativePath} 中的工作表`}
                        >
                          <div className="folder-sheet-header">
                            <button
                              type="button"
                              className="folder-back-button"
                              onClick={() => setFolderPickerView("files")}
                            >
                              ← 文件
                            </button>
                            <span>
                              <strong title={activeFolderFile.relativePath}>
                                {activeFolderFile.relativePath}
                              </strong>
                              <small>
                                已选 {folderSelectionCounts.get(activeFolderFile.id) ?? 0}/
                                {activeFolderSheets.length}
                              </small>
                            </span>
                          </div>

                          {activeFolderFile.error ? (
                            <p className="file-error">
                              {activeFolderFile.error}
                            </p>
                          ) : (
                            <>
                              <div className="sheet-search-row folder-sheet-search">
                                <input
                                  value={folderSheetSearch}
                                  onChange={(event) =>
                                    setFolderSheetSearch(event.target.value)
                                  }
                                  placeholder="搜索当前文件的工作表..."
                                  aria-label="搜索当前文件的工作表"
                                />
                                <span>{filteredFolderSheets.length} 项</span>
                              </div>
                              <div className="sheet-picker-toolbar">
                                <button
                                  type="button"
                                  onClick={() => { selectAllSheetsInFile(activeFolderFile.id); immediateScope.request(); }}
                                >
                                  全选本文件
                                </button>
                                <button
                                  type="button"
                                  onClick={clearActiveFolderFileSelection}
                                >
                                  清空本文件
                                </button>
                              </div>
                              <div className="folder-sheet-list">
                                {filteredFolderSheets.map((sheet) => {
                                  const key = folderSheetKey(
                                    activeFolderFile.id,
                                    sheet.name
                                  );
                                  const selected = folderSheetKeys.includes(key);
                                  const newlyAdded =
                                    recentlyAddedFolderSheetKeys.includes(key);
                                  return (
                                    <button
                                      type="button"
                                      key={sheet.name}
                                      className={selected ? "selected" : ""}
                                      onClick={() => { toggleFolderSheet(activeFolderFile.id, sheet.name); immediateScope.request(); }}
                                      aria-pressed={selected}
                                    >
                                      <i
                                        className="sheet-checkbox"
                                        aria-hidden="true"
                                      >
                                        {selected ? "✓" : ""}
                                      </i>
                                      <span>
                                        <strong>{sheet.name}</strong>
                                        <small>
                                          {sheet.rowCount} 行 · {sheet.columnCount} 列
                                        </small>
                                      </span>
                                      {newlyAdded && <em>新增</em>}
                                    </button>
                                  );
                                })}
                                {filteredFolderSheets.length === 0 && (
                                  <p className="folder-empty">
                                    {showSelectedFolderSheets
                                      ? "当前文件没有已选工作表"
                                      : "没有匹配的工作表"}
                                  </p>
                                )}
                              </div>
                            </>
                          )}
                        </section>
                      )}
                    </div>

                  </>
                )}
              </div>
            )}
            </fieldset>
            </div>

            {/*
              动作条是 `.scope-popover-scroll` 的**兄弟**节点，不是它的后代。
              这正是 Excel 与 WPS 「完成按钮所在的层不一样」的根因：放进滚动容器里，
              它就只是"内容末尾的一行 flex"，内容一滚、它跟着滚走；放在外面，才被
              `.scope-popover`（column flex + `overflow: hidden`）钉成固定页脚。
              所以这里的层不是靠多写几条 padding 对齐的，而是靠**结构**对齐：
              两端读同一份声明（`.scope-action-buttons, .wps-scope-actions` 写在
              共享样式表里），任一端把容器搬回去，两端就会立刻重新分叉。
              按钮集合同样对齐 WPS：刷新 / 确认。勾选自动核验，确认只关闭面板。
            */}
            {sourceMode === "workbook" && (
              <div className="scope-action-buttons">
                <button
                  className="scope-refresh"
                  disabled={scopeCheckingCommit}
                  onClick={() => { void refreshWorkbook({ ignoreBusy: true }); if (immediateScope.isPending()) immediateScope.request(); }}
                  title="刷新工作簿（更新工作簿名和工作表列表）"
                >
                  <i
                    className={`scope-refresh-glyph${refreshSpin ? " spinning" : ""}`}
                    key={refreshSpin}
                    aria-hidden="true"
                  >↻</i>
                  {" "}刷新
                </button>
                <button type="button" className="scope-close" data-scope-pending={immediateScope.pending} onClick={discardWorkbookScopeEdit}>确认</button>
              </div>
            )}
            {sourceMode === "workbook" && scopeConfirmationError &&
              <p role="alert" className="authorized-target-error">{targetGuard.error || scopeConfirmationError}</p>}

            {sourceMode === "folder" && folderCatalog && (
              <div className="folder-selection-footer">
                <div className="folder-selection-summary">
                  <strong>
                    已选 {selectedFolderFileCount} 个文件 · {folderSheetKeys.length} 张表
                  </strong>
                  <details className="folder-selection-help">
                    <summary
                      aria-label="查看工作表选择保留说明"
                      title="查看说明"
                    >
                      ⓘ
                    </summary>
                    <span role="note">
                      切换文件不会丢失已勾选的工作表
                    </span>
                  </details>
                </div>
                <button type="button" className="scope-close" data-scope-pending={immediateScope.pending} onClick={discardWorkbookScopeEdit}>确认</button>
              </div>
            )}
            {sourceMode === "folder" && <ScopeAdmissionSummary items={scopeAdmission.report?.items ?? []}
              count={scopeAdmissionIds.length} checking={scopeAdmission.checking} />}
            {sourceMode === "folder" && scopeConfirmationError && <p role="alert" className="authorized-target-error">{scopeConfirmationError}</p>}
          </section>
        )}

        <div className="scope-entry-row">
        <button
          className={`scope-trigger ${contextOpen ? "open" : ""}`}
          onClick={() => {
            if (!contextOpen) { void openScopePicker(); return; }
            const nextOpen = !contextOpen;
            setFolderBackupOpen(false);
            setSheetSearch("");
            setFolderSheetSearch("");
            setContextOpen(nextOpen);
            if (
              nextOpen &&
              sourceMode === "folder" &&
              folderCatalog &&
              !busy
            ) {
              void refreshCurrentFolder({ quiet: true });
            }
          }}
          aria-expanded={contextOpen}
        >
          <i aria-hidden="true">
            <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
              <path d="M19.2 12.1l-6.9 6.9a4.6 4.6 0 0 1-6.5-6.5l6.9-6.9a3 3 0 0 1 4.3 4.3l-6.9 6.9a1.4 1.4 0 0 1-2-2l6.2-6.2" />
            </svg>
          </i>
          <span>
            {workbook && sourceMode === "workbook" && (
              <>
                <small title={workbook.name}>{workbook.name}</small>
                <span className="scope-sep" aria-hidden="true">
                  |
                </span>
              </>
            )}
            <strong>
              {!workbook
                ? "正在读取工作簿"
                : sourceMode === "folder"
                  ? folderCatalog
                    ? `工作目录 ${folderCatalog.folderName} · 已选 ${folderSheetKeys.length} 张表`
                    : "设置工作目录"
                  : workbookScopeMode === "auto"
                    ? `当前表 ${workbook.activeWorksheet}`
                    : `已固定 ${selectedSheetNames.length} 个工作表`}
            </strong>
          </span>
          <b>{contextOpen ? "⌄" : "⌃"}</b>
        </button>
        </div>

        {/*
          对话历史的存储状态提示。四种形态由 `HistoryStorageNotice` 按状态给文案：
          "保存失败 + 旧记录保留" / "未保存且无旧记录" / "已保存但旧键待清理" / 无提示。
          不贴异常原文、不含任何历史内容。与"旧记录拦续聊"是两件事，所以不复用
          上面的 staleCredentials 气泡：那条是针对当前会话能不能续聊。
        */}
        <HistoryStorageNotice status={historyStorageStatus} />

        <div
          className={`composer-input-region ${
            draggingImage ? "dragging-image" : ""
          }`}
          onDragEnter={(event) => {
            event.preventDefault();
            setDraggingImage(true);
          }}
          onDragOver={(event) => event.preventDefault()}
          onDragLeave={(event) => {
            if (!event.currentTarget.contains(event.relatedTarget as Node)) {
              setDraggingImage(false);
            }
          }}
          onDrop={handleImageDrop}
        >
          <div className="composer-resize-handle" {...composerResizeHandleProps}>
            <i />
          </div>
          <TaskExecutionControls state={composerPauseButton}
            pausing={pauseRequestState?.phase === "requesting"}
            disabled={busy && composerPauseButton.kind === "continue"}
            onPause={() => stopTurn()}
            onResume={() => void resumeComposerTask()} />
          <div className="composer-box">
            <SlashCommandAutocomplete
              visible={showSlashAutocomplete}
              commands={slashMode === "model" ? slashModelCommands : slashCommands}
              onSelect={handleSlashCommandSelect}
              filter={slashFilter}
              title={slashMode === "model" ? "选择模型" : undefined}
            />
            <textarea
              ref={composerInputRef}
              aria-label="给 Excel Bro 发消息"
              value={prompt}
              onChange={handleComposerChange}
              onKeyDown={handleComposerKeyDown}
              onPaste={handleImagePaste}
              onBlur={() => setShowSlashAutocomplete(false)}
              placeholder={
                workbook
                  ? "描述你想查询、分析或修改的内容…"
                  : "可以先输入需求，工作簿读取完成后即可发送..."
              }
              rows={1}
            />
            {pendingImages.length > 0 && (
              <div className="composer-attachments">
                {pendingImages.map((image) => (
                  <figure key={image.id}>
                    <img src={image.previewUrl} alt={image.name} />
                    <figcaption title={image.name}>{image.name}</figcaption>
                    <button
                      disabled={busy || preparingImages}
                      onClick={() => {
                        setPendingImages((current) =>
                          current.filter((item) => item.id !== image.id)
                        );
                        setImageError("");
                      }}
                      aria-label={`移除图片 ${image.name}`}
                    >
                      ×
                    </button>
                  </figure>
                ))}
              </div>
            )}
            {preparingImages && <div className="composer-image-error" role="status">正在处理图片…</div>}
            {pendingImages.length > 0 && !supportsVision && (
              <div className="composer-image-error" role="status">当前模型不支持图片，请切换到支持视觉输入的模型后发送。附件不会自动发送。</div>
            )}
            {imageError && (
              <div className="composer-image-error">{imageError}</div>
            )}
            <div className="composer-toolbar">
              <input
                ref={imageInputRef}
                aria-label="选择图片附件"
                className="image-file-input"
                type="file"
                accept="image/png,image/jpeg,image/webp"
                multiple
                onChange={(event) => {
                  const files = [...(event.target.files ?? [])];
                  event.target.value = "";
                  void addImageFiles(files);
                }}
              />
              <div className="composer-tools-left">
                <button
                  className="attach-image-button"
                  disabled={busy || preparingImages}
                  onClick={() => imageInputRef.current?.click()}
                  title={
                    supportsVision
                      ? "添加图片，也可以直接粘贴或拖入截图"
                      : "添加图片；当前模型不支持视觉，发送前需切换模型"
                  }
                  aria-label="添加图片"
                >
                  <svg
                    viewBox="0 0 24 24"
                    aria-hidden="true"
                    focusable="false"
                  >
                    <rect x="3.5" y="4.5" width="17" height="15" rx="2.5" />
                    <circle cx="9" cy="9.5" r="1.5" />
                    <path d="m5.5 17 4.2-4.4 3.1 3 2.2-2.2 3.5 3.6" />
                  </svg>
                </button>
              </div>
              <span
                className="composer-shortcut-hint"
                title={
                  composerPauseButton.canPause
                    ? "运行中：输入新内容按 Enter 可打断并转向，点击上方 ■ 暂停"
                    : busy
                      ? (resultReviewState.phase === "reviewing" && !resultReviewState.pending
                        ? "等待审阅结果：可确认变更、放弃变更或打开双视图"
                        : "正在处理，请稍候")
                      : "Enter 发送 · Shift + Enter 换行 · 可粘贴截图"
                }
              >
                {composerPauseButton.canPause
                  ? "运行中：输入新内容按 Enter 可打断并转向，点击上方 ■ 暂停"
                  : busy
                    ? (resultReviewState.phase === "reviewing" && !resultReviewState.pending
                        ? "等待审阅结果：可确认变更、放弃变更或打开双视图"
                        : "正在处理，请稍候")
                    : "Enter 发送 · Shift + Enter 换行 · 可粘贴截图"}
              </span>
              <button className={`send-button ${prompt.trim() && running ? "is-steer" : ""}`}
                type="button"
                disabled={!workbook || preparingImages || (!prompt.trim() && pendingImages.length === 0) ||
                  (pendingImages.length > 0 && !supportsVision) || (running && !prompt.trim())}
                onClick={() => (running && prompt.trim() ? steerTurn(prompt) : void submitComposer())}
                aria-label={running && prompt.trim() ? "打断并补充" : "发送"}
                title={running && prompt.trim() ? "打断当前处理，带上这句话重新开始" : "发送这条需求"}>↑</button>
            </div>
            {/*
              暂停请求的阶段**如实**显示，挂在工具条**外面**：工具条是三列网格，
              多一个子元素会被挤到第二行，看起来像"又开了一块区域"。
              本端已停止接收 ≠ 服务端已暂停，文字必须分开。
            */}
            {pauseRequestState && (
              <span className="composer-pause-note" role="status" data-phase={pauseRequestState.phase}>
                {pauseRequestState.detail}
              </span>
            )}
          </div>
        </div>
        {(import.meta.env.DEV || import.meta.env.VITE_ENABLE_RECOVERY_PROBE === "true") && <RecoveryProbeControl identity={resultWorkbookIdentity} disabled={busy} />}
      </footer>
      {petVisible && <PetCompanion busy={busy} />}
      <ThemePanel
        open={themePanelOpen}
        onClose={() => setThemePanelOpen(false)}
        {...themeApi}
      />
      <RuleManager
        isOpen={isRuleManagerOpen}
        onClose={() => setIsRuleManagerOpen(false)}
      />
    </main>
  );
}
