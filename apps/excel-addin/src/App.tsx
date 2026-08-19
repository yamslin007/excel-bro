import {
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type ClipboardEvent,
  type DragEvent
} from "react";
import {
  checkIntent,
  streamAssistantResponse,
  createFolderSnapshot,
  executeFolderPlan,
  executeFolderQuery,
  generateFormula,
  isLocalServiceConnectionError,
  selectFolder,
} from "./api";
import type {
  AnalysisPlan,
  FolderCatalog,
  FolderSelection,
  DataToolResult,
  FormulaDictionarySheet,
  FormulaExtraSheet,
  IntentCheckResponse,
  IntentClarification,
  IntentMemory,
  IntentOption,
  IntentScopeContext,
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
import { demoWorkbook } from "./demo";
import {
  captureSelectionContext,
  captureWorkbook,
  captureWorkbookSourceFingerprint,
  captureWorkbookStructure,
  dataEpochsChanged,
  executePlan,
  isExplicitA1Address,
  isEBSystemSheet,
  isRunningInExcel,
  readSelectedRange,
  previewFormulaFirstCell,
  PlanExecutionError,
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
import { renderToolDsl } from "./toolDsl";
import {
  MAX_IMAGE_ATTACHMENTS,
  prepareImageFile,
  type PendingImage
} from "./imageAttachments";
import { extractWorkbookDataPeriod } from "./workbookIdentity";
import capabilities from "../../../config/capabilities.json";
import {
  buildCrossTableFormulas,
  buildCrossTableProposal,
  columnsFromRange,
  type CrossTableMatchProposal,
  type GeneratedFormulaPair
} from "./crossTableFormula";
import { formulaExternalFileNames } from "./formulaSafety";
import {
  currentModelCallCount,
  exportDiagnosticReport,
  recordDiagnosticEvent
} from "./diagnostics";
import {
  FOCUS_PAYLOAD_STORAGE_KEY,
  type FocusPayload
} from "./focusState";
import PetCompanion from "./PetCompanion";
import { RuleManager } from "./RuleManager";
import { SlashCommandAutocomplete, type SlashCommand } from "./SlashCommandAutocomplete";
import ThemePanel from "./components/ThemePanel";
import { BASE_MODE_DESCRIPTION, BASE_MODE_HELP_TEXT } from "./helpCommand";
import { useActivityProgress, type ActivityLog, type ActivityProgress } from "./hooks/useActivityProgress";
import { useConversation } from "./hooks/useConversation";
import { useCopyFeedback } from "./hooks/useCopyFeedback";
import { useExecutionApproval } from "./hooks/useExecutionApproval";
import { useModelManagement } from "./hooks/useModelManagement";
import { useServiceHealth } from "./hooks/useServiceHealth";
import { useTheme } from "./hooks/useTheme";
import { useToolManagement } from "./hooks/useToolManagement";
import { useUIState } from "./hooks/useUIState";
import { useScopeSelection } from "./hooks/useScopeSelection";
import { useUndoSnapshot } from "./hooks/useUndoSnapshot";
import { useLongPress } from "./hooks/useLongPress";
import { folderSheetKey, formulaTrialFailed } from "./utils";

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
const LEGACY_CHAT_STORAGE_KEY = "excel-bro.chat.v3";
const BASE_MODE_HELP_SHOWN_KEY = "ebBasicModeHelpShown";
const MAX_STORED_CONVERSATIONS =
  capabilities.conversation.maxStoredConversations;
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
const COMPOSER_MAX_HEIGHT = 156;
const COMPOSER_MIN_HEIGHT = 44;

function verificationSummary(report: VerificationReport): string {
  if (report.status === "verified") return "并通过独立验证";
  if (report.status === "executed_unverified") {
    return "；写入已完成，但部分操作暂时无法独立验证";
  }
  return "，但结果验证未通过";
}

export function normalizeStoredVerification(
  report: VerificationReport | undefined
): VerificationReport | undefined {
  if (!report) return undefined;
  return {
    ...report,
    status:
      report.status ??
      (report.passed ? "verified" : "failed"),
    unverifiedActions: Array.isArray(report.unverifiedActions)
      ? report.unverifiedActions
      : []
  };
}

function formatStepElapsed(ms: number): string {
  if (ms < 1000) return `${Math.max(0, Math.round(ms))}ms`;
  return `${(ms / 1000).toFixed(1)}s`;
}

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

// 把本地工具查询参数翻译成一句人能读懂的话，供「查看详情」里展示。
function describeQueryArguments(args: QueryTableArguments): string {
  const parts: string[] = [];
  const modeLabel =
    args.mode === "aggregate"
      ? "分组汇总"
      : args.mode === "profile"
        ? "字段画像"
        : "取明细行";
  parts.push(`方式：${modeLabel}`);
  if (args.fields?.length) parts.push(`字段：${args.fields.join("、")}`);
  if (args.groupBy?.length) parts.push(`分组：${args.groupBy.join("、")}`);
  if (args.metrics?.length) {
    parts.push(
      `指标：${args.metrics
        .map((metric) => `${metric.operation}(${metric.field ?? "*"})`)
        .join("、")}`
    );
  }
  if (args.filters?.length) parts.push(`筛选：${args.filters.length} 个条件`);
  if (args.combine?.mode) parts.push(`合并：${args.combine.mode}`);
  if (typeof args.limit === "number") parts.push(`上限：${args.limit} 行`);
  return parts.join("\n");
}

// describeIntentDecision：把需求确认的结构化结果翻译成「用户看得懂的判断」。
// 返回 label（步骤标题）、note（模型的真实理解，第一层默认可见）、
// detail（原始明细，第二层展开「查看详情」才显示），让用户判断模型是否偏离了轨迹。
function describeIntentDecision(intent: IntentCheckResponse): {
  label: string;
  note?: string;
  detail?: string;
} {
  const source = intent.provider === "model" ? "模型" : "本地规则";
  if (intent.kind === "clarification") {
    const clarification = intent.clarification;
    const optionText = clarification.options
      .map((option, index) => `${index + 1}. ${option.label}`)
      .join("\n");
    return {
      label: `${source}判断需要先澄清需求`,
      note: clarification.summary || clarification.question,
      detail: [
        `理解：${clarification.summary}`,
        `追问：${clarification.question}`,
        clarification.reason ? `原因：${clarification.reason}` : "",
        clarification.options.length ? `可选项：\n${optionText}` : ""
      ]
        .filter(Boolean)
        .join("\n\n")
    };
  }
  if (intent.kind === "tool_request") {
    return {
      label: `${source}决定先用本地工具取数`,
      note: intent.summary,
      detail: [
        `理解：${intent.summary}`,
        `锁定执行：${intent.confirmedPrompt}`,
        `本地查询：\n${describeQueryArguments(intent.request.arguments)}`
      ]
        .filter(Boolean)
        .join("\n\n")
    };
  }
  return {
    label: `${source}确认可以直接分析`,
    note: intent.summary,
    detail: [
      `理解：${intent.summary}`,
      `锁定执行：${intent.confirmedPrompt}`
    ]
      .filter(Boolean)
      .join("\n\n")
  };
}

function intentScopeFingerprint(scope: IntentScopeContext): string {
  // selectedRange (光标位置) 不纳入指纹：查询不依赖它，写入落点由预览确认兜底。
  // 把它算进来只会在用户移动光标时造成误判（"数据范围已经变化"）。
  // activeWorksheet 在 auto 模式下决定扫描哪张表，切表意味着数据真的变了，仍需保留。
  return JSON.stringify({
    workbookName: scope.workbookName,
    sourceMode: scope.sourceMode,
    selectionMode: scope.selectionMode,
    sheets: scope.sheets.map((sheet) => sheet.name),
    ...(scope.selectionMode === "auto"
      ? {
          activeWorksheet: scope.activeWorksheet
        }
      : {})
  });
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

function welcomeMessage(): ChatMessage {
  return {
    id: messageId(),
    role: "assistant",
    text:
      "你好。直接告诉我你想查询、分析或修改什么；涉及写入时，我会先给你预览。",
    createdAt: new Date().toISOString()
  };
}

function conversationTitle(messages: ChatMessage[]): string {
  const firstUserMessage = messages.find(
    (message) => message.role === "user" && message.text?.trim()
  )?.text;
  if (!firstUserMessage) return "新对话";
  const compact = firstUserMessage.replace(/\s+/g, " ").trim();
  return compact.length > 34 ? `${compact.slice(0, 34)}…` : compact;
}

function createConversation(messages = [welcomeMessage()]): ChatConversation {
  const now = new Date().toISOString();
  return {
    id: messageId(),
    title: conversationTitle(messages),
    messages,
    createdAt: now,
    updatedAt: now
  };
}

export function deleteConversationFromHistory(
  current: ChatHistoryState,
  conversationId: string
): ChatHistoryState {
  const remaining = current.conversations.filter(
    (conversation) => conversation.id !== conversationId
  );
  if (remaining.length === current.conversations.length) {
    return current;
  }
  if (remaining.length === 0) {
    const replacement = createConversation();
    return {
      activeConversationId: replacement.id,
      conversations: [replacement]
    };
  }
  return {
    activeConversationId:
      current.activeConversationId === conversationId
        ? remaining[0].id
        : current.activeConversationId,
    conversations: remaining
  };
}

function loadChatHistory(): ChatHistoryState {
  try {
    const stored = JSON.parse(
      localStorage.getItem(CHAT_STORAGE_KEY) ?? "null"
    ) as Partial<ChatHistoryState> | null;
    if (
      stored &&
      typeof stored.activeConversationId === "string" &&
      Array.isArray(stored.conversations) &&
      stored.conversations.length > 0
    ) {
      const conversations = stored.conversations
        .filter(
          (conversation): conversation is ChatConversation =>
            Boolean(
              conversation &&
                typeof conversation.id === "string" &&
                Array.isArray(conversation.messages)
            )
        )
        .map((conversation) => {
          const now = new Date().toISOString();
          return {
            ...conversation,
            messages: conversation.messages.map((message) => {
              const normalizedMessage = {
                ...message,
                verification: normalizeStoredVerification(
                  message.verification
                )
              };
              return message.clarification
                ? {
                    ...normalizedMessage,
                    clarification: {
                      ...message.clarification,
                      round:
                        typeof message.clarification.round === "number"
                          ? message.clarification.round
                          : 0,
                      status:
                        message.clarification.status === "resolving"
                          ? "pending"
                          : message.clarification.status
                    }
                  }
                : normalizedMessage;
            }),
            title:
              typeof conversation.title === "string"
                ? conversation.title
                : conversationTitle(conversation.messages),
            createdAt:
              typeof conversation.createdAt === "string"
                ? conversation.createdAt
                : now,
            updatedAt:
              typeof conversation.updatedAt === "string"
                ? conversation.updatedAt
                : now
          };
        });
      if (conversations.length > 0) {
        return {
          activeConversationId: conversations.some(
            (conversation) =>
              conversation.id === stored.activeConversationId
          )
            ? stored.activeConversationId
            : conversations[0].id,
          conversations
        };
      }
    }
  } catch {
    // Fall through to legacy migration.
  }

  try {
    const legacy = JSON.parse(
      localStorage.getItem(LEGACY_CHAT_STORAGE_KEY) ?? "[]"
    );
    const conversation = createConversation(
      Array.isArray(legacy) && legacy.length > 0
        ? (legacy as ChatMessage[])
        : [welcomeMessage()]
    );
    return {
      activeConversationId: conversation.id,
      conversations: [conversation]
    };
  } catch {
    const conversation = createConversation();
    return {
      activeConversationId: conversation.id,
      conversations: [conversation]
    };
  }
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

export default function App() {
  const [workbook, setWorkbook] = useState<WorkbookSnapshot | null>(null);
  const [prompt, setPrompt] = useState("");
  const [pendingImages, setPendingImages] = useState<PendingImage[]>([]);
  const [imageError, setImageError] = useState("");
  const [clarificationDrafts, setClarificationDrafts] = useState<
    Record<string, string>
  >({});
  const [draggingImage, setDraggingImage] = useState(false);

  // 对话管理（使用自定义 Hook）
  const conversationHook = useConversation();
  const {
    chatHistory,
    activeConversation,
    messages,
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
  // 活动进度管理（自定义 Hook）
  const activityHook = useActivityProgress({
    onPersistLog: (log) => {
      setMessages((messages) => {
        const lastAssistantIndex = [...messages]
          .reverse()
          .findIndex((message) => message.role === "assistant");
        if (lastAssistantIndex === -1) return messages;
        const index = messages.length - 1 - lastAssistantIndex;
        return messages.map((message, i) =>
          i === index ? { ...message, activityLog: log } : message
        );
      });
    }
  });
  const {
    activity,
    activitySeconds,
    startActivity: beginActivity,
    advanceActivity,
    updateActivityDetail,
    completeActivity: finishActivity
  } = activityHook;

  // 服务健康状态为模型管理提供模型目录刷新能力。
  const serviceHealthHook = useServiceHealth();
  const {
    serverOnline,
    serviceHealth,
    modelOptions,
    modelCatalogLoaded,
    refreshServiceState,
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
        text: "复制专家脚本失败，请手动选择脚本内容后复制。"
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
    setPendingToolDeletion
  } = toolManagementHook;

  const [contextOpen, setContextOpen] = useState(false);
  const [sheetSearch, setSheetSearch] = useState("");
  const [themePanelOpen, setThemePanelOpen] = useState(false);

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
    composerHeight, setComposerHeight,
    togglePetVisibility
  } = useUIState();

  const themeApi = useTheme();
  const logoLongPress = useLongPress(() => setThemePanelOpen(true));

  // 执行验收与工具固化批准状态（自定义 Hook）
  const {
    saveCandidate,
    setSaveCandidate,
    approveFixedContent,
    setApproveFixedContent,
    approveDestructive,
    setApproveDestructive,
    verifiedPlanIds,
    markPlanVerified,
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
    setSelectedSheetNames,
    selectionConfirmed,
    setSelectionConfirmed,
    sourceMode,
    workbookScopeMode,
    folderCatalog,
    folderSheetKeys,
    applyWorkbookSnapshotSelection,
    toggleSheet,
    toggleFolderSheet,
    selectAllSheetsInFile,
    clearSheetsInFile,
    applyFolderCatalog,
    chooseAutomaticScope,
    chooseManualScope,
    chooseFolderScope,
    folderSelections,
    selectedNamesFor,
    selectAllSheets,
    clearSelectedSheets
  } = useScopeSelection();
  const messageEndRef = useRef<HTMLDivElement>(null);
  const composerInputRef = useRef<HTMLTextAreaElement>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);
  const clarificationImagesRef = useRef<Map<string, PendingImage[]>>(
    new Map()
  );
  const composerResizeRef = useRef<{
    pointerId: number;
    startY: number;
    startHeight: number;
  } | null>(null);
  const focusDialogRef = useRef<Office.Dialog | null>(null);
  const queryAbortRef = useRef<AbortController | null>(null);
  // turnAbortRef：整个 turn（需求确认→本地取数→生成方案）共用的打断句柄。
  // 每个 turn 开始时新建；用户中途插话/停止时 abort，串在链上的模型请求会一起 reject。
  const turnAbortRef = useRef<AbortController | null>(null);
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

  // 一级斜杠命令定义
  const slashCommands: SlashCommand[] = useMemo(
    () => [
      {
        value: "help",
        description: "查看基础模式支持的命令（离线功能说明）"
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
        active: option.id === (selectedModelId || "local"),
        disabled: !option.available,
        description: option.available
          ? option.supportsVision
            ? "支持图片输入"
            : "文本模型"
          : "未配置或不可用"
      })),
    [modelOptions, selectedModelId]
  );

  const busy = status !== "idle";
  const isBaseMode = (selectedModelId || serviceHealth?.model || "local") === "local";

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

  // 撤销快照管理（自定义 Hook）
  const {
    lastUndoSnapshot,
    setLastUndoSnapshot,
    clearUndoSnapshot,
    undoLastExecution
  } = useUndoSnapshot({
    isBusy: busy,
    onStatusChange: setStatus,
    onMessage: (text) => {
      appendMessage({
        role: "system",
        text
      });
    },
    onAfterUndo: () => scan()
  });
  const selectedModel = useMemo(
    () =>
      modelOptions.find(
        (option) => option.id === (selectedModelId || "local")
      ) ?? modelOptions[0],
    [modelOptions, selectedModelId]
  );
  const supportsVision = selectedModel?.supportsVision === true;
  const headerStatusClassName = !serverOnline
    ? ""
    : isBaseMode
      ? "online base-mode"
      : serviceHealth?.configured
        ? "online model-online"
        : "online";
  const headerStatusText = !serverOnline
    ? "本地服务未连接"
    : isBaseMode
      ? "本地服务已连接 · 基础模式"
      : `模型：${selectedModel?.label ?? serviceHealth?.model}`;
  const hasConfiguredModel = modelOptions.some(
    (option) => option.provider === "model" && option.available
  );
  const showFirstModelGuide =
    serverOnline &&
    modelCatalogLoaded &&
    !hasConfiguredModel &&
    !modelGuideDismissed &&
    !settingsOpen;
  // 两步引导互斥、永不同屏：窄窗格且拉宽步未完成 → 只弹「拉宽」；
  // 拉过阈值(≥PANE_WIDEN_THRESHOLD)或点跳过、以及本就宽窗格 → 承接「模型」引导。
  const showWidenGuide =
    showFirstModelGuide && isNarrowPane && !widenStepDone;
  const showModelGuide =
    showFirstModelGuide && (!isNarrowPane || widenStepDone);
  const hasEnvironmentModel =
    Boolean(modelSettings?.baseUrl) && Boolean(modelSettings?.defaultModel);
  const hasManagedModels = (modelSettings?.connections.length ?? 0) > 0;
  const editingConnection = connectionDraft?.id
    ? modelSettings?.connections.find(
        (connection) => connection.id === connectionDraft.id
      )
    : null;
  const workbookDataPeriod = useMemo(
    () => (workbook ? extractWorkbookDataPeriod(workbook.name) : null),
    [workbook]
  );
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

  useEffect(() => {
    try {
      localStorage.setItem(CHAT_STORAGE_KEY, JSON.stringify(chatHistory));
    } catch {
      const compact: ChatHistoryState = {
        activeConversationId: chatHistory.activeConversationId,
        conversations: chatHistory.conversations
          .slice(0, 8)
          .map((conversation) => ({
            ...conversation,
            messages: conversation.messages.slice(
              -PERSISTED_MESSAGES_PER_CONVERSATION
            )
          }))
      };
      try {
        localStorage.setItem(CHAT_STORAGE_KEY, JSON.stringify(compact));
      } catch {
        // The current in-memory conversation remains available.
      }
    }
  }, [chatHistory]);

  useEffect(() => {
    messageEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  useEffect(
    () => () => {
      focusDialogRef.current?.close();
      focusDialogRef.current = null;
    },
    []
  );

  useLayoutEffect(() => {
    const input = composerInputRef.current;
    if (!input) return;
    if (composerHeight !== null) {
      input.style.height = `${composerHeight}px`;
      input.style.overflowY =
        input.scrollHeight > composerHeight ? "auto" : "hidden";
      return;
    }
    input.style.height = "auto";
    const nextHeight = Math.min(input.scrollHeight, COMPOSER_MAX_HEIGHT);
    input.style.height = `${nextHeight}px`;
    input.style.overflowY =
      input.scrollHeight > COMPOSER_MAX_HEIGHT ? "auto" : "hidden";
  }, [prompt, composerHeight]);

  function appendMessage(
    message: Omit<ChatMessage, "id" | "createdAt">
  ): string {
    const id = messageId();
    setMessages((current) => [
      ...current,
      {
        ...message,
        id,
        createdAt: new Date().toISOString()
      }
    ]);
    return id;
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

  async function scan(options?: { announce?: boolean }) {
    const diagnosticStartedAt = performance.now();
    setStatus("scanning");
    try {
      const snapshot = isRunningInExcel()
        ? await captureWorkbookStructure(
            sourceMode === "workbook" && workbookScopeMode === "manual"
              ? selectedSheetNames
              : undefined
          )
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
    } finally {
      setStatus("idle");
    }
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
    setModelMenuOpen(false);
    setMoreMenuOpen(false);
    if (
      activeConversation &&
      !activeConversation.messages.some(
        (message) => message.role === "user"
      )
    ) {
      setPrompt("");
      setPendingImages([]);
      setImageError("");
      setContextOpen(false);
      setHistoryOpen(false);
      setToolsOpen(false);
      closeSettings();
      closeSaveCandidate();
      return;
    }
    conversationHook.newChat();
    setPrompt("");
    setPendingImages([]);
    setImageError("");
    setContextOpen(false);
    setHistoryOpen(false);
    setToolsOpen(false);
    closeSettings();
    closeSaveCandidate();
  }

  function handleOpenConversation(conversationId: string) {
    if (busy) return;
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

  async function addImageFiles(files: File[]) {
    if (!supportsVision) {
      setImageError("当前模型不支持图片，请先切换到支持视觉输入的模型。");
      return;
    }
    const remaining = MAX_IMAGE_ATTACHMENTS - pendingImages.length;
    if (remaining <= 0) {
      setImageError(`一次最多添加 ${MAX_IMAGE_ATTACHMENTS} 张图片。`);
      return;
    }
    const accepted = files.slice(0, remaining);
    try {
      const prepared: PendingImage[] = [];
      for (const file of accepted) {
        prepared.push(await prepareImageFile(file));
      }
      setPendingImages((current) => [
        ...current,
        ...prepared.slice(0, MAX_IMAGE_ATTACHMENTS - current.length)
      ]);
      setImageError(
        files.length > remaining
          ? `一次最多添加 ${MAX_IMAGE_ATTACHMENTS} 张图片。`
          : ""
      );
    } catch (reason) {
      setImageError(
        reason instanceof Error ? reason.message : "添加图片失败"
      );
    }
  }

  function handleImagePaste(event: ClipboardEvent<HTMLTextAreaElement>) {
    const files = [...event.clipboardData.items]
      .filter((item) => item.kind === "file" && item.type.startsWith("image/"))
      .map((item) => item.getAsFile())
      .filter((file): file is File => file !== null);
    if (files.length === 0) return;
    event.preventDefault();
    void addImageFiles(files);
  }

  function handleImageDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    setDraggingImage(false);
    const files = [...event.dataTransfer.files].filter((file) =>
      file.type.startsWith("image/")
    );
    if (files.length === 0) {
      setImageError("请拖入 PNG、JPEG 或 WebP 图片。");
      return;
    }
    void addImageFiles(files);
  }

  async function browseFolder() {
    setStatus("scanning");
    try {
      const catalog = await selectFolder();
      if (!catalog) return;
      applyFolderCatalog(catalog);
      markServerOnline();
    } catch (reason) {
      appendMessage({
        role: "system",
        text: reason instanceof Error ? reason.message : "读取文件夹失败"
      });
    } finally {
      setStatus("idle");
    }
  }

  async function confirmSheetSelection() {
    if (sourceMode === "folder") {
      if (!folderCatalog || folderSheetKeys.length === 0) return;
      setStatus("scanning");
      try {
        const snapshot = await createFolderSnapshot(
          folderCatalog.sessionId,
          folderSelections()
        );
        setWorkbook(snapshot);
        setSelectedSheetNames(snapshot.worksheets.map((sheet) => sheet.name));
        setSelectionConfirmed(true);
        setContextOpen(false);
      } catch (reason) {
        appendMessage({
          role: "system",
          text: reason instanceof Error ? reason.message : "读取所选工作表失败"
        });
      } finally {
        setStatus("idle");
      }
      return;
    }

    if (!workbook || selectedSheetNames.length === 0) return;
    setSelectionConfirmed(true);
    setContextOpen(false);
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

  async function analyzeConfirmedIntent(
    confirmedPrompt: string,
    sentImages: PendingImage[],
    expectedScopeFingerprint?: string,
    clarificationMessageId?: string,
    dataResults: DataToolResult[] = [],
    preparedWorkbook?: WorkbookSnapshot,
    intentMemory?: IntentMemory,
    turnId?: string | null
  ) {
    if (!workbook) return;
    setStatus("planning");
    if (dataResults.length > 0) {
      const scannedRows = dataResults.reduce(
        (total, result) => total + result.scannedRows,
        0
      );
      advanceActivity(
        "正在生成回答",
        "模型只接收本地工具返回的紧凑结果，不接收完整工作表数据。",
        `本地计算完成，共扫描 ${scannedRows.toLocaleString()} 行`
      );
    } else {
      advanceActivity(
        "正在生成操作预览",
        "正在根据已确认需求生成受约束的 Excel 操作计划。"
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
        return;
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
      const requestWorkbook: WorkbookSnapshot = {
        ...latestWorkbook,
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
      const response = await streamAssistantResponse(
        {
          turnId,
          prompt: confirmedPrompt,
          workbook: requestWorkbook,
          lastResult,
          images: sentImages.map(({ name, mediaType, data }) => ({
            name,
            mediaType,
            data
          })),
          dataResults,
          modelId: selectedModelId || null
        },
        {
          signal: turnAbortRef.current?.signal,
          onStep: (step) =>
            advanceActivity(
              step.title,
              step.detail ?? "",
              step.completedStep ?? undefined
            )
        }
      );
      markServerOnline();
      if (response.kind === "answer") {
        const querySourceSheetIds = currentWorksheets
          .map((sheet) => sheet.sourceSheetId)
          .filter((value): value is string => Boolean(value));
        appendMessage({
          role: "assistant",
          text: response.message,
          resultContext: response.resultContext ?? undefined,
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
      } else {
        const sourceFingerprint =
          sourceMode === "workbook" && isRunningInExcel()
            ? await captureWorkbookSourceFingerprint(effectiveSheetNames)
            : latestWorkbook.sourceFingerprint;
        const plan: AnalysisPlan = {
          ...response.plan,
          sourceFingerprint,
          sourceFingerprintSheets: effectiveSheetNames
        };
        appendMessage({
          role: "assistant",
          text: plan.summary,
          plan,
          intentMemory,
          provider: response.provider
        });
      }
    } catch (reason) {
      // 用户在生成方案途中打断：静默退出，已吐出的 step 留作上下文。
      if (isAbortError(reason)) {
        setStatus("idle");
        return;
      }
      setPendingImages(sentImages);
      if (isLocalServiceConnectionError(reason)) {
        markServerOffline();
      } else {
        markServerOnline();
      }
      appendMessage({
        role: "system",
        text:
          reason instanceof Error
            ? reason.message
            : "没有连接到本地 AI 服务，请确认服务已经启动。"
      });
    } finally {
      setStatus("idle");
    }
  }

  async function executeRequestedDataTool(
    intent: Extract<IntentCheckResponse, { kind: "tool_request" }>,
    sentImages: PendingImage[],
    expectedScopeFingerprint: string,
    clarificationMessageId?: string,
    correctionAttempt = 0
  ) {
    if (!workbook) return;
    if (sourceMode === "folder") {
      if (!folderCatalog) return;
      setStatus("tooling");
      advanceActivity(
        "正在读取文件夹完整数据",
        "pandas 只会读取本次已选择的文件和工作表。"
      );
      try {
        const result = await executeFolderQuery(
          folderCatalog.sessionId,
          intent.request
        );
        await analyzeConfirmedIntent(
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
          intent.turnId
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
        finishActivity();
      }
      return;
    }
    if (sourceMode !== "workbook" || !isRunningInExcel()) {
      await analyzeConfirmedIntent(
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
        intent.turnId
      );
      return;
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
        return;
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
          return;
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
      await analyzeConfirmedIntent(
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
        intent.turnId
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
        return;
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
            modelId: selectedModelId || null
          }, turnAbortRef.current?.signal);
          await continueIntentDecision(
            corrected,
            intent.confirmedPrompt,
            sentImages,
            expectedScopeFingerprint,
            clarificationMessageId,
            1,
            correctionAttempt + 1
          );
          return;
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
    correctionAttempt = 0
  ) {
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
          originalPrompt,
          scopeFingerprint,
          hadImages: sentImages.length > 0,
          round: clarificationRound,
          status: "pending"
        }
      });
      return;
    }
    if (intent.kind === "tool_request") {
      await executeRequestedDataTool(
        intent,
        sentImages,
        scopeFingerprint,
        clarificationMessageId,
        correctionAttempt
      );
      return;
    }
    await analyzeConfirmedIntent(
      intent.confirmedPrompt,
      sentImages,
      scopeFingerprint,
      clarificationMessageId,
      [],
      undefined,
      {
        confirmedPrompt: intent.confirmedPrompt
      },
      intent.turnId
    );
  }

  async function resolveClarification(
    message: ChatMessage,
    resolution: string,
    label: string
  ) {
    const clarification = message.clarification;
    if (!clarification || clarification.status !== "pending" || busy) return;
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
    setMessages((current) =>
      current.map((item) =>
        item.id === message.id && item.clarification
          ? {
              ...item,
              clarification: {
                ...item.clarification,
                status: "resolving"
              }
            }
          : item
      )
    );
    setPendingImages([]);
    setImageError("");
    const confirmedPrompt =
      `${clarification.originalPrompt}\n\n用户确认：${resolution}`;
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
      const intent = await checkIntent({
        turnId: clarification.turnId,
        prompt: confirmedPrompt,
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
        modelId: selectedModelId || null
      }, turnController.signal);
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
      setStatus("idle");
      const decision = describeIntentDecision(intent);
      advanceActivity(
        decision.label,
        decision.note ?? "正在根据确认结果继续。",
        decision.label,
        { note: decision.note, detail: decision.detail }
      );
      await continueIntentDecision(
        intent,
        confirmedPrompt,
        sentImages,
        clarification.scopeFingerprint,
        message.id,
        Math.min(
          MAX_CLARIFICATION_ROUNDS,
          clarification.round + 1
        )
      );
      finishActivity();
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
        setStatus("idle");
        finishActivity();
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

  // 纯停止：掐掉当前 turn，不再重跑。清掉待转向文本与被打断意图，避免误触发重跑或串味。
  function stopTurn() {
    pendingSteerRef.current = null;
    steerBasePromptRef.current = null;
    turnAbortRef.current?.abort();
    queryAbortRef.current?.abort();
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
        .map((sheet) => folderSheetKey(sheet.sourceFileId ?? "", sheet.name)) ??
        []
    );
    const snapshot =
      workbook !== null &&
      snapshotKeys.size > 0 &&
      [...selectedKeys].every((key) => snapshotKeys.has(key))
        ? workbook
        : await createFolderSnapshot(
            folderCatalog.sessionId,
            folderSelections()
          );
    const extraSheets: FormulaExtraSheet[] = [];
    for (const sheet of snapshot.worksheets) {
      const fileId = sheet.sourceFileId ?? "";
      if (!selectedKeys.has(folderSheetKey(fileId, sheet.name))) continue;
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
        sheetName: sheet.name,
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
          modelId: selectedModelId || null
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
      markFunctionPreview(message.id, {
        targetError: reason instanceof Error ? reason.message : "生成公式失败"
      });
    } finally {
      finishActivity();
      setStatus("idle");
    }
  }

  // 用户觉得预填的跨表匹配不对：切到模型兜底，用原描述直接生成。
  async function switchToModelFunction(message: ChatMessage) {
    if (!message.functionPreview || busy) return;
    markFunctionPreview(message.id, {
      mode: "model",
      match: undefined,
      targetError: undefined
    });
    await confirmFunctionTarget(message, { forceModel: true });
  }

  async function sendMessage(options?: {
    forceRecompute?: boolean;
    overrideText?: string;
  }) {
    const enteredText = (options?.overrideText ?? prompt).trim();
    const text =
      enteredText ||
      (pendingImages.length > 0
        ? "请结合附件图片分析当前工作簿，并说明发现的问题。"
        : "");
    if (!workbook || !text || text.length < 2 || busy) return;
    // /function 短链：绕开 planner（checkIntent/streamAssistantResponse），
    // 命中式单发生成原生公式 + 首格真实试算 + 预览卡。
    if (enteredText.startsWith("/function ")) {
      const description = enteredText.slice("/function ".length).trim();
      if (description) {
        await handleFunctionCommand(description);
        return;
      }
    }
    // 转向重跑：把被打断那句的意图并进这句，让 checkIntent 收到完整承接式需求。
    // 仅影响送模型与缓存的意图文本（intentText），气泡仍只显示用户实际输入（text）。
    const steerBasePrompt = steerBasePromptRef.current;
    steerBasePromptRef.current = null;
    const intentText =
      steerBasePrompt && steerBasePrompt !== text
        ? `${steerBasePrompt}；补充：${text}`
        : text;
    // "仍要重新计算"绕过整个命中判定，并在重算后覆写缓存。
    forceRecomputeRef.current = options?.forceRecompute === true;
    rawPromptRef.current = intentText;
    if (pendingImages.length > 0 && !supportsVision) {
      setImageError("当前模型不支持图片，请切换模型或移除附件。");
      return;
    }

    const pendingClarification = [...messages]
      .reverse()
      .find(
        (message) =>
          message.clarification?.status === "pending"
      );
    if (pendingClarification && enteredText) {
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
      return;
    }
    if (sourceMode === "folder" && !selectionConfirmed) {
      setContextOpen(true);
      appendMessage({
        role: "system",
        text: "请先在上方“数据范围”中选择文件夹和工作表。"
      });
      return;
    }
    if (!selectionConfirmed) setSelectionConfirmed(true);

    let intentWorkbook = workbook;
    if (sourceMode === "workbook" && isRunningInExcel()) {
      setStatus("planning");
      beginActivity(
        "正在扫描字段结构",
        `正在本地检查 ${selectedNamesFor(workbook).length} 张工作表；不会上传数据行。`
      );
      try {
        intentWorkbook = await captureWorkbookStructure(
          workbookScopeMode === "manual" ? selectedSheetNames : undefined
        );
        setWorkbook(intentWorkbook);
      } catch (reason) {
        appendMessage({
          role: "system",
          text:
            reason instanceof Error
              ? `读取工作表结构失败：${reason.message}`
              : "读取工作表结构失败"
        });
        setStatus("idle");
        finishActivity();
        return;
      }
    }
    const effectiveSheetNames = selectedNamesFor(intentWorkbook);
    setSelectedSheetNames(effectiveSheetNames);
    const intentScope = buildIntentScope(
      intentWorkbook,
      effectiveSheetNames
    );
    const scopeFingerprint = intentScopeFingerprint(intentScope);
    const sentImages = pendingImages;
    appendMessage({
      role: "user",
      text: enteredText || "请结合附件图片分析当前工作簿。",
      attachmentNames: sentImages.map((image) => image.name)
    });
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
        setStatus("idle");
        return;
      }
    }
    setStatus("planning");
    // 开启一个新的 turn：建 controller 存进 ref，串到本 turn 所有模型请求上。
    const turnController = new AbortController();
    turnAbortRef.current = turnController;
    if (!(sourceMode === "workbook" && isRunningInExcel())) {
      beginActivity(
        "正在准备需求上下文",
        `已选择 ${effectiveSheetNames.length} 张工作表。`
      );
    }
    advanceActivity(
      "正在确认需求",
      "正在结合字段结构、最近对话和上一轮紧凑结果判断是否需要本地工具。",
      `已识别 ${intentScope.sheets.length} 张工作表的字段结构`
    );
    try {
      const intent = await checkIntent({
        prompt: intentText,
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
        modelId: selectedModelId || null
      }, turnController.signal);
      markServerOnline();
      setStatus("idle");
      const decision = describeIntentDecision(intent);
      advanceActivity(
        decision.label,
        decision.note ?? "正在根据确认结果继续。",
        decision.label,
        { note: decision.note, detail: decision.detail }
      );
      await continueIntentDecision(
        intent,
        intentText,
        sentImages,
        scopeFingerprint
      );
      finishActivity();
    } catch (reason) {
      // 用户主动打断：静默收尾，不弹错误。转向时新 turn 已在 steerTurn 里另起。
      if (isAbortError(reason)) {
        setStatus("idle");
        finishActivity();
        return;
      }
      setPendingImages(sentImages);
      if (isLocalServiceConnectionError(reason)) {
        markServerOffline();
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

  async function runPlan(
    plan: AnalysisPlan,
    allowedExternalFiles?: ReadonlySet<string>
  ) {
    if (sourceMode === "folder") {
      if (!folderCatalog) return;
      setStatus("executing");
      try {
        const result = await executeFolderPlan(folderCatalog.sessionId, plan);
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
          text: `「${plan.title}」已执行 ${
            result.actionResults.length
          } 步${verificationSummary(result.verification)}。已写入：${result.filesModified.join("、")}${
            result.backups.length > 0
              ? `；已备份：${result.backups.join("、")}`
              : ""
          }`,
          verification: result.verification
        });
        if (result.verification.status === "verified") {
          markPlanVerified(plan.id);
        }
      } catch (reason) {
        appendMessage({
          role: "system",
          text: reason instanceof Error ? reason.message : "执行文件夹计划失败"
        });
      } finally {
        setStatus("idle");
      }
      return;
    }

    if (!isRunningInExcel()) {
      appendMessage({
        role: "system",
        text: "当前是浏览器演示模式，不会写入文件。请在 Excel 中打开任务窗格。"
      });
      return;
    }

    setStatus("executing");
    try {
      const result = await executePlan(plan, { allowedExternalFiles });
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
            ? `已记录 ${result.undoSnapshot.ranges.length} 项本次执行撤销数据。`
            : ""
        }原始工作表没有被删除或清空。`,
        verification: result.verification,
        executedPlanId: plan.id
      });
      setLastUndoSnapshot(result.undoSnapshot ?? null);
      if (result.verification.status === "verified") {
        markPlanVerified(plan.id);
      }
      await scan();
    } catch (reason) {
      if (reason instanceof PlanExecutionError) {
        const succeeded = reason.actionResults.filter(
          (result) => result.status === "succeeded"
        ).length;
        const failed = reason.actionResults.find(
          (result) => result.status === "failed"
        );
        const notRun = reason.actionResults.filter(
          (result) => result.status === "not_run"
        ).length;
        appendMessage({
          role: "system",
          text: `${reason.message}。${
            succeeded > 0
              ? `已有 ${succeeded} 步成功写入当前工作簿；`
              : "本次没有步骤成功写入；"
          }${failed ? `失败步骤：${failed.type}（${failed.sheet}）；` : ""}${
            notRun > 0 ? `其余 ${notRun} 步未执行。` : ""
          }`
        });
        return;
      }
      appendMessage({
        role: "system",
        text: reason instanceof Error ? reason.message : "执行计划失败"
      });
    } finally {
      setStatus("idle");
    }
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
    const preview = message.functionPreview;
    if (
      !preview ||
      preview.phase !== "preview" ||
      preview.applied ||
      preview.cancelled ||
      busy
    )
      return;
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
    markFunctionPreview(message.id, {
      applied: true,
      appliedTarget: resolved.address,
      targetError: undefined,
      pickingTarget: false
    });
    await runPlan(
      plan,
      preview.externalFiles ? new Set(preview.externalFiles) : undefined
    );
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
      { height: 90, width: 90, displayInIframe: false },
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

  function openTools() {
    setModelMenuOpen(false);
    setMoreMenuOpen(false);
    setHistoryOpen(false);
    closeSettings();
    resetToolDrawer();
    setToolsOpen(true);
  }

  async function prepareToolRun(tool: SavedTool) {
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
    selectModel(modelId);
    setModelMenuOpen(false);
    if (modelId === "local") {
      try {
        if (localStorage.getItem(BASE_MODE_HELP_SHOWN_KEY) !== "1") {
          localStorage.setItem(BASE_MODE_HELP_SHOWN_KEY, "1");
          appendMessage({
            role: "assistant",
            text: BASE_MODE_HELP_TEXT
          });
        }
      } catch {
        // localStorage 不可用时跳过持久化，避免阻塞模型切换。
      }
    }
  }

  async function previewTool(tool: SavedTool) {
    try {
      const plan = instantiateTool(tool, toolParameterValues, workbook);
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
    if (!workbook) return;
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
          return executeFolderQuery(folderCatalog.sessionId, request);
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
          text: BASE_MODE_HELP_TEXT
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
        steerTurn(prompt);
        return;
      }
      void sendMessage();
    }
  }

  function maximumComposerHeight(): number {
    return Math.max(
      COMPOSER_MIN_HEIGHT,
      Math.min(360, Math.floor(window.innerHeight * 0.45))
    );
  }

  function startComposerResize(event: React.PointerEvent<HTMLDivElement>) {
    const input = composerInputRef.current;
    if (!input) return;
    event.preventDefault();
    event.currentTarget.setPointerCapture(event.pointerId);
    composerResizeRef.current = {
      pointerId: event.pointerId,
      startY: event.clientY,
      startHeight: input.getBoundingClientRect().height
    };
  }

  function moveComposerResize(event: React.PointerEvent<HTMLDivElement>) {
    const resize = composerResizeRef.current;
    if (!resize || resize.pointerId !== event.pointerId) return;
    const nextHeight = resize.startHeight + resize.startY - event.clientY;
    setComposerHeight(
      Math.max(
        COMPOSER_MIN_HEIGHT,
        Math.min(maximumComposerHeight(), Math.round(nextHeight))
      )
    );
  }

  function finishComposerResize(event: React.PointerEvent<HTMLDivElement>) {
    if (composerResizeRef.current?.pointerId !== event.pointerId) return;
    composerResizeRef.current = null;
    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId);
    }
  }

  function resizeComposerWithKeyboard(
    event: React.KeyboardEvent<HTMLDivElement>
  ) {
    if (event.key !== "ArrowUp" && event.key !== "ArrowDown") return;
    event.preventDefault();
    const current =
      composerHeight ??
      composerInputRef.current?.getBoundingClientRect().height ??
      COMPOSER_MIN_HEIGHT;
    const direction = event.key === "ArrowUp" ? 20 : -20;
    setComposerHeight(
      Math.max(
        COMPOSER_MIN_HEIGHT,
        Math.min(maximumComposerHeight(), Math.round(current + direction))
      )
    );
  }

  return (
    <main className="chat-shell">
      <header className="chat-header">
        <div className="brand-mark">EB</div>
        <div className="brand-copy">
          <strong className="app-logo" {...logoLongPress}>
            Excel Bro
          </strong>
          <span className={headerStatusClassName}>
            <i />
            {serverOnline && isBaseMode && (
              <svg
                className="base-mode-icon"
                viewBox="0 0 16 16"
                aria-hidden="true"
                focusable="false"
              >
                <circle cx="8" cy="8" r="6.25" />
                <path d="M5.5 8h5" />
              </svg>
            )}
            {headerStatusText}
            {serverOnline && isBaseMode && (
              <button
                type="button"
                className="base-mode-help-button"
                title="查看基础模式帮助"
                aria-label="查看基础模式帮助"
                onClick={() => handleSlashCommandSelect("help")}
              >
                <svg viewBox="0 0 16 16" aria-hidden="true" focusable="false">
                  <circle cx="8" cy="8" r="6.25" />
                  <path d="M6.15 6.15c.14-1.05 1.08-1.78 2.12-1.62 1.02.16 1.72 1.02 1.67 2.02-.04.83-.54 1.42-1.36 1.86-.38.2-.47.38-.43.78" />
                  <path d="M8 11.2h.01" />
                </svg>
              </button>
            )}
          </span>
        </div>
        <div
          className={`model-picker${showModelGuide ? " needs-model" : ""}`}
        >
          <button
            type="button"
            className="model-picker-trigger"
            disabled={!serverOnline || busy}
            aria-haspopup="menu"
            aria-expanded={modelMenuOpen}
            title="选择或添加模型"
            onClick={() => {
              setHistoryOpen(false);
              setToolsOpen(false);
              setMoreMenuOpen(false);
              closeSettings();
              setModelMenuOpen((current) => !current);
            }}
          >
            <span>{hasConfiguredModel ? "模型" : "添加模型"}</span>
            <svg viewBox="0 0 16 16" aria-hidden="true">
              <path d="m4 6 4 4 4-4" />
            </svg>
          </button>
          {modelMenuOpen && (
            <div className="model-menu" role="menu" aria-label="模型">
              <span className="model-menu-title">选择模型</span>
              {modelOptions.map((option) => (
                <button
                  type="button"
                  role="menuitemradio"
                  aria-checked={option.id === (selectedModelId || "local")}
                  key={option.id}
                  disabled={!option.available}
                  onClick={() => handleSelectModel(option.id)}
                >
                  {option.id === "local" ? (
                    <svg
                      className="base-mode-icon model-option-icon"
                      viewBox="0 0 16 16"
                      aria-hidden="true"
                      focusable="false"
                    >
                      <circle cx="8" cy="8" r="6.25" />
                      <path d="M5.5 8h5" />
                    </svg>
                  ) : (
                    <i />
                  )}
                  <span className="model-option-copy">
                    <span>{option.label}</span>
                    {option.id === "local" && (
                      <small>{BASE_MODE_DESCRIPTION}</small>
                    )}
                  </span>
                  {option.id === (selectedModelId || "local") && <b>✓</b>}
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
        <div className="header-actions">
          <button
            className="header-button labeled-header-button new-chat-entry"
            onClick={handleNewChat}
            disabled={busy}
            title="新对话"
            aria-label="新对话"
          >
            <span aria-hidden="true">＋</span>
            <span>新对话</span>
          </button>
          <button
            className="header-button labeled-header-button tools-entry"
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
            <span>工具</span>
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
                    <strong>{focusOpening ? "正在打开…" : "专注窗口"}</strong>
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
                正在读取本地模型连接…
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
                        {settingsSaving ? "正在保存…" : "保存并立即生效"}
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
                        selectedModelId === connection.catalogModelId
                          ? " current"
                          : ""
                      }`}
                      key={connection.id}
                    >
                      <div>
                        <div className="connection-card-title">
                          <strong>{connection.label}</strong>
                          {selectedModelId === connection.catalogModelId && (
                            <b>当前使用</b>
                          )}
                        </div>
                        <span>{connection.modelId}</span>
                        <small>{connection.baseUrl}</small>
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
                  <p>
                    兼容 OpenAI /chat/completions
                    的服务均可接入。可以先测试，确认无误后再保存。
                  </p>
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
                      placeholder="https://api.example.com/v1"
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
                      {settingsTesting ? "正在测试…" : "测试连接"}
                    </button>
                    <button
                      type="button"
                      className="tool-preview-button"
                      disabled={
                        settingsSaving || settingsTesting || !serverOnline
                      }
                      onClick={() => void saveConnection()}
                    >
                      {settingsSaving
                        ? "正在保存…"
                        : connectionDraft.id
                          ? "保存修改"
                          : "添加并使用"}
                    </button>
                  </div>
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
                      <option key={option.id} value={option.id}>
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
                              <span>{tool.description}</span>
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
                            <span>{tool.description}</span>
                          </span>
                        </button>
                        <div className="tool-card-footer">
                          <span className="tool-card-meta">
                            模型调用 0 次
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
                      </>
                    ) : (
                      <div className="tool-expert-view">
                        <div className="tool-expert-heading">
                          <div>
                            <small>CONTROLLED PLAN</small>
                            <strong>专家脚本</strong>
                            <span>由保存的白名单计划确定性生成</span>
                          </div>
                          <button
                            className="tool-copy-dsl"
                            onClick={() => void copyToolDsl(tool)}
                          >
                            {copiedToolDslId === tool.id ? "已复制" : "复制脚本"}
                          </button>
                        </div>
                        <div className="tool-expert-badges">
                          <span>只读展示</span>
                          <span>运行前预览</span>
                          <span>禁止任意代码</span>
                        </div>
                        <pre className="tool-dsl" tabIndex={0}>
                          <code>{renderToolDsl(tool.planTemplate)}</code>
                        </pre>
                        <p className="tool-expert-note">
                          这里和普通视图是同一份工具。复制脚本不会执行；实际运行仍使用受控
                          AnalysisPlan，并在写入前让你确认预览。
                        </p>
                      </div>
                    )}
                    <div className="tool-run-actions">
                      <button
                        className="tool-preview-button"
                        disabled={busy || !workbook}
                        onClick={() => void prepareToolRun(tool)}
                      >
                        使用这个工具
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
                      disabled={busy || !workbook}
                      onClick={() => void runQueryTool(tool)}
                    >
                      运行这个查询
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
                      <section className="tool-guide-step">
                        <span className="tool-guide-number">1</span>
                        <div className="tool-guide-content">
                          <div className="tool-guide-title">
                            <div>
                              <strong>选择数据来源</strong>
                              <span>这次要处理哪张工作表？</span>
                            </div>
                            <small>需要确认</small>
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
                            <small>最后确认</small>
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
        <div className="tool-dialog-backdrop" role="presentation">
          <section
            className="tool-dialog"
            role="dialog"
            aria-modal="true"
            aria-label="保存为工具"
          >
            <div className="tool-dialog-title">
              <div>
                <strong>保存为工具</strong>
                <span>验证后可从“我的工具”重复使用</span>
              </div>
              <button onClick={() => setSaveCandidate(null)}>×</button>
            </div>
            <label>
              <span>工具名称</span>
              <input
                value={toolName}
                maxLength={60}
                autoFocus
                onChange={(event) => setToolName(event.target.value)}
              />
            </label>
            <label>
              <span>用途说明</span>
              <textarea
                value={toolDescription}
                maxLength={240}
                onChange={(event) => setToolDescription(event.target.value)}
              />
            </label>
            <div className="tool-dialog-note">
              只保存白名单 Excel 操作，不会保存 API Key，也不会执行系统命令。
            </div>
            {saveEligibility && saveEligibility.issues.length > 0 && (
              <div className="tool-save-review">
                <strong>固化检查</strong>
                {saveEligibility.issues.map((issue) => (
                  <div
                    className={`tool-save-issue ${issue.severity}`}
                    key={issue.code}
                  >
                    <span>{issue.message}</span>
                    {issue.approval === "fixedContent" && (
                      <label>
                        <input
                          type="checkbox"
                          checked={approveFixedContent}
                          onChange={(event) =>
                            setApproveFixedContent(event.target.checked)
                          }
                        />
                        我确认以后仍写入这些固定内容
                      </label>
                    )}
                    {issue.approval === "destructive" && (
                      <label>
                        <input
                          type="checkbox"
                          checked={approveDestructive}
                          onChange={(event) =>
                            setApproveDestructive(event.target.checked)
                          }
                        />
                        我确认此工具包含覆盖或删除操作
                      </label>
                    )}
                  </div>
                ))}
              </div>
            )}
            <div className="tool-dialog-actions">
              <button
                className="secondary-button"
                onClick={() => setSaveCandidate(null)}
              >
                取消
              </button>
              <button
                className="primary-button"
                disabled={
                  !toolName.trim() ||
                  Boolean(saveEligibility?.blocked) ||
                  Boolean(
                    saveEligibility?.requiredApprovals.includes(
                      "fixedContent"
                    ) && !approveFixedContent
                  ) ||
                  Boolean(
                    saveEligibility?.requiredApprovals.includes(
                      "destructive"
                    ) && !approveDestructive
                  )
                }
                onClick={confirmSaveTool}
              >
                保存到我的工具
              </button>
            </div>
          </section>
        </div>
      )}

      <section className="message-stream" aria-live="polite">
        {messages.map((message) => (
          <article
            className={`message-row ${message.role}`}
            key={message.id}
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
                          : "基础模式"}
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
              {message.text && <p className="message-text">{message.text}</p>}
              {message.intentMemory?.toolRequest && message.resultContext && (
                <button
                  type="button"
                  className="secondary-button"
                  disabled={busy}
                  onClick={() => saveQueryFromMessage(message)}
                >
                  保存为固化查询
                </button>
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
                          ? "我需要确认你的需求"
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
                    <>
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
                                  option.label
                                );
                              }}
                            >
                              <strong>{option.label}</strong>
                              <span>{option.description}</span>
                            </button>
                          )
                        )}
                      </div>
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
                      <div className="clarification-actions">
                        <button
                          disabled={busy}
                          onClick={() => editClarificationScope(message)}
                        >
                          修改数据范围
                        </button>
                        <button
                          disabled={busy}
                          onClick={() =>
                            void resolveClarification(
                              message,
                              "按用户原始需求继续；若仍需自行判断，请在结果中明确列出采用的假设。",
                              "按原话继续"
                            )
                          }
                        >
                          按原话继续
                        </button>
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
              {message.activityLog &&
                message.activityLog.steps.length > 0 && (
                  <details className="activity-log-card">
                    <summary>
                      执行过程 · {message.activityLog.steps.length} 步 ·{" "}
                      {formatStepElapsed(message.activityLog.totalMs)}
                    </summary>
                    <ul className="activity-steps">
                      {message.activityLog.steps.map((step, index) => (
                        <li key={`${step.label}-${index}`}>
                          <span
                            className="activity-step-check"
                            aria-hidden="true"
                          >
                            ✓
                          </span>
                          <div className="activity-step-body">
                            <span className="activity-step-label">
                              {step.label}
                            </span>
                            {step.note && (
                              <span className="activity-step-note">
                                {step.note}
                              </span>
                            )}
                            {step.detail && (
                              <details className="activity-step-detail">
                                <summary>查看详情</summary>
                                <pre>{step.detail}</pre>
                              </details>
                            )}
                          </div>
                          <span className="activity-step-time">
                            {formatStepElapsed(step.elapsedMs)}
                          </span>
                        </li>
                      ))}
                    </ul>
                  </details>
                )}
              {message.verification && (
                <div
                  className={`verification-card ${message.verification.status}`}
                >
                  <strong>
                    {message.verification.status === "verified"
                      ? "验证通过"
                      : message.verification.status === "executed_unverified"
                        ? "已执行，部分操作未独立验证"
                        : "执行完成，但验证未通过"}
                  </strong>
                  <span>
                    {
                      message.verification.checks.filter((check) => check.passed)
                        .length
                    }
                    /{message.verification.checks.length} 项符合预期
                    {message.verification.unverifiedActions.length > 0
                      ? `；${message.verification.unverifiedActions.length} 步缺少独立验收`
                      : ""}
                  </span>
                  {message.verification.status === "failed" && (
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
                  {message.verification.status ===
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
                </div>
              )}
              {message.executedPlanId &&
                lastUndoSnapshot?.planId === message.executedPlanId && (
                  <div className="undo-row">
                    <button
                      type="button"
                      className="secondary-button"
                      disabled={busy}
                      onClick={() => void undoLastExecution()}
                    >
                      ↶ 撤销本次执行
                    </button>
                    <span className="undo-row-hint">
                      只保留最近一次执行的撤销数据
                    </span>
                  </div>
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
                          ? "读取中…"
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
                        <button
                          className="secondary-button"
                          disabled={busy}
                          onClick={() => void switchToModelFunction(message)}
                        >
                          这个不对，换 AI 生成
                        </button>
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
                        {status === "planning" ? "正在生成…" : "确定并生成"}
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
                      <div className="inline-notes">
                        已写入 {message.functionPreview.appliedTarget}。
                      </div>
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
                              ? "读取中…"
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
                            {status === "executing" ? "正在写入…" : "确认写入"}
                          </button>
                        </div>
                      </>
                    )}
                  </div>
                )}
              {message.plan && (
                <div className="inline-plan">
                  <div className="inline-plan-title">
                    <div>
                      <span>准备就绪</span>
                      <strong>{message.plan.title}</strong>
                    </div>
                  </div>

                  <div className="inline-plan-buttons">
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
                    <button
                      className="primary-button"
                      disabled={busy}
                      onClick={() => void runPlan(message.plan!)}
                    >
                      {status === "executing" ? "正在执行…" : "执行到 Excel"}
                    </button>
                  </div>

                  <details className="plan-details">
                    <summary>
                      查看执行细节
                      <span>{message.plan.actions.length} 步</span>
                    </summary>

                    {message.plan.assumptions.length > 0 && (
                      <div className="inline-notes">
                        <strong>采用的信息</strong>
                        <ul>
                          {message.plan.assumptions.map((item) => (
                            <li key={item}>{item}</li>
                          ))}
                        </ul>
                      </div>
                    )}

                    {message.plan.warnings.length > 0 && (
                      <div className="inline-notes warning">
                        <strong>补充说明</strong>
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
                  </details>
                </div>
              )}
              {message.text && (
                <div className="message-copy-row">
                  <button
                    className={`message-copy-button ${
                      copiedMessageId === message.id ? "copied" : ""
                    }`}
                    type="button"
                    title={
                      copiedMessageId === message.id
                        ? "已复制"
                        : "复制这条消息"
                    }
                    aria-label={
                      copiedMessageId === message.id
                        ? "消息已复制"
                        : "复制这条消息"
                    }
                    onClick={() =>
                      void copyMessageText(message.id, message.text?.trim() ?? "")
                    }
                  >
                    {copiedMessageId === message.id ? (
                      <>
                        <svg
                          viewBox="0 0 24 24"
                          aria-hidden="true"
                          focusable="false"
                        >
                          <path d="m5 12 4 4L19 6" />
                        </svg>
                        <span>已复制</span>
                      </>
                    ) : (
                      <svg
                        viewBox="0 0 24 24"
                        aria-hidden="true"
                        focusable="false"
                      >
                        <rect x="8" y="8" width="11" height="11" rx="2" />
                        <path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h2" />
                      </svg>
                    )}
                  </button>
                </div>
              )}
            </div>
          </article>
        ))}

        {(status === "planning" || status === "tooling") && (
          <article className="message-row assistant">
            <div className="activity-card" aria-live="polite">
              <div className="activity-heading">
                <span className="activity-pulse" aria-hidden="true">
                  <i />
                  <i />
                  <i />
                </span>
                <strong>
                  {activity?.title ??
                    (status === "tooling"
                      ? "正在本地读取并计算"
                      : "正在理解工作簿")}
                </strong>
                <time>{activitySeconds} 秒</time>
              </div>
              {activity?.detail && <p>{activity.detail}</p>}
              {busy && (
                <button
                  type="button"
                  className="activity-stop-button"
                  onClick={stopTurn}
                >
                  停止
                </button>
              )}
              {activity && activity.completed.length > 0 && (
                <ul className="activity-steps">
                  {activity.completed.map((step, index) => (
                    <li key={`${step.label}-${index}`}>
                      <span className="activity-step-check" aria-hidden="true">
                        ✓
                      </span>
                      <div className="activity-step-body">
                        <span className="activity-step-label">{step.label}</span>
                        {step.note && (
                          <span className="activity-step-note">{step.note}</span>
                        )}
                        {step.detail && (
                          <details className="activity-step-detail">
                            <summary>查看详情</summary>
                            <pre>{step.detail}</pre>
                          </details>
                        )}
                      </div>
                      <span className="activity-step-time">
                        {formatStepElapsed(step.elapsedMs)}
                      </span>
                    </li>
                  ))}
                </ul>
              )}
            </div>
          </article>
        )}
        <div ref={messageEndRef} />
      </section>

      <footer className="composer">
        {contextOpen && workbook && (
          <section className="scope-popover" aria-label="选择数据范围">
            <div className="scope-popover-header">
              <div>
                <strong>这次对话使用哪些数据？</strong>
                <span>默认跟随你当前正在查看的工作表</span>
              </div>
              <button
                onClick={() => setContextOpen(false)}
                aria-label="关闭数据范围选择"
              >
                ×
              </button>
            </div>

            {sourceMode === "workbook" && (
              <div className="workbook-identity">
                <i aria-hidden="true">XLSX</i>
                <span>
                  <small>当前文件</small>
                  <strong title={workbook.name}>{workbook.name}</strong>
                  <em>
                    {workbookDataPeriod
                      ? `报表日期 ${workbookDataPeriod}`
                      : "文件名中未识别到报表日期"}
                    {" · "}
                    当前工作表 {workbook.activeWorksheet}
                  </em>
                </span>
              </div>
            )}

            <div className="scope-mode-list">
              <button
                className={
                  sourceMode === "workbook" && workbookScopeMode === "auto"
                    ? "selected"
                    : ""
                }
                onClick={() => chooseAutomaticScope(workbook)}
              >
                <i>◎</i>
                <span>
                  <strong>跟随当前工作表</strong>
                  <small>发送时自动使用正在查看的工作表</small>
                </span>
              </button>
              <button
                className={
                  sourceMode === "workbook" && workbookScopeMode === "manual"
                    ? "selected"
                    : ""
                }
                onClick={() => chooseManualScope(workbook)}
              >
                <i>☷</i>
                <span>
                  <strong>选择多个工作表</strong>
                  <small>用于跨表查询、比较或汇总</small>
                </span>
              </button>
              <button
                className={sourceMode === "folder" ? "selected" : ""}
                onClick={() => chooseFolderScope()}
              >
                <i>⌑</i>
                <span>
                  <strong>选择文件夹</strong>
                  <small>批量处理多个本地工作簿</small>
                </span>
              </button>
            </div>

            {sourceMode === "workbook" && workbookScopeMode === "manual" && (
              <div className="workbook-sheet-picker">
                <div className="sheet-search-row">
                  <input
                    value={sheetSearch}
                    onChange={(event) => setSheetSearch(event.target.value)}
                    placeholder="搜索工作表…"
                    aria-label="搜索工作表"
                  />
                  <span>
                    已选 {selectedSheetNames.length}/{workbook.worksheets.length}
                  </span>
                </div>
                <div className="sheet-picker-toolbar">
                  <button
                    onClick={() => {
                      selectAllSheets(
                        filteredWorksheets.map((sheet) => sheet.name)
                      );
                    }}
                  >
                    全选搜索结果
                  </button>
                  <button
                    onClick={clearSelectedSheets}
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
                        onClick={() => toggleSheet(sheet.name)}
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
                            {sheet.name === workbook.activeWorksheet
                              ? "当前工作表 · "
                              : ""}
                            {sheet.rowCount} 行 · {sheet.columnCount} 列
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
            )}

            {sourceMode === "folder" && (
              <div className="folder-scope-picker">
                <button
                  className="browse-folder-button"
                  disabled={busy}
                  onClick={() => void browseFolder()}
                >
                  {folderCatalog
                    ? `更换文件夹 · ${folderCatalog.folderName}`
                    : "选择文件夹"}
                </button>

                {folderCatalog && (
                  <div className="folder-file-list">
                    {folderCatalog.files.map((file) => {
                      const selectableSheets = file.worksheets.filter(
                        (sheet) => !isEBSystemSheet(sheet.name)
                      );
                      return (
                        <div className="folder-file" key={file.id}>
                          <strong>{file.relativePath}</strong>
                          {file.error ? (
                            <small className="file-error">{file.error}</small>
                          ) : (
                            <>
                              <div className="sheet-picker-toolbar">
                                <button
                                  onClick={() => selectAllSheetsInFile(file.id)}
                                >
                                  全选本文件
                                </button>
                                <button
                                  onClick={() => clearSheetsInFile(file.id)}
                                >
                                  清空本文件
                                </button>
                              </div>
                              <div className="sheet-picker-options">
                                {selectableSheets.map((sheet) => {
                                  const selected = folderSheetKeys.includes(
                                    folderSheetKey(file.id, sheet.name)
                                  );
                                  return (
                                    <button
                                      key={sheet.name}
                                      className={selected ? "selected" : ""}
                                      onClick={() =>
                                        toggleFolderSheet(file.id, sheet.name)
                                      }
                                      aria-pressed={selected}
                                    >
                                      <i
                                        className="sheet-checkbox"
                                        aria-hidden="true"
                                      >
                                        {selected ? "✓" : ""}
                                      </i>
                                      <span>{sheet.name}</span>
                                      <small>
                                        {sheet.rowCount} 行 · {sheet.columnCount} 列
                                      </small>
                                    </button>
                                  );
                                })}
                                {selectableSheets.length === 0 && (
                                  <p>没有可选工作表</p>
                                )}
                              </div>
                            </>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            )}

            <div className="scope-popover-actions">
              {sourceMode === "workbook" ? (
                <button
                  disabled={busy}
                  onClick={() => void scan({ announce: true })}
                >
                  {status === "scanning" ? "读取中…" : "刷新工作簿"}
                </button>
              ) : (
                <span>已选 {folderSheetKeys.length} 个工作表</span>
              )}
              <button
                className="scope-confirm"
                disabled={
                  sourceMode === "workbook"
                    ? workbookScopeMode === "manual" &&
                      selectedSheetNames.length === 0
                    : folderSheetKeys.length === 0
                }
                onClick={() => void confirmSheetSelection()}
              >
                {status === "scanning" ? "读取中…" : "完成"}
              </button>
            </div>
          </section>
        )}

        <button
          className={`scope-trigger ${contextOpen ? "open" : ""}`}
          onClick={() => {
            setSheetSearch("");
            setContextOpen((value) => !value);
          }}
          aria-expanded={contextOpen}
        >
          <i aria-hidden="true">▦</i>
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
                    ? `文件夹 · 已选 ${folderSheetKeys.length} 个工作表`
                    : "选择文件夹"
                  : workbookScopeMode === "auto"
                    ? `${
                        workbookDataPeriod
                          ? `${workbookDataPeriod} · `
                          : ""
                      }当前表 ${workbook.activeWorksheet}`
                    : `${
                        workbookDataPeriod
                          ? `${workbookDataPeriod} · `
                          : ""
                      }已固定 ${selectedSheetNames.length} 个工作表`}
            </strong>
          </span>
          <b>{contextOpen ? "⌄" : "⌃"}</b>
        </button>

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
          <div
            className="composer-resize-handle"
            role="separator"
            aria-label="上下拖动调整输入框高度，双击恢复自动高度"
            aria-orientation="horizontal"
            tabIndex={0}
            onPointerDown={startComposerResize}
            onPointerMove={moveComposerResize}
            onPointerUp={finishComposerResize}
            onPointerCancel={finishComposerResize}
            onDoubleClick={() => setComposerHeight(null)}
            onKeyDown={resizeComposerWithKeyboard}
          >
            <i />
          </div>
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
                  : "可以先输入需求，工作簿读取完成后即可发送…"
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
            {imageError && (
              <div className="composer-image-error">{imageError}</div>
            )}
            <div className="composer-toolbar">
              <input
                ref={imageInputRef}
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
                  disabled={!supportsVision || busy}
                  onClick={() => imageInputRef.current?.click()}
                  title={
                    supportsVision
                      ? "添加图片，也可以直接粘贴或拖入截图"
                      : "当前模型不支持图片"
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
                  <span>图片</span>
                </button>
              </div>
              {busy ? (
                // 运行中：有字=转向（带话打断重跑），无字=纯停止。
                <button
                  className={`send-button ${
                    prompt.trim() ? "is-steer" : "is-stop"
                  }`}
                  onClick={() =>
                    prompt.trim() ? steerTurn(prompt) : stopTurn()
                  }
                  aria-label={prompt.trim() ? "打断并补充" : "停止"}
                  title={
                    prompt.trim()
                      ? "打断当前处理，带上这句话重新开始"
                      : "停止当前处理"
                  }
                >
                  {prompt.trim() ? "↑" : "■"}
                </button>
              ) : (
                <button
                  className="send-button"
                  disabled={
                    !workbook ||
                    (!prompt.trim() && pendingImages.length === 0) ||
                    (pendingImages.length > 0 && !supportsVision)
                  }
                  onClick={() => void sendMessage()}
                  aria-label="发送"
                >
                  ↑
                </button>
              )}
            </div>
          </div>
        </div>
        <span>
          {busy
            ? "运行中：输入新内容按 Enter 可打断并转向，留空点 ■ 停止"
            : "Enter 发送 · Shift + Enter 换行 · 可粘贴截图 · 写入操作会先预览"}
        </span>
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
