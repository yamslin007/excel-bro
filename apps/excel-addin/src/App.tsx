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
  checkHealth,
  checkIntent,
  streamAssistantResponse,
  createFolderSnapshot,
  deleteModelConnection,
  executeFolderPlan,
  executeFolderQuery,
  getModelSettings,
  isLocalServiceConnectionError,
  listModels,
  saveModelConnection,
  selectFolder,
  testModelConnection,
  updateModelSettings
} from "./api";
import type { ModelOption, ServiceHealth } from "./api";
import type {
  AnalysisPlan,
  FolderCatalog,
  FolderSelection,
  DataToolResult,
  ExecutionUndoSnapshot,
  IntentCheckResponse,
  IntentClarification,
  IntentMemory,
  IntentOption,
  IntentScopeContext,
  ModelSettings,
  QueryTableArguments,
  ResultContext,
  UpsertModelConnectionRequest,
  VerificationReport,
  WorkbookSnapshot
} from "./contracts";
import { demoWorkbook } from "./demo";
import {
  captureSelectionContext,
  captureWorkbook,
  captureWorkbookSourceFingerprint,
  captureWorkbookStructure,
  dataEpochsChanged,
  executePlan,
  isRunningInExcel,
  PlanExecutionError,
  snapshotDataEpochs,
  undoExecution,
  watchWorkbookStructureChanges
} from "./excel";
import {
  DataToolExecutionError,
  executeQueryTableTool
} from "./dataTools";
import {
  analyzeToolEligibility,
  createQueryTool,
  createTool,
  deleteTool,
  instantiateTool,
  loadQueryTools,
  loadTools,
  saveQueryTool,
  saveTool,
  type SavedQueryTool,
  type SavedTool,
  type ToolParameter
} from "./storage";
import {
  executeSavedQueryTool,
  SavedQueryToolFallbackError
} from "./deterministicTools";
import {
  MAX_IMAGE_ATTACHMENTS,
  prepareImageFile,
  type PendingImage
} from "./imageAttachments";
import { extractWorkbookDataPeriod } from "./workbookIdentity";
import { chooseAvailableModel } from "./modelSelection";
import capabilities from "../../../config/capabilities.json";
import {
  currentModelCallCount,
  exportDiagnosticReport,
  recordDiagnosticEvent
} from "./diagnostics";

type Status = "idle" | "scanning" | "planning" | "tooling" | "executing";
type MessageRole = "assistant" | "user" | "system";
type SourceMode = "workbook" | "folder";
type WorkbookScopeMode = "auto" | "manual";

interface ActivityStep {
  label: string;
  elapsedMs: number;
}

interface ActivityProgress {
  title: string;
  detail: string;
  completed: ActivityStep[];
  startedAt: number;
  lastStepAt: number;
}

interface ActivityLog {
  steps: ActivityStep[];
  totalMs: number;
}

interface ModelConnectionDraft {
  id: string | null;
  label: string;
  baseUrl: string;
  modelId: string;
  apiKey: string;
  clearApiKey: boolean;
  supportsVision: boolean;
}

interface MessageClarification extends IntentClarification {
  turnId?: string;
  originalPrompt: string;
  scopeFingerprint: string;
  hadImages?: boolean;
  round: number;
  status: "pending" | "resolving" | "resolved" | "cancelled" | "invalidated";
  resolvedLabel?: string;
}

interface ChatMessage {
  id: string;
  role: MessageRole;
  text?: string;
  attachmentNames?: string[];
  plan?: AnalysisPlan;
  resultContext?: ResultContext;
  intentMemory?: IntentMemory;
  verification?: VerificationReport;
  clarification?: MessageClarification;
  activityLog?: ActivityLog;
  reused?: boolean;
  provider?: "model" | "local";
  querySourceMode?: SourceMode;
  querySourceSheetNames?: string[];
  querySourceSheetIds?: string[];
  createdAt: string;
}

interface ChatConversation {
  id: string;
  title: string;
  messages: ChatMessage[];
  createdAt: string;
  updatedAt: string;
}

interface ChatHistoryState {
  activeConversationId: string;
  conversations: ChatConversation[];
}

const CHAT_STORAGE_KEY = "excel-bro.chat.v4";
const LEGACY_CHAT_STORAGE_KEY = "excel-bro.chat.v3";
const MODEL_STORAGE_KEY = "excel-bro.model.v2";
const PET_VISIBILITY_STORAGE_KEY = "excel-bro.pet.visibility.v1";
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

export function normalizePetVisibility(value: string | null): boolean {
  return value !== "hidden";
}

function emptyModelConnectionDraft(): ModelConnectionDraft {
  return {
    id: null,
    label: "",
    baseUrl: "",
    modelId: "",
    apiKey: "",
    clearApiKey: false,
    supportsVision: false
  };
}

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
  const [chatHistory, setChatHistory] =
    useState<ChatHistoryState>(loadChatHistory);
  const [status, setStatus] = useState<Status>("idle");
  const [activity, setActivity] = useState<ActivityProgress | null>(null);
  const [activitySeconds, setActivitySeconds] = useState(0);
  const [serverOnline, setServerOnline] = useState(false);
  const [serviceHealth, setServiceHealth] = useState<ServiceHealth | null>(null);
  const [modelOptions, setModelOptions] = useState<ModelOption[]>([
    {
      id: "local",
      label: "基础模式",
      provider: "local",
      available: true,
      supportsVision: false
    }
  ]);
  const [selectedModelId, setSelectedModelId] = useState(
    () => localStorage.getItem(MODEL_STORAGE_KEY) ?? ""
  );
  const [contextOpen, setContextOpen] = useState(false);
  const [sheetSearch, setSheetSearch] = useState("");
  const [composerHeight, setComposerHeight] = useState<number | null>(null);
  const [tools, setTools] = useState<SavedTool[]>(loadTools);
  const [queryTools, setQueryTools] =
    useState<SavedQueryTool[]>(loadQueryTools);
  const [toolsOpen, setToolsOpen] = useState(false);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [petVisible, setPetVisible] = useState(() =>
    normalizePetVisibility(
      localStorage.getItem(PET_VISIBILITY_STORAGE_KEY)
    )
  );
  const [modelMenuOpen, setModelMenuOpen] = useState(false);
  const [modelCatalogLoaded, setModelCatalogLoaded] = useState(false);
  const [modelGuideDismissed, setModelGuideDismissed] = useState(false);
  const [modelSettings, setModelSettings] = useState<ModelSettings | null>(null);
  const [apiKeyDraft, setApiKeyDraft] = useState("");
  const [showApiKey, setShowApiKey] = useState(false);
  const [settingsSaving, setSettingsSaving] = useState(false);
  const [settingsTesting, setSettingsTesting] = useState(false);
  const [settingsLoading, setSettingsLoading] = useState(false);
  const [settingsFeedback, setSettingsFeedback] = useState("");
  const [connectionDraft, setConnectionDraft] =
    useState<ModelConnectionDraft | null>(null);
  const [pendingDeleteConnectionId, setPendingDeleteConnectionId] =
    useState<string | null>(null);
  const [pendingDeleteConversationId, setPendingDeleteConversationId] =
    useState<string | null>(null);
  const [selectedToolId, setSelectedToolId] = useState<string | null>(null);
  const [toolParameterValues, setToolParameterValues] = useState<
    Record<string, string>
  >({});
  const [saveCandidate, setSaveCandidate] = useState<AnalysisPlan | null>(null);
  const [approveFixedContent, setApproveFixedContent] = useState(false);
  const [approveDestructive, setApproveDestructive] = useState(false);
  const [verifiedPlanIds, setVerifiedPlanIds] = useState<Set<string>>(
    () => new Set()
  );
  const [lastUndoSnapshot, setLastUndoSnapshot] =
    useState<ExecutionUndoSnapshot | null>(null);
  const [copiedMessageId, setCopiedMessageId] = useState<string | null>(null);
  const [toolName, setToolName] = useState("");
  const [toolDescription, setToolDescription] = useState("");
  const [selectedSheetNames, setSelectedSheetNames] = useState<string[]>([]);
  const [selectionConfirmed, setSelectionConfirmed] = useState(false);
  const [sourceMode, setSourceMode] = useState<SourceMode>("workbook");
  const [workbookScopeMode, setWorkbookScopeMode] =
    useState<WorkbookScopeMode>("auto");
  const [folderCatalog, setFolderCatalog] = useState<FolderCatalog | null>(null);
  const [folderSheetKeys, setFolderSheetKeys] = useState<string[]>([]);
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
  const copyFeedbackTimerRef = useRef<number | null>(null);
  const queryAbortRef = useRef<AbortController | null>(null);
  // 结论复用缓存（见 CachedConclusion）。intentKey → 结论；normalizedPrompt → intentKey。
  const resultCacheRef = useRef<Map<string, CachedConclusion>>(new Map());
  const promptKeyCacheRef = useRef<Map<string, string>>(new Map());
  // "仍要重新计算"按钮设为 true，绕过命中判定并在重算后覆写缓存。
  const forceRecomputeRef = useRef(false);
  // sendMessage 起点记录的原始用户文本，写缓存时作为一级 prompt key。
  const rawPromptRef = useRef("");

  const busy = status !== "idle";
  const activeConversation =
    chatHistory.conversations.find(
      (conversation) =>
        conversation.id === chatHistory.activeConversationId
    ) ?? chatHistory.conversations[0];
  const messages = activeConversation?.messages ?? [];
  const selectedModel = useMemo(
    () =>
      modelOptions.find(
        (option) => option.id === (selectedModelId || "local")
      ) ?? modelOptions[0],
    [modelOptions, selectedModelId]
  );
  const supportsVision = selectedModel?.supportsVision === true;
  const hasConfiguredModel = modelOptions.some(
    (option) => option.provider === "model" && option.available
  );
  const showFirstModelGuide =
    serverOnline &&
    modelCatalogLoaded &&
    !hasConfiguredModel &&
    !modelGuideDismissed &&
    !settingsOpen;
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
  const saveEligibility = useMemo(
    () => (saveCandidate ? analyzeToolEligibility(saveCandidate) : null),
    [saveCandidate]
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

  function beginActivity(title: string, detail: string) {
    const now = Date.now();
    setActivity({
      title,
      detail,
      completed: [],
      startedAt: now,
      lastStepAt: now
    });
    setActivitySeconds(0);
  }

  function togglePetVisibility() {
    setPetVisible((current) => {
      const next = !current;
      try {
        localStorage.setItem(
          PET_VISIBILITY_STORAGE_KEY,
          next ? "visible" : "hidden"
        );
      } catch {
        // The preference remains active for this session if storage is blocked.
      }
      return next;
    });
  }

  function advanceActivity(
    title: string,
    detail: string,
    completedStep?: string
  ) {
    setActivity((current) => {
      const now = Date.now();
      const startedAt = current?.startedAt ?? now;
      const lastStepAt = current?.lastStepAt ?? startedAt;
      const completed =
        completedStep && completedStep.trim()
          ? [
              ...(current?.completed ?? []),
              { label: completedStep, elapsedMs: now - lastStepAt }
            ]
          : current?.completed ?? [];
      return {
        title,
        detail,
        completed,
        startedAt,
        lastStepAt: completedStep && completedStep.trim() ? now : lastStepAt
      };
    });
  }

  // 把当前 activity 的已完成步骤固化成一条日志，挂到最近一条 assistant 消息上，
  // 让用户执行结束后仍能展开回看「做了哪些步骤、各花多久」。
  function persistActivityLog() {
    setActivity((current) => {
      if (current && current.completed.length > 0) {
        const log: ActivityLog = {
          steps: current.completed,
          totalMs: Date.now() - current.startedAt
        };
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
      return null;
    });
    setActivitySeconds(0);
  }

  function finishActivity() {
    persistActivityLog();
  }

  useEffect(() => {
    if (!activity) return;
    const updateElapsed = () =>
      setActivitySeconds(
        Math.max(0, Math.floor((Date.now() - activity.startedAt) / 1000))
      );
    updateElapsed();
    const timer = window.setInterval(updateElapsed, 1000);
    return () => window.clearInterval(timer);
  }, [activity?.startedAt]);

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
      if (copyFeedbackTimerRef.current !== null) {
        window.clearTimeout(copyFeedbackTimerRef.current);
      }
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

  function appendMessage(message: Omit<ChatMessage, "id" | "createdAt">) {
    setMessages((current) => [
      ...current,
      {
        ...message,
        id: messageId(),
        createdAt: new Date().toISOString()
      }
    ]);
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

  async function copyMessage(message: ChatMessage) {
    const text = message.text?.trim();
    if (!text) return;
    try {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(text);
      } else {
        const textarea = document.createElement("textarea");
        textarea.value = text;
        textarea.setAttribute("readonly", "");
        textarea.style.position = "fixed";
        textarea.style.opacity = "0";
        document.body.appendChild(textarea);
        textarea.select();
        const copied = document.execCommand("copy");
        textarea.remove();
        if (!copied) throw new Error("copy failed");
      }
      setCopiedMessageId(message.id);
      if (copyFeedbackTimerRef.current !== null) {
        window.clearTimeout(copyFeedbackTimerRef.current);
      }
      copyFeedbackTimerRef.current = window.setTimeout(() => {
        setCopiedMessageId((current) =>
          current === message.id ? null : current
        );
        copyFeedbackTimerRef.current = null;
      }, 1600);
    } catch {
      appendMessage({
        role: "system",
        text: "复制失败，请选中文字后手动复制。"
      });
    }
  }

  function setMessages(
    update:
      | ChatMessage[]
      | ((current: ChatMessage[]) => ChatMessage[])
  ) {
    setChatHistory((current) => {
      const now = new Date().toISOString();
      return {
        ...current,
        conversations: current.conversations.map((conversation) => {
          if (conversation.id !== current.activeConversationId) {
            return conversation;
          }
          const nextMessages =
            typeof update === "function"
              ? update(conversation.messages)
              : update;
          const trimmedMessages = nextMessages.slice(
            -MAX_MESSAGES_PER_CONVERSATION
          );
          return {
            ...conversation,
            title: conversationTitle(trimmedMessages),
            messages: trimmedMessages,
            updatedAt: now
          };
        })
      };
    });
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
      setSelectedSheetNames((current) => {
        if (sourceMode === "workbook" && workbookScopeMode === "auto") {
          return [snapshot.activeWorksheet];
        }
        const available = new Set(snapshot.worksheets.map((sheet) => sheet.name));
        const preserved = current.filter((name) => available.has(name));
        return preserved.length > 0 ? preserved : [snapshot.activeWorksheet];
      });
      setSelectionConfirmed(true);
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
    let active = true;
    async function refreshServiceState() {
      const health = await checkHealth();
      if (!active) return;
      setServiceHealth(health);
      setServerOnline(health !== null);
      if (!health) return;
      try {
        const catalog = await listModels();
        if (!active) return;
        setModelOptions(catalog.models);
        setModelCatalogLoaded(true);
        setSelectedModelId((current) => {
          const next = catalog.models.some(
            (option) => option.id === current && option.available
          )
            ? current
            : catalog.defaultModelId;
          localStorage.setItem(MODEL_STORAGE_KEY, next);
          return next;
        });
      } catch {
        // The health indicator remains useful if an older service lacks /api/models.
      }
    }

    void refreshServiceState();
    const intervalId = window.setInterval(() => {
      void refreshServiceState();
    }, 5000);
    const handleFocus = () => void refreshServiceState();
    window.addEventListener("focus", handleFocus);
    return () => {
      active = false;
      window.clearInterval(intervalId);
      window.removeEventListener("focus", handleFocus);
    };
  }, []);

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
          setLastUndoSnapshot(null);
        });
      }
    });
    return () => dispose?.();
  }, []);

  function newChat() {
    if (busy) return;
    setModelMenuOpen(false);
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
      setSaveCandidate(null);
      return;
    }
    const conversation = createConversation();
    setChatHistory((current) => ({
      activeConversationId: conversation.id,
      conversations: [
        conversation,
        ...current.conversations
      ].slice(0, MAX_STORED_CONVERSATIONS)
    }));
    setPrompt("");
    setPendingImages([]);
    setImageError("");
    setContextOpen(false);
    setHistoryOpen(false);
    setToolsOpen(false);
    closeSettings();
    setSaveCandidate(null);
  }

  function openConversation(conversationId: string) {
    if (busy) return;
    setChatHistory((current) => ({
      ...current,
      activeConversationId: conversationId
    }));
    setPrompt("");
    setPendingImages([]);
    setImageError("");
    setContextOpen(false);
    setHistoryOpen(false);
    closeSettings();
    setSaveCandidate(null);
  }

  function deleteConversation(conversationId: string) {
    if (busy) return;
    if (
      !chatHistory.conversations.some(
        (conversation) => conversation.id === conversationId
      )
    ) {
      return;
    }
    setPendingDeleteConversationId(conversationId);
  }

  function confirmDeleteConversation() {
    if (!pendingDeleteConversationId) return;
    setChatHistory((current) =>
      deleteConversationFromHistory(current, pendingDeleteConversationId)
    );
    setPendingDeleteConversationId(null);
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

  function toggleSheet(sheetName: string) {
    setSelectedSheetNames((current) =>
      current.includes(sheetName)
        ? current.filter((name) => name !== sheetName)
        : [...current, sheetName]
    );
    setSelectionConfirmed(false);
  }

  function folderSheetKey(fileId: string, sheetName: string) {
    return `${fileId}\u0000${sheetName}`;
  }

  function toggleFolderSheet(fileId: string, sheetName: string) {
    const key = folderSheetKey(fileId, sheetName);
    setFolderSheetKeys((current) =>
      current.includes(key)
        ? current.filter((item) => item !== key)
        : [...current, key]
    );
    setSelectionConfirmed(false);
  }

  async function browseFolder() {
    setStatus("scanning");
    try {
      const catalog = await selectFolder();
      if (!catalog) return;
      setFolderCatalog(catalog);
      setFolderSheetKeys([]);
      setSelectionConfirmed(false);
      setServerOnline(true);
    } catch (reason) {
      appendMessage({
        role: "system",
        text: reason instanceof Error ? reason.message : "读取文件夹失败"
      });
    } finally {
      setStatus("idle");
    }
  }

  function chooseAutomaticScope() {
    setSourceMode("workbook");
    setWorkbookScopeMode("auto");
    if (workbook) setSelectedSheetNames([workbook.activeWorksheet]);
    setSelectionConfirmed(true);
  }

  function chooseManualScope() {
    setSourceMode("workbook");
    setWorkbookScopeMode("manual");
    if (selectedSheetNames.length === 0 && workbook) {
      setSelectedSheetNames([workbook.activeWorksheet]);
    }
    setSelectionConfirmed(selectedSheetNames.length > 0 || workbook !== null);
  }

  function chooseFolderScope() {
    setSourceMode("folder");
    setSelectionConfirmed(folderSheetKeys.length > 0);
  }

  function folderSelections(): FolderSelection[] {
    if (!folderCatalog) return [];
    return folderCatalog.files
      .map((file) => ({
        fileId: file.id,
        sheets: file.worksheets
          .filter((sheet) =>
            folderSheetKeys.includes(folderSheetKey(file.id, sheet.name))
          )
          .map((sheet) => sheet.name)
      }))
      .filter((selection) => selection.sheets.length > 0);
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

  function selectedNamesFor(snapshot: WorkbookSnapshot): string[] {
    if (sourceMode === "workbook" && workbookScopeMode === "auto") {
      return [snapshot.activeWorksheet];
    }
    const available = new Set(snapshot.worksheets.map((sheet) => sheet.name));
    const selected = selectedSheetNames.filter((name) => available.has(name));
    return selected.length > 0 ? selected : [snapshot.activeWorksheet];
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
          onStep: (step) =>
            advanceActivity(
              step.title,
              step.detail ?? "",
              step.completedStep ?? undefined
            )
        }
      );
      setServerOnline(true);
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
      setPendingImages(sentImages);
      if (isLocalServiceConnectionError(reason)) {
        setServerOnline(false);
        setServiceHealth(null);
      } else {
        setServerOnline(true);
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
      const controller = new AbortController();
      queryAbortRef.current = controller;
      const result = await executeQueryTableTool(
        intent.request,
        effectiveSheetNames,
        selection.activeWorksheet,
        {
          signal: controller.signal,
          onProgress: ({ scannedRows, totalRows, sheet }) => {
            setActivity((current) =>
              current
                ? {
                    ...current,
                    detail: `正在读取「${sheet}」：${scannedRows.toLocaleString()} / ${totalRows.toLocaleString()} 行`
                  }
                : current
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
          });
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
      });
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
    if (!workbook || !text || busy) return;
    // "仍要重新计算"绕过整个命中判定，并在重算后覆写缓存。
    forceRecomputeRef.current = options?.forceRecompute === true;
    rawPromptRef.current = text;
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
        normalizePrompt(text)
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
        prompt: text,
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
      });
      setServerOnline(true);
      setStatus("idle");
      await continueIntentDecision(
        intent,
        text,
        sentImages,
        scopeFingerprint
      );
      finishActivity();
    } catch (reason) {
      setPendingImages(sentImages);
      if (isLocalServiceConnectionError(reason)) {
        setServerOnline(false);
        setServiceHealth(null);
      } else {
        setServerOnline(true);
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
    }
  }

  async function runPlan(plan: AnalysisPlan) {
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
          setVerifiedPlanIds((current) => new Set(current).add(plan.id));
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
      const result = await executePlan(plan);
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
        verification: result.verification
      });
      setLastUndoSnapshot(result.undoSnapshot ?? null);
      if (result.verification.status === "verified") {
        setVerifiedPlanIds((current) => new Set(current).add(plan.id));
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

  async function undoLastExecution() {
    if (!lastUndoSnapshot || busy || !isRunningInExcel()) return;
    if (!window.confirm("撤销上一次 Excel Bro 执行中可恢复的单元格更改？")) {
      return;
    }
    setStatus("executing");
    try {
      await undoExecution(lastUndoSnapshot);
      setLastUndoSnapshot(null);
      appendMessage({
        role: "system",
        text: "已撤销上一次执行中记录的单元格值、公式和常用格式。"
      });
      await scan();
    } catch (reason) {
      appendMessage({
        role: "system",
        text: reason instanceof Error ? reason.message : "撤销上一次执行失败"
      });
    } finally {
      setStatus("idle");
    }
  }

  function beginSaveTool(plan: AnalysisPlan) {
    setSaveCandidate(plan);
    setToolName(plan.title);
    setToolDescription(plan.summary);
    setApproveFixedContent(false);
    setApproveDestructive(false);
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
      setTools(saveTool(tool));
      setSaveCandidate(null);
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

  function selectTool(
    tool: SavedTool,
    snapshot: WorkbookSnapshot | null = workbook
  ) {
    setSelectedToolId(tool.id);
    const available = new Set(
      (snapshot?.worksheets ?? []).map((sheet) => sheet.name)
    );
    const fallback =
      snapshot?.activeWorksheet ?? snapshot?.worksheets[0]?.name ?? "";
    const suggestOutputName = (requested: string) => {
      if (
        ![...available].some(
          (name) =>
            name.toLocaleLowerCase() === requested.toLocaleLowerCase()
        )
      ) {
        return requested;
      }
      for (let index = 2; index < 1000; index += 1) {
        const suffix = ` (${index})`;
        const candidate = `${requested.slice(0, 31 - suffix.length)}${suffix}`;
        if (
          ![...available].some(
            (name) =>
              name.toLocaleLowerCase() === candidate.toLocaleLowerCase()
          )
        ) {
          return candidate;
        }
      }
      return requested.slice(0, 27) + " 副本";
    };
    const values: Record<string, string> = {};
    for (const parameter of tool.parameters) {
      if (parameter.type === "worksheet") {
        values[parameter.id] = available.has(parameter.defaultValue)
          ? parameter.defaultValue
          : fallback;
      } else if (parameter.type === "outputWorksheet") {
        values[parameter.id] = suggestOutputName(parameter.defaultValue);
      }
    }
    for (const parameter of tool.parameters) {
      if (parameter.type === "range") {
        const source = parameter.sourceParameterId
          ? values[parameter.sourceParameterId]
          : "";
        values[parameter.id] =
          snapshot?.worksheets.find((sheet) => sheet.name === source)
            ?.usedRange ??
          parameter.defaultValue;
      } else if (parameter.type !== "field") {
        continue;
      }
      if (parameter.type === "field") {
        const source = values[parameter.sourceParameterId];
        const headers = (
          snapshot?.worksheets.find((sheet) => sheet.name === source)
            ?.headers ?? []
        )
          .map((header) => String(header ?? "").trim())
          .filter(Boolean);
        values[parameter.id] = parameter.defaultValue;
        values[parameter.id] =
          headers.find(
            (header) =>
              header.toLocaleLowerCase() ===
              parameter.defaultValue.toLocaleLowerCase()
          ) ??
          headers[0] ??
          parameter.defaultValue;
      }
    }
    setToolParameterValues(values);
  }

  function fieldOptions(
    tool: SavedTool,
    parameter: Extract<ToolParameter, { type: "field" }>
  ): string[] {
    const sourceParameter = tool.parameters.find(
      (candidate) => candidate.id === parameter.sourceParameterId
    );
    const sourceSheet =
      toolParameterValues[parameter.sourceParameterId] ??
      sourceParameter?.defaultValue;
    const sheet = workbook?.worksheets.find(
      (candidate) => candidate.name === sourceSheet
    );
    const detected = [
      ...new Set(
        (sheet?.headers ?? [])
          .map((header) => String(header ?? "").trim())
          .filter(Boolean)
      )
    ];
    return detected.length > 0 ? detected : [parameter.defaultValue];
  }

  function updateToolParameter(
    tool: SavedTool,
    parameter: ToolParameter,
    value: string
  ) {
    setToolParameterValues((current) => {
      const next = { ...current, [parameter.id]: value };
      if (parameter.type !== "worksheet") return next;
      for (const candidate of tool.parameters) {
        if (
          candidate.type === "range" &&
          candidate.sourceParameterId === parameter.id
        ) {
          next[candidate.id] =
            workbook?.worksheets.find((sheet) => sheet.name === value)
              ?.usedRange ??
            candidate.defaultValue;
          continue;
        }
        if (
          candidate.type !== "field" ||
          candidate.sourceParameterId !== parameter.id
        ) {
          continue;
        }
        const headers = (
          workbook?.worksheets.find((sheet) => sheet.name === value)?.headers ??
          []
        )
          .map((header) => String(header ?? "").trim())
          .filter(Boolean);
        next[candidate.id] =
          headers.find(
            (header) =>
              header.toLocaleLowerCase() ===
              candidate.defaultValue.toLocaleLowerCase()
          ) ??
          headers[0] ??
          "";
      }
      return next;
    });
  }

  async function openTools() {
    setModelMenuOpen(false);
    setHistoryOpen(false);
    closeSettings();
    setToolsOpen(true);
    const selected =
      tools.find((tool) => tool.id === selectedToolId) ?? tools[0];
    if (!selected) return;
    if (!isRunningInExcel() || !workbook) {
      selectTool(selected);
      return;
    }
    setStatus("scanning");
    try {
      const snapshot = await captureWorkbookStructure(
        workbookScopeMode === "manual" ? selectedSheetNames : undefined
      );
      setWorkbook(snapshot);
      selectTool(selected, snapshot);
    } catch (reason) {
      appendMessage({
        role: "system",
        text:
          reason instanceof Error
            ? `读取工具所需字段失败：${reason.message}`
            : "读取工具所需字段失败"
      });
      selectTool(selected);
    } finally {
      setStatus("idle");
    }
  }

  function openHistory() {
    setModelMenuOpen(false);
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

  async function openSettings(): Promise<boolean> {
    setModelMenuOpen(false);
    setToolsOpen(false);
    setHistoryOpen(false);
    setSettingsOpen(true);
    setApiKeyDraft("");
    setShowApiKey(false);
    setSettingsFeedback("");
    setConnectionDraft(null);
    setPendingDeleteConnectionId(null);
    setSettingsLoading(true);
    try {
      setModelSettings(await getModelSettings());
      return true;
    } catch (reason) {
      setModelSettings(null);
      setSettingsFeedback(
        reason instanceof Error
          ? reason.message
          : "无法读取模型设置，请确认本地服务已经启动。"
      );
      return false;
    } finally {
      setSettingsLoading(false);
    }
  }

  async function openConnectionCreator() {
    setModelMenuOpen(false);
    if (await openSettings()) {
      setConnectionDraft(emptyModelConnectionDraft());
    }
  }

  function selectModel(modelId: string) {
    setSelectedModelId(modelId);
    localStorage.setItem(MODEL_STORAGE_KEY, modelId);
    setModelMenuOpen(false);
  }

  function dismissModelGuide() {
    setModelGuideDismissed(true);
  }

  async function saveApiKey() {
    const apiKey = apiKeyDraft.trim();
    if (!apiKey) {
      setSettingsFeedback("请输入新的 API Key。");
      return;
    }
    setSettingsSaving(true);
    setSettingsFeedback("");
    try {
      const saved = await updateModelSettings({ apiKey });
      setModelSettings(saved);
      setApiKeyDraft("");
      setShowApiKey(false);
      setSettingsFeedback("API Key 已保存并立即生效。");
      const [health, catalog] = await Promise.all([
        checkHealth(),
        listModels()
      ]);
      setServiceHealth(health);
      setServerOnline(health !== null);
      setModelOptions(catalog.models);
      setSelectedModelId((current) => {
        const next = catalog.models.some(
          (option) => option.id === current && option.available
        )
          ? current
          : catalog.defaultModelId;
        localStorage.setItem(MODEL_STORAGE_KEY, next);
        return next;
      });
    } catch (reason) {
      setSettingsFeedback(
        reason instanceof Error ? reason.message : "API Key 保存失败。"
      );
    } finally {
      setSettingsSaving(false);
    }
  }

  function editModelConnection(connectionId: string) {
    const connection = modelSettings?.connections.find(
      (item) => item.id === connectionId
    );
    if (!connection) return;
    setShowApiKey(false);
    setConnectionDraft({
      id: connection.id,
      label: connection.label,
      baseUrl: connection.baseUrl,
      modelId: connection.modelId,
      apiKey: "",
      clearApiKey: false,
      supportsVision: connection.supportsVision
    });
    setSettingsFeedback("");
  }

  async function refreshModelsAfterSettings(
    saved: ModelSettings,
    preferredModelId?: string
  ) {
    setModelSettings(saved);
    const [health, catalog] = await Promise.all([
      checkHealth(),
      listModels()
    ]);
    setServiceHealth(health);
    setServerOnline(health !== null);
    setModelOptions(catalog.models);
    setSelectedModelId((current) => {
      const next = chooseAvailableModel(
        catalog.models,
        current,
        catalog.defaultModelId,
        preferredModelId
      );
      localStorage.setItem(MODEL_STORAGE_KEY, next);
      return next;
    });
  }

  function connectionRequest(): UpsertModelConnectionRequest | null {
    if (!connectionDraft) return null;
    const label = connectionDraft.label.trim();
    const baseUrl = connectionDraft.baseUrl.trim();
    const modelId = connectionDraft.modelId.trim();
    if (!label || !baseUrl || !modelId) {
      setSettingsFeedback("请填写连接名称、服务地址和模型 ID。");
      return null;
    }
    return {
      id: connectionDraft.id,
      label,
      baseUrl,
      modelId,
      apiKey: connectionDraft.apiKey.trim() || null,
      clearApiKey: connectionDraft.clearApiKey,
      supportsVision: connectionDraft.supportsVision
    };
  }

  async function verifyConnection() {
    const request = connectionRequest();
    if (!request) return;
    setSettingsTesting(true);
    setSettingsFeedback("");
    try {
      const result = await testModelConnection(request);
      setSettingsFeedback(result.message);
    } catch (reason) {
      setSettingsFeedback(
        reason instanceof Error ? reason.message : "模型连接测试失败。"
      );
    } finally {
      setSettingsTesting(false);
    }
  }

  async function saveConnection() {
    const request = connectionRequest();
    if (!request || !connectionDraft) return;
    const wasCreating = !connectionDraft.id;
    const previousConnectionIds = new Set(
      modelSettings?.connections.map((connection) => connection.id) ?? []
    );
    setSettingsSaving(true);
    setSettingsFeedback("");
    try {
      const saved = await saveModelConnection(request);
      const createdConnection = wasCreating
        ? saved.connections.find(
            (connection) => !previousConnectionIds.has(connection.id)
          )
        : null;
      await refreshModelsAfterSettings(
        saved,
        createdConnection?.catalogModelId
      );
      setShowApiKey(false);
      setConnectionDraft(null);
      setSettingsFeedback(
        wasCreating
          ? "模型连接已添加，并已切换为当前模型。"
          : "模型连接已更新。"
      );
    } catch (reason) {
      setSettingsFeedback(
        reason instanceof Error ? reason.message : "模型连接保存失败。"
      );
    } finally {
      setSettingsSaving(false);
    }
  }

  async function removeConnection(connectionId: string) {
    setSettingsSaving(true);
    setSettingsFeedback("");
    try {
      const saved = await deleteModelConnection(connectionId);
      await refreshModelsAfterSettings(saved);
      setPendingDeleteConnectionId(null);
      setConnectionDraft((current) =>
        current?.id === connectionId ? null : current
      );
      setSettingsFeedback("模型连接已删除。");
    } catch (reason) {
      setSettingsFeedback(
        reason instanceof Error ? reason.message : "模型连接删除失败。"
      );
    } finally {
      setSettingsSaving(false);
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
        plan.sourceFingerprint = await captureWorkbookSourceFingerprint(
          sourceSheets
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
    setQueryTools(saveQueryTool(tool));
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

  function removeTool(tool: SavedTool) {
    if (!window.confirm(`确定删除工具「${tool.name}」吗？`)) return;
    const next = deleteTool(tool.id);
    setTools(next);
    const replacement = next[0] ?? null;
    setSelectedToolId(replacement?.id ?? null);
    if (replacement) selectTool(replacement);
  }

  function handleComposerKeyDown(event: React.KeyboardEvent<HTMLTextAreaElement>) {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
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
          <strong>Excel Bro</strong>
          <span
            className={
              serverOnline
                ? serviceHealth?.configured
                  ? "online model-online"
                  : "online"
                : ""
            }
          >
            <i />
            {!serverOnline
              ? "本地服务未连接"
              : (selectedModelId || serviceHealth?.model || "local") === "local"
                ? "本地服务已连接 · 基础模式"
                : `模型：${selectedModel?.label ?? serviceHealth?.model}`}
          </span>
        </div>
        <div
          className={`model-picker${showFirstModelGuide ? " needs-model" : ""}`}
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
              closeSettings();
              setModelMenuOpen((current) => !current);
            }}
          >
            <span>{hasConfiguredModel ? selectedModel?.label : "添加模型"}</span>
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
                  onClick={() => selectModel(option.id)}
                >
                  <i />
                  <span>{option.label}</span>
                  {option.id === (selectedModelId || "local") && <b>✓</b>}
                </button>
              ))}
              <div className="model-menu-actions">
                <button
                  type="button"
                  role="menuitem"
                  onClick={() => void openConnectionCreator()}
                >
                  ＋ 添加模型连接
                </button>
                <button
                  type="button"
                  role="menuitem"
                  onClick={() => void openSettings()}
                >
                  管理模型连接
                </button>
              </div>
            </div>
          )}
        </div>
        <button
          className="header-button labeled-header-button history-entry"
          onClick={openHistory}
          disabled={busy}
          title="历史对话"
          aria-label="历史对话"
        >
          <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
            <path d="M4.7 8.2A8 8 0 1 1 4 12" />
            <path d="M4.7 3.8v4.4H9M12 7.5V12l3 1.8" />
          </svg>
          <span>历史</span>
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
          title={petVisible ? "隐藏宠物" : "显示宠物"}
          aria-label={petVisible ? "隐藏宠物" : "显示宠物"}
          aria-pressed={petVisible}
        >
          <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
            <circle cx="7.1" cy="8.1" r="2.1" />
            <circle cx="12" cy="5.8" r="2.1" />
            <circle cx="16.9" cy="8.1" r="2.1" />
            <path d="M6.2 15.6c0-3.2 2.6-5.7 5.8-5.7s5.8 2.5 5.8 5.7c0 2-1.4 3.3-3.2 3.3-.9 0-1.8-.3-2.6-.9-.8.6-1.7.9-2.6.9-1.8 0-3.2-1.3-3.2-3.3Z" />
          </svg>
        </button>
        <button
          className="header-button"
          onClick={newChat}
          disabled={busy}
          title="新对话"
          aria-label="新对话"
        >
          ＋
        </button>
      </header>

      {showFirstModelGuide && (
        <section className="first-model-guide" aria-label="首次模型设置引导">
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
              onClick={() => void openConnectionCreator()}
            >
              现在添加
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
                    onClick={() => {
                      setShowApiKey(false);
                      setConnectionDraft(emptyModelConnectionDraft())
                    }}
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
              <strong>我的工具</strong>
              <span>已审核并固化的个人工作流</span>
            </div>
            <button onClick={() => setToolsOpen(false)} aria-label="关闭">
              ×
            </button>
          </div>
          {tools.length === 0 ? (
            <div className="tool-empty">
              <i>▦</i>
              <strong>还没有保存工具</strong>
              <span>在执行计划卡片中点击“保存为工具”。</span>
            </div>
          ) : (
            <div className="tool-layout">
              <div className="tool-list">
                {tools.map((tool) => (
                  <button
                    key={tool.id}
                    className={selectedToolId === tool.id ? "selected" : ""}
                    onClick={() => selectTool(tool)}
                  >
                    <strong>{tool.name}</strong>
                    <span>{tool.description}</span>
                  </button>
                ))}
              </div>
              {tools
                .filter((tool) => tool.id === selectedToolId)
                .map((tool) => (
                  <section className="tool-detail" key={tool.id}>
                    <div className="tool-detail-title">
                      <div>
                        <strong>{tool.name}</strong>
                        <span>版本 {tool.version} · {tool.planTemplate.actions.length} 步</span>
                      </div>
                      <button onClick={() => removeTool(tool)}>删除</button>
                    </div>
                    <p>{tool.description}</p>
                    {tool.parameters.length > 0 ? (
                      <div className="tool-parameters">
                        <strong>本次运行参数</strong>
                        {tool.parameters.map((parameter) => (
                          <label key={parameter.id}>
                            <span>{parameter.label}</span>
                            {parameter.type === "outputWorksheet" ||
                            parameter.type === "range" ? (
                              <input
                                value={
                                  toolParameterValues[parameter.id] ??
                                  parameter.defaultValue
                                }
                                placeholder={
                                  parameter.type === "outputWorksheet"
                                    ? "输入新的工作表名称"
                                    : "例如 A1:E812"
                                }
                                onChange={(event) =>
                                  updateToolParameter(
                                    tool,
                                    parameter,
                                    event.target.value
                                  )
                                }
                              />
                            ) : (
                              <select
                                value={
                                  toolParameterValues[parameter.id] ??
                                  parameter.defaultValue
                                }
                                onChange={(event) =>
                                  updateToolParameter(
                                    tool,
                                    parameter,
                                    event.target.value
                                  )
                                }
                              >
                                {parameter.type === "worksheet"
                                  ? (workbook?.worksheets ?? []).map((sheet) => (
                                      <option key={sheet.name} value={sheet.name}>
                                        {sheet.name}
                                      </option>
                                    ))
                                  : fieldOptions(tool, parameter).map((field) => (
                                      <option key={field} value={field}>
                                        {field}
                                      </option>
                                    ))}
                              </select>
                            )}
                          </label>
                        ))}
                      </div>
                    ) : (
                      <div className="tool-fixed-note">
                        此工具不需要运行参数，将按保存时的固定设置执行。
                      </div>
                    )}
                    <details className="tool-steps">
                      <summary>查看固化逻辑</summary>
                      <ol>
                        {tool.planTemplate.actions.map((action, index) => (
                          <li key={`${tool.id}-${index}`}>
                            {actionLabel(action)}
                          </li>
                        ))}
                      </ol>
                    </details>
                    <button
                      className="tool-preview-button"
                      disabled={busy || !workbook}
                      onClick={() => void previewTool(tool)}
                    >
                      生成执行预览
                    </button>
                  </section>
                ))}
            </div>
          )}
          {queryTools.length > 0 && (
            <section className="tool-detail">
              <div className="tool-detail-title">
                <div>
                  <strong>固化查询</strong>
                  <span>本地重复执行 · 模型调用 0 次</span>
                </div>
              </div>
              <p>字段或来源发生变化时会停止本地执行并要求重新确认。</p>
              {queryTools.map((tool) => (
                <button
                  key={tool.id}
                  className="tool-preview-button"
                  disabled={busy || !workbook}
                  onClick={() => void runQueryTool(tool)}
                >
                  运行 · {tool.name}
                </button>
              ))}
            </section>
          )}
        </aside>
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
                      onClick={() => openConversation(conversation.id)}
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
                      onClick={() => deleteConversation(conversation.id)}
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
            onClick={newChat}
            disabled={busy}
          >
            ＋ 新建对话
          </button>
        </aside>
      )}

      {pendingDeleteConversationId &&
        chatHistory.conversations
          .filter(
            (conversation) =>
              conversation.id === pendingDeleteConversationId
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
                    onClick={() => setPendingDeleteConversationId(null)}
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
                    onClick={() => setPendingDeleteConversationId(null)}
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
        {petVisible && (
          <section className="pet-home-placeholder" aria-label="宠物小屋">
            <div className="pet-home-mark" aria-hidden="true">
              <i />
              <i />
              <i />
              <b />
            </div>
            <div>
              <strong>宠物小屋准备中</strong>
              <span>角色方案已保留，选定后会住进这里</span>
            </div>
            <button
              type="button"
              onClick={togglePetVisibility}
              aria-label="隐藏宠物"
              title="隐藏宠物"
            >
              ×
            </button>
          </section>
        )}
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
                          <span className="activity-step-label">
                            {step.label}
                          </span>
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
                    onClick={() => void copyMessage(message)}
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
              {status === "tooling" && (
                <button
                  type="button"
                  onClick={() => queryAbortRef.current?.abort()}
                >
                  取消本地查询
                </button>
              )}
              {activity && activity.completed.length > 0 && (
                <ul className="activity-steps">
                  {activity.completed.map((step, index) => (
                    <li key={`${step.label}-${index}`}>
                      <span className="activity-step-check" aria-hidden="true">
                        ✓
                      </span>
                      <span className="activity-step-label">{step.label}</span>
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
                onClick={chooseAutomaticScope}
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
                onClick={chooseManualScope}
              >
                <i>☷</i>
                <span>
                  <strong>选择多个工作表</strong>
                  <small>用于跨表查询、比较或汇总</small>
                </span>
              </button>
              <button
                className={sourceMode === "folder" ? "selected" : ""}
                onClick={chooseFolderScope}
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
                      setSelectedSheetNames((current) => [
                        ...new Set([
                          ...current,
                          ...filteredWorksheets.map((sheet) => sheet.name)
                        ])
                      ]);
                      setSelectionConfirmed(false);
                    }}
                  >
                    全选搜索结果
                  </button>
                  <button
                    onClick={() => {
                      setSelectedSheetNames([]);
                      setSelectionConfirmed(false);
                    }}
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
                        <i aria-hidden="true">{selected ? "✓" : ""}</i>
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
                    {folderCatalog.files.map((file) => (
                      <div className="folder-file" key={file.id}>
                        <strong>{file.relativePath}</strong>
                        {file.error ? (
                          <small className="file-error">{file.error}</small>
                        ) : (
                          <div className="sheet-picker-options">
                            {file.worksheets.map((sheet) => {
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
                                  <span>{sheet.name}</span>
                                  <small>
                                    {sheet.rowCount} 行 · {sheet.columnCount} 列
                                  </small>
                                </button>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    ))}
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
            <small title={workbook?.name}>
              {workbook && sourceMode === "workbook"
                ? `当前文件 · ${workbook.name}`
                : "数据范围"}
            </small>
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
                      }当前工作表 · ${workbook.activeWorksheet}`
                    : `${
                        workbookDataPeriod
                          ? `${workbookDataPeriod} · `
                          : ""
                      }已固定选择 ${selectedSheetNames.length} 个工作表`}
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
            <textarea
              ref={composerInputRef}
              aria-label="给 Excel Bro 发消息"
              value={prompt}
              onChange={(event) => setPrompt(event.target.value)}
              onKeyDown={handleComposerKeyDown}
              onPaste={handleImagePaste}
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
              {lastUndoSnapshot && (
                <button
                  className="attach-image-button"
                  disabled={busy}
                  onClick={() => void undoLastExecution()}
                  title="撤销上一次 Excel Bro 执行"
                >
                  ↶ <span>撤销</span>
                </button>
              )}
              <button
                className="send-button"
                disabled={
                  busy ||
                  !workbook ||
                  (!prompt.trim() && pendingImages.length === 0) ||
                  (pendingImages.length > 0 && !supportsVision)
                }
                onClick={() => void sendMessage()}
                aria-label="发送"
              >
                ↑
              </button>
            </div>
          </div>
        </div>
        <span>
          Enter 发送 · Shift + Enter 换行 · 可粘贴截图 · 写入操作会先预览
        </span>
      </footer>
    </main>
  );
}
