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
  createAssistantResponse,
  createFolderSnapshot,
  executeFolderPlan,
  listModels,
  selectFolder
} from "./api";
import type { ModelOption, ServiceHealth } from "./api";
import type {
  AnalysisPlan,
  FolderCatalog,
  FolderSelection,
  DataToolResult,
  IntentCheckResponse,
  IntentClarification,
  IntentMemory,
  IntentOption,
  IntentScopeContext,
  ResultContext,
  VerificationReport,
  WorkbookSnapshot
} from "./contracts";
import { demoWorkbook } from "./demo";
import {
  captureSelectionContext,
  captureWorkbook,
  captureWorkbookStructure,
  executePlan,
  isRunningInExcel
} from "./excel";
import {
  DataToolExecutionError,
  executeQueryTableTool
} from "./dataTools";
import {
  analyzeToolEligibility,
  createTool,
  deleteTool,
  instantiateTool,
  loadTools,
  saveTool,
  type SavedTool,
  type ToolParameter
} from "./storage";
import {
  MAX_IMAGE_ATTACHMENTS,
  prepareImageFile,
  type PendingImage
} from "./imageAttachments";
import { extractWorkbookDataPeriod } from "./workbookIdentity";
import capabilities from "../../../config/capabilities.json";

type Status = "idle" | "scanning" | "planning" | "tooling" | "executing";
type MessageRole = "assistant" | "user" | "system";
type SourceMode = "workbook" | "folder";
type WorkbookScopeMode = "auto" | "manual";

interface ActivityProgress {
  title: string;
  detail: string;
  completed: string[];
  startedAt: number;
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
  provider?: "model" | "local";
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

function intentScopeFingerprint(scope: IntentScopeContext): string {
  return JSON.stringify({
    workbookName: scope.workbookName,
    sourceMode: scope.sourceMode,
    selectionMode: scope.selectionMode,
    sheets: scope.sheets.map((sheet) => sheet.name),
    ...(scope.selectionMode === "auto"
      ? {
          activeWorksheet: scope.activeWorksheet,
          selectedRange: scope.selectedRange ?? null
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
            messages: conversation.messages.map((message) =>
              message.clarification
                ? {
                    ...message,
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
                : message
            ),
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
  const [toolsOpen, setToolsOpen] = useState(false);
  const [historyOpen, setHistoryOpen] = useState(false);
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
    setActivity({
      title,
      detail,
      completed: [],
      startedAt: Date.now()
    });
    setActivitySeconds(0);
  }

  function advanceActivity(
    title: string,
    detail: string,
    completedStep?: string
  ) {
    setActivity((current) => ({
      title,
      detail,
      completed: completedStep
        ? [...(current?.completed ?? []), completedStep]
        : current?.completed ?? [],
      startedAt: current?.startedAt ?? Date.now()
    }));
  }

  function finishActivity() {
    setActivity(null);
    setActivitySeconds(0);
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
    setStatus("scanning");
    try {
      const snapshot = isRunningInExcel()
        ? await captureWorkbook(
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
    } catch (reason) {
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
        setSelectedModelId((current) =>
          catalog.models.some(
            (option) => option.id === current && option.available
          )
            ? current
            : catalog.defaultModelId
        );
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
    Office.onReady(() => void scan());
  }, []);

  function newChat() {
    if (busy) return;
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
    setSaveCandidate(null);
  }

  function deleteConversation(conversationId: string) {
    if (busy) return;
    const target = chatHistory.conversations.find(
      (conversation) => conversation.id === conversationId
    );
    if (!target || !window.confirm(`确定删除历史对话「${target.title}」吗？`)) {
      return;
    }
    setChatHistory((current) => {
      const remaining = current.conversations.filter(
        (conversation) => conversation.id !== conversationId
      );
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
    });
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
      const response = await createAssistantResponse({
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
      });
      setServerOnline(true);
      if (response.kind === "answer") {
        appendMessage({
          role: "assistant",
          text: response.message,
          resultContext: response.resultContext ?? undefined,
          intentMemory,
          provider: response.provider
        });
      } else {
        appendMessage({
          role: "assistant",
          text: response.plan.summary,
          plan: response.plan,
          intentMemory,
          provider: response.provider
        });
      }
    } catch (reason) {
      setPendingImages(sentImages);
      setServerOnline(false);
      setServiceHealth(null);
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
    advanceActivity(
      "正在本地读取并计算",
      `将扫描 ${selectedNamesFor(workbook).length} 张已选工作表；完整数据只在 Excel 本地处理。`,
      "需求已确认，已选择本地数据工具"
    );
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
      const result = await executeQueryTableTool(
        intent.request,
        effectiveSheetNames,
        selection.activeWorksheet
      );
      advanceActivity(
        "本地计算完成",
        `已扫描 ${result.scannedRows.toLocaleString()} 行，正在准备紧凑结果。`,
        `完成 ${effectiveSheetNames.length} 张工作表的本地查询`
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

  async function sendMessage() {
    const enteredText = prompt.trim();
    const text =
      enteredText ||
      (pendingImages.length > 0
        ? "请结合附件图片分析当前工作簿，并说明发现的问题。"
        : "");
    if (!workbook || !text || busy) return;
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
      setServerOnline(false);
      setServiceHealth(null);
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
        appendMessage({
          role: "assistant",
          text: `「${plan.title}」已执行 ${
            result.actionResults.length
          } 步并完成验证。已写入：${result.filesModified.join("、")}${
            result.backups.length > 0
              ? `；已备份：${result.backups.join("、")}`
              : ""
          }`,
          verification: result.verification
        });
        if (result.verification.passed) {
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
      appendMessage({
        role: "assistant",
        text: `「${plan.title}」已执行 ${result.actionResults.length} 步并完成验证。原始工作表没有被删除或清空。`,
        verification: result.verification
      });
      if (result.verification.passed) {
        setVerifiedPlanIds((current) => new Set(current).add(plan.id));
      }
      await scan();
    } catch (reason) {
      appendMessage({
        role: "system",
        text: reason instanceof Error ? reason.message : "执行计划失败"
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
    setHistoryOpen(false);
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
    setToolsOpen(false);
    setHistoryOpen(true);
  }

  function previewTool(tool: SavedTool) {
    try {
      const plan = instantiateTool(tool, toolParameterValues, workbook);
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
                : `模型：${selectedModelId || serviceHealth?.model}`}
          </span>
        </div>
        <label className="model-picker" title="选择本次对话使用的模型">
          <span>模型</span>
          <select
            value={selectedModelId || "local"}
            disabled={!serverOnline || busy}
            onChange={(event) => {
              setSelectedModelId(event.target.value);
              localStorage.setItem(MODEL_STORAGE_KEY, event.target.value);
            }}
          >
            {modelOptions.map((option) => (
              <option
                key={option.id}
                value={option.id}
                disabled={!option.available}
              >
                {option.label}
              </option>
            ))}
          </select>
        </label>
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
          className="header-button"
          onClick={newChat}
          disabled={busy}
          title="新对话"
          aria-label="新对话"
        >
          ＋
        </button>
      </header>

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
                      onClick={() => previewTool(tool)}
                    >
                      生成执行预览
                    </button>
                  </section>
                ))}
            </div>
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
              {message.text && <p className="message-text">{message.text}</p>}
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
              {message.verification && (
                <div
                  className={`verification-card ${
                    message.verification.passed ? "passed" : "failed"
                  }`}
                >
                  <strong>
                    {message.verification.passed
                      ? "验证通过"
                      : "执行完成，但验证未通过"}
                  </strong>
                  <span>
                    {
                      message.verification.checks.filter((check) => check.passed)
                        .length
                    }
                    /{message.verification.checks.length} 项符合预期
                  </span>
                  {!message.verification.passed && (
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
              {activity && activity.completed.length > 0 && (
                <ul>
                  {activity.completed.map((step, index) => (
                    <li key={`${step}-${index}`}>{step}</li>
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
