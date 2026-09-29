import { Fragment, type ReactNode, useCallback, useEffect, useMemo, useRef, useState } from "react";
import { redactCredentials } from "@excel-bro/core";
import {
  containsCredentialText,
  isLegacyCredentialConversation,
  redactHistoryCredentials,
  remainingLegacyCredentialConversationIds,
  replaceRestoredSecrets
} from "../../excel-addin/src/conversation";
import {
  apiErrorCategory,
  listModels,
  validatePlan,
  type ModelCatalogStatus,
  type ModelOption
} from "../../excel-addin/src/api";
import { assertEditablePlan, type IntentOption, type ResultContext } from "../../excel-addin/src/contracts";
import { useTheme } from "../../excel-addin/src/hooks/useTheme";
import { useLongPress } from "../../excel-addin/src/hooks/useLongPress";
import { useModelManagement } from "../../excel-addin/src/hooks/useModelManagement";
import {
  chooseAvailableModel,
  isRealModelId
} from "../../excel-addin/src/modelSelection";
import ThemePanel from "../../excel-addin/src/components/ThemePanel";
import PetCompanion from "../../excel-addin/src/PetCompanion";
import { RuleManager } from "../../excel-addin/src/RuleManager";
import {
  instantiateTool,
  instantiateQueryTool,
  loadQueryTools,
  loadTools,
  type SavedQueryTool,
  type SavedTool
} from "../../excel-addin/src/storage";
import { executeQueryTableData } from "../../excel-addin/src/dataTools";
import TaskProgressTimeline from "../../excel-addin/src/components/TaskProgressTimeline";
import type { ActivityLog, ActivityOutcome } from "../../excel-addin/src/hooks/useActivityProgress";
import { useMessageAutoFollow } from "../../excel-addin/src/hooks/useMessageAutoFollow";
import { copyTextToClipboard } from "../../excel-addin/src/utils";
import WpsModelSettingsPanel from "./WpsModelSettingsPanel";
import {
  PET_VISIBILITY_STORAGE_KEY,
  normalizePetVisibility
} from "../../excel-addin/src/conversation";
import { getWpsApplication } from "./wpsHost";
import {
  authorizeWpsSheets,
  listWpsWorksheets,
  readAuthorizedWpsSheets
} from "./wpsWorkbookReader";
import { probeWpsWriteCapabilities } from "./wpsWriteCapabilities";
import {
  executeWpsPlan,
  latestWpsUndo,
  prepareWpsPlan,
  undoWpsExecution,
  type WpsPreparedPlan,
  type WpsUndoRecord
} from "./wpsPlanExecutor";
import {
  advanceWpsActivity,
  executionActivityUpdate,
  finishWpsActivity,
  readonlyActivityUpdate,
  startWpsActivity,
  type WpsActivityUpdate
} from "./wpsActivityLog";
import {
  runWpsReadonlyTurn,
  type WpsReadonlyTurnProgress
} from "./wpsReadonlyAgent";

type ChatMessage = {
  id: string;
  role: "user" | "assistant" | "system";
  text: string;
  resultContext?: ResultContext | null;
  activityLog?: ActivityLog;
  planPreview?: { title: string; status: "pending" | "executed" | "cancelled" };
  verification?: {
    executionId: string;
    actionCount: number;
    cellCount: number;
    status: "verified" | "undone";
    targets: { type: string; sheetName: string; range: string }[];
  };
  clarification?: {
    summary: string;
    question: string;
    reason: string;
    scopeLabel: string;
    options: { label: string; description: string }[];
    status: "pending" | "resolved" | "invalidated" | "cancelled";
    answer?: string;
  };
  createdAt: string;
};

type WpsConversation = {
  id: string;
  title: string;
  messages: ChatMessage[];
  createdAt: string;
  updatedAt: string;
};

type WpsChatHistory = {
  activeConversationId: string;
  conversations: WpsConversation[];
  /**
   * 旧记录里曾带真凭据、已被读取路径脱敏的会话 id。独立且可持久化，
   * 用于在脱敏之后继续拦截这些会话的续聊（判据不能是"还能不能匹配密钥"）。
   */
  legacyCredentialConversationIds?: string[];
};

type PendingClarification = {
  messageId: string;
  originalPrompt: string;
  round: number;
  options: IntentOption[];
};

type WpsAgentPanelProps = {
  selectedSheetIndexes: readonly number[];
  resolveSelectedSheetIndexes?: () => readonly number[];
  scopeDock: ReactNode;
};

const WPS_CHAT_STORAGE_KEY = "excel-bro.wps.chat.v1";
const LEGACY_MODEL_CATALOG_NOTICE =
  "暂时无法读取模型配置，请点击上方模型菜单里的「重试」。已输入的内容仍保留在输入框中。";

function messageId(): string {
  return `wps-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function turnErrorMessage(error: unknown): string {
  if (apiErrorCategory(error) === "local_service") {
    return "无法连接 Excel Bro 后端。请确认本地服务已启动，然后重试；本次未修改工作表。";
  }
  // 非连接类错误的 message 会被写进对话历史并随之持久化。服务端错误文本源于
  // Python 侧 `str(error)`，而请求里带着 apiKey —— 原样透出就等于把凭据存进
  // 宿主 localStorage。先过共享脱敏，再截断，避免长文把历史撑爆。
  const raw = error instanceof Error ? error.message : "任务处理失败；本次未修改工作表。";
  const redacted = redactCredentials(raw).trim();
  if (!redacted) return "任务处理失败；本次未修改工作表。";
  return redacted.length > 500 ? `${redacted.slice(0, 500)}…` : redacted;
}

function conversationTitle(messages: readonly ChatMessage[]): string {
  const firstQuestion = messages.find((message) => message.role === "user")?.text.trim();
  if (!firstQuestion) return "新对话";
  const compact = firstQuestion.replace(/\s+/g, " ");
  return compact.length > 34 ? `${compact.slice(0, 34)}…` : compact;
}

function createWpsConversation(messages: ChatMessage[] = []): WpsConversation {
  const now = new Date().toISOString();
  return {
    id: messageId(),
    title: conversationTitle(messages),
    messages,
    createdAt: now,
    updatedAt: now
  };
}

function restoreActivityLog(log: ActivityLog | undefined): ActivityLog | undefined {
  if (!log || !Array.isArray(log.steps)) return undefined;
  if (log.outcome) return log;
  return {
    ...log,
    outcome: "interrupted",
    connectionState: "interrupted",
    connectionMessage: "任务窗格曾被关闭或刷新；此过程未取得终态，请核对工作表和恢复点。",
    steps: log.steps.map((step) => step.status === "started" || step.status === "progress"
      ? { ...step, status: "warning" as const }
      : step)
  };
}

function invalidatePendingClarifications(messages: ChatMessage[]): ChatMessage[] {
  return messages.map((message) => message.clarification?.status === "pending"
    ? { ...message, clarification: { ...message.clarification, status: "invalidated" as const } }
    : message);
}

function loadWpsChatHistory(): WpsChatHistory {
  try {
    const stored = JSON.parse(localStorage.getItem(WPS_CHAT_STORAGE_KEY) ?? "null") as
      | Partial<WpsChatHistory>
      | null;
    const conversations = Array.isArray(stored?.conversations)
      ? stored.conversations.filter((conversation): conversation is WpsConversation =>
          Boolean(
            conversation &&
              typeof conversation.id === "string" &&
              Array.isArray(conversation.messages)
          )
        ).map((conversation) => ({
          ...conversation,
          messages: conversation.messages
            .filter((message) =>
              !(message.role === "system" && message.text === LEGACY_MODEL_CATALOG_NOTICE)
            )
            .map((message) => ({
              ...message,
              activityLog: restoreActivityLog(message.activityLog),
              clarification: message.clarification?.status === "pending"
                ? { ...message.clarification, status: "invalidated" as const }
                : message.clarification
            }))
        }))
      : [];
    if (conversations.length > 0) {
      // 读取路径与 Excel 侧保持一致：先在原始存储上检测并脱敏，再进入可渲染
      // 状态（理由详见 conversation.ts 的 `loadChatHistory`）。凡是原始文本里
      // 检出凭据的会话，都登记到一个**独立且可持久化**的标记里 —— 脱敏之后
      // 就再也匹配不到密钥了，靠"还能不能匹配"做发送闸会在重开后自动松开。
      const legacyIds = new Set<string>(
        Array.isArray(stored?.legacyCredentialConversationIds)
          ? stored.legacyCredentialConversationIds.filter(
              (id): id is string => typeof id === "string"
            )
          : []
      );
      for (const conversation of conversations) {
        const carriedCredential =
          containsCredentialText(conversation.title) ||
          conversation.messages.some((message) =>
            containsCredentialText(message.text)
          );
        if (carriedCredential) legacyIds.add(conversation.id);
      }
      const sanitized = redactHistoryCredentials({
        activeConversationId: String(stored?.activeConversationId ?? ""),
        conversations,
        legacyCredentialConversationIds: [...legacyIds]
      } as never) as unknown as WpsChatHistory;

      // 这里**不做**措辞还原（`replaceRestoredSecrets`）：返回对象会被落盘
      // effect 再 `redactHistoryCredentials` 一次，若现在把 `【已隐藏】` 换成
      // 人话，落盘时匹配不到占位符，下一轮读取又会把真凭据渲染出来。措辞还原
      // 由 `renderWpsRestoredSecrets` 在挂载时做一次。
      const activeConversationId = sanitized.conversations.some(
        (conversation) => conversation.id === stored?.activeConversationId
      )
        ? String(stored?.activeConversationId)
        : sanitized.conversations[0].id;
      return { ...sanitized, activeConversationId };
    }
  } catch {
    // 损坏或不可用的宿主存储不应阻断任务窗格启动。
  }
  const conversation = createWpsConversation();
  return {
    activeConversationId: conversation.id,
    conversations: [conversation],
    legacyCredentialConversationIds: []
  };
}

/**
 * 把「落盘占位符」还原成界面措辞 —— 只做一次，且只在挂载时。
 *
 * 与 Excel 端同一考量：落盘 effect 每次都会把 state 重新脱敏写回存储，所以在
 * 它之前保留原样占位符是安全的；把还原放到读取路径里则会让占位符在落盘时消失，
 * 脱敏再也匹配不到，真凭据会在下一轮循环里重新渲染。
 */
function renderWpsRestoredSecrets(history: WpsChatHistory): WpsChatHistory {
  return {
    ...history,
    conversations: history.conversations.map((conversation) => ({
      ...conversation,
      title: replaceRestoredSecrets(conversation.title ?? ""),
      messages: conversation.messages.map((message) => ({
        ...message,
        text: replaceRestoredSecrets(message.text ?? "")
      }))
    }))
  };
}

function formatConversationTime(value: string): string {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return date.toLocaleString("zh-CN", {
    month: "numeric",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  });
}

export default function WpsAgentPanel({
  selectedSheetIndexes,
  resolveSelectedSheetIndexes,
  scopeDock
}: WpsAgentPanelProps) {
  // 模型目录只包含用户真实配置的连接。初始列表为空、状态单独表达，
  // 不再注入虚拟的 “基础模式” 占位项——目录尚未加载或读取失败时，
  // 那会让人误以为“已经选好模型，可以开始对话”。
  const [models, setModels] = useState<ModelOption[]>([]);
  const [defaultModelId, setDefaultModelId] = useState<string | null>(null);
  const [catalogStatus, setCatalogStatus] = useState<ModelCatalogStatus>("idle");
  const [prompt, setPrompt] = useState("");
  const [chatHistory, setChatHistory] = useState<WpsChatHistory>(() =>
    renderWpsRestoredSecrets(loadWpsChatHistory())
  );
  const messages = useMemo(
    () =>
      chatHistory.conversations.find(
        (conversation) => conversation.id === chatHistory.activeConversationId
      )?.messages ?? [],
    [chatHistory]
  );
  /**
   * 当前对话是否被标为"旧记录里曾带真凭据"。
   *
   * 判据是**持久化标记**（`legacyCredentialConversationIds`），不是"文本里还能
   * 不能匹配到密钥"。读取路径已经把这些会话脱敏，脱敏后当然匹配不到；拿"匹配
   * 不到"当"已安全"会让重开之后发送闸自动打开。
   */
  const staleCredentials = useMemo(
    () =>
      isLegacyCredentialConversation(
        chatHistory as never,
        chatHistory.activeConversationId
      ),
    [chatHistory]
  );
  const [progress, setProgress] = useState<WpsReadonlyTurnProgress | null>(null);
  const [activeActivityMessageId, setActiveActivityMessageId] = useState<string | null>(null);
  const [writeProgress, setWriteProgress] = useState<string | null>(null);
  const [pendingPlan, setPendingPlan] = useState<WpsPreparedPlan | null>(null);
  const [pendingPlanMessageId, setPendingPlanMessageId] = useState<string | null>(null);
  const [planEditDraft, setPlanEditDraft] = useState<string | null>(null);
  const [planEditError, setPlanEditError] = useState<string | null>(null);
  useEffect(() => {
    setPlanEditDraft(null);
    setPlanEditError(null);
  }, [pendingPlan?.previewRevision]);
  const [undoRecord, setUndoRecord] = useState<WpsUndoRecord | null>(null);
  const [clarification, setClarification] = useState<PendingClarification | null>(null);
  const [clarificationDraft, setClarificationDraft] = useState("");
  const [themeOpen, setThemeOpen] = useState(false);
  const [petVisible, setPetVisible] = useState(() => {
    try {
      return normalizePetVisibility(localStorage.getItem(PET_VISIBILITY_STORAGE_KEY));
    } catch {
      return true;
    }
  });
  const [modelMenuOpen, setModelMenuOpen] = useState(false);
  const [modelSettingsOpen, setModelSettingsOpen] = useState(false);
  const [toolsOpen, setToolsOpen] = useState(false);
  const [backupOpen, setBackupOpen] = useState(false);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [moreMenuOpen, setMoreMenuOpen] = useState(false);
  const [ruleManagerOpen, setRuleManagerOpen] = useState(false);
  const [tools] = useState(loadTools);
  const [queryTools] = useState(loadQueryTools);
  const [copiedMessageId, setCopiedMessageId] = useState<string | null>(null);
  const [copyError, setCopyError] = useState(false);
  const [newConversationNotice, setNewConversationNotice] = useState<{
    id: number;
    previousSaved: boolean;
  } | null>(null);
  const [newConversationConfirmOpen, setNewConversationConfirmOpen] = useState(false);
  const controllerRef = useRef<AbortController | null>(null);
  const composerInputRef = useRef<HTMLTextAreaElement | null>(null);
  const authorizationKey = useMemo(
    () => [...selectedSheetIndexes].sort((left, right) => left - right).join(","),
    [selectedSheetIndexes]
  );
  const authorizationKeyRef = useRef(authorizationKey);
  const writeCapability = useMemo(() => {
    try {
      return probeWpsWriteCapabilities(getWpsApplication(), selectedSheetIndexes);
    } catch {
      return null;
    }
  }, [authorizationKey, selectedSheetIndexes]);
  const canWrite = writeCapability?.writeEnabled === true;
  const developmentWrite = writeCapability?.versionStatus === "development";
  const writeStatusReason = writeCapability?.checks.find((check) => !check.passed)?.detail
    ?? "确认任务时会检查当前工作簿是否满足写入条件";
  const latestActivity = [...messages].reverse().find((message) => message.activityLog)?.activityLog;
  const scrollSignal = [
    chatHistory.activeConversationId,
    messages.length,
    latestActivity?.totalMs ?? 0,
    latestActivity?.steps.length ?? 0,
    latestActivity?.steps.at(-1)?.current ?? "",
    latestActivity?.outcome ?? "",
    progress?.message ?? "",
    writeProgress ?? ""
  ].join(":");
  const {
    containerRef: messageStreamRef,
    hasUnseenProgress,
    handleScroll: handleMessageStreamScroll,
    scrollToLatest
  } = useMessageAutoFollow(scrollSignal);
  const theme = useTheme();
  const logoLongPress = useLongPress(() => setThemeOpen(true));
  const busy = progress !== null || writeProgress !== null;

  /**
   * 读取模型目录。
   *
   * 返回 null 表示本次读取失败：调用方必须保留原有选择，不能把失败当成
   * “用户没有配置模型”——那会把一次后台不可用误报成用户从没配过。
   */
  const refreshModelCatalog = useCallback(async () => {
    setCatalogStatus((current) => (current === "ready" ? current : "loading"));
    try {
      const catalog = await listModels();
      setModels(catalog.models);
      setDefaultModelId(catalog.defaultModelId);
      setCatalogStatus("ready");
      return {
        modelOptions: catalog.models,
        defaultModelId: catalog.defaultModelId
      };
    } catch {
      setCatalogStatus("failed");
      return null;
    }
  }, []);

  const modelManagement = useModelManagement({
    refreshServiceHealth: refreshModelCatalog
  });

  /**
   * 界面展示的模型必须与请求实际发送的 modelId 完全一致：
   * 这里和 Excel 端共用同一份选择规则，解析不出有效连接时就是「未配置」，
   * 绝不回退到目录首项，也不再有任何虚拟的基础模式占位项。
   */
  const effectiveModelId = useMemo(
    () => chooseAvailableModel(models, modelManagement.selectedModelId, defaultModelId),
    [models, modelManagement.selectedModelId, defaultModelId]
  );

  const selectedModel = useMemo(
    () => models.find((model) => model.id === effectiveModelId) ?? null,
    [effectiveModelId, models]
  );

  const modelCatalogFailed = catalogStatus === "failed";
  const modelCatalogLoading = catalogStatus === "loading" || catalogStatus === "idle";
  const modelNotConfigured = catalogStatus === "ready" && models.length === 0;
  const brandStatusText = modelCatalogFailed
    ? isRealModelId(modelManagement.selectedModelId)
      ? "等待 Excel Bro 后台"
      : "点击添加模型开始"
    : modelCatalogLoading
      ? "正在读取模型配置"
      : modelNotConfigured
        ? "尚未配置模型"
        : `模型：${selectedModel?.label ?? effectiveModelId}`;

  useEffect(() => {
    let active = true;
    void refreshModelCatalog().then(() => {
      if (!active) return;
    });
    return () => {
      active = false;
      controllerRef.current?.abort();
    };
  }, [refreshModelCatalog]);

  useEffect(() => {
    try {
      setUndoRecord(latestWpsUndo(getWpsApplication()));
    } catch {
      setUndoRecord(null);
    }
  }, []);

  useEffect(() => {
    if (authorizationKeyRef.current === authorizationKey) return;
    authorizationKeyRef.current = authorizationKey;
    setPendingPlan(null);
    setPendingPlanMessageId(null);
    setChatHistory((current) => ({
      ...current,
      conversations: current.conversations.map((conversation) =>
        conversation.id === current.activeConversationId
          ? { ...conversation, messages: invalidatePendingClarifications(conversation.messages) }
          : conversation
      )
    }));
    setClarification(null);
  }, [authorizationKey]);

  // 目录成功加载后才迁移旧记录：本机存过已被移除的 local、或指向已删除连接的
  // 选择，会在这里被换成当前有效模型；目录读取失败时不迁移，以免依据空列表
  // 删掉仍然有效的选择。
  useEffect(() => {
    if (catalogStatus !== "ready") return;
    modelManagement.migrateModelSelection({
      modelOptions: models,
      defaultModelId
    });
  }, [
    catalogStatus,
    models,
    defaultModelId,
    modelManagement.migrateModelSelection
  ]);

  useEffect(() => {
    if (!newConversationNotice) return;
    const timer = window.setTimeout(() => setNewConversationNotice(null), 1500);
    return () => window.clearTimeout(timer);
  }, [newConversationNotice]);

  useEffect(() => {
    if (!copiedMessageId && !copyError) return;
    const timer = window.setTimeout(() => {
      setCopiedMessageId(null);
      setCopyError(false);
    }, 1600);
    return () => window.clearTimeout(timer);
  }, [copiedMessageId, copyError]);

  /**
   * 落盘失败提示的展示状态。与 Excel 端一致：失败时明确告知用户，不静默降级。
   *
   * 用 state 而不是直接 append 系统消息，理由有两条：
   * 1. 落盘失败本身也是"这条历史不会被保留"，把它写进历史会在下一次落盘时
   *    又被写丢，自相矛盾；提示应当独立于可持久化的内容存在。
   * 2. 避免"提示写进历史 → 触发落盘 → 又失败 → 再提示"的循环。
   */
  const [historyNotice, setHistoryNotice] = useState<string | null>(null);
  const storageFailedRef = useRef(false);

  useEffect(() => {
    try {
      const active = chatHistory.activeConversationId;
      const compact: WpsChatHistory = {
        activeConversationId: active,
        legacyCredentialConversationIds: remainingLegacyCredentialConversationIds(
          chatHistory as never
        ) as string[],
        conversations: [...chatHistory.conversations]
          .sort((left, right) => right.updatedAt.localeCompare(left.updatedAt))
          .slice(0, 30)
          .map((conversation) => ({
            ...conversation,
            messages: conversation.messages.slice(-120)
          }))
      };
      // 落盘前统一脱敏：对话文本可能来自服务端错误、模型回答或用户粘贴，
      // 任何一条都不该把凭据带进 localStorage。与 Excel 端同一实现。
      // `legacyCredentialConversationIds` 一并写下去：脱敏之后就只能靠这个标记
      // 判断哪些会话是旧记录，靠它重开后发送闸才仍然有效。
      localStorage.setItem(
        WPS_CHAT_STORAGE_KEY,
        JSON.stringify(redactHistoryCredentials(compact as never))
      );
    } catch {
      // 宿主写入失败：内存里的历史在读取路径已经脱敏（原文此时已不存在），
      // 所以这里要做的不是"再脱一遍原文"，而是明确告知用户本次会话不会被保留。
      // 关键：**不把错误对象原文写进提示或日志** —— 异常里可能带着写入失败
      // 的上下文（键名、配额乃至被拒绝的值），一律折叠成固定文案。
      //
      // 提示只提示一次；发送闸不受影响（判据是内存里的 `legacyCredentialConversationIds`，
      // 与是否写进存储无关），所以"落盘失败"不会让旧记录被误当成正常历史续聊。
      if (!storageFailedRef.current) {
        storageFailedRef.current = true;
        setHistoryNotice(
          "无法把对话历史写入本机存储（宿主可能禁止了本地存储或空间已满）。" +
            "本次会话仍可继续，但关掉任务窗格后不会被保留。"
        );
      }
      return;
    }
    storageFailedRef.current = false;
    setHistoryNotice(null);
  }, [chatHistory]);

  useEffect(() => {
    try {
      localStorage.setItem(
        PET_VISIBILITY_STORAGE_KEY,
        petVisible ? "visible" : "hidden"
      );
    } catch {
      // 存储不可用时，本次会话仍保留当前选择。
    }
  }, [petVisible]);

  /**
   * 更新当前对话的消息。
   *
   * `clearLegacyMark` 决定这次写入是否算"用户明确放弃旧记录拦截"。
   *
   * **默认必须为 false。** 起初这里无条件清标记，结果是：用户对一段旧记录
   * 提问 → 发送闸拦下并追加一条系统提示 → 那条提示本身把标记清掉了 →
   * 第二次提问就畅通无阻。等于"拦一次就失效"，发送闸形同虚设。
   *
   * 只有真正把用户的新提问（`runTurn` 的 `append({role:"user"})`）写进历史时，
   * 才显式传 true —— 那一刻才是"这段历史已经翻篇，后续是新内容"。
   */
  const updateMessages = useCallback(
    (
      updater: (current: ChatMessage[]) => ChatMessage[],
      clearLegacyMark = false
    ) => {
      setChatHistory((current) => {
        const now = new Date().toISOString();
        const legacyIds = (current as WpsChatHistory).legacyCredentialConversationIds;
        return {
          ...current,
          legacyCredentialConversationIds: clearLegacyMark
            ? (legacyIds ?? []).filter(
                (id) => id !== current.activeConversationId
              )
            : legacyIds,
          conversations: current.conversations.map((conversation) => {
            if (conversation.id !== current.activeConversationId) return conversation;
            const nextMessages = updater(conversation.messages);
            return {
              ...conversation,
              title: conversationTitle(nextMessages),
              messages: nextMessages,
              updatedAt: now
            };
          })
        };
      });
    },
    []
  );

  const append = (message: Omit<ChatMessage, "id" | "createdAt">): string => {
    const id = messageId();
    updateMessages((current) => [
      ...current,
      { id, createdAt: new Date().toISOString(), ...message }
    ]);
    return id;
  };

  const beginActivity = (step: WpsActivityUpdate) => {
    const startedAt = Date.now();
    const id = append({ role: "assistant", text: "", activityLog: startWpsActivity(step, startedAt) });
    setActiveActivityMessageId(id);
    return { id, startedAt };
  };

  const updateActivity = (activity: { id: string; startedAt: number }, step: WpsActivityUpdate) => {
    updateMessages((current) => current.map((message) => message.id === activity.id && message.activityLog
      ? { ...message, activityLog: advanceWpsActivity(message.activityLog, step, activity.startedAt) }
      : message));
  };

  const endActivity = (activity: { id: string; startedAt: number }, outcome: ActivityOutcome) => {
    updateMessages((current) => current.map((message) => message.id === activity.id && message.activityLog
      ? { ...message, activityLog: finishWpsActivity(message.activityLog, outcome, activity.startedAt) }
      : message));
    setActiveActivityMessageId((current) => current === activity.id ? null : current);
  };

  const markPlanPreview = (status: "executed" | "cancelled") => {
    if (!pendingPlanMessageId) return;
    updateMessages((current) => current.map((message) =>
      message.id === pendingPlanMessageId && message.planPreview
        ? { ...message, planPreview: { ...message.planPreview, status } }
        : message
    ));
  };

  const completeNewConversation = () => {
    const previousSaved = messages.some((message) => message.role === "user");
    if (previousSaved) {
      const conversation = createWpsConversation();
      setChatHistory((current) => ({
        activeConversationId: conversation.id,
        conversations: [
          conversation,
          ...current.conversations.map((item) => item.id === current.activeConversationId
            ? { ...item, messages: invalidatePendingClarifications(item.messages) }
            : item)
        ].slice(0, 30),
        // 标记跟着会话走：只保留切到最后 30 条里仍然存在的那些。
        legacyCredentialConversationIds: (
          current.legacyCredentialConversationIds ?? []
        ).filter((id) =>
          current.conversations.slice(0, 29).some((item) => item.id === id)
        )
      }));
    }
    setPrompt("");
    setProgress(null);
    setActiveActivityMessageId(null);
    setClarification(null);
    setPendingPlan(null);
    setPendingPlanMessageId(null);
    setCopiedMessageId(null);
    setHistoryOpen(false);
    setToolsOpen(false);
    setBackupOpen(false);
    setMoreMenuOpen(false);
    setModelMenuOpen(false);
    setNewConversationConfirmOpen(false);
    setNewConversationNotice({ id: Date.now(), previousSaved });
    window.requestAnimationFrame(() => composerInputRef.current?.focus());
  };

  const startNewConversation = () => {
    if (busy) return;
    if (pendingPlan) {
      setHistoryOpen(false);
      setToolsOpen(false);
      setBackupOpen(false);
      setMoreMenuOpen(false);
      setModelMenuOpen(false);
      setNewConversationConfirmOpen(true);
      return;
    }
    completeNewConversation();
  };

  const openConversation = (conversationId: string) => {
    setChatHistory((current) =>
      current.conversations.some((conversation) => conversation.id === conversationId)
        ? {
            ...current,
            activeConversationId: conversationId,
            conversations: current.conversations.map((conversation) =>
              conversation.id === current.activeConversationId
                ? { ...conversation, messages: invalidatePendingClarifications(conversation.messages) }
                : conversation
            )
          }
        : current
    );
    setPrompt("");
    setClarification(null);
    setClarificationDraft("");
    setPendingPlan(null);
    setPendingPlanMessageId(null);
    setCopiedMessageId(null);
    setHistoryOpen(false);
  };

  const deleteConversation = (conversationId: string) => {
    setChatHistory((current) => {
      const remaining = current.conversations.filter(
        (conversation) => conversation.id !== conversationId
      );
      if (remaining.length === current.conversations.length) return current;
      if (remaining.length === 0) {
        const replacement = createWpsConversation();
        return { activeConversationId: replacement.id, conversations: [replacement] };
      }
      return {
        activeConversationId:
          current.activeConversationId === conversationId
            ? remaining[0].id
            : current.activeConversationId,
        conversations: remaining
      };
    });
    setClarification(null);
    setPendingPlan(null);
    setPendingPlanMessageId(null);
  };

  const closeHeaderSurfaces = () => {
    setModelMenuOpen(false);
    setToolsOpen(false);
    setBackupOpen(false);
    setHistoryOpen(false);
    setMoreMenuOpen(false);
  };

  const copyMessage = async (message: ChatMessage) => {
    const copied = await copyTextToClipboard(message.text);
    if (copied) {
      setCopiedMessageId(message.id);
      setCopyError(false);
    } else {
      setCopyError(true);
    }
  };

  const runTurn = async (
    question: string,
    options: { intentConfirmed?: boolean; clarificationRound?: number; showUser?: boolean } = {}
  ) => {
    const text = question.trim();
    const turnSheetIndexes = [
      ...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)
    ];
    if (!text || busy || turnSheetIndexes.length === 0) return;
    if (pendingPlan) {
      append({ role: "system", text: "请先确认或取消当前写入计划，再开始新的任务。" });
      return;
    }
    // 续聊前的旧记录闸：与 Excel 端同一判据 —— 查的是读取时登记的持久标记
    // （`legacyCredentialConversationIds`），不是"文本里还能不能匹配到凭据"。
    // 这段历史在读取时已被脱敏，真凭据不会进入界面、内存或请求；标记只是为了
    // 让用户明确知道这是一段旧记录、并据此新开一段对话。
    if (staleCredentials) {
      append({
        role: "system",
        text:
          "这段历史里的凭据是在旧版本写进去的，已被拦下，不会随本次提问发送。" +
          "请点上方「新对话」开始一段干净的对话，或先删除这段历史。"
      });
      setPrompt(text);
      return;
    }
    // 没有可用模型时必须明确反馈，并保留用户已经输入的任务文本：
    // 不清空输入框、不伪造一次成功回答，也不擅自改用别的连接。
    if (!effectiveModelId) {
      setModelMenuOpen(true);
      return;
    }
    if (options.showUser !== false) {
      // 明确放弃"旧记录"拦截：用户真的开始输入新内容了，这段历史翻篇，
      // 后续是新写的、不再携带原始凭据。**只有这里**会清掉陈旧标记 ——
      // 系统提示之类的追加写入不能清，否则发送闸被拦一次就失效。
      updateMessages(
        (current) => [
          ...current,
          { id: messageId(), role: "user", text, createdAt: new Date().toISOString() }
        ],
        true
      );
    }
    const activity = beginActivity({
      id: "read-schemas",
      label: "读取已授权表结构",
      phase: "preflight",
      detail: "只读取字段结构，不向意图模型发送原始数据行。"
    });
    setPrompt("");
    if (clarification && options.showUser !== false) {
      updateMessages((current) => current.map((message) =>
        message.id === clarification.messageId && message.clarification
          ? { ...message, clarification: { ...message.clarification, status: "invalidated" } }
          : message
      ));
    }
    setClarification(null);
    const controller = new AbortController();
    controllerRef.current?.abort();
    controllerRef.current = controller;
    setProgress({ stage: "reading", message: "正在读取表结构并理解消息…" });
    try {
      const result = await runWpsReadonlyTurn({
        application: getWpsApplication(),
        selectedSheetIndexes: turnSheetIndexes,
        prompt: text,
        modelId: effectiveModelId || null,
        intentConfirmed: options.intentConfirmed,
        clarificationRound: options.clarificationRound,
        conversation: messages
          .filter((message) => message.role !== "system" && message.text.trim().length > 0)
          .slice(-6)
          .map((message) => ({
            role: message.role as "user" | "assistant",
            text: message.text.slice(0, 1000)
          })),
        signal: controller.signal,
        onProgress: (update) => {
          setProgress(update);
          updateActivity(activity, readonlyActivityUpdate(update));
        }
      });
      if (result.kind === "clarification") {
        endActivity(activity, "awaiting_input");
        setClarificationDraft("");
        const clarificationMessageId = append({
          role: "assistant",
          text: "",
          clarification: {
            summary: redactCredentials(result.response.clarification.summary),
            question: redactCredentials(result.response.clarification.question),
            reason: redactCredentials(result.response.clarification.reason),
            scopeLabel: redactCredentials(result.response.clarification.scopeLabel),
            options: result.response.clarification.options.map((option) => ({
              label: redactCredentials(option.label),
              description: redactCredentials(option.description)
            })),
            status: "pending"
          }
        });
        setClarification({
          messageId: clarificationMessageId,
          originalPrompt: text,
          round: options.clarificationRound ?? 0,
          options: result.response.clarification.options
        });
      } else if (result.kind === "plan") {
        setProgress({ stage: "answering", message: "正在校验写入计划和授权边界…" });
        updateActivity(activity, {
          id: "preflight-plan",
          label: "校验写入计划与授权",
          phase: "preflight",
          detail: "准备可供用户检查的计划预览；此时不写入工作簿。"
        });
        const application = getWpsApplication();
        const metadata = listWpsWorksheets(application);
        const selectedNames = metadata.sheets
          .filter((sheet) => turnSheetIndexes.includes(sheet.index))
          .map((sheet) => sheet.name);
        // 本端算出的计划不送服务端校验：验收条件里带着完整值矩阵，上传等于
        // 把刚才刻意留在本地的源值送出去。理由详见 WpsTurnResult.computedLocally。
        const normalizedPlan = result.computedLocally
          ? result.plan
          : await validatePlan(
              result.plan,
              selectedNames,
              metadata.sheets.map((sheet) => sheet.name)
            );
        if (result.computedLocally && !result.sourceSnapshots?.length) {
          throw new Error("本地映射计划缺少源区域快照，不能安全预览写入。");
        }
        const prepared = prepareWpsPlan(application, turnSheetIndexes, normalizedPlan, result.sourceSnapshots);
        endActivity(activity, "awaiting_input");
        setPendingPlan(prepared);
        setPendingPlanMessageId(append({
          role: "assistant",
          text: "",
          planPreview: { title: prepared.plan.title, status: "pending" }
        }));
      } else {
        endActivity(activity, "completed");
        append({
          role: "assistant",
          text: result.message,
          resultContext: result.resultContext
        });
      }
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") {
        endActivity(activity, "cancelled");
        append({ role: "system", text: "任务已停止；未确认的计划不会修改工作表。" });
      } else {
        endActivity(activity, "failed");
        if (apiErrorCategory(error) === "local_service" && options.showUser !== false) setPrompt(text);
        append({
          role: "system",
          text: turnErrorMessage(error)
        });
      }
    } finally {
      if (controllerRef.current === controller) controllerRef.current = null;
      setProgress(null);
    }
  };

  const confirmPendingPlan = async () => {
    if (!pendingPlan || busy || planEditDraft !== null) return;
    const plan = pendingPlan;
    const activity = beginActivity({
      id: "preflight",
      label: "复核计划与授权范围",
      phase: "preflight",
      detail: `${plan.actions.length} 个动作，目标 ${plan.totalCells} 个单元格；执行前重新核对工作簿与授权。`
    });
    setWriteProgress("正在保存执行前恢复点…");
    try {
      const currentSheetIndexes = [...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)].sort((a, b) => a - b);
      const previewSheetIndexes = plan.authorization.sheets.map((sheet) => sheet.index).sort((a, b) => a - b);
      if (JSON.stringify(currentSheetIndexes) !== JSON.stringify(previewSheetIndexes)) {
        throw new Error("授权工作表范围在预览后发生变化，请重新预览并确认。");
      }
      const record = await executeWpsPlan(getWpsApplication(), plan, (update) => {
        setWriteProgress(update.message);
        updateActivity(activity, executionActivityUpdate(update));
      });
      updateActivity(activity, {
        id: "verified-result",
        label: "写入与回读核对完成",
        phase: "verifying",
        note: "所有目标均按计划回读匹配；恢复点仅保证当前加载项会话内可用。",
        facts: [
          { label: "操作", value: `${record.actions.length} 步` },
          { label: "已核对", value: `${plan.totalCells} 个单元格` },
          { label: "恢复", value: "会话内可撤销" }
        ]
      });
      endActivity(activity, "completed");
      setUndoRecord(record);
      markPlanPreview("executed");
      setPendingPlan(null);
      setPendingPlanMessageId(null);
      append({
        role: "assistant",
        text: `已执行「${plan.plan.title}」，并完成目标区域回读核对。`,
        verification: {
          executionId: record.executionId,
          actionCount: plan.actions.length,
          cellCount: plan.totalCells,
          status: "verified",
          targets: plan.actions.map((action) => ({
            type: action.type,
            sheetName: redactCredentials(action.sheetName),
            range: action.range
          }))
        }
      });
    } catch (error) {
      let latest: WpsUndoRecord | null = null;
      try {
        latest = latestWpsUndo(getWpsApplication());
        setUndoRecord(latest);
      } catch {
        setUndoRecord(null);
      }
      const currentRecovery = latest?.planId === plan.plan.id &&
        Date.parse(latest.createdAt) >= activity.startedAt;
      const partiallyWritten = currentRecovery && latest?.state === "partial_write_recoverable";
      endActivity(activity, partiallyWritten ? "warning" : "failed");
      append({
        role: "system",
        text: `${error instanceof Error ? error.message : "WPS 写入失败。"}${partiallyWritten
          ? " 本次可能已部分写入，恢复资料已保留；请检查执行过程与撤销入口。"
          : ""}`
      });
    } finally {
      setWriteProgress(null);
    }
  };

  const applyEditedPlan = async () => {
    if (!pendingPlan || planEditDraft === null || busy) return;
    setWriteProgress("正在重新校验修改后的计划…");
    try {
      const parsed: unknown = JSON.parse(planEditDraft);
      assertEditablePlan(parsed);
      const application = getWpsApplication();
      const currentSheetIndexes = [...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)];
      const metadata = listWpsWorksheets(application);
      const selectedNames = metadata.sheets
        .filter((sheet) => currentSheetIndexes.includes(sheet.index))
        .map((sheet) => sheet.name);
      // Locally computed mapping values must never be sent to plan validation.
      const validated = pendingPlan.sourceSnapshots.length
        ? parsed
        : await validatePlan(parsed, selectedNames, metadata.sheets.map((sheet) => sheet.name));
      const updated = prepareWpsPlan(application, currentSheetIndexes, validated, pendingPlan.sourceSnapshots);
      setPendingPlan(updated);
      if (pendingPlanMessageId) {
        updateMessages((current) => current.map((message) =>
          message.id === pendingPlanMessageId && message.planPreview
            ? { ...message, planPreview: { ...message.planPreview, title: updated.plan.title } }
            : message
        ));
      }
      setPlanEditDraft(null);
      setPlanEditError(null);
    } catch (error) {
      setPlanEditError(error instanceof Error ? error.message : "计划编辑内容无效。");
    } finally {
      setWriteProgress(null);
    }
  };

  const undoLatestExecution = async () => {
    if (!undoRecord || busy) return;
    const activity = beginActivity({
      id: "undo-preflight",
      label: "核对撤销条件",
      phase: "preflight",
      detail: "先确认目标仍是上次写入后的状态，避免覆盖后续人工修改。"
    });
    setWriteProgress("正在核对撤销条件…");
    try {
      await undoWpsExecution(getWpsApplication(), undoRecord.executionId, (update) => {
        setWriteProgress(update.message);
        updateActivity(activity, executionActivityUpdate(update));
      });
      updateActivity(activity, {
        id: "undo-verified",
        label: "恢复结果已回读核对",
        phase: "verifying",
        detail: "目标区域与执行前快照一致。"
      });
      endActivity(activity, "completed");
      setUndoRecord(null);
      updateMessages((current) => current.map((message) =>
        message.verification?.executionId === undoRecord.executionId
          ? { ...message, verification: { ...message.verification, status: "undone" } }
          : message
      ));
      append({ role: "assistant", text: `已撤销「${undoRecord.planTitle}」，并回读确认恢复成功。` });
    } catch (error) {
      endActivity(activity, "failed");
      try {
        setUndoRecord(latestWpsUndo(getWpsApplication()));
      } catch {
        setUndoRecord(null);
      }
      append({ role: "system", text: error instanceof Error ? error.message : "WPS 撤销失败。" });
    } finally {
      setWriteProgress(null);
    }
  };

  const prepareSavedWorkflow = async (tool: SavedTool) => {
    if (busy) return;
    try {
      const toolSheetIndexes = [
        ...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)
      ];
      if (toolSheetIndexes.length === 0) {
        throw new Error("请先选择这个工具可以使用的工作表");
      }
      const application = getWpsApplication();
      const metadata = listWpsWorksheets(application);
      const selectedNames = metadata.sheets
        .filter((sheet) => toolSheetIndexes.includes(sheet.index))
        .map((sheet) => sheet.name);
      const values = Object.fromEntries(
        tool.parameters.map((parameter) => [parameter.id, parameter.defaultValue])
      );
      const plan = instantiateTool(tool, values);
      const normalizedPlan = await validatePlan(
        plan,
        selectedNames,
        metadata.sheets.map((sheet) => sheet.name)
      );
      const prepared = prepareWpsPlan(application, toolSheetIndexes, normalizedPlan);
      setPendingPlan(prepared);
      setPendingPlanMessageId(append({
        role: "assistant",
        text: "",
        planPreview: { title: prepared.plan.title, status: "pending" }
      }));
      setToolsOpen(false);
    } catch (error) {
      append({
        role: "system",
        text:
          error instanceof Error
            ? `这个工具暂时不能在 WPS 中运行：${error.message}`
            : "这个工具暂时不能在 WPS 中运行。"
      });
      setToolsOpen(false);
    }
  };

  const runSavedQuery = async (tool: SavedQueryTool) => {
    if (busy) return;
    const querySheetIndexes = [
      ...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)
    ];
    if (querySheetIndexes.length === 0) {
      append({ role: "system", text: "请先选择这个固化查询可以使用的工作表。" });
      return;
    }
    const controller = new AbortController();
    controllerRef.current?.abort();
    controllerRef.current = controller;
    const activity = beginActivity({
      id: "saved-query-preflight",
      label: "检查固化查询的数据范围",
      phase: "preflight",
      detail: "只使用本次明确授权的工作表。"
    });
    try {
      if (tool.sourceMode !== "workbook") {
        throw new Error("这个固化查询来自工作目录；WPS 当前只支持已打开的工作簿");
      }
      const application = getWpsApplication();
      const metadata = listWpsWorksheets(application);
      const authorizedSheets = metadata.sheets.filter((sheet) =>
        querySheetIndexes.includes(sheet.index)
      );
      const requiredNames = new Set(tool.sourceSheetNames);
      const missingNames = tool.sourceSheetNames.filter(
        (name) => !authorizedSheets.some((sheet) => sheet.name === name)
      );
      if (missingNames.length > 0) {
        throw new Error(
          `当前授权范围缺少：${missingNames.join("、")}。请在数据范围中明确选择后再运行`
        );
      }
      const targetIndexes = (requiredNames.size > 0
        ? authorizedSheets.filter((sheet) => requiredNames.has(sheet.name))
        : authorizedSheets
      ).map((sheet) => sheet.index);
      setProgress({ stage: "reading", message: "正在读取固化查询已授权的数据源…" });
      updateActivity(activity, {
        id: "saved-query-read",
        label: "读取授权数据",
        phase: "preflight",
        detail: "读取固化查询实际需要的工作表。"
      });
      const authorization = authorizeWpsSheets(application, targetIndexes);
      const results = await readAuthorizedWpsSheets(application, authorization, {
        signal: controller.signal,
        onProgress: ({ sheet, rowsRead, totalRows }) => {
          setProgress({
            stage: "reading",
            message: `正在读取「${sheet.name}」：${rowsRead} / ${totalRows} 行`
          });
          updateActivity(activity, {
            id: `saved-query-read:${sheet.name}`,
            label: `读取「${sheet.name}」`,
            phase: "preflight",
            target: { sheet: sheet.name },
            current: rowsRead,
            total: totalRows,
            unit: "rows"
          });
        }
      });
      setProgress({ stage: "calculating", message: "正在本地确定性执行固化查询…" });
      updateActivity(activity, {
        id: "saved-query-calculate",
        label: "执行固化查询",
        phase: "executing",
        toolName: "query_table",
        detail: "在本机对完整授权数据执行确定性计算。"
      });
      const result = executeQueryTableData(
        instantiateQueryTool(tool),
        results.map((item) => ({ name: item.sheet.name, values: item.values }))
      );
      endActivity(activity, "completed");
      append({
        role: "assistant",
        text: `固化查询「${tool.name}」已完成：扫描 ${result.scannedRows.toLocaleString()} 行，返回 ${result.rows.length} 行；模型调用 0 次。`,
        resultContext: {
          kind: "table",
          title: result.title,
          headers: result.headers.map(String),
          rows: result.rows,
          sourceSheets: result.sourceSheets,
          warnings: result.warnings
        }
      });
      setToolsOpen(false);
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") {
        endActivity(activity, "cancelled");
        append({ role: "system", text: "固化查询已停止，没有修改任何工作表。" });
      } else {
        endActivity(activity, "failed");
        append({
          role: "system",
          text: error instanceof Error ? error.message : "固化查询运行失败。"
        });
      }
    } finally {
      if (controllerRef.current === controller) controllerRef.current = null;
      setProgress(null);
    }
  };

  const resolveClarification = (optionIndex: number) => {
    if (!clarification || busy) return;
    const option = clarification.options[optionIndex];
    if (!option) return;
    updateMessages((current) => current.map((message) =>
      message.id === clarification.messageId && message.clarification
        ? { ...message, clarification: { ...message.clarification, status: "resolved", answer: redactCredentials(option.label) } }
        : message
    ));
    append({ role: "user", text: option.label });
    const resolvedPrompt =
      option.resolutionMode === "replace"
        ? option.resolution
        : `${clarification.originalPrompt}\n补充确认：${option.resolution}`;
    void runTurn(resolvedPrompt, {
      intentConfirmed: true,
      clarificationRound: clarification.round + 1,
      showUser: false
    });
  };

  const resolveClarificationWithText = () => {
    if (!clarification || !clarificationDraft.trim() || busy) return;
    const answer = clarificationDraft.trim();
    updateMessages((current) => current.map((message) =>
      message.id === clarification.messageId && message.clarification
        ? { ...message, clarification: { ...message.clarification, status: "resolved", answer: redactCredentials(answer) } }
        : message
    ));
    append({ role: "user", text: answer });
    void runTurn(`${clarification.originalPrompt}\n补充确认：${answer}`, {
      intentConfirmed: true,
      clarificationRound: clarification.round + 1,
      showUser: false
    });
  };

  const cancelClarification = () => {
    if (!clarification || busy) return;
    updateMessages((current) => current.map((message) =>
      message.id === clarification.messageId && message.clarification
        ? { ...message, clarification: { ...message.clarification, status: "cancelled" } }
        : message
    ));
    setClarification(null);
    setClarificationDraft("");
  };

  const closeModelSettings = () => {
    setModelSettingsOpen(false);
    modelManagement.setApiKeyDraft("");
    modelManagement.setShowApiKey(false);
    modelManagement.setConnectionDraft(null);
    modelManagement.setPendingDeleteConnectionId(null);
    modelManagement.setSettingsFeedback("");
  };

  const openModelSettings = async (createConnection: boolean) => {
    setModelMenuOpen(false);
    setModelSettingsOpen(true);
    if (createConnection) {
      // 先展示可填写的表单，再异步发现 Excel/WPS 共用的已有连接。
      modelManagement.openConnectionCreator();
      void modelManagement.checkExistingConnections();
    } else {
      await modelManagement.openSettings();
    }
  };

  const planCard = pendingPlan && (
    <section className="wps-plan-card" aria-label="WPS 写入计划预览">
      <span className="clarification-eyebrow">写入前确认</span>
      <strong>{pendingPlan.plan.title}</strong>
      <p>{pendingPlan.plan.summary}</p>
      {developmentWrite && <p>当前 WPS JSAPI 版本仅在开发调试模式开放写入；请先在独立测试工作簿验收写入与撤销。</p>}
      <div className="wps-plan-meta">
        <span>{pendingPlan.actions.length} 个操作</span>
        <span>{pendingPlan.totalCells} 个单元格</span>
        <span>可撤销</span>
      </div>
      <div className="wps-plan-actions-list">
        {pendingPlan.actions.map((action) => (
          <div key={`${action.actionIndex}:${action.sheetName}:${action.range}`}>
            <strong>{action.type === "clearRange" ? "清空内容" : action.type === "writeFormulas" ? "写入公式" : "写入值"}</strong>
            <span>{action.sheetName}!{action.range}</span>
            <small>{action.rowCount} 行 × {action.columnCount} 列</small>
            <div className="wps-plan-values" aria-label={`${action.sheetName}!${action.range} 待写内容预览`}>
              <table><tbody>
                {action.preview.map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    <th scope="row">{rowIndex + 1}</th>
                    {row.map((value, columnIndex) => (
                      <td key={columnIndex}>{value === null ? "空单元格" : value === "" ? "空字符串" : String(value)}</td>
                    ))}
                  </tr>
                ))}
              </tbody></table>
              {(action.rowCount > 5 || action.columnCount > 5) && <small>此处仅显示前 5 行、前 5 列；编辑计划可查看完整矩阵。</small>}
            </div>
          </div>
        ))}
      </div>
      {pendingPlan.plan.warnings.length > 0 && (
        <ul className="wps-plan-warnings">
          {pendingPlan.plan.warnings.map((warning) => <li key={warning}>{warning}</li>)}
        </ul>
      )}
      {planEditDraft !== null && (
        <div className="wps-plan-editor">
          <label htmlFor="wps-plan-edit-json">编辑写入计划 JSON</label>
          <textarea
            id="wps-plan-edit-json"
            aria-label="编辑写入计划 JSON"
            rows={12}
            value={planEditDraft}
            disabled={busy}
            onChange={(event) => { setPlanEditDraft(event.target.value); setPlanEditError(null); }}
          />
          {planEditError && <p role="alert">{planEditError}</p>}
        </div>
      )}
      <div className="wps-plan-buttons">
        {planEditDraft === null ? (
          <>
            <button type="button" disabled={busy} onClick={() => { setPlanEditDraft(JSON.stringify(pendingPlan.plan, null, 2)); setPlanEditError(null); }}>编辑计划</button>
            <button type="button" disabled={busy} onClick={() => { markPlanPreview("cancelled"); setPendingPlan(null); setPendingPlanMessageId(null); }}>取消</button>
            <button className="primary" type="button" disabled={busy} onClick={() => void confirmPendingPlan()}>确认写入</button>
          </>
        ) : (
          <>
            <button type="button" disabled={busy} onClick={() => { setPlanEditDraft(null); setPlanEditError(null); }}>取消编辑</button>
            <button className="primary" type="button" disabled={busy} onClick={() => void applyEditedPlan()}>应用修改并重新预览</button>
          </>
        )}
      </div>
    </section>
  );

  const progressCard = (progress || writeProgress) && (
    <section className="agent-progress" aria-label="任务执行过程">
      <div className="agent-progress-heading">
        <strong>{writeProgress ? "写入过程" : "任务过程"}</strong>
        <span>{writeProgress ? "可撤销" : "处理中"}</span>
      </div>
      <div className="agent-progress-current">
        <span className="progress-dot" aria-hidden="true" />
        <strong>{writeProgress ? "执行与校验" : "处理中"}</strong>
        <span>{writeProgress ?? progress?.message}</span>
        {!writeProgress && <button type="button" onClick={() => controllerRef.current?.abort()}>停止</button>}
      </div>
    </section>
  );

  return (
    <section className="wps-chat-shell chat-shell">
      <header
        className={`chat-header${modelMenuOpen || moreMenuOpen ? " menu-open" : ""}`}
      >
        <div className="brand-mark">EB</div>
        <div className="brand-copy">
            <strong className="app-logo" {...logoLongPress}>Excel Bro</strong>
            <span className={modelCatalogFailed ? "offline" : "model-online online"}>
              <i aria-hidden="true" />
              {brandStatusText}
            </span>
        </div>
          <div
            className="model-picker"
            onBlur={(event) => {
              if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
                setModelMenuOpen(false);
              }
            }}
          >
            <button
              className="model-picker-trigger"
              type="button"
              disabled={busy}
              aria-haspopup="menu"
              aria-expanded={modelMenuOpen}
              title="选择或添加模型"
              onClick={() => {
                const nextOpen = !modelMenuOpen;
                closeHeaderSurfaces();
                setModelMenuOpen(nextOpen);
              }}
            >
              <span>{modelNotConfigured || modelCatalogFailed ? "添加模型" : "模型"}</span>
              <svg viewBox="0 0 16 16" aria-hidden="true"><path d="m4 6 4 4 4-4" /></svg>
            </button>
            {modelMenuOpen && (
              <div className="model-menu" role="menu" aria-label="模型">
                <span className="model-menu-title">选择模型</span>
                {modelCatalogLoading && (
                  <p className="eb-model-menu-note">正在读取模型配置…</p>
                )}
                {modelCatalogFailed && (
                  <p className="eb-model-menu-note">
                    暂时无法取得已保存的模型连接。可先填写模型信息；测试和保存前需启动 Excel Bro 后台。已有配置会在连接恢复后显示。
                  </p>
                )}
                {modelNotConfigured && (
                  <p className="eb-model-menu-note">
                    尚未配置模型。添加一个模型连接后，即可通过对话查询和分析表格。
                  </p>
                )}
                {models.map((model) => (
                  <button
                    key={model.id}
                    type="button"
                    role="menuitemradio"
                    aria-checked={model.id === effectiveModelId}
                    disabled={!model.available}
                    onClick={() => {
                      if (!isRealModelId(model.id)) return;
                      modelManagement.selectModel(model.id);
                      setModelMenuOpen(false);
                    }}
                  >
                    <i aria-hidden="true" />
                    <span>{model.label}{model.available ? "" : "（未配置）"}</span>
                    {model.id === effectiveModelId && <b aria-hidden="true">✓</b>}
                  </button>
                ))}
                <div className="model-menu-actions">
                  <button type="button" role="menuitem" onClick={() => void openModelSettings(true)}>
                    ＋ 添加模型连接
                  </button>
                  <button type="button" role="menuitem" onClick={() => void openModelSettings(false)}>
                    管理模型连接
                  </button>
                  {modelCatalogFailed && (
                    <button
                      type="button"
                      role="menuitem"
                      onClick={() => void refreshModelCatalog()}
                    >
                      重试读取模型配置
                    </button>
                  )}
                </div>
              </div>
            )}
          </div>
        <div className="header-actions">
          <button
            className="header-button labeled-header-button new-chat-entry"
            type="button"
            disabled={busy}
            title="新对话"
            aria-label="新对话"
            onClick={startNewConversation}
          >
            <span aria-hidden="true">＋</span><span>新对话</span>
          </button>
          <button
            className="header-button labeled-header-button tools-entry"
            type="button"
            disabled={busy}
            title="我的工具"
            aria-label="我的工具"
            onClick={() => {
              const nextOpen = !toolsOpen;
              closeHeaderSurfaces();
              setToolsOpen(nextOpen);
            }}
          >
            <svg viewBox="0 0 24 24" aria-hidden="true">
              <path d="M8.5 7V5.8c0-1 .8-1.8 1.8-1.8h3.4c1 0 1.8.8 1.8 1.8V7" />
              <rect x="3" y="7" width="18" height="12.5" rx="2.5" />
              <path d="M3.5 11.5h17M10 11.5v2h4v-2" />
            </svg>
            <span>工具</span>
          </button>
          <button
            className="header-button labeled-header-button folder-backup-entry"
            type="button"
            disabled={busy}
            title="备份与恢复"
            aria-label="备份与恢复"
            onClick={() => {
              const nextOpen = !backupOpen;
              closeHeaderSurfaces();
              setBackupOpen(nextOpen);
            }}
          >
            <span aria-hidden="true">↶</span><span>备份</span>
            {undoRecord && <b>1</b>}
          </button>
          <button
            className={`header-button pet-toggle${petVisible ? " active" : ""}`}
            type="button"
            aria-label={petVisible ? "隐藏格仔" : "显示格仔"}
            aria-pressed={petVisible}
            title={petVisible ? "隐藏格仔" : "显示格仔"}
            onClick={() => setPetVisible((visible) => !visible)}
          >
            <svg aria-hidden="true" viewBox="0 0 24 24">
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
              type="button"
              aria-label="更多"
              title="更多"
              aria-haspopup="menu"
              aria-expanded={moreMenuOpen}
              onClick={() => {
                const nextOpen = !moreMenuOpen;
                closeHeaderSurfaces();
                setMoreMenuOpen(nextOpen);
              }}
            >
              <svg viewBox="0 0 24 24" aria-hidden="true">
                <circle cx="5" cy="12" r="1.6" />
                <circle cx="12" cy="12" r="1.6" />
                <circle cx="19" cy="12" r="1.6" />
              </svg>
            </button>
            {moreMenuOpen && (
              <>
                <div className="more-menu-backdrop" aria-hidden="true" onClick={() => setMoreMenuOpen(false)} />
                <div className="view-menu" role="menu" aria-label="更多">
                  <button type="button" role="menuitem" aria-label="历史对话" onClick={() => { setMoreMenuOpen(false); setHistoryOpen(true); }}>
                    <strong>历史对话</strong><span>查看当前 WPS 会话记录</span>
                  </button>
                  <button type="button" role="menuitem" onClick={() => { setMoreMenuOpen(false); setRuleManagerOpen(true); }}>
                    <strong>EB 函数说明</strong><span>查看内置 =EB() 函数用法</span>
                  </button>
                  <button type="button" role="menuitem" onClick={() => { setMoreMenuOpen(false); setThemeOpen(true); }}>
                    <strong>外观与壁纸</strong><span>设置主题或自定义背景</span>
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      </header>

      {toolsOpen && (
        <aside className="tool-drawer" aria-label="我的工具">
          <div className="tool-drawer-header">
            <div><strong>我的工具</strong><span>选择后仍会先校验并预览，不会立即修改工作簿</span></div>
            <button type="button" aria-label="关闭" onClick={() => setToolsOpen(false)}>×</button>
          </div>
          <div className="tool-library">
            {tools.length === 0 && queryTools.length === 0 ? (
              <div className="tool-empty">
                <i>▦</i><strong>工具箱还是空的</strong><span>执行并验证一个计划后，可以把它保存到这里。</span>
              </div>
            ) : (
              <>
                <div className="tool-library-intro">
                  <span>工具箱</span><strong>今天想用哪一个？</strong>
                  <p>这里显示当前 WPS 加载项保存的工具；运行前仍会按 WPS 白名单和本次授权重新校验。</p>
                </div>
                <div className="tool-card-grid">
                  {tools.map((tool) => (
                    <article className="tool-card-wrap" key={tool.id}>
                      <button className="tool-card" type="button" onClick={() => void prepareSavedWorkflow(tool)}>
                        <span className="tool-card-icon tone-green">▦</span>
                        <span className="tool-card-copy">
                          <small>工作流</small><strong>{tool.name}</strong><span>{tool.description}</span>
                        </span>
                      </button>
                      <div className="tool-card-footer"><span className="tool-card-meta">{tool.planTemplate.actions.length} 个步骤 · 点击后预览</span></div>
                    </article>
                  ))}
                  {queryTools.map((tool) => (
                    <article className="tool-card-wrap" key={tool.id}>
                      <button
                        className="tool-card"
                        type="button"
                        onClick={() => void runSavedQuery(tool)}
                      >
                        <span className="tool-card-icon tone-blue">查</span>
                        <span className="tool-card-copy">
                          <small>固化查询</small><strong>{tool.name}</strong><span>{tool.description}</span>
                        </span>
                      </button>
                      <div className="tool-card-footer"><span className="tool-card-meta">本地确定性执行 · 模型调用 0 次</span></div>
                    </article>
                  ))}
                </div>
              </>
            )}
          </div>
        </aside>
      )}

      {backupOpen && (
        <aside className="tool-drawer" aria-label="备份与恢复">
          <div className="tool-drawer-header">
            <div><strong>备份与恢复</strong><span>WPS 写入前保存完整快照，并在撤销前核对后续修改</span></div>
            <button type="button" aria-label="关闭" onClick={() => setBackupOpen(false)}>×</button>
          </div>
          <div className="wps-backup-content">
            {undoRecord ? (
              <section className="wps-backup-card">
                <span>最近恢复点</span><strong>{undoRecord.planTitle}</strong>
                <small>{undoRecord.state === "undo_conflict" ? "检测到后续修改，需要先处理冲突" : "当前可安全撤销"}</small>
                <button type="button" disabled={busy || undoRecord.state === "undo_conflict"} onClick={() => void undoLatestExecution()}>↶ 撤销本次执行</button>
              </section>
            ) : (
              <div className="tool-empty"><i>↶</i><strong>还没有恢复点</strong><span>只有确认并完成写入后才会创建，不会生成空备份。</span></div>
            )}
          </div>
        </aside>
      )}

      {historyOpen && (
        <aside className="tool-drawer history-drawer" aria-label="历史对话">
          <div className="tool-drawer-header">
            <div><strong>历史对话</strong><span>新建对话不会覆盖之前的记录</span></div>
            <button type="button" aria-label="关闭" onClick={() => setHistoryOpen(false)}>×</button>
          </div>
          <div className="history-list">
            {[...chatHistory.conversations]
              .sort((left, right) => right.updatedAt.localeCompare(left.updatedAt))
              .map((conversation) => {
                const active = conversation.id === chatHistory.activeConversationId;
                const questionCount = conversation.messages.filter(
                  (message) => message.role === "user"
                ).length;
                return (
                  <article className={`history-item${active ? " active" : ""}`} key={conversation.id}>
                    <button
                      className="history-open-button"
                      type="button"
                      disabled={busy}
                      onClick={() => openConversation(conversation.id)}
                    >
                      <span className="history-title-row">
                        <strong>{conversation.title}</strong>
                        {active && <em>当前</em>}
                      </span>
                      <span>
                        {formatConversationTime(conversation.updatedAt)}
                        {questionCount > 0 ? ` · ${questionCount} 条提问` : " · 尚未提问"}
                      </span>
                    </button>
                    <button
                      className="history-delete-button"
                      type="button"
                      disabled={busy}
                      title={`删除「${conversation.title}」`}
                      aria-label={`删除「${conversation.title}」`}
                      onClick={() => deleteConversation(conversation.id)}
                    >
                      删除
                    </button>
                  </article>
                );
              })}
          </div>
          <button className="history-new-button" type="button" disabled={busy} onClick={startNewConversation}>＋ 新建对话</button>
        </aside>
      )}

      <div className={`message-stream-shell${newConversationNotice ? " conversation-starting" : ""}`}>
      {newConversationNotice && (
        <div className="new-conversation-toast" role="status">
          {newConversationNotice.previousSaved ? "已新建对话" : "当前已是新对话"}
        </div>
      )}
      {copyError && (
        <div className="new-conversation-toast" role="status">
          复制未成功；可以选中文字后使用 Ctrl+C
        </div>
      )}
      <main ref={messageStreamRef} className="message-stream" aria-live="polite" onScroll={handleMessageStreamScroll}>
        {messages.length === 0 && (
          <article className="message-row assistant welcome">
            <div className="message-content">
              <div className="message-author">Excel Bro</div>
              <p className="message-text">
              {modelCatalogFailed
                ? "你好。点击上方「添加模型」开始连接；如果你已配置过模型，Excel Bro 后台连接后会自动显示。"
                : modelNotConfigured
                ? "你好。添加一个模型连接后，即可查询、分析或修改已授权的工作表。"
                : canWrite
                  ? "你好。直接告诉我你想查询、分析或修改什么；涉及写入时，我会先展示计划并等你确认。"
                  : `你好。你可以查询和分析工作表；写入前还需通过当前文件检查：${writeStatusReason}`}
              </p>
            </div>
          </article>
        )}

        {messages.map((message) => (
          <Fragment key={message.id}>
            <article className={`message-row ${message.role}`}>
              <div className="message-content">
              {message.role === "assistant" && (
                <div className="message-author">
                  Excel Bro <em>{message.resultContext ? "本地工具" : "Agent"}</em>
                </div>
              )}
              {message.resultContext ? (
                <section className="query-result-card">
                  <p className="message-text">{message.text}</p>
                  <div className="wps-query-result-meta">
                    <strong>{message.resultContext.title}</strong>
                    <span>显示 {Math.min(20, message.resultContext.rows.length)} / {message.resultContext.rows.length} 行</span>
                    {message.resultContext.sourceSheets.length > 0 && (
                      <small>数据来源：{message.resultContext.sourceSheets.join("、")}</small>
                    )}
                  </div>
                  <div className="preview-table-wrap">
                    <table>
                      <thead>
                        <tr>
                          {message.resultContext.headers.map((header) => <th key={header}>{header}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {message.resultContext.rows.slice(0, 20).map((row, rowIndex) => (
                          <tr key={rowIndex}>
                            {row.map((cell, columnIndex) => (
                              <td key={columnIndex}>{cell == null ? "" : String(cell)}</td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {message.resultContext.warnings.length > 0 && (
                    <ul className="wps-query-result-warnings">
                      {message.resultContext.warnings.map((warning, index) => <li key={`${index}:${warning}`}>{warning}</li>)}
                    </ul>
                  )}
                </section>
              ) : message.text ? (
                <p className="message-text">{message.text}</p>
              ) : null}
              {message.clarification && (
                <section className={`clarification-card ${message.clarification.status}`} aria-label="需求澄清">
                  <div className="clarification-heading">
                    <div>
                      <span>需要确认</span>
                      <strong>{message.clarification.summary}</strong>
                    </div>
                    <em>{message.clarification.scopeLabel}</em>
                  </div>
                  <p className="clarification-question">{message.clarification.question}</p>
                  <p className="clarification-reason">{message.clarification.reason}</p>
                  {message.clarification.status === "pending" && clarification?.messageId === message.id ? (
                    <>
                      <div className="clarification-options">
                        {message.clarification.options.map((option, index) => (
                          <button type="button" key={index} disabled={busy} onClick={() => resolveClarification(index)}>
                            <strong>{option.label}</strong>
                            <span>{option.description}</span>
                          </button>
                        ))}
                      </div>
                      <span className="clarification-custom-hint">其他想法也可以直接说明</span>
                      <div className="clarification-custom-input">
                        <textarea
                          aria-label="其他想法"
                          value={clarificationDraft}
                          placeholder="补充你希望我怎么处理…"
                          onChange={(event) => setClarificationDraft(event.target.value)}
                          onKeyDown={(event) => {
                            if (event.key === "Enter" && !event.shiftKey) {
                              event.preventDefault();
                              resolveClarificationWithText();
                            }
                          }}
                        />
                        <button type="button" disabled={!clarificationDraft.trim() || busy} onClick={resolveClarificationWithText}>发送</button>
                      </div>
                      <div className="clarification-actions">
                        <button type="button" disabled={busy} onClick={cancelClarification}>取消本次澄清</button>
                      </div>
                    </>
                  ) : (
                    <div className="clarification-outcome">
                      {message.clarification.status === "resolved"
                        ? `已按「${message.clarification.answer ?? "补充说明"}」继续`
                        : message.clarification.status === "cancelled"
                          ? "已取消本次澄清，未修改工作表"
                          : "这张澄清卡已失效；如需继续，请重新发送需求"}
                    </div>
                  )}
                </section>
              )}
              {message.planPreview && (
                pendingPlan && pendingPlanMessageId === message.id
                  ? planCard
                  : <div className="wps-plan-outcome" role="status">
                      <strong>{message.planPreview.title}</strong>
                      <span>{message.planPreview.status === "executed"
                        ? "计划已执行并回读确认"
                        : message.planPreview.status === "cancelled"
                          ? "计划已取消，未写入工作表"
                          : "计划预览已失效；如需写入，请重新提出任务"}</span>
                    </div>
              )}
              {message.verification && (
                <section className={`verification-card wps-verification-card ${message.verification.status}`} aria-label="写入验证结果">
                  <strong>{message.verification.status === "undone" ? "已撤销并核对恢复" : "写入验证通过"}</strong>
                  <span>{message.verification.actionCount} 个动作 · {message.verification.cellCount} 个单元格已回读核对</span>
                  <ul>
                    {message.verification.targets.slice(0, 5).map((target, index) => (
                      <li key={`${message.id}-target-${index}`}>
                        {target.type === "clearRange" ? "清空内容" : target.type === "writeFormulas" ? "写入公式" : "写入值"} · {target.sheetName}!{target.range}
                      </li>
                    ))}
                  </ul>
                  {message.verification.targets.length > 5 && (
                    <small>另有 {message.verification.targets.length - 5} 个目标，详见执行过程。</small>
                  )}
                  <small>恢复点仅在当前加载项会话内可用；关闭后不能从此卡片撤销。</small>
                </section>
              )}
              {message.activityLog && message.activityLog.steps.length > 0 && (
                <TaskProgressTimeline
                  log={message.activityLog}
                  live={activeActivityMessageId === message.id}
                  onStop={activeActivityMessageId === message.id && progress
                    ? () => controllerRef.current?.abort()
                    : undefined}
                />
              )}
              {message.text.trim() && (
                <div className="message-copy-row">
                  <button
                    className={`message-copy-button${copiedMessageId === message.id ? " copied" : ""}`}
                    type="button"
                    aria-label={copiedMessageId === message.id ? "消息已复制" : "复制这条消息"}
                    onClick={() => void copyMessage(message)}
                  >
                    {copiedMessageId === message.id ? "✓ 已复制" : "复制"}
                  </button>
                </div>
              )}
              </div>
            </article>
          </Fragment>
        ))}

        {undoRecord && (
          <section className="wps-undo-card" aria-label="WPS 撤销恢复点">
            <div>
              <strong>{undoRecord.state === "undo_conflict" ? "撤销需要处理冲突" : "已保存恢复点"}</strong>
              <span>{undoRecord.planTitle}</span>
            </div>
            <button type="button" disabled={busy} onClick={() => void undoLatestExecution()}>↶ 撤销本次执行</button>
          </section>
        )}

        {(progress || writeProgress) && !activeActivityMessageId && progressCard}
      </main>
      {hasUnseenProgress && (
        <button type="button" className="new-progress-button" onClick={scrollToLatest} aria-label="查看最新执行进度">
          <span aria-hidden="true">↓</span><span>查看最新进度</span>
        </button>
      )}
      </div>

      {historyNotice && (
        <div className="storage-write-warning" role="alert">
          <strong>对话历史没能保存到本机</strong>
          <span>{historyNotice}</span>
        </div>
      )}

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
        {scopeDock}
        <form
          className="composer-box"
          onSubmit={(event) => {
            event.preventDefault();
            void runTurn(prompt);
          }}
        >
          <textarea
            ref={composerInputRef}
            aria-label="给 Excel Bro 发消息"
            value={prompt}
            disabled={busy}
            placeholder="描述你想查询、分析或修改的内容…"
            onChange={(event) => setPrompt(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === "Enter" && !event.shiftKey) {
                event.preventDefault();
                void runTurn(prompt);
              }
            }}
          />
          <div className="composer-toolbar">
            <button
              className="attach-image-button composer-image-button"
              type="button"
              disabled
              title="WPS 当前暂未开放图片输入"
            >
              ▧ 图片
            </button>
            <span className="composer-shortcut-hint">Enter 发送 · Shift + Enter 换行</span>
            {progress ? (
              <button
                className="send-button composer-send-button wps-stop-button"
                type="button"
                aria-label="停止任务"
                title="停止当前读取或规划任务"
                onClick={() => controllerRef.current?.abort()}
              >
                ■
              </button>
            ) : (
              <button
                className="send-button composer-send-button"
                type="submit"
                aria-label={writeProgress ? "正在执行写入与回读" : "发送"}
                title={writeProgress ? "写入与回读期间不能中断，以免结果状态不明" : undefined}
                disabled={Boolean(writeProgress) || prompt.trim().length < 2 || selectedSheetIndexes.length === 0}
              >
                {writeProgress ? "…" : "↑"}
              </button>
            )}
          </div>
        </form>
        {(developmentWrite || !canWrite) && (
          <p className="wps-capability-note">
            {developmentWrite
              ? "开发验收模式：仅在独立测试工作簿中确认写入与撤销。"
              : `当前仍可查询；写入将在满足条件后开放：${writeStatusReason}`}
          </p>
        )}
      </footer>

      {newConversationConfirmOpen && (
        <div className="tool-dialog-backdrop" onMouseDown={() => setNewConversationConfirmOpen(false)}>
          <section
            className="tool-dialog history-delete-dialog"
            role="dialog"
            aria-modal="true"
            aria-labelledby="new-conversation-confirm-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <div className="tool-dialog-title-row">
              <strong id="new-conversation-confirm-title">放弃待确认的写入计划？</strong>
              <button type="button" aria-label="关闭" onClick={() => setNewConversationConfirmOpen(false)}>×</button>
            </div>
            <p>当前对话还有一份尚未执行的写入计划。新建对话会放弃这份计划，但不会修改工作簿。</p>
            <div className="tool-dialog-actions">
              <button type="button" onClick={() => setNewConversationConfirmOpen(false)}>继续当前对话</button>
              <button className="danger-button" type="button" onClick={completeNewConversation}>放弃计划并新建</button>
            </div>
          </section>
        </div>
      )}

      {petVisible && <PetCompanion busy={busy} />}
      <WpsModelSettingsPanel
        open={modelSettingsOpen}
        models={models}
        currentModelId={effectiveModelId}
        management={modelManagement}
        onClose={closeModelSettings}
        onAddConnection={() => void openModelSettings(true)}
      />
      <ThemePanel open={themeOpen} onClose={() => setThemeOpen(false)} {...theme} />
      <RuleManager isOpen={ruleManagerOpen} onClose={() => setRuleManagerOpen(false)} />
    </section>
  );
}
