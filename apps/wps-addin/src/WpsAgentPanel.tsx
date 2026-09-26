import { type ReactNode, useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  listModels,
  validatePlan,
  type ModelCatalogStatus,
  type ModelOption
} from "../../excel-addin/src/api";
import type { IntentOption, ResultContext } from "../../excel-addin/src/contracts";
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
  runWpsReadonlyTurn,
  type WpsReadonlyTurnProgress
} from "./wpsReadonlyAgent";

type ChatMessage = {
  id: string;
  role: "user" | "assistant" | "system";
  text: string;
  resultContext?: ResultContext | null;
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
};

type PendingClarification = {
  originalPrompt: string;
  round: number;
  question: string;
  reason: string;
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

const progressLabels: Record<WpsReadonlyTurnProgress["stage"], string> = {
  reading: "读取工作表",
  intent: "理解问题",
  calculating: "本地统计",
  answering: "整理回答"
};

function messageId(): string {
  return `wps-${Date.now()}-${Math.random().toString(16).slice(2)}`;
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
          messages: conversation.messages.filter((message) =>
            !(message.role === "system" && message.text === LEGACY_MODEL_CATALOG_NOTICE)
          )
        }))
      : [];
    if (conversations.length > 0) {
      const activeConversationId = conversations.some(
        (conversation) => conversation.id === stored?.activeConversationId
      )
        ? String(stored?.activeConversationId)
        : conversations[0].id;
      return { activeConversationId, conversations };
    }
  } catch {
    // 损坏或不可用的宿主存储不应阻断任务窗格启动。
  }
  const conversation = createWpsConversation();
  return { activeConversationId: conversation.id, conversations: [conversation] };
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
  const [chatHistory, setChatHistory] = useState<WpsChatHistory>(loadWpsChatHistory);
  const messages = useMemo(
    () =>
      chatHistory.conversations.find(
        (conversation) => conversation.id === chatHistory.activeConversationId
      )?.messages ?? [],
    [chatHistory]
  );
  const [progress, setProgress] = useState<WpsReadonlyTurnProgress | null>(null);
  const [writeProgress, setWriteProgress] = useState<string | null>(null);
  const [pendingPlan, setPendingPlan] = useState<WpsPreparedPlan | null>(null);
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
  const writeStatusReason = writeCapability?.checks.find((check) => !check.passed)?.detail
    ?? "确认任务时会检查当前工作簿是否满足写入条件";
  const messagesEndRef = useRef<HTMLDivElement | null>(null);
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
    messagesEndRef.current?.scrollIntoView({ block: "end" });
  }, [messages, progress, clarification]);

  useEffect(() => {
    if (!newConversationNotice) return;
    const timer = window.setTimeout(() => setNewConversationNotice(null), 1500);
    return () => window.clearTimeout(timer);
  }, [newConversationNotice]);

  useEffect(() => {
    try {
      const active = chatHistory.activeConversationId;
      const compact: WpsChatHistory = {
        activeConversationId: active,
        conversations: [...chatHistory.conversations]
          .sort((left, right) => right.updatedAt.localeCompare(left.updatedAt))
          .slice(0, 30)
          .map((conversation) => ({
            ...conversation,
            messages: conversation.messages.slice(-120)
          }))
      };
      localStorage.setItem(WPS_CHAT_STORAGE_KEY, JSON.stringify(compact));
    } catch {
      // 宿主禁用 localStorage 时只影响跨重启历史，不影响当前会话。
    }
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

  const updateMessages = useCallback(
    (updater: (current: ChatMessage[]) => ChatMessage[]) => {
      setChatHistory((current) => {
        const now = new Date().toISOString();
        return {
          ...current,
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

  const append = (message: Omit<ChatMessage, "id" | "createdAt">) => {
    updateMessages((current) => [
      ...current,
      { id: messageId(), createdAt: new Date().toISOString(), ...message }
    ]);
  };

  const completeNewConversation = () => {
    const previousSaved = messages.some((message) => message.role === "user");
    if (previousSaved) {
      const conversation = createWpsConversation();
      setChatHistory((current) => ({
        activeConversationId: conversation.id,
        conversations: [conversation, ...current.conversations].slice(0, 30)
      }));
    }
    setPrompt("");
    setProgress(null);
    setClarification(null);
    setPendingPlan(null);
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
        ? { ...current, activeConversationId: conversationId }
        : current
    );
    setPrompt("");
    setClarification(null);
    setClarificationDraft("");
    setPendingPlan(null);
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
  };

  const closeHeaderSurfaces = () => {
    setModelMenuOpen(false);
    setToolsOpen(false);
    setBackupOpen(false);
    setHistoryOpen(false);
    setMoreMenuOpen(false);
  };

  const copyMessage = async (message: ChatMessage) => {
    try {
      await navigator.clipboard.writeText(message.text);
      setCopiedMessageId(message.id);
    } catch {
      append({ role: "system", text: "复制失败，请选中文字后手动复制。" });
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
    // 没有可用模型时必须明确反馈，并保留用户已经输入的任务文本：
    // 不清空输入框、不伪造一次成功回答，也不擅自改用别的连接。
    if (!effectiveModelId) {
      setModelMenuOpen(true);
      return;
    }
    if (options.showUser !== false) append({ role: "user", text });
    setPrompt("");
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
          .filter((message) => message.role !== "system")
          .slice(-6)
          .map((message) => ({
            role: message.role as "user" | "assistant",
            text: message.text.slice(0, 1000)
          })),
        signal: controller.signal,
        onProgress: setProgress
      });
      if (result.kind === "clarification") {
        setClarificationDraft("");
        setClarification({
          originalPrompt: text,
          round: options.clarificationRound ?? 0,
          question: result.response.clarification.question,
          reason: result.response.clarification.reason,
          options: result.response.clarification.options
        });
      } else if (result.kind === "plan") {
        setProgress({ stage: "answering", message: "正在校验写入计划和授权边界…" });
        const application = getWpsApplication();
        const metadata = listWpsWorksheets(application);
        const selectedNames = metadata.sheets
          .filter((sheet) => turnSheetIndexes.includes(sheet.index))
          .map((sheet) => sheet.name);
        const normalizedPlan = await validatePlan(
          result.plan,
          selectedNames,
          metadata.sheets.map((sheet) => sheet.name)
        );
        setPendingPlan(prepareWpsPlan(application, turnSheetIndexes, normalizedPlan));
      } else {
        append({
          role: "assistant",
          text: result.message,
          resultContext: result.resultContext
        });
      }
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") {
        append({ role: "system", text: "任务已停止；未确认的计划不会修改工作表。" });
      } else {
        append({
          role: "system",
          text: error instanceof Error ? error.message : "任务处理失败。"
        });
      }
    } finally {
      if (controllerRef.current === controller) controllerRef.current = null;
      setProgress(null);
    }
  };

  const confirmPendingPlan = async () => {
    if (!pendingPlan || busy) return;
    const plan = pendingPlan;
    setWriteProgress("正在保存执行前恢复点…");
    try {
      const record = await executeWpsPlan(getWpsApplication(), plan, ({ message }) => {
        setWriteProgress(message);
      });
      setUndoRecord(record);
      setPendingPlan(null);
      append({
        role: "assistant",
        text: `已执行「${plan.plan.title}」，并回读确认 ${plan.totalCells} 个单元格。已保存恢复点，可撤销本次执行。`
      });
    } catch (error) {
      try {
        setUndoRecord(latestWpsUndo(getWpsApplication()));
      } catch {
        setUndoRecord(null);
      }
      append({
        role: "system",
        text: error instanceof Error ? error.message : "WPS 写入失败。"
      });
    } finally {
      setWriteProgress(null);
    }
  };

  const undoLatestExecution = async () => {
    if (!undoRecord || busy) return;
    setWriteProgress("正在核对撤销条件…");
    try {
      await undoWpsExecution(getWpsApplication(), undoRecord.executionId, ({ message }) => {
        setWriteProgress(message);
      });
      setUndoRecord(null);
      append({ role: "assistant", text: `已撤销「${undoRecord.planTitle}」，并回读确认恢复成功。` });
    } catch (error) {
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
      setPendingPlan(prepareWpsPlan(application, toolSheetIndexes, normalizedPlan));
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
      const authorization = authorizeWpsSheets(application, targetIndexes);
      const results = await readAuthorizedWpsSheets(application, authorization, {
        signal: controller.signal,
        onProgress: ({ sheet, rowsRead, totalRows }) => {
          setProgress({
            stage: "reading",
            message: `正在读取「${sheet.name}」：${rowsRead} / ${totalRows} 行`
          });
        }
      });
      setProgress({ stage: "calculating", message: "正在本地确定性执行固化查询…" });
      const result = executeQueryTableData(
        instantiateQueryTool(tool),
        results.map((item) => ({ name: item.sheet.name, values: item.values }))
      );
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
        append({ role: "system", text: "固化查询已停止，没有修改任何工作表。" });
      } else {
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

  const resolveClarification = (option: IntentOption) => {
    if (!clarification) return;
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
    append({ role: "user", text: answer });
    void runTurn(`${clarification.originalPrompt}\n补充确认：${answer}`, {
      intentConfirmed: true,
      clarificationRound: clarification.round + 1,
      showUser: false
    });
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
      <main className="message-stream" aria-live="polite">
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
          <article className={`message-row ${message.role}`} key={message.id}>
            <div className="message-content">
              {message.role === "assistant" && (
                <div className="message-author">
                  Excel Bro <em>{message.resultContext ? "本地工具" : "Agent"}</em>
                </div>
              )}
              {message.resultContext ? (
                <section className="query-result-card">
                  <p className="message-text">{message.text}</p>
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
                </section>
              ) : (
                <p className="message-text">{message.text}</p>
              )}
              {message.role === "assistant" && (
                <button
                  className={`message-copy-button${copiedMessageId === message.id ? " copied" : ""}`}
                  type="button"
                  aria-label={copiedMessageId === message.id ? "消息已复制" : "复制这条消息"}
                  onClick={() => void copyMessage(message)}
                >
                  {copiedMessageId === message.id ? "✓ 已复制" : "复制"}
                </button>
              )}
            </div>
          </article>
        ))}

        {clarification && (
          <section className="clarification-card">
            <span className="clarification-eyebrow">确认你的需求</span>
            <strong>{clarification.question}</strong>
            <p>{clarification.reason}</p>
            <div className="clarification-options">
              {clarification.options.map((option) => (
                <button type="button" key={option.id} onClick={() => resolveClarification(option)}>
                  <span>{option.label}</span>
                  <small>{option.description}</small>
                </button>
              ))}
            </div>
            <div className="clarification-custom">
              <input
                aria-label="其他想法"
                value={clarificationDraft}
                placeholder="其他想法：直接说明你的需求"
                onChange={(event) => setClarificationDraft(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === "Enter") resolveClarificationWithText();
                }}
              />
              <button type="button" disabled={!clarificationDraft.trim() || busy} onClick={resolveClarificationWithText}>发送</button>
            </div>
          </section>
        )}

        {pendingPlan && (
          <section className="wps-plan-card" aria-label="WPS 写入计划预览">
            <span className="clarification-eyebrow">写入前确认</span>
            <strong>{pendingPlan.plan.title}</strong>
            <p>{pendingPlan.plan.summary}</p>
            <div className="wps-plan-meta">
              <span>{pendingPlan.actions.length} 个操作</span>
              <span>{pendingPlan.totalCells} 个单元格</span>
              <span>可撤销</span>
            </div>
            <div className="wps-plan-actions-list">
              {pendingPlan.actions.map((action) => (
                <div key={`${action.actionIndex}:${action.sheetName}:${action.range}`}>
                  <strong>{action.type === "clearRange" ? "清空内容" : "写入值"}</strong>
                  <span>{action.sheetName}!{action.range}</span>
                  <small>{action.rowCount} 行 × {action.columnCount} 列</small>
                </div>
              ))}
            </div>
            {pendingPlan.plan.warnings.length > 0 && (
              <ul className="wps-plan-warnings">
                {pendingPlan.plan.warnings.map((warning) => <li key={warning}>{warning}</li>)}
              </ul>
            )}
            <div className="wps-plan-buttons">
              <button type="button" disabled={busy} onClick={() => setPendingPlan(null)}>取消</button>
              <button className="primary" type="button" disabled={busy} onClick={() => void confirmPendingPlan()}>
                确认写入
              </button>
            </div>
          </section>
        )}

        {undoRecord && (
          <section className="wps-undo-card" aria-label="WPS 撤销恢复点">
            <div>
              <strong>{undoRecord.state === "undo_conflict" ? "撤销需要处理冲突" : "已保存恢复点"}</strong>
              <span>{undoRecord.planTitle}</span>
            </div>
            <button type="button" disabled={busy} onClick={() => void undoLatestExecution()}>↶ 撤销本次执行</button>
          </section>
        )}

        {(progress || writeProgress) && (
          <section className="agent-progress" aria-label="任务执行过程">
            <div className="agent-progress-heading">
              <strong>{writeProgress ? "写入过程" : "任务过程"}</strong>
              <span>{writeProgress ? "可撤销" : "处理中"}</span>
            </div>
            <div className="agent-progress-current">
              <span className="progress-dot" aria-hidden="true" />
              <strong>{writeProgress ? "执行与校验" : progress ? progressLabels[progress.stage] : "处理中"}</strong>
              <span>{writeProgress ?? progress?.message}</span>
              {!writeProgress && <button type="button" onClick={() => controllerRef.current?.abort()}>停止</button>}
            </div>
          </section>
        )}
        <div ref={messagesEndRef} />
      </main>
      </div>

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
            <button
              className="send-button composer-send-button"
              type="submit"
              aria-label="发送"
              disabled={busy || prompt.trim().length < 2 || selectedSheetIndexes.length === 0}
            >
              ↑
            </button>
          </div>
        </form>
        <p className="wps-capability-note">
          {canWrite
            ? "修改会严格限制在已授权工作表，先预览确认、再执行；当前支持写入值、清空内容和撤销。"
            : `当前仍可查询；写入将在满足条件后开放：${writeStatusReason}`}
        </p>
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
