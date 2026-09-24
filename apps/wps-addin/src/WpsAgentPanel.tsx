import { type ReactNode, useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  listModels,
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
import { emptyModelConnectionDraft } from "../../excel-addin/src/utils";
import ThemePanel from "../../excel-addin/src/components/ThemePanel";
import PetCompanion from "../../excel-addin/src/PetCompanion";
import WpsModelSettingsPanel from "./WpsModelSettingsPanel";
import {
  PET_VISIBILITY_STORAGE_KEY,
  normalizePetVisibility
} from "../../excel-addin/src/conversation";
import { getWpsApplication } from "./wpsHost";
import {
  runWpsReadonlyTurn,
  type WpsReadonlyTurnProgress
} from "./wpsReadonlyAgent";

type ChatMessage = {
  id: string;
  role: "user" | "assistant" | "system";
  text: string;
  resultContext?: ResultContext | null;
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
  scopeDock: ReactNode;
};

const progressLabels: Record<WpsReadonlyTurnProgress["stage"], string> = {
  reading: "读取工作表",
  intent: "理解问题",
  calculating: "本地统计",
  answering: "整理回答"
};

function messageId(): string {
  return `wps-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

export default function WpsAgentPanel({
  selectedSheetIndexes,
  scopeDock
}: WpsAgentPanelProps) {
  // 模型目录只包含用户真实配置的连接。初始列表为空、状态单独表达，
  // 不再注入虚拟的 “基础模式” 占位项——目录尚未加载或读取失败时，
  // 那会让人误以为“已经选好模型，可以开始对话”。
  const [models, setModels] = useState<ModelOption[]>([]);
  const [defaultModelId, setDefaultModelId] = useState<string | null>(null);
  const [catalogStatus, setCatalogStatus] = useState<ModelCatalogStatus>("idle");
  const [catalogError, setCatalogError] = useState("");
  const [prompt, setPrompt] = useState("");
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [progress, setProgress] = useState<WpsReadonlyTurnProgress | null>(null);
  const [clarification, setClarification] = useState<PendingClarification | null>(null);
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
  const [copiedMessageId, setCopiedMessageId] = useState<string | null>(null);
  const controllerRef = useRef<AbortController | null>(null);
  const messagesEndRef = useRef<HTMLDivElement | null>(null);
  const theme = useTheme();
  const logoLongPress = useLongPress(() => setThemeOpen(true));
  const busy = progress !== null;

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
      setCatalogError("");
      return {
        modelOptions: catalog.models,
        defaultModelId: catalog.defaultModelId
      };
    } catch (error) {
      setCatalogStatus("failed");
      setCatalogError(
        error instanceof Error
          ? `无法读取模型配置：${error.message}`
          : "无法读取模型配置，请确认 Excel Bro 后台正在运行。"
      );
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
    ? "无法读取模型配置"
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
    if (catalogStatus !== "failed") return;

    const retry = () => void refreshModelCatalog();
    const retryTimer = window.setTimeout(retry, 5000);
    window.addEventListener("focus", retry);
    return () => {
      window.clearTimeout(retryTimer);
      window.removeEventListener("focus", retry);
    };
  }, [catalogStatus, refreshModelCatalog]);

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
    try {
      localStorage.setItem(
        PET_VISIBILITY_STORAGE_KEY,
        petVisible ? "visible" : "hidden"
      );
    } catch {
      // 存储不可用时，本次会话仍保留当前选择。
    }
  }, [petVisible]);

  const append = (message: Omit<ChatMessage, "id">) => {
    setMessages((current) => [...current, { id: messageId(), ...message }]);
  };

  const startNewConversation = () => {
    setMessages([]);
    setPrompt("");
    setProgress(null);
    setClarification(null);
    setCopiedMessageId(null);
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
    if (!text || busy || selectedSheetIndexes.length === 0) return;
    // 没有可用模型时必须明确反馈，并保留用户已经输入的任务文本：
    // 不清空输入框、不伪造一次成功回答，也不擅自改用别的连接。
    if (!effectiveModelId) {
      setModelMenuOpen(false);
      append({
        role: "system",
        text:
          modelCatalogFailed
            ? "暂时无法读取模型配置，请点击上方模型菜单里的「重试」。已输入的内容仍保留在输入框中。"
            : "尚未配置模型。请点击上方「添加模型」，填写服务地址、模型 ID 和 API Key 后重试；已输入的内容仍保留在输入框中。"
      });
      return;
    }
    if (options.showUser !== false) append({ role: "user", text });
    setPrompt("");
    setClarification(null);
    const controller = new AbortController();
    controllerRef.current?.abort();
    controllerRef.current = controller;
    setProgress({ stage: "reading", message: "正在准备只读查询…" });
    try {
      const result = await runWpsReadonlyTurn({
        application: getWpsApplication(),
        selectedSheetIndexes,
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
        setClarification({
          originalPrompt: text,
          round: options.clarificationRound ?? 0,
          question: result.response.clarification.question,
          reason: result.response.clarification.reason,
          options: result.response.clarification.options
        });
      } else if (result.kind === "blocked") {
        append({ role: "system", text: result.message });
      } else {
        append({
          role: "assistant",
          text: result.message,
          resultContext: result.resultContext
        });
      }
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") {
        append({ role: "system", text: "只读查询已停止，没有修改任何工作表。" });
      } else {
        append({
          role: "system",
          text: error instanceof Error ? error.message : "只读查询失败。"
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
      // 模型来自用户配置的外部 API。后台暂时不可用时仍允许填写配置，
      // 测试或保存会执行真实请求并把连接错误反馈给用户。
      await modelManagement.openSettings();
      modelManagement.setConnectionDraft(emptyModelConnectionDraft());
    } else {
      await modelManagement.openSettings();
    }
  };

  return (
    <section className="wps-chat-shell chat-shell">
      <header className="eb-chat-header">
        <div className="eb-brand">
          <span className="eb-brand-mark">EB</span>
          <div className="eb-brand-copy">
            <strong className="app-logo" {...logoLongPress}>Excel Bro</strong>
            <span className={modelCatalogFailed ? "offline" : "online"}>
              <i aria-hidden="true" />
              {brandStatusText}
            </span>
          </div>
        </div>
        <div className="eb-header-actions">
          <div
            className="eb-model-picker"
            onBlur={(event) => {
              if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
                setModelMenuOpen(false);
              }
            }}
          >
            <button
              className="eb-model-trigger"
              type="button"
              disabled={busy}
              aria-haspopup="menu"
              aria-expanded={modelMenuOpen}
              title="选择或添加模型"
              onClick={() => setModelMenuOpen((open) => !open)}
            >
              <span>{modelNotConfigured || modelCatalogFailed ? "添加模型" : "模型"}</span>
              <svg viewBox="0 0 16 16" aria-hidden="true"><path d="m4 6 4 4 4-4" /></svg>
            </button>
            {modelMenuOpen && (
              <div className="eb-model-menu" role="menu" aria-label="模型">
                <span>选择模型</span>
                {modelCatalogLoading && (
                  <p className="eb-model-menu-note">正在读取模型配置…</p>
                )}
                {modelCatalogFailed && (
                  <p className="eb-model-menu-note">
                    无法读取模型配置。这不代表你没有配置模型，请恢复后台连接后重试。
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
                <div className="eb-model-menu-actions">
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
          <button className="eb-new-chat" type="button" disabled={busy} onClick={startNewConversation}>
            <span aria-hidden="true">＋</span> 新对话
          </button>
          <button
            className={`eb-pet-button${petVisible ? " active" : ""}`}
            type="button"
            aria-label={petVisible ? "隐藏格仔" : "显示格仔"}
            aria-pressed={petVisible}
            title={petVisible ? "隐藏格仔" : "显示格仔"}
            onClick={() => setPetVisible((visible) => !visible)}
          >
            <svg aria-hidden="true" viewBox="0 0 24 24">
              <circle cx="7.2" cy="7.2" r="2.2" />
              <circle cx="12" cy="5.5" r="2.25" />
              <circle cx="16.8" cy="7.2" r="2.2" />
              <circle cx="19" cy="11.6" r="2" />
              <path d="M12 10.2c-3.8 0-6.8 3.2-6.8 6.2 0 2.1 1.7 3.6 3.8 3.6 1.2 0 2-.6 3-.6s1.8.6 3 .6c2.1 0 3.8-1.5 3.8-3.6 0-3-3-6.2-6.8-6.2Z" />
            </svg>
          </button>
          <span className="eb-readonly-chip" title="WPS 当前仅开放只读查询">只读</span>
        </div>
      </header>

      <main className="eb-message-stream" aria-live="polite">
        {messages.length === 0 && (
          <article className="eb-message assistant welcome">
            <div className="eb-message-author">Excel Bro</div>
            <div className="eb-message-bubble">
              {modelNotConfigured
                ? "你好。添加一个模型连接后，即可通过对话查询和分析表格；WPS 当前处于只读模式，不会修改工作簿。"
                : "你好。直接告诉我你想查询或分析什么；WPS 当前处于只读模式，不会修改工作簿。"}
            </div>
          </article>
        )}

        {modelCatalogFailed && (
          <article className="eb-message system">
            <div className="eb-message-author">系统</div>
            <div className="eb-message-bubble">
              <p>
                {catalogError || "无法读取模型配置，请确认 Excel Bro 后台正在运行。"}
              </p>
              <button
                className="eb-model-retry"
                type="button"
                onClick={() => void refreshModelCatalog()}
              >
                重试
              </button>
            </div>
          </article>
        )}

        {messages.map((message) => (
          <article className={`eb-message ${message.role}`} key={message.id}>
            <div className="eb-message-author">
              {message.role === "user" ? "你" : message.role === "assistant" ? "Excel Bro" : "系统"}
            </div>
            <div className="eb-message-bubble">
              <p>{message.text}</p>
              {message.resultContext && (
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
              )}
            </div>
            {message.role === "assistant" && (
              <button
                className="eb-message-copy"
                type="button"
                aria-label={copiedMessageId === message.id ? "消息已复制" : "复制这条消息"}
                onClick={() => void copyMessage(message)}
              >
                {copiedMessageId === message.id ? "✓ 已复制" : "复制"}
              </button>
            )}
          </article>
        ))}

        {clarification && (
          <section className="clarification-card">
            <span className="clarification-eyebrow">执行前确认</span>
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
          </section>
        )}

        {progress && (
          <section className="agent-progress" aria-label="只读查询过程">
            <div className="agent-progress-heading">
              <strong>查询过程</strong>
              <span>只读</span>
            </div>
            <div className="agent-progress-current">
              <span className="progress-dot" aria-hidden="true" />
              <strong>{progressLabels[progress.stage]}</strong>
              <span>{progress.message}</span>
              <button type="button" onClick={() => controllerRef.current?.abort()}>停止</button>
            </div>
          </section>
        )}
        <div ref={messagesEndRef} />
      </main>

      <footer className="eb-composer-zone">
        {scopeDock}
        <form
          className="eb-composer-box"
          onSubmit={(event) => {
            event.preventDefault();
            void runTurn(prompt);
          }}
        >
          <textarea
            aria-label="输入消息"
            value={prompt}
            disabled={busy}
            placeholder="描述你想查询、分析的内容…"
            onChange={(event) => setPrompt(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === "Enter" && !event.shiftKey) {
                event.preventDefault();
                void runTurn(prompt);
              }
            }}
          />
          <div className="eb-composer-actions">
            <button
              className="composer-image-button"
              type="button"
              disabled
              title="WPS 只读版暂未开放图片输入"
            >
              ▧ 图片
            </button>
            <span>Enter 发送 · Shift + Enter 换行</span>
            <button
              className="composer-send-button"
              type="submit"
              aria-label="发送"
              disabled={busy || prompt.trim().length < 2 || selectedSheetIndexes.length === 0}
            >
              ↑
            </button>
          </div>
        </form>
        <p className="eb-readonly-note">WPS 当前仅支持只读查询；写入、删除、拆表、撤销和恢复均不会执行。</p>
      </footer>

      {petVisible && <PetCompanion busy={busy} />}
      <WpsModelSettingsPanel
        open={modelSettingsOpen}
        backendAvailable={!modelCatalogFailed}
        models={models}
        currentModelId={effectiveModelId}
        management={modelManagement}
        onClose={closeModelSettings}
        onAddConnection={() => void openModelSettings(true)}
      />
      <ThemePanel open={themeOpen} onClose={() => setThemeOpen(false)} {...theme} />
    </section>
  );
}
