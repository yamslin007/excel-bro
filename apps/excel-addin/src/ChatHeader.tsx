// 顶部聊天头：品牌、模型选择、首次引导、新对话 / 工具 / 宠物 / 更多菜单。
// 纯展示组件，JSX 从 App.tsx 逐字搬移；所有状态与回调都由 props 传入。
import type { Dispatch, SetStateAction } from "react";
import type { ModelOption, ServiceHealth } from "./api";
import { BASE_MODE_DESCRIPTION } from "./helpCommand";

interface ChatHeaderProps {
  serverOnline: boolean;
  serviceHealth: ServiceHealth | null;
  selectedModelId: string;
  selectedModel: ModelOption | null;
  modelOptions: ModelOption[];
  modelMenuOpen: boolean;
  setModelMenuOpen: Dispatch<SetStateAction<boolean>>;
  showModelGuide: boolean;
  hasConfiguredModel: boolean;
  busy: boolean;
  petVisible: boolean;
  moreMenuOpen: boolean;
  setMoreMenuOpen: Dispatch<SetStateAction<boolean>>;
  setHistoryOpen: Dispatch<SetStateAction<boolean>>;
  setToolsOpen: Dispatch<SetStateAction<boolean>>;
  focusOpening: boolean;
  selectModel: (modelId: string) => void;
  onShowBaseModeHelp: () => void;
  openConnectionCreator: () => void;
  openSettings: () => void;
  dismissModelGuide: () => void;
  newChat: () => void;
  openTools: () => void;
  togglePetVisibility: () => void;
  openHistory: () => void;
  setIsRuleManagerOpen: Dispatch<SetStateAction<boolean>>;
  closeSettings: () => void;
  openFocusWindow: () => void;
}

export default function ChatHeader({
  serverOnline,
  serviceHealth,
  selectedModelId,
  selectedModel,
  modelOptions,
  modelMenuOpen,
  setModelMenuOpen,
  showModelGuide,
  hasConfiguredModel,
  busy,
  petVisible,
  moreMenuOpen,
  setMoreMenuOpen,
  setHistoryOpen,
  setToolsOpen,
  focusOpening,
  selectModel,
  onShowBaseModeHelp,
  openConnectionCreator,
  openSettings,
  dismissModelGuide,
  newChat,
  openTools,
  togglePetVisibility,
  openHistory,
  setIsRuleManagerOpen,
  closeSettings,
  openFocusWindow
}: ChatHeaderProps) {
  const activeModelId = selectedModelId || serviceHealth?.model || "local";
  const isBaseMode = activeModelId === "local";
  const statusClassName = !serverOnline
    ? ""
    : isBaseMode
      ? "online base-mode"
      : serviceHealth?.configured
        ? "online model-online"
        : "online";
  const statusText = !serverOnline
    ? "本地服务未连接"
    : isBaseMode
      ? "本地服务已连接 · 基础模式"
      : `模型：${selectedModel?.label ?? serviceHealth?.model}`;

  return (
    <header className="chat-header">
      <div className="brand-mark">EB</div>
      <div className="brand-copy">
        <strong>Excel Bro</strong>
        <span className={statusClassName}>
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
          {statusText}
          {serverOnline && isBaseMode && (
            <button
              type="button"
              className="base-mode-help-button"
              title="查看基础模式帮助"
              aria-label="查看基础模式帮助"
              onClick={onShowBaseModeHelp}
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
                onClick={() => selectModel(option.id)}
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
                onClick={() => void openConnectionCreator()}
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
          onClick={newChat}
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
  );
}
