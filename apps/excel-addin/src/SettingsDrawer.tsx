// 模型设置抽屉：默认连接 API Key、模型连接管理、/function 公式模型。
// 纯展示组件，JSX 从 App.tsx 逐字搬移；所有状态与回调都由 props 传入。
import type { Dispatch, SetStateAction } from "react";
import type { ModelOption } from "./api";
import type { ManagedModelConnection, ModelSettings } from "./contracts";
import type { ModelConnectionDraft } from "./hooks/useModelManagement";
import { emptyModelConnectionDraft } from "./utils";

interface SettingsDrawerProps {
  settingsOpen: boolean;
  settingsLoading: boolean;
  modelSettings: ModelSettings | null;
  hasEnvironmentModel: boolean;
  hasManagedModels: boolean;
  apiKeyDraft: string;
  setApiKeyDraft: Dispatch<SetStateAction<string>>;
  showApiKey: boolean;
  setShowApiKey: Dispatch<SetStateAction<boolean>>;
  settingsSaving: boolean;
  settingsTesting: boolean;
  settingsFeedback: string;
  setSettingsFeedback: Dispatch<SetStateAction<string>>;
  serverOnline: boolean;
  connectionDraft: ModelConnectionDraft | null;
  setConnectionDraft: Dispatch<SetStateAction<ModelConnectionDraft | null>>;
  pendingDeleteConnectionId: string | null;
  setPendingDeleteConnectionId: Dispatch<SetStateAction<string | null>>;
  selectedModelId: string;
  modelOptions: ModelOption[];
  editingConnection: ManagedModelConnection | null | undefined;
  exportDiagnosticReport: () => void;
  closeSettings: () => void;
  saveApiKey: () => void;
  removeConnection: (connectionId: string) => void;
  editModelConnection: (connectionId: string) => void;
  verifyConnection: () => void;
  saveConnection: () => void;
  saveFormulaModel: (modelId: string) => void;
}

export default function SettingsDrawer({
  settingsOpen,
  settingsLoading,
  modelSettings,
  hasEnvironmentModel,
  hasManagedModels,
  apiKeyDraft,
  setApiKeyDraft,
  showApiKey,
  setShowApiKey,
  settingsSaving,
  settingsTesting,
  settingsFeedback,
  setSettingsFeedback,
  serverOnline,
  connectionDraft,
  setConnectionDraft,
  pendingDeleteConnectionId,
  setPendingDeleteConnectionId,
  selectedModelId,
  modelOptions,
  editingConnection,
  exportDiagnosticReport,
  closeSettings,
  saveApiKey,
  removeConnection,
  editModelConnection,
  verifyConnection,
  saveConnection,
  saveFormulaModel
}: SettingsDrawerProps) {
  if (!settingsOpen) return null;
  return (
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
  );
}
