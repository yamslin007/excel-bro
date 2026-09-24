import type { ModelOption } from "../../excel-addin/src/api";
import type { useModelManagement } from "../../excel-addin/src/hooks/useModelManagement";

type ModelManagement = ReturnType<typeof useModelManagement>;

const SUCCESSFUL_SETTINGS_FEEDBACK = new Set([
  "API Key 已保存并立即生效。",
  "模型连接已添加，并已切换为当前模型。",
  "模型连接已更新。",
  "模型连接已删除。",
  "/function 公式模型已更新。",
  "/function 公式模型已恢复为跟随全局选择。"
]);

function isSuccessfulSettingsFeedback(message: string): boolean {
  return SUCCESSFUL_SETTINGS_FEEDBACK.has(message);
}

type WpsModelSettingsPanelProps = {
  open: boolean;
  backendAvailable: boolean;
  models: readonly ModelOption[];
  currentModelId: string;
  management: ModelManagement;
  onClose: () => void;
  onAddConnection: () => void;
};

export default function WpsModelSettingsPanel({
  open,
  backendAvailable,
  models,
  currentModelId,
  management,
  onClose,
  onAddConnection
}: WpsModelSettingsPanelProps) {
  const {
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
    setSettingsFeedback
  } = management;

  if (!open) return null;

  const hasEnvironmentModel =
    Boolean(modelSettings?.baseUrl) && Boolean(modelSettings?.defaultModel);
  const editingConnection = connectionDraft?.id
    ? modelSettings?.connections.find((item) => item.id === connectionDraft.id)
    : null;

  return (
    <>
      <button
        className="model-settings-backdrop"
        type="button"
        aria-label="关闭模型设置"
        onClick={onClose}
      />
      <aside className="wps-model-settings" aria-label="模型设置">
        <header className="wps-model-settings-header">
          <div>
            <strong>模型设置</strong>
            <span>API Key 安全保存在这台电脑上</span>
          </div>
          <button type="button" onClick={onClose} aria-label="关闭模型设置">×</button>
        </header>

        <div className="wps-model-settings-body">
          {settingsLoading ? (
            <div className="model-settings-loading"><i />正在读取模型连接…</div>
          ) : (
            <>
              {hasEnvironmentModel && (
                <section className="model-settings-summary">
                  <div><span>服务地址</span><strong>{modelSettings?.baseUrl}</strong></div>
                  <div><span>默认模型</span><strong>{modelSettings?.defaultModel}</strong></div>
                  <div>
                    <span>API Key</span>
                    <strong>
                      {modelSettings?.apiKeyConfigured
                        ? `已配置 ${modelSettings.apiKeyHint ?? ""}`
                        : "未配置（API 服务无需鉴权时可留空）"}
                    </strong>
                  </div>
                </section>
              )}

              {hasEnvironmentModel && (
                <section className="model-settings-section api-key-settings">
                  <div className="model-settings-section-heading">
                    <div>
                      <strong>更换默认连接的 API Key</strong>
                      <span>完整密钥不会从后端读回，也不会保存到聊天记录或浏览器存储。</span>
                    </div>
                  </div>
                  <label>
                    <span>新的 API Key</span>
                    <div className="model-key-row">
                      <input
                        type={showApiKey ? "text" : "password"}
                        value={apiKeyDraft}
                        disabled={settingsSaving}
                        autoComplete="off"
                        spellCheck={false}
                        placeholder={modelSettings?.apiKeyConfigured ? "留空不会覆盖现有密钥" : "粘贴 API Key"}
                        onChange={(event) => {
                          setApiKeyDraft(event.target.value);
                          setSettingsFeedback("");
                        }}
                      />
                      <button
                        type="button"
                        disabled={!apiKeyDraft || settingsSaving}
                        onClick={() => setShowApiKey((visible) => !visible)}
                      >
                        {showApiKey ? "隐藏" : "显示"}
                      </button>
                    </div>
                  </label>
                  <button
                    className="model-primary-action"
                    type="button"
                    disabled={!apiKeyDraft.trim() || settingsSaving}
                    onClick={() => void saveApiKey()}
                  >
                    {settingsSaving ? "正在保存…" : "保存并立即生效"}
                  </button>
                </section>
              )}

              <section className="model-settings-section">
                <div className="model-settings-section-heading">
                  <div>
                    <strong>{hasEnvironmentModel ? "其他模型连接" : "模型连接"}</strong>
                    <span>每个连接使用独立的服务地址、模型 ID 和 API Key。</span>
                  </div>
                  {!connectionDraft && (
                    <button type="button" disabled={settingsSaving} onClick={onAddConnection}>
                      ＋ 添加
                    </button>
                  )}
                </div>

                {(modelSettings?.connections.length ?? 0) === 0 && !connectionDraft ? (
                  <div className="model-connection-empty">
                    还没有模型连接。可添加 DeepSeek、OpenAI 或其他兼容 OpenAI /chat/completions 的服务。
                  </div>
                ) : (
                  <div className="model-connection-list">
                    {(modelSettings?.connections ?? []).map((connection) => (
                      <article
                        className={`model-connection-card${currentModelId === connection.catalogModelId ? " current" : ""}`}
                        key={connection.id}
                      >
                        <div className="model-connection-copy">
                          <div>
                            <strong>{connection.label}</strong>
                            {currentModelId === connection.catalogModelId && <b>当前使用</b>}
                          </div>
                          <span>{connection.modelId}</span>
                          <small>{connection.baseUrl}</small>
                          <small>
                            Key：{connection.apiKeyConfigured ? `已配置 ${connection.apiKeyHint ?? ""}` : "未配置"}
                            {connection.supportsVision ? " · 支持图片" : ""}
                          </small>
                        </div>
                        {pendingDeleteConnectionId === connection.id ? (
                          <div className="model-connection-actions">
                            <button type="button" onClick={() => setPendingDeleteConnectionId(null)}>取消</button>
                            <button
                              className="danger"
                              type="button"
                              disabled={settingsSaving || settingsTesting}
                              onClick={() => void removeConnection(connection.id)}
                            >
                              确认删除
                            </button>
                          </div>
                        ) : (
                          <div className="model-connection-actions">
                            <button type="button" disabled={settingsSaving || settingsTesting} onClick={() => editModelConnection(connection.id)}>
                              编辑
                            </button>
                            <button className="danger" type="button" disabled={settingsSaving} onClick={() => setPendingDeleteConnectionId(connection.id)}>
                              删除
                            </button>
                          </div>
                        )}
                      </article>
                    ))}
                  </div>
                )}

                {connectionDraft && (
                  <div className="model-connection-editor">
                    <div className="model-editor-title">
                      <strong>{connectionDraft.id ? "编辑模型连接" : "添加模型连接"}</strong>
                      <button type="button" aria-label="关闭连接编辑" onClick={() => {
                        setShowApiKey(false);
                        setConnectionDraft(null);
                      }}>×</button>
                    </div>
                    <p>兼容 OpenAI /chat/completions 的服务均可接入。连接测试成功后才能保存。</p>
                    <label>
                      <span>连接名称</span>
                      <input
                        value={connectionDraft.label}
                        placeholder="例如 DeepSeek"
                        onChange={(event) => setConnectionDraft((draft) => draft ? { ...draft, label: event.target.value } : draft)}
                      />
                    </label>
                    <label>
                      <span>服务地址</span>
                      <input
                        value={connectionDraft.baseUrl}
                        placeholder="https://api.example.com/v1"
                        onChange={(event) => setConnectionDraft((draft) => draft ? { ...draft, baseUrl: event.target.value } : draft)}
                      />
                    </label>
                    <label>
                      <span>模型 ID</span>
                      <input
                        value={connectionDraft.modelId}
                        placeholder="供应商实际接受的模型 ID"
                        onChange={(event) => setConnectionDraft((draft) => draft ? { ...draft, modelId: event.target.value } : draft)}
                      />
                    </label>
                    <label>
                      <span>API Key</span>
                      <div className="model-key-row">
                        <input
                          type={showApiKey ? "text" : "password"}
                          value={connectionDraft.apiKey}
                          disabled={connectionDraft.clearApiKey}
                          autoComplete="off"
                          spellCheck={false}
                          placeholder={connectionDraft.id ? "留空保留现有密钥" : "本地无鉴权服务可留空"}
                          onChange={(event) => setConnectionDraft((draft) => draft ? {
                            ...draft,
                            apiKey: event.target.value,
                            clearApiKey: false
                          } : draft)}
                        />
                        <button
                          type="button"
                          disabled={!connectionDraft.apiKey || connectionDraft.clearApiKey}
                          onClick={() => setShowApiKey((visible) => !visible)}
                        >
                          {showApiKey ? "隐藏 Key" : "显示 Key"}
                        </button>
                      </div>
                    </label>
                    {connectionDraft.id && editingConnection?.apiKeyConfigured && (
                      <label className="model-checkbox-row">
                        <input
                          type="checkbox"
                          checked={connectionDraft.clearApiKey}
                          onChange={(event) => setConnectionDraft((draft) => draft ? {
                            ...draft,
                            apiKey: "",
                            clearApiKey: event.target.checked
                          } : draft)}
                        />
                        <span>清除现有 API Key（仅适用于无需鉴权的 API 服务）</span>
                      </label>
                    )}
                    <label className="model-checkbox-row">
                      <input
                        type="checkbox"
                        checked={connectionDraft.supportsVision}
                        onChange={(event) => setConnectionDraft((draft) => draft ? {
                          ...draft,
                          supportsVision: event.target.checked
                        } : draft)}
                      />
                      <span>这个模型支持图片输入</span>
                    </label>
                    <div className="model-editor-actions">
                      <button type="button" disabled={settingsSaving || settingsTesting} onClick={() => void verifyConnection()}>
                        {settingsTesting ? "正在测试…" : "测试连接"}
                      </button>
                      <button
                        className="model-primary-action"
                        type="button"
                        disabled={settingsSaving || settingsTesting || !connectionVerified}
                        title={!connectionVerified ? "请先测试连接并确认成功" : undefined}
                        onClick={() => void saveConnection()}
                      >
                        {settingsSaving ? "正在保存…" : connectionDraft.id ? "保存修改" : "添加并使用"}
                      </button>
                    </div>
                    {connectionTestFeedback && (
                      <div className={`model-connection-feedback ${connectionTestFeedback.kind}`} role={connectionTestFeedback.kind === "error" ? "alert" : "status"}>
                        <b aria-hidden="true">{connectionTestFeedback.kind === "success" ? "✓" : "!"}</b>
                        <span>{connectionTestFeedback.message}</span>
                      </div>
                    )}
                  </div>
                )}
              </section>

              <section className="model-settings-section">
                <div className="model-settings-section-heading">
                  <div>
                    <strong>/function 公式模型</strong>
                    <span>短链公式任务可使用独立模型，与聊天模型互不影响。</span>
                  </div>
                </div>
                <label>
                  <span>用于生成公式的模型</span>
                  <select
                    value={modelSettings?.formulaModelId ?? ""}
                    disabled={settingsSaving}
                    onChange={(event) => void saveFormulaModel(event.target.value)}
                  >
                    <option value="">跟随全局选择（默认）</option>
                    {models.map((model) => (
                      <option key={model.id} value={model.id} disabled={!model.available}>{model.label}</option>
                    ))}
                  </select>
                </label>
              </section>
            </>
          )}

          {settingsFeedback && (
            <div
              className={
                isSuccessfulSettingsFeedback(settingsFeedback)
                  ? "model-settings-feedback"
                  : "model-settings-feedback error"
              }
              role={
                isSuccessfulSettingsFeedback(settingsFeedback) ? "status" : "alert"
              }
            >
              {settingsFeedback}
            </div>
          )}
          {!backendAvailable && (
            <div className="model-settings-feedback error" role="alert">Excel Bro 后台暂时不可用；你仍可填写配置，测试或保存时会重新尝试连接。</div>
          )}
        </div>
      </aside>
    </>
  );
}
