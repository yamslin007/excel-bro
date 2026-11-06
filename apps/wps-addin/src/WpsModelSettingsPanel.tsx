import { ModelConnectionProtocolFields } from "../../excel-addin/src/components/ModelConnectionProtocolFields";
import type { ModelOption } from "../../excel-addin/src/api";
import type { useModelManagement } from "../../excel-addin/src/hooks/useModelManagement";
import { agentModelAvailable, agentModelUnavailableReason, formulaModelAvailable, connectionModelSuffix } from "../../excel-addin/src/modelSelection";

type ModelManagement = ReturnType<typeof useModelManagement>;

const SUCCESSFUL_SETTINGS_FEEDBACK = new Set([
  "API Key 已保存并立即生效。",
  "模型连接已添加。",
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
  models: readonly ModelOption[];
  currentModelId: string;
  management: ModelManagement;
  onClose: () => void;
  onAddConnection: () => void;
};

export default function WpsModelSettingsPanel({
  open,
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
    existingConnectionsStatus,
    connectionTestFeedback,
    connectionVerified,
    saveApiKey,
    editModelConnection,
    verifyConnection,
    saveConnection,
    checkExistingConnections,
    selectModel,
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
  const leaveEditor = () => {
    setShowApiKey(false);
    setConnectionDraft(null);
  };
  const connectionEditor = connectionDraft ? (
    <div className="model-connection-editor">
      <div className="model-editor-title">
        <strong>{connectionDraft.id
          ? `编辑 ${editingConnection?.label ?? connectionDraft.label}`
          : "添加模型连接"}</strong>
        <button
          type="button"
          disabled={settingsSaving || settingsTesting}
          onClick={leaveEditor}
        >取消</button>
      </div>
      {!connectionDraft.id && (existingConnectionsStatus === "checking" ||
        existingConnectionsStatus === "ready" ||
        existingConnectionsStatus === "failed") && (
        <div className="model-connection-discovery" role="status">
          {existingConnectionsStatus === "checking" && <span>正在检查这台电脑已保存的模型连接…</span>}
          {existingConnectionsStatus === "ready" && (
            <span>{modelSettings?.connections.length
              ? `已找到 ${modelSettings.connections.length} 个已保存的模型连接；Excel 与 WPS 共用这些配置。`
              : "已检查：这台电脑上暂无已保存的模型连接。"}</span>
          )}
          {existingConnectionsStatus === "failed" && (
            <>
              <span>暂时无法检查已有连接；这不代表没有配置。可继续填写并测试，新连接会独立保存。</span>
              <button type="button" onClick={() => void checkExistingConnections()}>重试检查</button>
            </>
          )}
        </div>
      )}
      <ModelConnectionProtocolFields draft={connectionDraft} onChange={setConnectionDraft} />
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
          placeholder={connectionDraft.protocol === "anthropic" ? "供应商提供的 Anthropic 兼容地址" : "https://api.example.com/v1"}
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
        {connectionDraft.id && (
          <span className="model-saved-key-status">
            <strong>{editingConnection?.apiKeyConfigured
              ? `已保存 ${editingConnection.apiKeyHint ?? "（密钥已配置）"}`
              : "尚未配置密钥"}</strong>
            <small>{connectionDraft.clearApiKey
              ? "保存后将清除现有密钥。"
              : editingConnection?.apiKeyConfigured
                ? "为安全起见，现有密钥不回显；留空会保留原密钥，输入新 Key 才会替换。"
                : "如服务需要鉴权，请在下方输入 API Key。"}</small>
          </span>
        )}
        <div className={`model-key-row${!connectionDraft.apiKey || connectionDraft.clearApiKey ? " no-key-reveal" : ""}`}>
          <input
            type={showApiKey ? "text" : "password"}
            value={connectionDraft.apiKey}
            disabled={connectionDraft.clearApiKey}
            autoComplete="off"
            spellCheck={false}
            placeholder={connectionDraft.id
              ? editingConnection?.apiKeyConfigured ? "输入新 Key 以替换；留空保持不变" : "输入 API Key"
              : "无鉴权服务可留空"}
            onChange={(event) => setConnectionDraft((draft) => draft ? {
              ...draft,
              apiKey: event.target.value,
              clearApiKey: false
            } : draft)}
          />
          {connectionDraft.apiKey && !connectionDraft.clearApiKey && (
            <button
              type="button"
              onClick={() => setShowApiKey((visible) => !visible)}
            >{showApiKey ? "隐藏新 Key" : "显示新 Key"}</button>
          )}
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
        <button
          type="button"
          disabled={settingsSaving || settingsTesting}
          onClick={() => void verifyConnection()}
        >{settingsTesting ? "正在测试…" : "测试连接"}</button>
        <button
          className="model-primary-action"
          type="button"
          disabled={settingsSaving || settingsTesting || !connectionVerified}
          title={!connectionVerified ? "请先测试连接并确认成功" : undefined}
          onClick={() => void saveConnection()}
        >{settingsSaving ? "正在保存…" : connectionDraft.id ? "保存修改" : "添加连接"}</button>
      </div>
      {connectionTestFeedback && (
        <div className={`model-connection-feedback ${connectionTestFeedback.kind}`} role={connectionTestFeedback.kind === "error" ? "alert" : "status"}>
          <b aria-hidden="true">{connectionTestFeedback.kind === "success" ? "✓" : "!"}</b>
          <span>{connectionTestFeedback.message}</span>
        </div>
      )}
    </div>
  ) : null;

  return (
    <>
      <aside className="wps-model-settings tool-drawer settings-drawer" aria-label="模型设置">
        <header className="wps-model-settings-header tool-drawer-header">
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

              {!hasEnvironmentModel && !connectionDraft && (
                <section
                  className="settings-first-run"
                  role="button"
                  tabIndex={0}
                  onClick={() => {
                    onAddConnection();
                  }}
                  onKeyDown={(event) => {
                    if (event.key === "Enter" || event.key === " ") {
                      event.preventDefault();
                      onAddConnection();
                    }
                  }}
                >
                  <div className="settings-first-run-icon">＋</div>
                  <div>
                    <strong>添加模型连接</strong>
                    <span>填写服务地址、模型 ID 和自己的 API Key，即可添加一个模型连接。</span>
                  </div>
                </section>
              )}

              <section className="model-settings-section">
                <div className="model-settings-section-heading">
                  <div>
                    <strong>{hasEnvironmentModel ? "其他模型连接" : "模型连接"}</strong>
                    <span>点击模型即可切换；编辑会在所选连接下方展开。</span>
                  </div>
                  {hasEnvironmentModel && (
                    <button type="button" disabled={settingsSaving || settingsTesting || Boolean(connectionDraft)} onClick={onAddConnection}>
                      ＋ 添加
                    </button>
                  )}
                </div>

                {connectionDraft && !connectionDraft.id && connectionEditor}
                {(modelSettings?.connections.length ?? 0) === 0 && !connectionDraft && (
                  <div className="model-connection-empty">
                    还没有模型连接。可以添加 OpenAI 或 Anthropic 兼容的模型连接。
                  </div>
                )}
                {modelSettings?.connections.some((connection) =>
                  !models.some((model) => model.id === connection.catalogModelId && model.available)
                ) && (
                  <div className="model-connection-discovery" role="status">
                    <span>部分连接暂不可切换，请检查后台或重新读取模型目录。</span>
                    <button
                      type="button"
                      disabled={existingConnectionsStatus === "checking"}
                      onClick={() => void checkExistingConnections()}
                    >{existingConnectionsStatus === "checking" ? "正在读取…" : "重试读取"}</button>
                  </div>
                )}
                {(modelSettings?.connections.length ?? 0) > 0 && (
                  <div className="model-connection-list">
                    {(modelSettings?.connections ?? []).map((connection) => {
                      const isCurrent = currentModelId === connection.catalogModelId;
                      const catalogModel = models.find((model) => model.id === connection.catalogModelId);
                      const canSelect = agentModelAvailable(catalogModel);
                      return (
                        <article
                          className={`model-connection-card${isCurrent ? " current" : ""}`}
                          key={connection.id}
                        >
                          <button
                            className="model-connection-select"
                            type="button"
                            aria-label={`${isCurrent ? "当前使用模型" : "切换为模型"} ${connection.label} · ${connection.modelId}${connectionModelSuffix(modelSettings?.connections ?? [], connection.id)}`}
                            disabled={isCurrent || !canSelect || settingsSaving || settingsTesting || Boolean(connectionDraft)}
                            title={!canSelect && catalogModel ? agentModelUnavailableReason(catalogModel) : undefined}
                            onClick={() => {
                              selectModel(connection.catalogModelId);
                            }}
                          >
                            <span className="model-connection-copy">
                              <span className="model-connection-name">
                                <strong>{connection.label}</strong>
                                {isCurrent && <b>当前使用</b>}
                              </span>
                              <span>{connection.modelId}{connectionModelSuffix(modelSettings?.connections ?? [], connection.id)}</span>
                              <small>{connection.baseUrl}</small>
                              <small>
                                Key：{connection.apiKeyConfigured ? `已配置 ${connection.apiKeyHint ?? ""}` : "未配置"}
                                {connection.supportsVision ? " · 支持图片" : ""}
                              </small>
                              {!isCurrent && <em>{canSelect ? "点击切换使用" : catalogModel ? agentModelUnavailableReason(catalogModel) : "暂不可用，请检查连接"}</em>}
                            </span>
                          </button>
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
                              <button type="button" disabled={settingsSaving || settingsTesting || Boolean(connectionDraft)} onClick={() => editModelConnection(connection.id)}>
                                编辑
                              </button>
                              <button className="danger" type="button" disabled={settingsSaving || settingsTesting || Boolean(connectionDraft)} onClick={() => setPendingDeleteConnectionId(connection.id)}>
                                删除
                              </button>
                            </div>
                          )}
                          {connectionDraft?.id === connection.id && connectionEditor}
                        </article>
                      );
                    })}
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
                      <option key={model.id} value={model.id} disabled={!formulaModelAvailable(model)}>{model.label}</option>
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
        </div>
      </aside>
    </>
  );
}
