import { useState, useCallback } from "react";
import type { ModelOption } from "../api";
import {
  getModelSettings,
  updateModelSettings,
  saveModelConnection,
  deleteModelConnection,
  testModelConnection,
  setFormulaModel,
  checkHealth,
  listModels
} from "../api";
import type {
  ModelSettings,
  UpsertModelConnectionRequest
} from "../contracts";
import { emptyModelConnectionDraft } from "../utils";
import { MODEL_STORAGE_KEY } from "../conversation";
import { chooseAvailableModel } from "../modelSelection";

export interface ModelConnectionDraft {
  id: string | null;
  label: string;
  baseUrl: string;
  modelId: string;
  apiKey: string;
  clearApiKey: boolean;
  supportsVision: boolean;
}

interface UseModelManagementOptions {
  /**
   * 刷新服务健康状态的回调
   * 由 useServiceHealth Hook 提供
   */
  refreshServiceHealth: () => Promise<{
    modelOptions: ModelOption[];
    defaultModelId?: string;
  } | null>;
}

/**
 * 模型管理 Hook
 *
 * 职责：
 * - 管理选中的模型 ID
 * - 管理模型设置（API Key、连接列表）
 * - 管理连接编辑状态
 * - 提供模型选择、连接 CRUD、测试等操作
 *
 * 依赖：
 * - 需要 useServiceHealth 的 refreshServiceHealth 函数来同步模型目录
 */
export function useModelManagement(options: UseModelManagementOptions) {
  const { refreshServiceHealth } = options;

  // 模型选择状态
  const [selectedModelId, setSelectedModelId] = useState(
    () => localStorage.getItem(MODEL_STORAGE_KEY) ?? ""
  );

  // 模型设置状态
  const [modelSettings, setModelSettings] = useState<ModelSettings | null>(null);
  const [apiKeyDraft, setApiKeyDraft] = useState("");
  const [showApiKey, setShowApiKey] = useState(false);

  // 连接编辑状态
  const [connectionDraft, setConnectionDraft] = useState<ModelConnectionDraft | null>(null);
  const [pendingDeleteConnectionId, setPendingDeleteConnectionId] = useState<string | null>(null);

  // UI 状态
  const [settingsSaving, setSettingsSaving] = useState(false);
  const [settingsTesting, setSettingsTesting] = useState(false);
  const [settingsLoading, setSettingsLoading] = useState(false);
  const [settingsFeedback, setSettingsFeedback] = useState("");
  const [modelGuideDismissed, setModelGuideDismissed] = useState(false);

  /**
   * 选择模型
   */
  const selectModel = useCallback((modelId: string) => {
    setSelectedModelId(modelId);
    localStorage.setItem(MODEL_STORAGE_KEY, modelId);
  }, []);

  /**
   * 关闭模型引导
   */
  const dismissModelGuide = useCallback(() => {
    setModelGuideDismissed(true);
  }, []);

  /**
   * 打开设置并加载模型设置
   */
  const openSettings = useCallback(async () => {
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
  }, []);

  /**
   * 打开连接创建器
   */
  const openConnectionCreator = useCallback(async () => {
    if (await openSettings()) {
      setConnectionDraft(emptyModelConnectionDraft());
    }
  }, [openSettings]);

  /**
   * 刷新模型列表（在保存设置后）
   * 调用 useServiceHealth 的 refreshServiceHealth 来更新模型目录
   */
  const refreshModelsAfterSettings = useCallback(async (
    saved: ModelSettings,
    preferredModelId?: string
  ) => {
    setModelSettings(saved);

    // 刷新服务健康状态和模型目录
    const catalog = await refreshServiceHealth();
    if (!catalog) return;

    setSelectedModelId((current) => {
      const next = chooseAvailableModel(
        catalog.modelOptions,
        current,
        catalog.defaultModelId ?? "local",
        preferredModelId
      );
      localStorage.setItem(MODEL_STORAGE_KEY, next);
      return next;
    });
  }, [refreshServiceHealth]);

  /**
   * 保存 API Key
   */
  const saveApiKey = useCallback(async () => {
    const apiKey = apiKeyDraft.trim();
    if (!apiKey) {
      setSettingsFeedback("请输入新的 API Key。");
      return;
    }
    setSettingsSaving(true);
    setSettingsFeedback("");
    try {
      const saved = await updateModelSettings({ apiKey });
      setApiKeyDraft("");
      setShowApiKey(false);
      setSettingsFeedback("API Key 已保存并立即生效。");
      await refreshModelsAfterSettings(saved);
    } catch (reason) {
      setSettingsFeedback(
        reason instanceof Error ? reason.message : "API Key 保存失败。"
      );
    } finally {
      setSettingsSaving(false);
    }
  }, [apiKeyDraft, refreshModelsAfterSettings]);

  /**
   * 编辑模型连接
   */
  const editModelConnection = useCallback((connectionId: string) => {
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
  }, [modelSettings]);

  /**
   * 构建连接请求对象
   */
  const connectionRequest = useCallback((): UpsertModelConnectionRequest | null => {
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
  }, [connectionDraft]);

  /**
   * 测试连接
   */
  const verifyConnection = useCallback(async () => {
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
  }, [connectionRequest]);

  /**
   * 保存连接
   */
  const saveConnection = useCallback(async () => {
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
  }, [connectionRequest, connectionDraft, modelSettings, refreshModelsAfterSettings]);

  /**
   * 删除连接
   */
  const removeConnection = useCallback(async (connectionId: string) => {
    setSettingsSaving(true);
    setSettingsFeedback("");
    try {
      const saved = await deleteModelConnection(connectionId);
      await refreshModelsAfterSettings(saved);
      setPendingDeleteConnectionId(null);
      setConnectionDraft((current: ModelConnectionDraft | null) =>
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
  }, [refreshModelsAfterSettings]);

  /**
   * 保存公式模型设置
   */
  const saveFormulaModel = useCallback(async (modelId: string) => {
    setSettingsSaving(true);
    setSettingsFeedback("");
    try {
      const saved = await setFormulaModel({ modelId });
      await refreshModelsAfterSettings(saved);
      setSettingsFeedback(
        modelId
          ? "/function 公式模型已更新。"
          : "/function 公式模型已恢复为跟随全局选择。"
      );
    } catch (reason) {
      setSettingsFeedback(
        reason instanceof Error ? reason.message : "公式模型保存失败。"
      );
    } finally {
      setSettingsSaving(false);
    }
  }, [refreshModelsAfterSettings]);

  return {
    // 状态
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

    // 操作
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

    // 状态设置（给特殊场景使用）
    setSelectedModelId,
    setModelSettings,
    setApiKeyDraft,
    setShowApiKey,
    setConnectionDraft,
    setPendingDeleteConnectionId,
    setSettingsSaving,
    setSettingsTesting,
    setSettingsLoading,
    setSettingsFeedback,
    setModelGuideDismissed,
  };
}
