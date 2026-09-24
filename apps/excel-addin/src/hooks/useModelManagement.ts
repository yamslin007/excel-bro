import { useState, useCallback, useRef } from "react";
import type { Dispatch, SetStateAction } from "react";
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
import {
  chooseAvailableModel,
  shouldPersistSelection
} from "../modelSelection";

export interface ModelConnectionDraft {
  id: string | null;
  label: string;
  baseUrl: string;
  modelId: string;
  apiKey: string;
  clearApiKey: boolean;
  supportsVision: boolean;
}

export interface ConnectionTestFeedback {
  kind: "success" | "error";
  message: string;
}

interface UseModelManagementOptions {
  /**
   * 刷新服务健康状态的回调
   * 由 useServiceHealth Hook 提供
   */
  refreshServiceHealth: () => Promise<{
    modelOptions: ModelOption[];
    defaultModelId?: string | null;
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
  const [connectionDraft, setConnectionDraftState] = useState<ModelConnectionDraft | null>(null);
  const [pendingDeleteConnectionId, setPendingDeleteConnectionId] = useState<string | null>(null);
  const [connectionTestFeedback, setConnectionTestFeedback] =
    useState<ConnectionTestFeedback | null>(null);
  const connectionDraftVersionRef = useRef(0);
  const verifiedDraftVersionRef = useRef<number | null>(null);

  // UI 状态
  const [settingsSaving, setSettingsSaving] = useState(false);
  const [settingsTesting, setSettingsTesting] = useState(false);
  const [settingsLoading, setSettingsLoading] = useState(false);
  const [settingsFeedback, setSettingsFeedback] = useState("");
  const [modelGuideDismissed, setModelGuideDismissed] = useState(false);

  const setConnectionDraft = useCallback<
    Dispatch<SetStateAction<ModelConnectionDraft | null>>
  >((nextDraft) => {
    connectionDraftVersionRef.current += 1;
    verifiedDraftVersionRef.current = null;
    setConnectionTestFeedback(null);
    setConnectionDraftState(nextDraft);
  }, []);

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
   *
   * 读取失败不返回 false 也不再阻止调用方：配置表单必须能在后台不可达时
   * 打开，让用户先填写。失败信息写在 settingsFeedback 里，测试与保存会
   * 发起真实请求并以本次结果更新状态。
   */
  const openSettings = useCallback(async () => {
    setShowApiKey(false);
    setSettingsFeedback("");
    setConnectionDraft(null);
    setPendingDeleteConnectionId(null);
    setSettingsLoading(true);
    try {
      setModelSettings(await getModelSettings());
      // 读取设置成功已经证明后台可达；同时刷新健康状态/模型目录，
      // 清除页面初始化阶段可能留下的“后台不可用”旧状态。
      try {
        await refreshServiceHealth();
      } catch {
        // 设置本身已经读取成功，目录刷新失败不应把设置页重新判成加载失败。
      }
      return true;
    } catch (reason) {
      setModelSettings(null);
      setSettingsFeedback(
        reason instanceof Error
          ? `暂时无法连接 Excel Bro 后台：${reason.message}。已填写内容仍保留在当前表单中，请恢复连接后重试。`
          : "暂时无法连接 Excel Bro 后台。已填写内容仍保留在当前表单中，请恢复连接后重试。"
      );
      return false;
    } finally {
      setSettingsLoading(false);
    }
  }, [refreshServiceHealth]);

  /**
   * 打开连接创建器
   *
   * 不等待后台读取成功：即使 openSettings 失败也进入表单。用户填写完
   * 再点“测试连接”或“添加并使用”时才发起真实请求。
   */
  const openConnectionCreator = useCallback(async () => {
    await openSettings();
    setConnectionDraft(emptyModelConnectionDraft());
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
    // 目录读取失败时必须保留原选择，不能依据空目录删掉仍然有效的记录。
    if (!catalog) return;

    setSelectedModelId((current) => {
      const next = chooseAvailableModel(
        catalog.modelOptions,
        current,
        catalog.defaultModelId,
        preferredModelId
      );
      if (shouldPersistSelection(current, next)) {
        localStorage.setItem(MODEL_STORAGE_KEY, next);
      }
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
    if (!request) {
      setConnectionTestFeedback({
        kind: "error",
        message: "请先填写连接名称、服务地址和模型 ID。"
      });
      return;
    }
    const draftVersion = connectionDraftVersionRef.current;
    verifiedDraftVersionRef.current = null;
    setSettingsTesting(true);
    setSettingsFeedback("");
    setConnectionTestFeedback(null);
    try {
      const result = await testModelConnection(request);
      if (connectionDraftVersionRef.current !== draftVersion) return;
      setConnectionTestFeedback({
        kind: result.ok ? "success" : "error",
        message: result.ok ? "连接成功，当前模型可用。" : result.message
      });
      if (result.ok) {
        verifiedDraftVersionRef.current = draftVersion;
        // 连接测试成功后同步刷新目录状态。否则 WPS 首次预检失败留下的
        // 红色“后台不可用”提示会与成功结果同时存在。
        try {
          await refreshServiceHealth();
        } catch {
          // 模型测试结果有效；辅助刷新失败不应覆盖真实测试结果。
        }
      }
    } catch (reason) {
      if (connectionDraftVersionRef.current !== draftVersion) return;
      verifiedDraftVersionRef.current = null;
      setConnectionTestFeedback({
        kind: "error",
        message:
          reason instanceof Error ? reason.message : "模型连接测试失败。"
      });
    } finally {
      setSettingsTesting(false);
    }
  }, [connectionRequest, refreshServiceHealth]);

  /**
   * 保存连接
   */
  const saveConnection = useCallback(async () => {
    const request = connectionRequest();
    if (!request || !connectionDraft) {
      setSettingsFeedback("请填写连接名称、服务地址和模型 ID。");
      return;
    }
    if (verifiedDraftVersionRef.current !== connectionDraftVersionRef.current) {
      setConnectionTestFeedback({
        kind: "error",
        message: "请先测试连接并确认成功，再添加或保存。"
      });
      return;
    }
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

  /**
   * 按最新目录迁移已保存的选择。
   *
   * 目录**加载成功后**才调用：此时旧的 `local` 记录或指向已删除连接的选择
   * 会被迁移到当前有效模型；目录为空则回到未配置状态（空选择）。
   * 目录加载失败时调用方不得调用本函数。
   */
  const migrateModelSelection = useCallback(
    (catalog: { modelOptions: ModelOption[]; defaultModelId?: string | null }) => {
      setSelectedModelId((current) => {
        const next = chooseAvailableModel(
          catalog.modelOptions,
          current,
          catalog.defaultModelId
        );
        if (shouldPersistSelection(current, next)) {
          try {
            localStorage.setItem(MODEL_STORAGE_KEY, next);
          } catch {
            // 存储不可用时本次会话仍使用解析后的选择。
          }
        }
        return next;
      });
    },
    []
  );

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
    connectionTestFeedback,
    connectionVerified:
      connectionTestFeedback?.kind === "success" &&
      verifiedDraftVersionRef.current === connectionDraftVersionRef.current,
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
    migrateModelSelection,

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
