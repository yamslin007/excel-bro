import { useState, useEffect, useCallback, useRef } from "react";
import { checkHealth, listModels, type ModelOption } from "../api";
import type { ModelCatalogStatus, ServiceHealth } from "../api";

export interface ModelCatalogRefreshResult {
  modelOptions: ModelOption[];
  defaultModelId: string | null;
}

/**
 * 服务健康检查 Hook
 *
 * 职责：
 * - 管理服务在线状态和健康信息
 * - 定期轮询服务健康状态（5 秒间隔）
 * - 窗口获得焦点时刷新状态
 * - 加载模型目录，并区分「加载中 / 加载失败 / 成功但为空 / 已有真实连接」
 *
 * 旧版本这里会把一个虚拟的“基础模式”项作为初始模型列表，导致目录尚未加载
 * 或读取失败时看起来“已经选好模型”。现在初始列表为空，目录状态单独表达。
 */
export function useServiceHealth() {
  const [serverOnline, setServerOnline] = useState(false);
  const [serviceHealth, setServiceHealth] = useState<ServiceHealth | null>(null);
  const [modelOptions, setModelOptions] = useState<ModelOption[]>([]);
  const [modelCatalogStatus, setModelCatalogStatus] =
    useState<ModelCatalogStatus>("idle");
  const [modelCatalogError, setModelCatalogError] = useState("");
  const [defaultCatalogModelId, setDefaultCatalogModelId] = useState<
    string | null
  >(null);
  const hasLoadedCatalogRef = useRef(false);

  /**
   * 读取模型目录。
   *
   * 返回 null 表示本次读取失败：调用方必须保留原有选择，不能把失败
   * 当成“用户没有配置模型”。
   */
  const loadModelCatalog = useCallback(async () => {
    if (!hasLoadedCatalogRef.current) {
      setModelCatalogStatus("loading");
    }
    try {
      const catalog = await listModels();
      hasLoadedCatalogRef.current = true;
      setModelOptions(catalog.models);
      setDefaultCatalogModelId(catalog.defaultModelId);
      setModelCatalogStatus("ready");
      setModelCatalogError("");
      return {
        modelOptions: catalog.models,
        defaultModelId: catalog.defaultModelId
      } satisfies ModelCatalogRefreshResult;
    } catch (reason) {
      setModelCatalogStatus("failed");
      setModelCatalogError(
        reason instanceof Error
          ? reason.message
          : "无法读取模型配置，请确认 Excel Bro 后台正在运行。"
      );
      return null;
    }
  }, []);

  /**
   * 刷新服务状态和模型目录
   */
  const refreshServiceState = useCallback(async () => {
    const health = await checkHealth();
    setServiceHealth(health);
    setServerOnline(health !== null);

    if (!health) return null;

    return await loadModelCatalog();
  }, [loadModelCatalog]);

  /**
   * 重试模型目录（供“加载失败 → 重试”入口使用）。
   * 不清空当前选择，也不改变已就绪的列表直到新结果返回。
   */
  const retryModelCatalog = useCallback(
    () => loadModelCatalog(),
    [loadModelCatalog]
  );

  /**
   * 手动设置服务在线状态
   * 用于其他组件在成功调用 API 后更新状态
   */
  const markServerOnline = useCallback(() => {
    setServerOnline(true);
  }, []);

  /**
   * 手动设置服务离线状态
   * 用于其他组件在 API 调用失败后更新状态
   */
  const markServerOffline = useCallback(() => {
    setServerOnline(false);
    setServiceHealth(null);
  }, []);

  // 定期轮询服务健康状态
  useEffect(() => {
    let active = true;

    async function poll() {
      const health = await checkHealth();
      if (!active) return;
      setServiceHealth(health);
      setServerOnline(health !== null);

      if (!health) return;
      await loadModelCatalog();
    }

    // 初始检查
    void poll();

    // 每 5 秒轮询一次
    const intervalId = window.setInterval(() => {
      void poll();
    }, 5000);

    // 窗口获得焦点时刷新
    const handleFocus = () => void poll();
    window.addEventListener("focus", handleFocus);

    return () => {
      active = false;
      window.clearInterval(intervalId);
      window.removeEventListener("focus", handleFocus);
    };
  }, [loadModelCatalog]);

  return {
    // 状态
    serverOnline,
    serviceHealth,
    modelOptions,
    modelCatalogStatus,
    modelCatalogError,
    defaultCatalogModelId,

    // 操作
    refreshServiceState,
    retryModelCatalog,
    markServerOnline,
    markServerOffline,

    // 内部状态设置（给特殊场景使用）
    setServerOnline,
    setServiceHealth,
    setModelOptions,
    setModelCatalogStatus,
  };
}
