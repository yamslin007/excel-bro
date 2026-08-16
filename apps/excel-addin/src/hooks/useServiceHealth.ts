import { useState, useEffect, useCallback } from "react";
import { checkHealth, listModels, type ModelOption } from "../api";
import type { ServiceHealth } from "../api";

/**
 * 服务健康检查 Hook
 *
 * 职责：
 * - 管理服务在线状态和健康信息
 * - 定期轮询服务健康状态（5 秒间隔）
 * - 窗口获得焦点时刷新状态
 * - 加载模型目录
 */
export function useServiceHealth() {
  const [serverOnline, setServerOnline] = useState(false);
  const [serviceHealth, setServiceHealth] = useState<ServiceHealth | null>(null);
  const [modelOptions, setModelOptions] = useState<ModelOption[]>([
    {
      id: "local",
      label: "基础模式",
      provider: "local",
      available: true,
      supportsVision: false
    }
  ]);
  const [modelCatalogLoaded, setModelCatalogLoaded] = useState(false);

  /**
   * 刷新服务状态和模型目录
   */
  const refreshServiceState = useCallback(async () => {
    const health = await checkHealth();
    setServiceHealth(health);
    setServerOnline(health !== null);

    if (!health) return null;

    try {
      const catalog = await listModels();
      setModelOptions(catalog.models);
      setModelCatalogLoaded(true);
      return {
        modelOptions: catalog.models,
        defaultModelId: catalog.defaultModelId
      };
    } catch {
      // 如果旧版本服务缺少 /api/models 端点，健康指示器仍然有用
      return null;
    }
  }, []);

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

      try {
        const catalog = await listModels();
        if (!active) return;
        setModelOptions(catalog.models);
        setModelCatalogLoaded(true);
      } catch {
        // The health indicator remains useful if an older service lacks /api/models.
      }
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
  }, []);

  return {
    // 状态
    serverOnline,
    serviceHealth,
    modelOptions,
    modelCatalogLoaded,

    // 操作
    refreshServiceState,
    markServerOnline,
    markServerOffline,

    // 内部状态设置（给特殊场景使用）
    setServerOnline,
    setServiceHealth,
    setModelOptions,
    setModelCatalogLoaded,
  };
}
