// @vitest-environment jsdom
import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { renderHook, act } from "@testing-library/react";
import { checkHealth, listModels, type ModelOption } from "../api";
import { useServiceHealth } from "./useServiceHealth";

vi.mock("../api", () => ({
  checkHealth: vi.fn(),
  listModels: vi.fn()
}));

const checkHealthMock = vi.mocked(checkHealth);
const listModelsMock = vi.mocked(listModels);

const model: ModelOption = {
  id: "gpt-4.1",
  label: "GPT-4.1",
  provider: "model",
  available: true,
  supportsVision: true
};

const online = {
  status: "ok" as const,
  configured: true,
  mode: "model" as const,
  model: model.id
};

describe("useServiceHealth", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.clearAllMocks();
    checkHealthMock.mockResolvedValue(null);
    listModelsMock.mockResolvedValue({
      defaultModelId: model.id,
      models: [model]
    });
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  async function flushInitialPoll() {
    await act(async () => {
      await Promise.resolve();
    });
  }

  it("starts with an empty model list instead of a virtual base-mode option", async () => {
    const { result } = renderHook(() => useServiceHealth());
    await flushInitialPoll();

    expect(result.current.serverOnline).toBe(false);
    expect(result.current.serviceHealth).toBeNull();
    // 目录尚未加载时列表为空；这里以前会注入一个 id 为 local 的“基础模式”。
    expect(result.current.modelOptions).toEqual([]);
    expect(result.current.modelCatalogStatus).toBe("idle");
  });

  it("reports the catalog as ready and exposes real connections", async () => {
    checkHealthMock.mockResolvedValue(online);
    const { result } = renderHook(() => useServiceHealth());

    let refreshResult: unknown;
    await act(async () => {
      refreshResult = await result.current.refreshServiceState();
    });

    expect(result.current.serverOnline).toBe(true);
    expect(result.current.modelCatalogStatus).toBe("ready");
    expect(result.current.modelOptions).toEqual([model]);
    expect(result.current.modelCatalogError).toBe("");
    expect(refreshResult).toEqual({
      modelOptions: [model],
      defaultModelId: model.id
    });
  });

  it("treats a successful empty catalog as configured-but-empty, not as failure", async () => {
    checkHealthMock.mockResolvedValue({
      status: "ok",
      configured: false,
      mode: "model",
      model: null
    });
    listModelsMock.mockResolvedValue({ defaultModelId: null, models: [] });

    const { result } = renderHook(() => useServiceHealth());
    let refreshResult: unknown;
    await act(async () => {
      refreshResult = await result.current.refreshServiceState();
    });

    expect(result.current.modelCatalogStatus).toBe("ready");
    expect(result.current.modelOptions).toEqual([]);
    expect(result.current.modelCatalogError).toBe("");
    expect(refreshResult).toEqual({
      modelOptions: [],
      defaultModelId: null
    });
  });

  it("keeps the previous catalog and reports failure when the listing request fails", async () => {
    checkHealthMock.mockResolvedValue(online);
    const { result } = renderHook(() => useServiceHealth());
    await act(async () => {
      await result.current.refreshServiceState();
    });
    expect(result.current.modelOptions).toEqual([model]);

    // 目录读取失败：不得清空列表，也不得断言用户没有连接。
    listModelsMock.mockRejectedValue(new Error("目录请求失败"));
    let refreshResult: unknown;
    await act(async () => {
      refreshResult = await result.current.refreshServiceState();
    });

    expect(refreshResult).toBeNull();
    expect(result.current.modelCatalogStatus).toBe("failed");
    expect(result.current.modelCatalogError).toContain("目录请求失败");
    expect(result.current.modelOptions).toEqual([model]);
  });

  it("recovers from a failed catalog through retryModelCatalog", async () => {
    checkHealthMock.mockResolvedValue(online);
    const { result } = renderHook(() => useServiceHealth());

    // 让本次读取失败（挂载时的初始 poll 已经消费掉默认实现）。
    listModelsMock.mockRejectedValue(new Error("后台未就绪"));
    await act(async () => {
      await result.current.refreshServiceState();
    });
    expect(result.current.modelCatalogStatus).toBe("failed");
    expect(result.current.modelOptions).toEqual([]);
    expect(result.current.modelCatalogError).toContain("后台未就绪");

    listModelsMock.mockResolvedValue({
      defaultModelId: model.id,
      models: [model]
    });
    await act(async () => {
      await result.current.retryModelCatalog();
    });

    expect(result.current.modelCatalogStatus).toBe("ready");
    expect(result.current.modelOptions).toEqual([model]);
    expect(result.current.modelCatalogError).toBe("");
  });

  it("returns null when health is unavailable", async () => {
    checkHealthMock.mockResolvedValue(null);
    const { result } = renderHook(() => useServiceHealth());

    let refreshResult: unknown;
    await act(async () => {
      refreshResult = await result.current.refreshServiceState();
    });

    expect(refreshResult).toBeNull();
    expect(result.current.serverOnline).toBe(false);
    expect(result.current.serviceHealth).toBeNull();
  });

  it("should mark server online and offline manually", async () => {
    const { result } = renderHook(() => useServiceHealth());
    await flushInitialPoll();

    act(() => result.current.markServerOnline());
    expect(result.current.serverOnline).toBe(true);

    act(() => result.current.markServerOffline());
    expect(result.current.serverOnline).toBe(false);
    expect(result.current.serviceHealth).toBeNull();
  });

  it("should poll health on an interval", async () => {
    checkHealthMock.mockResolvedValue({
      status: "ok",
      configured: false,
      mode: "model",
      model: null
    });
    const { result } = renderHook(() => useServiceHealth());
    await flushInitialPoll();

    const initialCalls = checkHealthMock.mock.calls.length;

    await act(async () => {
      vi.advanceTimersByTime(5000);
      await Promise.resolve();
    });

    expect(checkHealthMock.mock.calls.length).toBeGreaterThan(initialCalls);
    expect(result.current.serverOnline).toBe(true);
  });

  it("should refresh when window gains focus", async () => {
    checkHealthMock.mockResolvedValue(online);
    const { result } = renderHook(() => useServiceHealth());
    await flushInitialPoll();

    const before = checkHealthMock.mock.calls.length;

    await act(async () => {
      window.dispatchEvent(new Event("focus"));
      await Promise.resolve();
    });

    expect(checkHealthMock.mock.calls.length).toBeGreaterThan(before);
    expect(result.current.serverOnline).toBe(true);
  });
});
