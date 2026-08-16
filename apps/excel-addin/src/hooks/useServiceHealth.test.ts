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

describe("useServiceHealth", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.clearAllMocks();
    checkHealthMock.mockResolvedValue(null);
    listModelsMock.mockResolvedValue({
      defaultModelId: "local",
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

  it("should initialize with offline state and local-only model", async () => {
    const { result } = renderHook(() => useServiceHealth());
    await flushInitialPoll();

    expect(result.current.serverOnline).toBe(false);
    expect(result.current.serviceHealth).toBeNull();
    expect(result.current.modelOptions).toHaveLength(1);
    expect(result.current.modelOptions[0].id).toBe("local");
    expect(result.current.modelCatalogLoaded).toBe(false);
  });

  it("should refresh service state and model catalog", async () => {
    checkHealthMock.mockResolvedValue({
      status: "ok",
      configured: true,
      mode: "model",
      model: model.id
    });
    const { result } = renderHook(() => useServiceHealth());

    let refreshResult: unknown;
    await act(async () => {
      refreshResult = await result.current.refreshServiceState();
    });

    expect(result.current.serverOnline).toBe(true);
    expect(result.current.serviceHealth).toMatchObject({
      status: "ok",
      configured: true,
      mode: "model"
    });
    expect(result.current.modelCatalogLoaded).toBe(true);
    expect(result.current.modelOptions).toEqual([model]);
    expect(refreshResult).toEqual({
      modelOptions: [model],
      defaultModelId: "local"
    });
  });

  it("should return null when health is unavailable", async () => {
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
      mode: "local",
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
    checkHealthMock.mockResolvedValue({
      status: "ok",
      configured: true,
      mode: "model",
      model: model.id
    });
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
