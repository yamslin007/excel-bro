// @vitest-environment jsdom
import { describe, it, expect, vi, beforeEach } from "vitest";
import { renderHook, act } from "@testing-library/react";
import type { ModelSettings } from "../contracts";
import {
  deleteModelConnection,
  getModelSettings,
  saveModelConnection,
  setFormulaModel,
  testModelConnection,
  updateModelSettings
} from "../api";
import { emptyModelConnectionDraft } from "../utils";
import { MODEL_STORAGE_KEY } from "../conversation";
import { chooseAvailableModel } from "../modelSelection";
import {
  useModelManagement,
  type ModelConnectionDraft
} from "./useModelManagement";

vi.mock("../api", () => ({
  getModelSettings: vi.fn(),
  updateModelSettings: vi.fn(),
  saveModelConnection: vi.fn(),
  deleteModelConnection: vi.fn(),
  testModelConnection: vi.fn(),
  setFormulaModel: vi.fn(),
  checkHealth: vi.fn(),
  listModels: vi.fn()
}));

vi.mock("../utils", () => ({
  emptyModelConnectionDraft: vi.fn()
}));

vi.mock("../conversation", () => ({
  MODEL_STORAGE_KEY: "excel-bro.model.v2"
}));

vi.mock("../modelSelection", () => ({
  chooseAvailableModel: vi.fn()
}));

const getModelSettingsMock = vi.mocked(getModelSettings);
const updateModelSettingsMock = vi.mocked(updateModelSettings);
const saveModelConnectionMock = vi.mocked(saveModelConnection);
const deleteModelConnectionMock = vi.mocked(deleteModelConnection);
const testModelConnectionMock = vi.mocked(testModelConnection);
const setFormulaModelMock = vi.mocked(setFormulaModel);
const emptyModelConnectionDraftMock = vi.mocked(emptyModelConnectionDraft);
const chooseAvailableModelMock = vi.mocked(chooseAvailableModel);

const modelSettings: ModelSettings = {
  baseUrl: "https://api.example.com",
  defaultModel: "gpt-4.1",
  apiKeyConfigured: true,
  apiKeyHint: "sk-***",
  formulaModelId: "",
  connections: [
    {
      id: "conn-1",
      catalogModelId: "gpt-4.1",
      label: "GPT 4.1",
      baseUrl: "https://api.example.com",
      modelId: "gpt-4.1",
      supportsVision: true,
      apiKeyConfigured: true,
      apiKeyHint: "sk-***"
    }
  ]
};

const emptyDraft: ModelConnectionDraft = {
  id: null,
  label: "",
  baseUrl: "",
  modelId: "",
  apiKey: "",
  clearApiKey: false,
  supportsVision: false
};

function renderModelManagement() {
  const refreshServiceHealth = vi.fn().mockResolvedValue({
    modelOptions: [
      {
        id: "gpt-4.1",
        label: "GPT 4.1",
        provider: "model",
        available: true,
        supportsVision: true
      }
    ],
    defaultModelId: "gpt-4.1"
  });

  const view = renderHook(() =>
    useModelManagement({ refreshServiceHealth })
  );

  return { ...view, refreshServiceHealth };
}

describe("useModelManagement", () => {
  beforeEach(() => {
    localStorage.clear();
    vi.clearAllMocks();
    getModelSettingsMock.mockResolvedValue(modelSettings);
    updateModelSettingsMock.mockResolvedValue(modelSettings);
    saveModelConnectionMock.mockResolvedValue(modelSettings);
    deleteModelConnectionMock.mockResolvedValue(modelSettings);
    testModelConnectionMock.mockResolvedValue({
      ok: true,
      message: "连接成功"
    });
    setFormulaModelMock.mockResolvedValue(modelSettings);
    emptyModelConnectionDraftMock.mockReturnValue(emptyDraft);
    chooseAvailableModelMock.mockReturnValue("gpt-4.1");
  });

  it("should initialize selected model from localStorage", () => {
    localStorage.setItem(MODEL_STORAGE_KEY, "gpt-4.1");
    const { result } = renderModelManagement();

    expect(result.current.selectedModelId).toBe("gpt-4.1");
    expect(result.current.modelGuideDismissed).toBe(false);
  });

  it("should select a model and persist it", () => {
    const { result } = renderModelManagement();

    act(() => result.current.selectModel("gpt-4.1"));

    expect(result.current.selectedModelId).toBe("gpt-4.1");
    expect(localStorage.getItem(MODEL_STORAGE_KEY)).toBe("gpt-4.1");
  });

  it("should dismiss the model guide", () => {
    const { result } = renderModelManagement();

    act(() => result.current.dismissModelGuide());

    expect(result.current.modelGuideDismissed).toBe(true);
  });

  it("should load model settings", async () => {
    const { result } = renderModelManagement();

    await act(async () => {
      await result.current.openSettings();
    });

    expect(result.current.modelSettings).toEqual(modelSettings);
    expect(result.current.settingsLoading).toBe(false);
  });

  it("should open connection creator after loading settings", async () => {
    const { result } = renderModelManagement();

    await act(async () => {
      await result.current.openConnectionCreator();
    });

    expect(result.current.connectionDraft).toEqual(emptyDraft);
  });

  it("should populate draft when editing an existing connection", async () => {
    const { result } = renderModelManagement();

    await act(async () => {
      await result.current.openSettings();
    });

    act(() => result.current.editModelConnection("conn-1"));

    expect(result.current.connectionDraft).toMatchObject({
      id: "conn-1",
      label: "GPT 4.1",
      baseUrl: "https://api.example.com",
      modelId: "gpt-4.1",
      supportsVision: true,
      apiKey: ""
    });
  });

  it("should test a connection", async () => {
    const { result } = renderModelManagement();

    act(() =>
      result.current.setConnectionDraft({
        id: null,
        label: "New",
        baseUrl: "https://api.example.com",
        modelId: "gpt-4.1",
        apiKey: "sk-test",
        clearApiKey: false,
        supportsVision: true
      })
    );

    await act(async () => {
      await result.current.verifyConnection();
    });

    expect(testModelConnectionMock).toHaveBeenCalled();
    expect(result.current.settingsFeedback).toBe("连接成功");
  });

  it("should save an API key and refresh models", async () => {
    const { result, refreshServiceHealth } = renderModelManagement();

    act(() => result.current.setApiKeyDraft("sk-new"));
    await act(async () => {
      await result.current.saveApiKey();
    });

    expect(updateModelSettingsMock).toHaveBeenCalledWith({ apiKey: "sk-new" });
    expect(refreshServiceHealth).toHaveBeenCalled();
    expect(result.current.apiKeyDraft).toBe("");
    expect(result.current.settingsFeedback).toContain("API Key 已保存");
  });

  it("should save a new model connection", async () => {
    const { result } = renderModelManagement();

    act(() =>
      result.current.setConnectionDraft({
        id: null,
        label: "New",
        baseUrl: "https://api.example.com",
        modelId: "gpt-4.1",
        apiKey: "sk-test",
        clearApiKey: false,
        supportsVision: true
      })
    );

    await act(async () => {
      await result.current.saveConnection();
    });

    expect(saveModelConnectionMock).toHaveBeenCalled();
    expect(result.current.connectionDraft).toBeNull();
    expect(result.current.settingsFeedback).toContain("模型连接已添加");
  });

  it("should remove a model connection", async () => {
    const { result } = renderModelManagement();

    await act(async () => {
      await result.current.removeConnection("conn-1");
    });

    expect(deleteModelConnectionMock).toHaveBeenCalledWith("conn-1");
    expect(result.current.settingsFeedback).toBe("模型连接已删除。");
  });

  it("should save formula model setting", async () => {
    const { result } = renderModelManagement();

    await act(async () => {
      await result.current.saveFormulaModel("gpt-4.1");
    });

    expect(setFormulaModelMock).toHaveBeenCalledWith({ modelId: "gpt-4.1" });
    expect(result.current.settingsFeedback).toContain(
      "/function 公式模型已更新"
    );
  });
});
