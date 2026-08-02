import {
  assertTestModelConnectionResponse,
  assertIntentCheckResponse,
  assertAssistantResponse,
  assertDataToolResult,
  assertModelSettings,
  parseTurnStepEvent,
  type AssistantResponse,
  type TurnStepEvent,
  type AnalysisPlan,
  type DataToolRequest,
  type DataToolResult,
  type FolderCatalog,
  type FolderExecuteResult,
  type FolderSelection,
  type IntentCheckRequest,
  type IntentCheckResponse,
  type ModelSettings,
  type PlanRequest,
  type TestModelConnectionResponse,
  type UpdateModelSettingsRequest,
  type UpsertModelConnectionRequest,
  type WorkbookSnapshot
} from "./contracts";
import { recordModelCall } from "./diagnostics";

const API_BASE_URL =
  import.meta.env.VITE_API_BASE_URL ?? "http://127.0.0.1:8765";

export interface ServiceHealth {
  status: "ok";
  configured: boolean;
  mode: "model" | "local";
  model: string | null;
}

export interface ModelOption {
  id: string;
  label: string;
  provider: "model" | "local";
  available: boolean;
  supportsVision: boolean;
}

export interface ModelCatalog {
  defaultModelId: string;
  models: ModelOption[];
}

export class ApiRequestError extends Error {
  constructor(
    public readonly status: number,
    public readonly code: string,
    message: string,
    public readonly retryable: boolean
  ) {
    super(message);
    this.name = "ApiRequestError";
  }
}

export function isLocalServiceConnectionError(reason: unknown): boolean {
  if (reason instanceof ApiRequestError) return false;
  return (
    reason instanceof TypeError &&
    /fetch|network|connection|load failed/i.test(reason.message)
  );
}

export function apiErrorCategory(
  reason: unknown
): "local_service" | "model" | "service" {
  if (isLocalServiceConnectionError(reason)) return "local_service";
  if (
    reason instanceof ApiRequestError &&
    (reason.code.startsWith("MODEL_") || reason.status === 502)
  ) {
    return "model";
  }
  return "service";
}

interface ErrorEnvelope {
  detail?: {
    code?: unknown;
    message?: unknown;
    retryable?: unknown;
  } | string;
}

async function requestError(
  response: Response,
  fallbackLabel: string
): Promise<ApiRequestError> {
  const raw = await response.text();
  let code = "SERVICE_ERROR";
  let message = raw || fallbackLabel;
  let retryable = response.status >= 500;

  try {
    const envelope = JSON.parse(raw) as ErrorEnvelope;
    if (typeof envelope.detail === "string") {
      message = envelope.detail;
    } else if (envelope.detail && typeof envelope.detail === "object") {
      if (typeof envelope.detail.code === "string") {
        code = envelope.detail.code;
      }
      if (typeof envelope.detail.message === "string") {
        message = envelope.detail.message;
      }
      if (typeof envelope.detail.retryable === "boolean") {
        retryable = envelope.detail.retryable;
      }
    }
  } catch {
    // 非 JSON 响应保留服务端原文，便于诊断代理或网关错误。
  }

  return new ApiRequestError(response.status, code, message, retryable);
}

export async function createAssistantResponse(
  request: PlanRequest,
  signal?: AbortSignal
): Promise<AssistantResponse> {
  recordModelCall();
  const response = await fetch(`${API_BASE_URL}/api/turn`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request),
    signal
  });

  if (!response.ok) {
    throw await requestError(response, "本地 AI 服务请求失败");
  }

  const value: unknown = await response.json();
  assertAssistantResponse(value);
  return value;
}

interface StreamHooks {
  onStep?: (event: TurnStepEvent) => void;
  // signal：打断/转向时掐掉流式请求。已吐出的 step 由调用方保留做上下文。
  signal?: AbortSignal;
}

function parseSseFrame(frame: string): { event: string; data: string } | null {
  let event = "message";
  const dataLines: string[] = [];
  for (const line of frame.split("\n")) {
    if (line.startsWith("event:")) {
      event = line.slice(6).trim();
    } else if (line.startsWith("data:")) {
      dataLines.push(line.slice(5).replace(/^ /, ""));
    }
  }
  if (dataLines.length === 0) {
    return null;
  }
  return { event, data: dataLines.join("\n") };
}

export async function streamAssistantResponse(
  request: PlanRequest,
  hooks: StreamHooks = {}
): Promise<AssistantResponse> {
  recordModelCall();

  const signal = hooks.signal;
  let response: Response;
  try {
    response = await fetch(`${API_BASE_URL}/api/turn/stream`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(request),
      signal
    });
  } catch (reason) {
    // 网络/连接错误发生在收到任何事件之前：回退到非流式端点。
    if (isLocalServiceConnectionError(reason)) {
      return createAssistantResponse(request, signal);
    }
    throw reason;
  }

  if (!response.ok || !response.body) {
    // 旧后端无此端点（404）或代理不支持流式：回退到非流式端点。
    if (response.status === 404 || response.status === 405 || !response.body) {
      return createAssistantResponse(request, signal);
    }
    throw await requestError(response, "本地 AI 服务流式请求失败");
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let result: AssistantResponse | null = null;

  const handleFrame = (frame: string): void => {
    const parsed = parseSseFrame(frame);
    if (!parsed) {
      return;
    }
    const payload: unknown = JSON.parse(parsed.data);
    if (parsed.event === "step") {
      const step = parseTurnStepEvent(payload);
      if (step && hooks.onStep) {
        hooks.onStep(step);
      }
      return;
    }
    if (parsed.event === "error") {
      const detail = (payload ?? {}) as {
        status?: unknown;
        code?: unknown;
        message?: unknown;
        retryable?: unknown;
      };
      throw new ApiRequestError(
        typeof detail.status === "number" ? detail.status : 502,
        typeof detail.code === "string" ? detail.code : "SERVICE_ERROR",
        typeof detail.message === "string" ? detail.message : "本地 AI 服务请求失败",
        typeof detail.retryable === "boolean" ? detail.retryable : false
      );
    }
    if (parsed.event === "result") {
      assertAssistantResponse(payload);
      result = payload;
    }
  };

  for (;;) {
    const { done, value } = await reader.read();
    if (value) {
      buffer += decoder.decode(value, { stream: true });
    }
    let separator = buffer.indexOf("\n\n");
    while (separator !== -1) {
      const frame = buffer.slice(0, separator);
      buffer = buffer.slice(separator + 2);
      if (frame.trim()) {
        handleFrame(frame);
      }
      separator = buffer.indexOf("\n\n");
    }
    if (done) {
      break;
    }
  }

  const tail = buffer.trim();
  if (tail) {
    handleFrame(tail);
  }

  if (!result) {
    throw new ApiRequestError(
      502,
      "SERVICE_ERROR",
      "本地 AI 服务未返回完整结果",
      true
    );
  }
  return result;
}

export async function checkIntent(
  request: IntentCheckRequest,
  signal?: AbortSignal
): Promise<IntentCheckResponse> {
  recordModelCall();
  const response = await fetch(`${API_BASE_URL}/api/turn`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request),
    signal
  });

  if (!response.ok) {
    throw await requestError(response, "需求确认服务请求失败");
  }

  const value: unknown = await response.json();
  assertIntentCheckResponse(value);
  return value;
}

export async function checkHealth(): Promise<ServiceHealth | null> {
  try {
    const response = await fetch(`${API_BASE_URL}/health`);
    if (!response.ok) return null;
    const value = (await response.json()) as Partial<ServiceHealth>;
    if (
      value.status !== "ok" ||
      typeof value.configured !== "boolean" ||
      (value.mode !== "model" && value.mode !== "local")
    ) {
      return null;
    }
    return {
      status: "ok",
      configured: value.configured,
      mode: value.mode,
      model: typeof value.model === "string" ? value.model : null
    };
  } catch {
    return null;
  }
}

async function responseJson<T>(response: Response): Promise<T> {
  if (!response.ok) {
    throw await requestError(response, "本地服务请求失败");
  }
  return (await response.json()) as T;
}

export async function listModels(): Promise<ModelCatalog> {
  const response = await fetch(`${API_BASE_URL}/api/models`);
  return responseJson<ModelCatalog>(response);
}

export async function getModelSettings(): Promise<ModelSettings> {
  const response = await fetch(`${API_BASE_URL}/api/settings/model`);
  const value: unknown = await responseJson(response);
  assertModelSettings(value);
  return value;
}

export async function updateModelSettings(
  request: UpdateModelSettingsRequest
): Promise<ModelSettings> {
  const response = await fetch(`${API_BASE_URL}/api/settings/model`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request)
  });
  const value: unknown = await responseJson(response);
  assertModelSettings(value);
  return value;
}

export async function saveModelConnection(
  request: UpsertModelConnectionRequest
): Promise<ModelSettings> {
  const response = await fetch(
    `${API_BASE_URL}/api/settings/model/connections`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(request)
    }
  );
  const value: unknown = await responseJson(response);
  assertModelSettings(value);
  return value;
}

export async function testModelConnection(
  request: UpsertModelConnectionRequest
): Promise<TestModelConnectionResponse> {
  const response = await fetch(
    `${API_BASE_URL}/api/settings/model/connections/test`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(request)
    }
  );
  const value: unknown = await responseJson(response);
  assertTestModelConnectionResponse(value);
  return value;
}

export async function deleteModelConnection(
  connectionId: string
): Promise<ModelSettings> {
  const response = await fetch(
    `${API_BASE_URL}/api/settings/model/connections/${encodeURIComponent(connectionId)}`,
    { method: "DELETE" }
  );
  const value: unknown = await responseJson(response);
  assertModelSettings(value);
  return value;
}

export async function selectFolder(): Promise<FolderCatalog | null> {
  const response = await fetch(`${API_BASE_URL}/api/folders/select`, {
    method: "POST"
  });
  return responseJson<FolderCatalog | null>(response);
}

export async function createFolderSnapshot(
  sessionId: string,
  selections: FolderSelection[]
): Promise<WorkbookSnapshot> {
  const response = await fetch(`${API_BASE_URL}/api/folders/snapshot`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ sessionId, selections })
  });
  return responseJson<WorkbookSnapshot>(response);
}

export async function executeFolderPlan(
  sessionId: string,
  plan: AnalysisPlan
): Promise<FolderExecuteResult> {
  const response = await fetch(`${API_BASE_URL}/api/folders/execute`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ sessionId, plan })
  });
  return responseJson<FolderExecuteResult>(response);
}

export async function executeFolderQuery(
  sessionId: string,
  request: DataToolRequest
): Promise<DataToolResult> {
  const response = await fetch(`${API_BASE_URL}/api/folders/query`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ sessionId, request })
  });
  if (!response.ok) {
    throw await requestError(response, "文件夹数据工具请求失败");
  }
  const value: unknown = await response.json();
  assertDataToolResult(value);
  return value;
}
