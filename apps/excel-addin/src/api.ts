import {
  assertIntentCheckResponse,
  assertAssistantResponse,
  type AssistantResponse,
  type AnalysisPlan,
  type FolderCatalog,
  type FolderExecuteResult,
  type FolderSelection,
  type IntentCheckRequest,
  type IntentCheckResponse,
  type PlanRequest,
  type WorkbookSnapshot
} from "./contracts";

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
  request: PlanRequest
): Promise<AssistantResponse> {
  const response = await fetch(`${API_BASE_URL}/api/turn`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request)
  });

  if (!response.ok) {
    throw await requestError(response, "本地 AI 服务请求失败");
  }

  const value: unknown = await response.json();
  assertAssistantResponse(value);
  return value;
}

export async function checkIntent(
  request: IntentCheckRequest
): Promise<IntentCheckResponse> {
  const response = await fetch(`${API_BASE_URL}/api/turn`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request)
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
