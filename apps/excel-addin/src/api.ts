import capabilities from "../../../config/capabilities.json";
import {
  assertTestModelConnectionResponse,
  assertIntentCheckResponse,
  assertAssistantResponse,
  assertPlanningPauseResult,
  assertDataToolResult,
  assertModelSettings,
  assertOfficeExecutionState,
  assertWpsExecutionState,
  assertWpsExecutionResult,
  assertWpsRecoveryDescriptor,
  assertWpsEvidence,
  parseTaskProgressEvent,
  parseTurnStepEvent,
  type AssistantResponse,
  type PlanningPauseResult,
  type PlanningStreamResult,
  type PlanningTaskStatusResult,
  type TaskProgressEvent,
  type TurnStepEvent,
  type AnalysisPlan,
  type DataToolRequest,
  type DataToolResult,
  type FolderCatalog,
  type FolderBackupCleanupResult,
  type FolderBackupDeleteResult,
  type FolderBackupExportResult,
  type FolderBackupOpenResult,
  type FolderBackupPinResult,
  type FolderBackupsResult,
  type FolderExecuteResult,
  type FolderExecutionStatusResult,
  type FolderManagedRecoveryResult,
  type ManagedFolderRecoveryDirectoriesResult,
  type FolderLegacyCollectResult,
  type FolderUndoResult,
  type FolderSelection,
  type IntentCheckRequest,
  type IntentCheckResponse,
  type ModelSettings,
  type OfficeExecutionState,
  type WpsExecutionState,
  type WpsRecoveryDescriptor,
  type PlanExecutionResult,
  type ExecutionUndoSnapshot,
  type UndoRangeSnapshot,
  type PlanRequest,
  type PlanValidateRequest,
  type TestModelConnectionResponse,
  type UpdateModelSettingsRequest,
  type UpsertModelConnectionRequest,
  type SetFormulaModelRequest,
  type WorkbookSnapshot,
  type GenerateFormulaRequest,
  type GenerateFormulaResponse
} from "./contracts";
import type { ExcelReviewLayoutPrepareRequest, ExcelReviewLayoutTicket, ExcelReviewLayoutResult, ExcelReviewPaneMeasurement } from "./contracts";
import { recordModelCall } from "./diagnostics";

export interface PlanChatRequest {
  plan: AnalysisPlan;
  message: string;
  modelId?: string;
}

export interface PlanChatResponse {
  explanation: string;
  modifiedPlan: AnalysisPlan;
}

const API_BASE_URL =
  import.meta.env.VITE_API_BASE_URL ?? "http://127.0.0.1:8765";

export class ExcelReviewLayoutVersionError extends Error {
  constructor() {
    super('本地窗口排列服务版本过旧，请重启 Python 服务、刷新加载项并重新生成副本。');
    this.name = 'ExcelReviewLayoutVersionError';
  }
}

async function reviewLayoutRequest(path: "prepare" | "arrange" | "measure", body: ExcelReviewLayoutPrepareRequest | Pick<ExcelReviewLayoutTicket, 'token'>): Promise<unknown> {
  const endpoint = new URL(API_BASE_URL);
  if (!['localhost', '127.0.0.1', '[::1]'].includes(endpoint.hostname) || !['http:', 'https:'].includes(endpoint.protocol)) {
    throw new Error('自动窗口排列仅能连接本机服务，请手动并排查看。');
  }
  const response = await fetch(`${API_BASE_URL}/api/office/review-layout/${path}`, {
    method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(body),
    signal: AbortSignal.timeout(capabilities.excelWindowLayout.requestTimeoutMs), redirect: 'error'
  });
  if (!response.ok) throw new Error('自动窗口排列未完成，请重启新版本地服务或手动并排查看。');
  return response.json();
}
export async function measureExcelReviewPane(request: ExcelReviewLayoutPrepareRequest): Promise<ExcelReviewPaneMeasurement> {
  const value = await reviewLayoutRequest('measure', {workbookPathHash: request.workbookPathHash,
    worksheetNameHash: request.worksheetNameHash, copySheetName: request.copySheetName});
  if (!value || typeof value !== 'object' || Object.keys(value).length !== 1 ||
    !('clientWidthCssPx' in value) || typeof value.clientWidthCssPx !== 'number' ||
    !Number.isFinite(value.clientWidthCssPx) || value.clientWidthCssPx <= 0 || value.clientWidthCssPx > 32768)
    throw new Error('审阅窗口宽度未核实，请手动调整。');
  return value as ExcelReviewPaneMeasurement;
}

export async function prepareExcelReviewLayout(request: ExcelReviewLayoutPrepareRequest): Promise<ExcelReviewLayoutTicket> {
  // Explicit projection, not a spread: extra runtime plan/value/path fields never leave the pane.
  const value = await reviewLayoutRequest('prepare', {workbookPathHash: request.workbookPathHash,
    worksheetNameHash: request.worksheetNameHash, copySheetName: request.copySheetName});
    if (value && typeof value === 'object' && (!('layoutVersion' in value) || value.layoutVersion !== 3)) throw new ExcelReviewLayoutVersionError();
  if (!value || typeof value !== 'object' || Object.keys(value).length !== 2 ||
      !('token' in value) || typeof value.token !== 'string' || !/^[A-Za-z0-9_-]{43}$/.test(value.token)) throw new Error('窗口排列凭据无效，请手动并排查看。');
  return value as ExcelReviewLayoutTicket;
}
export async function arrangeExcelReviewLayout(ticket: ExcelReviewLayoutTicket): Promise<ExcelReviewLayoutResult> {
  if (ticket.layoutVersion !== 3) throw new ExcelReviewLayoutVersionError();
  const value = await reviewLayoutRequest('arrange', {token: ticket.token});
  if (value && typeof value === 'object' && (!('layoutVersion' in value) || value.layoutVersion !== 3)) throw new ExcelReviewLayoutVersionError();
  if (!value || typeof value !== 'object' || Object.keys(value).length !== 5 ||
      !('status' in value) || value.status !== 'arranged' ||
      !('copySide' in value) || value.copySide !== 'left' ||
      !('originalSide' in value) || value.originalSide !== 'right' ||
      !('synchronousScrolling' in value) || typeof value.synchronousScrolling !== 'string' || !['enabled', 'different_instance', 'unavailable'].includes(value.synchronousScrolling)) throw new Error('窗口排列结果未核实，请手动并排查看。');
  return value as ExcelReviewLayoutResult;
}

export interface ServiceHealth {
  status: "ok";
  configured: boolean;
  mode: "model";
  model: string | null;
}

/**
 * 模型目录里的一个真实连接。
 *
 * 目录只包含用户实际配置的模型；未配置时 `models` 为空数组、
 * `defaultModelId` 为 `null`。早期版本会在目录里注入一个
 * `id: "local"` 的“基础模式”占位项，该入口已移除；此处仍保留
 * `"local"` 只用于识别并迁移旧的本地选择记录。
 */
export interface ModelOption {
  id: string;
  label: string;
  provider: "model";
  available: boolean;
  supportsVision: boolean;
}

export interface ModelCatalog {
  defaultModelId: string | null;
  models: ModelOption[];
}

/**
 * 模型目录的加载状态。
 * - `idle`：首次目录尚未加载
 * - `loading`：正在读取模型配置
 * - `ready`：读取成功（`models` 可能为空，代表确实没有配置模型）
 * - `failed`：读取失败；不得当作“没有模型”，也不得覆盖已选模型
 */
export type ModelCatalogStatus = "idle" | "loading" | "ready" | "failed";

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
  onProgress?: (event: TaskProgressEvent) => void;
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

async function consumePlanningStream(
  response: Response,
  turnId: string,
  hooks: StreamHooks
): Promise<PlanningStreamResult> {
  if (!response.body) {
    throw new ApiRequestError(502, "SERVICE_ERROR", "本地 AI 服务没有返回数据流", true);
  }
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let result: PlanningStreamResult | null = null;
  let legacySequence = 0;

  const handleFrame = (frame: string): void => {
    const parsed = parseSseFrame(frame);
    if (!parsed) return;
    const payload: unknown = JSON.parse(parsed.data);
    if (parsed.event === "step" || parsed.event === "progress") {
      legacySequence += 1;
      const progress = parseTaskProgressEvent(payload, turnId, legacySequence);
      if (progress && hooks.onProgress) hooks.onProgress(progress);
      const step = parseTurnStepEvent(payload);
      if (step && hooks.onStep) hooks.onStep(step);
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
    if (parsed.event === "paused") {
      assertPlanningPauseResult(payload);
      result = payload;
      return;
    }
    if (parsed.event === "result") {
      assertAssistantResponse(payload);
      result = payload;
    }
  };

  for (;;) {
    const { done, value } = await reader.read();
    if (value) buffer += decoder.decode(value, { stream: true });
    let separator = buffer.indexOf("\n\n");
    while (separator !== -1) {
      const frame = buffer.slice(0, separator);
      buffer = buffer.slice(separator + 2);
      if (frame.trim()) handleFrame(frame);
      separator = buffer.indexOf("\n\n");
    }
    if (done) break;
  }

  const tail = buffer.trim();
  if (tail) handleFrame(tail);
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

export async function streamAssistantResponse(
  request: PlanRequest,
  hooks: StreamHooks = {}
): Promise<PlanningStreamResult> {
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

  return consumePlanningStream(
    response,
    request.turnId ?? "legacy-turn",
    hooks
  );
}

export async function getPlanningTaskStatus(
  turnId: string,
  signal?: AbortSignal
): Promise<PlanningTaskStatusResult> {
  const response = await fetch(
    `${API_BASE_URL}/api/turn/${encodeURIComponent(turnId)}/status`,
    { signal }
  );
  if (!response.ok) {
    throw await requestError(response, "读取规划任务状态失败");
  }
  return (await response.json()) as PlanningTaskStatusResult;
}

export async function continuePlanningTask(
  request: {
    turnId: string;
    expectedRevision: number;
    requestId: string;
    afterSequence: number;
    modelId?: string;
    confirmModelSwitch?: true;
  },
  hooks: StreamHooks = {}
): Promise<PlanningStreamResult> {
  recordModelCall();
  const response = await fetch(`${API_BASE_URL}/api/turn/continue/stream`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ...request, confirmed: true }),
    signal: hooks.signal
  });
  if (!response.ok || !response.body) {
    throw await requestError(response, "继续规划任务失败");
  }
  return consumePlanningStream(response, request.turnId, hooks);
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
      value.mode !== "model"
    ) {
      return null;
    }
    return {
      status: "ok",
      configured: value.configured,
      mode: "model",
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

export async function setFormulaModel(
  request: SetFormulaModelRequest
): Promise<ModelSettings> {
  const response = await fetch(`${API_BASE_URL}/api/settings/model/formula`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request)
  });
  const value: unknown = await responseJson(response);
  assertModelSettings(value);
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

/**
 * 把手工编辑过的计划送去服务端校验，拿回**规范化之后**的计划。
 *
 * 验收条件由服务端的 AnalysisPlan 校验器重算 —— 前端绝不自己拼，否则计划内容
 * 和验收条件会各说各话，执行时才暴露。
 */
export async function validatePlan(
  plan: AnalysisPlan,
  authorizedSheets: string[],
  existingSheets: string[] = authorizedSheets
): Promise<AnalysisPlan> {
  const request: PlanValidateRequest = {
    plan,
    authorized_sheets: authorizedSheets,
    existing_sheets: existingSheets
  };
  const response = await fetch(`${API_BASE_URL}/api/plans/validate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request)
  });
  if (!response.ok) {
    throw await requestError(response, "计划校验失败");
  }
  return (await response.json()) as AnalysisPlan;
}

async function officeExecutionResponse(
  response: Response,
  fallback: string
): Promise<OfficeExecutionState> {
  if (!response.ok) throw await requestError(response, fallback);
  const payload: unknown = await response.json();
  assertOfficeExecutionState(payload);
  return payload;
}

export async function registerOfficeExecution(input: {
  operationId: string;
  plan: AnalysisPlan;
  authorizedSheetNames: string[];
  allowedExternalFiles: string[];
  workbookIdentity: string;
}): Promise<OfficeExecutionState> {
  const response = await fetch(`${API_BASE_URL}/api/office/executions/register`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ...input, confirmed: true })
  });
  return officeExecutionResponse(response, "无法登记 Office 写入现场");
}

export async function getOfficeExecutionStatus(
  operationId: string
): Promise<OfficeExecutionState> {
  const response = await fetch(
    `${API_BASE_URL}/api/office/executions/${encodeURIComponent(operationId)}`
  );
  return officeExecutionResponse(response, "无法读取 Office 写入现场");
}

export async function startOfficeAction(
  operationId: string,
  input: {
    expectedRevision: number;
    ownerToken: string;
    actionIndex: number;
    beforeSnapshot: UndoRangeSnapshot;
  }
): Promise<OfficeExecutionState> {
  const response = await fetch(
    `${API_BASE_URL}/api/office/executions/${encodeURIComponent(operationId)}/actions/start`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(input)
    }
  );
  return officeExecutionResponse(response, "无法登记 Office 动作写前状态");
}

export async function completeOfficeAction(
  operationId: string,
  input: {
    expectedRevision: number;
    ownerToken: string;
    actionIndex: number;
    checkpoint: ExecutionUndoSnapshot;
  }
): Promise<OfficeExecutionState> {
  const response = await fetch(
    `${API_BASE_URL}/api/office/executions/${encodeURIComponent(operationId)}/actions/complete`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(input)
    }
  );
  return officeExecutionResponse(response, "无法提交 Office 动作完成回执");
}

export async function suspendOfficeExecution(
  operationId: string,
  input: {
    expectedRevision: number;
    ownerToken: string;
    code: string;
    message: string;
  }
): Promise<OfficeExecutionState> {
  const response = await fetch(
    `${API_BASE_URL}/api/office/executions/${encodeURIComponent(operationId)}/suspend`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(input)
    }
  );
  return officeExecutionResponse(response, "无法挂起 Office 写入现场");
}

export async function claimOfficeExecutionResume(
  operationId: string,
  expectedRevision: number
): Promise<OfficeExecutionState> {
  const response = await fetch(
    `${API_BASE_URL}/api/office/executions/${encodeURIComponent(operationId)}/resume`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ expectedRevision, confirmed: true })
    }
  );
  return officeExecutionResponse(response, "无法领取 Office 断点任务");
}

export async function reconcileOfficeActionReceipt(
  operationId: string,
  input: {
    expectedRevision: number;
    actionIndex: number;
    outcome: "applied" | "not_applied";
    checkpoint?: ExecutionUndoSnapshot;
  }
): Promise<OfficeExecutionState> {
  const response = await fetch(
    `${API_BASE_URL}/api/office/executions/${encodeURIComponent(operationId)}/actions/reconcile`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ...input, confirmed: true })
    }
  );
  return officeExecutionResponse(response, "无法核对 Office 在途动作");
}

export async function finishOfficeExecution(
  operationId: string,
  input: {
    expectedRevision: number;
    ownerToken: string;
    result: PlanExecutionResult;
  }
): Promise<OfficeExecutionState> {
  const response = await fetch(
    `${API_BASE_URL}/api/office/executions/${encodeURIComponent(operationId)}/complete`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(input)
    }
  );
  return officeExecutionResponse(response, "无法提交 Office 任务完成回执");
}

async function wpsExecutionResponse(
  response: Response,
  fallback: string
): Promise<WpsExecutionState> {
  if (!response.ok) throw await requestError(response, fallback);
  const payload: unknown = await response.json();
  assertWpsExecutionState(payload);
  return payload;
}

/**
 * WPS 首次原表写入前，把**执行身份、授权与恢复描述**持久登记到服务端账本。
 *
 * 载荷刻意只带 `WpsRecoveryDescriptor`（动作类型 / 范围 / 形状 / 目标指纹），
 * **不带 `AnalysisPlan`** —— 后者含 `writeValues` 的目标值矩阵，一旦随登记请求
 * 出站就会落进服务端账本。这里只接受描述类型，是类型层面的出站闸：
 * 想传整份计划在编译期就过不去。
 */
export async function registerWpsExecution(input: {
  operationId: string;
  descriptor: WpsRecoveryDescriptor;
  authorizedSheetNames: string[];
  workbookIdentity: string;
  taskId: string;
}): Promise<WpsExecutionState> {
  assertWpsRecoveryDescriptor(input.descriptor);
  const response = await fetch(`${API_BASE_URL}/api/wps/executions/register`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ operationId: input.operationId, descriptor: input.descriptor,
      authorizedSheetNames: input.authorizedSheetNames, workbookIdentity: input.workbookIdentity,
      taskId: input.taskId, confirmed: true })
  });
  return wpsExecutionResponse(response, "无法登记 WPS 写入现场");
}

export async function getWpsExecutionStatus(
  operationId: string
): Promise<WpsExecutionState> {
  const response = await fetch(
    `${API_BASE_URL}/api/wps/executions/${encodeURIComponent(operationId)}`
  );
  return wpsExecutionResponse(response, "无法读取 WPS 写入现场");
}

export async function startWpsAction(
  operationId: string,
  input: {
    expectedRevision: number;
    ownerToken: string;
    actionIndex: number;
    beforeSnapshot: Record<string, unknown>;
  }
): Promise<WpsExecutionState> {
  assertWpsEvidence(input.beforeSnapshot);
  const response = await fetch(
    `${API_BASE_URL}/api/wps/executions/${encodeURIComponent(operationId)}/actions/start`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ expectedRevision: input.expectedRevision, ownerToken: input.ownerToken,
        actionIndex: input.actionIndex, beforeSnapshot: input.beforeSnapshot })
    }
  );
  return wpsExecutionResponse(response, "无法登记 WPS 动作写前状态");
}

export async function completeWpsAction(
  operationId: string,
  input: {
    expectedRevision: number;
    ownerToken: string;
    actionIndex: number;
    checkpoint: Record<string, unknown>;
  }
): Promise<WpsExecutionState> {
  assertWpsEvidence(input.checkpoint);
  const response = await fetch(
    `${API_BASE_URL}/api/wps/executions/${encodeURIComponent(operationId)}/actions/complete`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ expectedRevision: input.expectedRevision, ownerToken: input.ownerToken,
        actionIndex: input.actionIndex, checkpoint: input.checkpoint })
    }
  );
  return wpsExecutionResponse(response, "无法提交 WPS 动作完成回执");
}

export async function suspendWpsExecution(
  operationId: string,
  input: {
    expectedRevision: number;
    ownerToken: string;
    code: string;
    message: string;
  }
): Promise<WpsExecutionState> {
  const response = await fetch(
    `${API_BASE_URL}/api/wps/executions/${encodeURIComponent(operationId)}/suspend`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ expectedRevision: input.expectedRevision, ownerToken: input.ownerToken,
        code: "WPS_WRITE_INTERRUPTED", message: "WPS 写入中断，需要核对现场。" })
    }
  );
  return wpsExecutionResponse(response, "无法挂起 WPS 写入现场");
}

export async function claimWpsExecutionResume(
  operationId: string,
  expectedRevision: number
): Promise<WpsExecutionState> {
  const response = await fetch(
    `${API_BASE_URL}/api/wps/executions/${encodeURIComponent(operationId)}/resume`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ expectedRevision, confirmed: true })
    }
  );
  return wpsExecutionResponse(response, "无法领取 WPS 恢复处置权");
}

export async function reconcileWpsActionReceipt(
  operationId: string,
  input: {
    expectedRevision: number;
    actionIndex: number;
    outcome: "applied" | "not_applied";
    checkpoint?: Record<string, unknown>;
  }
): Promise<WpsExecutionState> {
  if (input.checkpoint) assertWpsEvidence(input.checkpoint);
  const response = await fetch(
    `${API_BASE_URL}/api/wps/executions/${encodeURIComponent(operationId)}/actions/reconcile`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ expectedRevision: input.expectedRevision, actionIndex: input.actionIndex,
        outcome: input.outcome, checkpoint: input.checkpoint, confirmed: true })
    }
  );
  return wpsExecutionResponse(response, "无法核对 WPS 在途动作");
}

export async function finishWpsExecution(
  operationId: string,
  input: {
    expectedRevision: number;
    ownerToken: string;
    result: Record<string, unknown>;
  }
): Promise<WpsExecutionState> {
  assertWpsExecutionResult(input.result);
  const response = await fetch(
    `${API_BASE_URL}/api/wps/executions/${encodeURIComponent(operationId)}/complete`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ expectedRevision: input.expectedRevision, ownerToken: input.ownerToken,
        result: input.result })
    }
  );
  return wpsExecutionResponse(response, "无法提交 WPS 任务完成回执");
}

export async function selectFolder(): Promise<FolderCatalog | null> {
  const response = await fetch(`${API_BASE_URL}/api/folders/select`, {
    method: "POST"
  });
  return responseJson<FolderCatalog | null>(response);
}

export async function scanFolderByPath(folderPath: string): Promise<FolderCatalog> {
  const response = await fetch(`${API_BASE_URL}/api/folders/scan`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ folderPath })
  });
  return responseJson<FolderCatalog>(response);
}

export async function refreshFolder(sessionId: string): Promise<FolderCatalog> {
  const response = await fetch(`${API_BASE_URL}/api/folders/refresh`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ sessionId })
  });
  return responseJson<FolderCatalog>(response);
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

interface FolderExecuteHooks {
  operationId?: string;
  onProgress?: (event: TaskProgressEvent) => void;
  onConnectionState?: (
    state: "connected" | "reconnecting" | "interrupted",
    message?: string
  ) => void;
  signal?: AbortSignal;
}

export async function executeFolderPlan(
  sessionId: string,
  plan: AnalysisPlan,
  hooks: FolderExecuteHooks = {}
): Promise<FolderExecuteResult> {
  const operationId =
    hooks.operationId ?? `folder-${Date.now().toString(36)}-${Math.random().toString(16).slice(2)}`;
  let lastSequence = 0;

  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      const response = await fetch(`${API_BASE_URL}/api/folders/execute/stream`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sessionId,
          plan,
          operationId,
          afterSequence: lastSequence
        }),
        signal: hooks.signal
      });
      if (!response.ok) {
        throw await requestError(response, "文件夹计划流式执行失败");
      }
      if (!response.body) {
        throw new ApiRequestError(
          502,
          "EMPTY_EXECUTION_STREAM",
          "文件夹执行进度流未返回响应体；为避免重复写入，已停止自动回退",
          false
        );
      }
      hooks.onConnectionState?.("connected");

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer = "";
      let result: FolderExecuteResult | null = null;
      let legacySequence = 0;
      const handleFrame = (frame: string): void => {
        const parsed = parseSseFrame(frame);
        if (!parsed) return;
        const payload: unknown = JSON.parse(parsed.data);
        if (parsed.event === "progress" || parsed.event === "step") {
          legacySequence += 1;
          const progress = parseTaskProgressEvent(
            payload,
            operationId,
            legacySequence
          );
          if (progress) {
            lastSequence = Math.max(lastSequence, progress.sequence);
            hooks.onProgress?.(progress);
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
            typeof detail.message === "string"
              ? detail.message
              : "文件夹计划执行失败",
            typeof detail.retryable === "boolean" ? detail.retryable : false
          );
        }
        if (parsed.event === "result") {
          if (!payload || typeof payload !== "object") {
            throw new Error("文件夹执行结果格式无效");
          }
          result = payload as FolderExecuteResult;
        }
      };

      for (;;) {
        const { done, value } = await reader.read();
        if (value) buffer += decoder.decode(value, { stream: true });
        let separator = buffer.indexOf("\n\n");
        while (separator !== -1) {
          const frame = buffer.slice(0, separator);
          buffer = buffer.slice(separator + 2);
          if (frame.trim()) handleFrame(frame);
          separator = buffer.indexOf("\n\n");
        }
        if (done) break;
      }
      if (buffer.trim()) handleFrame(buffer.trim());
      if (!result) {
        throw new ApiRequestError(
          502,
          "SERVICE_ERROR",
          "文件夹执行流已结束，但没有返回完整结果",
          true
        );
      }
      return result;
    } catch (reason) {
      if (
        attempt === 0 &&
        !(reason instanceof ApiRequestError) &&
        !hooks.signal?.aborted
      ) {
        // 使用相同 operationId 重连；服务端会追赶历史事件而不会重复写文件。
        hooks.onConnectionState?.(
          "reconnecting",
          "连接短暂中断，正在使用同一任务标识恢复进度。"
        );
        continue;
      }
      if (!hooks.signal?.aborted) {
        hooks.onConnectionState?.(
          "interrupted",
          "连接中断，执行结果待确认。"
        );
      }
      throw reason;
    }
  }
  throw new Error("文件夹计划执行失败");
}

export async function getFolderExecutionStatus(
  operationId: string
): Promise<FolderExecutionStatusResult> {
  const params = new URLSearchParams({ operationId });
  const response = await fetch(
    `${API_BASE_URL}/api/folders/execute/status?${params.toString()}`
  );
  return responseJson<FolderExecutionStatusResult>(response);
}

export async function resumeFolderExecution(
  operationId: string,
  expectedRevision: number,
  hooks: FolderExecuteHooks = {}
): Promise<FolderExecuteResult> {
  const response = await fetch(`${API_BASE_URL}/api/folders/execute/resume/stream`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      operationId,
      expectedRevision,
      confirmed: true,
      afterSequence: 0
    }),
    signal: hooks.signal
  });
  if (!response.ok) {
    throw await requestError(response, "文件夹任务断点续跑失败");
  }
  if (!response.body) {
    throw new ApiRequestError(
      502,
      "EMPTY_EXECUTION_STREAM",
      "断点续跑进度流未返回响应体",
      false
    );
  }
  hooks.onConnectionState?.("connected");
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let result: FolderExecuteResult | null = null;
  let legacySequence = 0;
  const handleFrame = (frame: string): void => {
    const parsed = parseSseFrame(frame);
    if (!parsed) return;
    const payload: unknown = JSON.parse(parsed.data);
    if (parsed.event === "progress" || parsed.event === "step") {
      legacySequence += 1;
      const progress = parseTaskProgressEvent(
        payload,
        operationId,
        legacySequence
      );
      if (progress) hooks.onProgress?.(progress);
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
        typeof detail.message === "string" ? detail.message : "断点续跑失败",
        typeof detail.retryable === "boolean" ? detail.retryable : false
      );
    }
    if (parsed.event === "result") {
      if (!payload || typeof payload !== "object") {
        throw new Error("断点续跑结果格式无效");
      }
      result = payload as FolderExecuteResult;
    }
  };
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (value) buffer += decoder.decode(value, { stream: true });
      let separator = buffer.indexOf("\n\n");
      while (separator !== -1) {
        const frame = buffer.slice(0, separator);
        buffer = buffer.slice(separator + 2);
        if (frame.trim()) handleFrame(frame);
        separator = buffer.indexOf("\n\n");
      }
      if (done) break;
    }
    if (buffer.trim()) handleFrame(buffer.trim());
  } catch (reason) {
    if (!hooks.signal?.aborted) {
      hooks.onConnectionState?.("interrupted", "连接中断，续跑任务状态待确认。");
    }
    throw reason;
  }
  if (!result) {
    throw new ApiRequestError(
      502,
      "SERVICE_ERROR",
      "断点续跑进度流已结束，但没有返回完整结果",
      true
    );
  }
  return result;
}

export async function undoFolderExecution(
  sessionId: string,
  undoId: string
): Promise<FolderUndoResult> {
  const response = await fetch(`${API_BASE_URL}/api/folders/undo`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ sessionId, undoId })
  });
  return responseJson<FolderUndoResult>(response);
}

export async function listFolderBackups(
  target: { sessionId: string } | { folderPath: string }
): Promise<FolderBackupsResult> {
  const params = new URLSearchParams();
  if ("sessionId" in target) {
    params.set("sessionId", target.sessionId);
  } else {
    params.set("folderPath", target.folderPath);
  }
  const response = await fetch(
    `${API_BASE_URL}/api/folders/backups?${params.toString()}`
  );
  return responseJson<FolderBackupsResult>(response);
}

export async function listManagedFolderRecoveryDirectories(): Promise<ManagedFolderRecoveryDirectoriesResult> {
  const response = await fetch(`${API_BASE_URL}/api/folders/recovery/directories`);
  return responseJson<ManagedFolderRecoveryDirectoriesResult>(response);
}

export async function removeManagedFolderRecoveryDirectory(
  directoryId: string
): Promise<{ directoryId: string; removed: boolean }> {
  const response = await fetch(
    `${API_BASE_URL}/api/folders/recovery/directories/${encodeURIComponent(directoryId)}`,
    { method: "DELETE" }
  );
  return responseJson<{ directoryId: string; removed: boolean }>(response);
}

export async function restoreManagedFolderBackup(input: {
  directoryId: string;
  undoId: string;
  expectedRevision: number;
  operationId: string;
  confirmed: true;
}): Promise<FolderManagedRecoveryResult> {
  const response = await fetch(`${API_BASE_URL}/api/folders/recovery/restore`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input)
  });
  return responseJson<FolderManagedRecoveryResult>(response);
}

export async function openFolderBackupDirectory(
  directoryToken: string
): Promise<FolderBackupOpenResult> {
  const response = await fetch(`${API_BASE_URL}/api/folders/backups/open`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ directoryToken })
  });
  return responseJson<FolderBackupOpenResult>(response);
}

export async function exportFolderBackup(
  target: { sessionId: string } | { backupManagementToken: string },
  undoId: string,
  operationId?: string
): Promise<FolderBackupExportResult> {
  const response = await fetch(`${API_BASE_URL}/api/folders/backups/export`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ...target, undoId, ...(operationId ? { operationId } : {}) })
  });
  return responseJson<FolderBackupExportResult>(response);
}

export async function deleteFolderBackup(
  backupManagementToken: string,
  undoId: string,
  expectedRevision?: number,
  expectedInvalidatedUndoIds?: string[]
): Promise<FolderBackupDeleteResult> {
  const response = await fetch(`${API_BASE_URL}/api/folders/backups/delete`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      backupManagementToken,
      undoId,
      confirmed: true,
      ...(expectedRevision !== undefined ? { expectedRevision } : {}),
      ...(expectedInvalidatedUndoIds ? { expectedInvalidatedUndoIds } : {})
    })
  });
  return responseJson<FolderBackupDeleteResult>(response);
}

export async function setFolderBackupPinned(
  backupManagementToken: string,
  undoId: string,
  pinned: boolean
): Promise<FolderBackupPinResult> {
  const response = await fetch(`${API_BASE_URL}/api/folders/backups/pin`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ backupManagementToken, undoId, pinned })
  });
  return responseJson<FolderBackupPinResult>(response);
}

export async function cleanupFolderBackups(
  target: { sessionId: string } | { backupManagementToken: string }
): Promise<FolderBackupCleanupResult> {
  const response = await fetch(`${API_BASE_URL}/api/folders/backups/cleanup`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(target)
  });
  return responseJson<FolderBackupCleanupResult>(response);
}

export async function collectLegacyFolderBackups(
  target: { sessionId: string } | { backupManagementToken: string }
): Promise<FolderLegacyCollectResult> {
  const response = await fetch(
    `${API_BASE_URL}/api/folders/backups/legacy/collect`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ...target, confirmed: true })
    }
  );
  return responseJson<FolderLegacyCollectResult>(response);
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

export async function generateFormula(
  request: GenerateFormulaRequest
): Promise<GenerateFormulaResponse> {
  const response = await fetch(`${API_BASE_URL}/api/formulas/generate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request)
  });
  return responseJson<GenerateFormulaResponse>(response);
}

export async function chatModifyPlan(
  request: PlanChatRequest,
  signal?: AbortSignal
): Promise<PlanChatResponse> {
  const response = await fetch(`${API_BASE_URL}/api/plans/chat-modify`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request),
    signal
  });
  if (!response.ok) {
    throw await requestError(response, "AI 辅助修改计划失败");
  }
  return responseJson<PlanChatResponse>(response);
}
