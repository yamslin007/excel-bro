import { assertWorksheetCheckpointResult, assertWorksheetRestoreResult } from "./contracts";
import { waitForPlanning } from "@excel-bro/core/planningReplay";
import type { PlanningPauseAck } from "./contracts";
import type { ModelCatalog } from "./contracts";
import capabilities from "../../../config/capabilities.json";
import { modelUsageNotice } from "@excel-bro/core/modelUsageNotice";
import {
  assertTestModelConnectionResponse,
  assertIntentCheckResponse,
  assertMessageUnderstandingResponse,
  type MessageUnderstandingRequest,
  type MessageUnderstandingResponse,
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
  type TestModelConnectionRequest,
  type UpdateModelSettingsRequest,
  type UpsertModelConnectionRequest,
  type SetFormulaModelRequest,
  type WorkbookSnapshot,
  type GenerateFormulaRequest,
  type GenerateFormulaResponse,
  type SdkBridgeBindingRequest,
  type SdkBridgePollResponse,
  type SdkBridgeReceiptRequest,
  type SdkBridgeReceiptResponse,
  type SdkBridgeRegisterRequest,
  type SdkBridgeRegisterResponse,
  type SdkBridgeResultRequest,
  type SdkBridgeResultResponse
} from "./contracts";
import type { ExcelReviewLayoutPrepareRequest, ExcelReviewLayoutTicket, ExcelReviewLayoutResult, ExcelReviewPaneMeasurement,
  ExcelReviewCopyCloseRequest, ExcelReviewCopyCloseResult, ExcelReviewCopyCloseStatus } from "./contracts";
import { recordModelCall } from "./diagnostics";
import { requiresStreamingToolDelivery } from "./modelSelection";

export interface PlanChatRequest {
  plan: AnalysisPlan;
  message: string;
  modelId?: string;
}

export interface PlanChatResponse {
  explanation: string;
  modifiedPlan: AnalysisPlan;
}

/**
 * 本地服务基址。导出是为了让 WPS 端的规划流客户端复用**同一份**解析规则
 * （含 `VITE_API_BASE_URL` 覆盖），而不是各自再写一遍环境变量读取。
 */
export const API_BASE_URL =
  import.meta.env.VITE_API_BASE_URL ?? "http://127.0.0.1:8765";

/** 代码正文取回的结果：要么真的拿到了正文，要么明确说没拿到以及为什么。 */
export type PythonCodeBodyResult =
  | { status: "ready"; code: string; chars: number }
  | { status: "unavailable"; reason: "missing" | "offline" };

export class ExcelReviewLayoutVersionError extends Error {
  constructor() {
    super('本地窗口排列服务版本过旧，请重启 Python 服务、刷新加载项并重新生成副本。');
    this.name = 'ExcelReviewLayoutVersionError';
  }
}

async function reviewLayoutRequest(path: "prepare" | "arrange" | "measure" | "close", body: ExcelReviewCopyCloseRequest | ExcelReviewLayoutPrepareRequest | Pick<ExcelReviewLayoutTicket, 'token'>): Promise<unknown> {
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
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('审阅窗口宽度未核实，请手动调整。');
  const record = value as Record<string, unknown>;
  const width = (field: 'clientWidthCssPx' | 'workAreaWidthCssPx') => {
    const candidate = record[field];
    return typeof candidate === 'number' && Number.isFinite(candidate) && candidate > 0 && candidate <= 32768;
  };
  // 必须**两个**字段齐备且合法：AI 比例只认工作区宽度，缺了它宁可降级为手动，
  // 也不能拿原表窗口的内容区宽度顶替（那正是此前把原表挤没的根因）。
  if (Object.keys(record).length !== 2 || !width('clientWidthCssPx') || !width('workAreaWidthCssPx'))
    throw new Error('审阅窗口宽度未核实，请手动调整。');
  // 显式投影，不把响应的其他键带出去。
  return { clientWidthCssPx: record.clientWidthCssPx as number, workAreaWidthCssPx: record.workAreaWidthCssPx as number };
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
const CLOSE_STATUSES: readonly ExcelReviewCopyCloseStatus[] =
  ['closed', 'already_closed', 'close_unconfirmed', 'user_saved', 'identity_lost', 'ambiguous', 'unsupported', 'failed'];
export async function closeExcelReviewCopy(request: ExcelReviewCopyCloseRequest): Promise<ExcelReviewCopyCloseResult> {
  // Required v4 discriminator: older servers reject it before any close mutation.
  const value = await reviewLayoutRequest('close', {closeVersion: 4, workbookPathHash: request.workbookPathHash,
    worksheetNameHash: request.worksheetNameHash, copySheetName: request.copySheetName});
  if (value && typeof value === 'object' && (!('layoutVersion' in value) || value.layoutVersion !== 4)) throw new ExcelReviewLayoutVersionError();
  if (!value || typeof value !== 'object' || Object.keys(value).length !== 4 ||
      !('status' in value) || typeof value.status !== 'string' || !CLOSE_STATUSES.includes(value.status as ExcelReviewCopyCloseStatus) ||
      !('copyClosed' in value) || typeof value.copyClosed !== 'boolean' ||
      !('recoveryCopyPath' in value) || !(value.recoveryCopyPath === null ||
        (typeof value.recoveryCopyPath === 'string' && value.recoveryCopyPath.trim().length > 0)) ||
      (value.status === 'closed' && !value.recoveryCopyPath)) throw new Error('副本关闭结果未核实，请检查副本窗口。');
  return {status: value.status as ExcelReviewCopyCloseStatus, layoutVersion: 4,
    copyClosed: value.copyClosed, recoveryCopyPath: value.recoveryCopyPath as string | null};
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
export type { ModelOption, ModelCatalog } from "./contracts";

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
    const usageNotice = modelUsageNotice(code);
    super(usageNotice ?? message);
    // No automatic retry for a hard limit, even when an older server labels it retryable.
    if (usageNotice) this.retryable = false;
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

/**
 * 当前链路无法把模型发起的工具请求投递到本机（缺流式通道）。
 *
 * 这不是"网络不通可以重试"那类错误：退回非流式端点后，服务端仍会启动这一轮、
 * 仍会等本机回执，而投递工具请求的唯一通道就是被判定不可用的那条 SSE —— 于是
 * 服务端干等到 `agentToolBridge.waitTimeoutSeconds` 才拿到一个 `timeout`。
 * 与其进入这种"无人回执的等待"，不如在**退回之前**就明确拒绝。
 */
export class StreamingToolDeliveryUnavailableError extends Error {
  readonly code = "streaming_tool_delivery_unavailable";
  readonly retryable = false;
  constructor() {
    super(
      "当前链路无法把模型发起的表格查询请求投递到本机：该后端未提供流式接口，或代理/网关没有透传流式响应。" +
        "请更新并重启本地 Python 服务，并确认代理不缓冲 SSE。"
    );
    this.name = "StreamingToolDeliveryUnavailableError";
  }
}

/**
 * 流式通道不可用时的**唯一**出口。
 *
 * 需要工具投递的请求（SDK 适配器）绝不能退回非流式端点 —— 那会让服务端为一轮
 * 送不到本机的工具请求空等；其余请求保持原来的回退行为，旧后端照旧可用。
 */
function fallbackOrRefuseToolDelivery(
  request: PlanRequest,
  signal?: AbortSignal
): Promise<PlanningStreamResult> {
  if (requiresStreamingToolDelivery(request.modelId)) {
    throw new StreamingToolDeliveryUnavailableError();
  }
  return createAssistantResponse(request, signal);
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

  const cancelReader = () => { void reader.cancel().catch(() => undefined); };
  hooks.signal?.addEventListener("abort", cancelReader, { once: true });
  try {
  for (;;) {
    const { done, value } = await waitForPlanning(reader.read(), hooks.signal);
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

  } finally {
    hooks.signal?.removeEventListener("abort", cancelReader);
    if (hooks.signal?.aborted) cancelReader();
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
    // 但需要工具投递的请求不能回退（服务端这一轮根本还没开始，回退毫无意义）。
    if (isLocalServiceConnectionError(reason)) {
      return fallbackOrRefuseToolDelivery(request, signal);
    }
    throw reason;
  }

  if (!response.ok || !response.body) {
    // 旧后端无此端点（404）或代理不支持流式：回退到非流式端点。
    if (response.status === 404 || response.status === 405 || (response.ok && !response.body)) {
      // SDK 适配器**只能**经流式通道拿工具请求：回退 = 服务端空等，故明确拒绝。
      return fallbackOrRefuseToolDelivery(request, signal);
    }
    throw await requestError(response, "本地 AI 服务流式请求失败");
  }

  return consumePlanningStream(
    response,
    request.turnId ?? "legacy-turn",
    hooks
  );
}

export async function pausePlanningTask(turnId: string): Promise<PlanningPauseAck> {
  const response = await fetch(`${API_BASE_URL}/api/turn/${encodeURIComponent(turnId)}/pause`, { method: "POST" });
  if (!response.ok) throw await requestError(response, "未能向原任务发送暂停请求");
  return await response.json() as PlanningPauseAck;
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

/**
 * 补发**已经落盘**的规划进度事件（只读，不重新规划）。
 *
 * 服务端仍在跑时「继续」会被 409 挡住 —— 那是对的，再发一次就是重复规划。
 * 想接回**正在运行**的原任务，只能读它已保存的事件。
 */
export async function getPlanningTaskEvents(
  turnId: string,
  afterSequence = 0,
  signal?: AbortSignal
): Promise<{
  turnId: string;
  status: string;
  active: boolean;
  lastSequence: number;
  events: unknown[];
}> {
  const query = afterSequence > 0 ? `?afterSequence=${afterSequence}` : "";
  const response = await fetch(
    `${API_BASE_URL}/api/turn/${encodeURIComponent(turnId)}/events${query}`,
    { signal }
  );
  if (!response.ok) {
    throw await requestError(response, "读取原任务进度事件失败");
  }
  const value = (await response.json()) as Record<string, unknown>;
  if (!Array.isArray(value.events)) {
    throw new ApiRequestError(502, "PLANNING_BAD_EVENT_PAGE", "原任务进度事件格式不正确", false);
  }
  return {
    turnId,
    status: typeof value.status === "string" ? value.status : "",
    active: value.active === true,
    lastSequence: Number(value.lastSequence ?? 0) || 0,
    events: value.events
  };
}

/**
 * 按 `codeId` 取回这次任务生成过的 Python 代码正文。
 *
 * 代码正文**不进持久账本**（会随 `planning_task_events` 落进 SQLite），只在服务端
 * 进程内短期登记。因此界面拿到的事件只有"状态 + 规模"，正文要用这个端点按需取回。
 *
 * 取不回来（登记表过期 / 被淘汰 / 任务对不上 / 本机服务没开）时**不抛错**：
 * 对界面来说这三种情况是同一件事——"手上没有这段代码"，必须照实说，
 * 不能拿状态标签假装代码还在。
 */
export async function fetchPythonCodeBody(
  request: { taskId: string; codeId: string },
  signal?: AbortSignal
): Promise<PythonCodeBodyResult> {
  let response: Response;
  try {
    response = await fetch(`${API_BASE_URL}/api/sdk/python-code`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(request),
      signal
    });
  } catch {
    return { status: "unavailable", reason: "offline" };
  }
  if (response.status === 404) return { status: "unavailable", reason: "missing" };
  if (!response.ok) return { status: "unavailable", reason: "offline" };
  let value: unknown;
  try {
    value = await response.json();
  } catch {
    return { status: "unavailable", reason: "offline" };
  }
  const payload = (value ?? {}) as Record<string, unknown>;
  // 只认"确实给出了正文"的响应：`code` 缺失时回 `unavailable`，
  // 绝不用空串顶替（空串会被上层当成"拿到了代码"）。
  if (typeof payload.code !== "string" || payload.code.length === 0) {
    return { status: "unavailable", reason: "missing" };
  }
  return {
    status: "ready",
    code: payload.code,
    chars: typeof payload.chars === "number" ? payload.chars : payload.code.length
  };
}

export async function continuePlanningTask(
  request: {
    turnId: string;
    expectedRevision: number;
    requestId: string;
    afterSequence: number;
    modelId?: string;
    confirmModelSwitch?: true;
    /**
     * 本轮**重新读到的**现场证据。三项必填：服务端不再接受"没传就跳过比对"的
     * 请求，所以缺证据应该在发请求之前停下，而不是发一个省略字段的请求碰运气。
     */
    workbookIdentity: string;
    authorizedSheetNames: string[];
    /** 来源核实的**三态**；"读取失败"与"尚未采集"必须分开表达。 */
    sourceVerification: "verified" | "not_collected" | "failed";
    /** 仅 `sourceVerification === "verified"` 时提供。 */
    sourceFingerprint?: string;
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
    // 续跑没有非流式回退通道；对 SDK 适配器给出可执行的明确原因，而不是笼统一句失败。
    if ((response.status === 404 || response.status === 405 || response.ok) &&
        requiresStreamingToolDelivery(request.modelId)) {
      throw new StreamingToolDeliveryUnavailableError();
    }
    throw await requestError(response, "继续规划任务失败");
  }
  return consumePlanningStream(response, request.turnId, hooks);
}

export async function understandMessage(
  request: MessageUnderstandingRequest,
  signal?: AbortSignal
): Promise<MessageUnderstandingResponse> {
  recordModelCall();
  const response = await fetch(`${API_BASE_URL}/api/intent/message`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    // Project explicitly: even an over-wide runtime object cannot leak workbook context.
    body: JSON.stringify({
      prompt: request.prompt,
      imageCount: request.imageCount ?? 0,
      conversation: request.conversation?.map(({ role, text }) => ({ role, text })) ?? [],
      modelId: request.modelId ?? null,
      ...(request.recoveryTasks?.length ? { recoveryTasks: request.recoveryTasks
        .filter((id) => typeof id === "string" && id.length > 0 && id.length <= 200).slice(-6) } : {})
    }),
    signal
  });
  if (!response.ok) throw await requestError(response, "消息理解服务请求失败；请确认后端已更新并重启。");
  const value: unknown = await response.json();
  assertMessageUnderstandingResponse(value);
  if (value.kind === "task_recovery" &&
      (request.imageCount || !request.recoveryTasks?.includes(value.turnId))) {
    throw new Error("消息理解返回了未提供的恢复任务；未查询、未继续。");
  }
  return value;
}

export async function checkIntent(
  request: IntentCheckRequest,
  signal?: AbortSignal
): Promise<IntentCheckResponse> {
  recordModelCall();
  const response = await waitForPlanning(fetch(`${API_BASE_URL}/api/turn`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request),
    signal
  }), signal);

  if (!response.ok) {
    throw await requestError(response, "需求确认服务请求失败");
  }

  const value: unknown = await waitForPlanning(response.json(), signal);
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
  request: TestModelConnectionRequest
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
  existingSheets: string[] = authorizedSheets,
  signal?: AbortSignal
): Promise<AnalysisPlan> {
  const request: PlanValidateRequest = {
    plan,
    authorized_sheets: authorizedSheets,
    existing_sheets: existingSheets
  };
  const response = await fetch(`${API_BASE_URL}/api/plans/validate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request), signal
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
  /**
   * 产生这份计划的规划任务。带上它，服务端才能在这个任务"其实已经写过原表"
   * 时拒绝「继续分析」；不知道来源（历史计划重放、`/function` 预览）时留空，
   * 空只是不声称关联，不会让写入被放行。
   */
  taskId?: string;
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

export async function checkFolderAdmission(sessionId: string, selections: FolderSelection[]): Promise<import("./contracts").FolderAdmissionResponse> {
  const response = await fetch(`${API_BASE_URL}/api/folders/admission`, {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ sessionId, selections: selections.map((s) => ({ fileId: s.fileId, sheets: [...s.sheets] })) })
  });
  return responseJson<import("./contracts").FolderAdmissionResponse>(response);
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

// ---------------------------------------------------------------------------
// 任务工具桥（本机回环）：请求参数与值矩阵只走这条通道，不进模型/历史/账本。
// ---------------------------------------------------------------------------

function bridgeBody(binding: SdkBridgeBindingRequest): SdkBridgeBindingRequest {
  return {
    taskId: binding.taskId,
    hostId: binding.hostId,
    workbookName: binding.workbookName,
    workbookIdentity: binding.workbookIdentity ?? null,
    sourceFingerprint: binding.sourceFingerprint ?? null
  };
}

/** 拉取本次会话待执行的本地查询调用。 */
export async function pollSdkBridgeTool(
  binding: SdkBridgeBindingRequest,
  signal?: AbortSignal
): Promise<SdkBridgePollResponse> {
  const response = await fetch(`${API_BASE_URL}/api/sdk/tool-bridge/poll`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(bridgeBody(binding)),
    signal
  });
  return responseJson<SdkBridgePollResponse>(response);
}

/** 回执：本机已在真实宿主上跑完确定性查询。 */
export async function postSdkBridgeReceipt(
  request: SdkBridgeReceiptRequest,
  signal?: AbortSignal
): Promise<SdkBridgeReceiptResponse> {
  const response = await fetch(`${API_BASE_URL}/api/sdk/tool-bridge/receipt`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ...bridgeBody(request), operationId: request.operationId,
      requestFingerprint: request.requestFingerprint, receipt: request.receipt }),
    signal
  });
  return responseJson<SdkBridgeReceiptResponse>(response);
}

/**
 * 登记**已经算好**的本地结果（不重跑查询）。
 *
 * 服务端要求存在活跃且归属一致的会话：只凭可猜的 taskId 不能替别人登记值。
 */
export async function registerSdkBridgeResult(
  request: SdkBridgeRegisterRequest,
  signal?: AbortSignal
): Promise<SdkBridgeRegisterResponse> {
  const response = await fetch(`${API_BASE_URL}/api/sdk/tool-bridge/register`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ...bridgeBody(request), result: request.result }),
    signal
  });
  return responseJson<SdkBridgeRegisterResponse>(response);
}

/**
 * 按引用取回结果值（本地结果卡渲染路径）。
 *
 * 跨轮读取：此时会话可能已关闭，服务端校验的是**授权归属**而不是活跃会话；
 * 授权变化 / 换簿 / 同名不同工作簿都会被拒绝。
 */
export async function fetchSdkBridgeResult(
  request: SdkBridgeResultRequest,
  signal?: AbortSignal
): Promise<SdkBridgeResultResponse> {
  const response = await fetch(`${API_BASE_URL}/api/sdk/tool-bridge/result`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ...bridgeBody(request), resultRef: request.resultRef }),
    signal
  });
  return responseJson<SdkBridgeResultResponse>(response);
}


export async function executeWorksheetDeletion(input: import("./contracts").WorksheetDeleteCheckpointRequest) {
  const result = await responseJson<unknown>(await fetch(`${API_BASE_URL}/api/worksheet-checkpoints/execute`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(input)
  }));
  assertWorksheetCheckpointResult(result, input.operationId);
  return result;
}
export async function worksheetCheckpointStatus(operationId: string) {
  const result = await responseJson<unknown>(await fetch(`${API_BASE_URL}/api/worksheet-checkpoints/status`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ operationId })
  }));
  assertWorksheetCheckpointResult(result, operationId);
  return result;
}
export async function openWorksheetCheckpoint(operationId: string) {
  const result = await responseJson<import("./contracts").WorksheetCheckpointOpenResult>(await fetch(`${API_BASE_URL}/api/worksheet-checkpoints/open`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ operationId })
  }));
  if (result.opened !== true || typeof result.recoveryPath !== "string" || !result.recoveryPath) throw new Error("恢复副本打开结果尚未确认；当前原表保留。");
  return result;
}

/**
 * **原地**撤销本次删除（把表放回当前打开的原工作簿）。
 *
 * 与 `openWorksheetCheckpoint` 是两条路：这里不打开任何新文件、不改文件身份。
 * 回执必须自带覆盖关系（见 `assertWorksheetRestoreResult`），否则不当作已完成。
 */
export async function restoreWorksheetCheckpoint(operationId: string, deletedSheets: readonly string[]) {
  const result = await responseJson<unknown>(await fetch(`${API_BASE_URL}/api/worksheet-checkpoints/restore`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ operationId })
  }));
  assertWorksheetRestoreResult(result, operationId, deletedSheets);
  return result;
}

/**
 * 按**工作簿身份**列出恢复点（只读）。
 *
 * 旧记录只有请求元数据、没有会话归属：这里按服务端保存的工作簿路径把它们找回来，
 * 但不会替它们编一个会话。列表返回的是摘要，不含单元格内容。
 */
export async function listWorksheetCheckpoints(input: import("./contracts").WorksheetCheckpointListRequest) {
  const result = await responseJson<import("./contracts").WorksheetCheckpointListResult>(
    await fetch(`${API_BASE_URL}/api/worksheet-checkpoints/list`, {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(input)
    })
  );
  const checkpoints = Array.isArray(result?.checkpoints) ? result.checkpoints : null;
  if (!checkpoints) throw new Error("恢复点列表不完整，请稍后重试。");
  return checkpoints.map(item => {
    if (typeof item.operationId !== "string" || !/^[a-f0-9]{32}$/.test(item.operationId) ||
        typeof item.message !== "string" || typeof item.recoveryAvailable !== "boolean" ||
        !Array.isArray(item.deletedSheets) || !Array.isArray(item.remainingSheets)) {
      throw new Error("恢复点列表不完整，请稍后重试。");
    }
    return item;
  });
}

export async function confirmWorksheetCheckpointBackup(input: import("./contracts").WorksheetCheckpointBackupConsent) {
  const result = await responseJson<unknown>(await fetch(`${API_BASE_URL}/api/worksheet-checkpoints/confirm-backup`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(input)
  }));
  assertWorksheetCheckpointResult(result, input.operationId);
  return result;
}
