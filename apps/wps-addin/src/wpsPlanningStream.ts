import { waitForPlanning } from "@excel-bro/core/planningReplay";
/**
 * WPS 端的**流式规划客户端**（CAP-02）。
 *
 * 与 Excel 端共用同一套服务端协议与同一份共享语义
 * （`@excel-bro/core` 的 `planningStream`），但宿主侧适配留在本端：
 *
 *  - 事件归并、游标判定、按钮推导 → 共享层 `planningStream`；
 *  - HTTP 传输、SSE 分帧、宿主身份与授权检查 → 本模块 / 调用方。
 *
 * 三条纪律（与项目既有边界一致）：
 *
 * 1. **流式与非流式必须走同一结构化出口。** 本模块的 `streamWpsPlanning` 与
 *    `runWpsTurn` 里那次 `createAssistantResponse` 返回的是**同一个**
 *    `AssistantResponse` 联合类型，由 `finalizeResponse` 归一。事件展示是增量，
 *    但最终结果不因传输方式不同而变形。
 *
 * 2. **断线不重建任务。** 连不上时**不**退回"再发一次需求"，而是把 `turnId`
 *    交给调用方去查 `getPlanningTaskStatus` —— 那才是权威现场。
 *
 * 3. **继续只带身份与游标。** `continueWpsPlanning` 只发送 `turnId` /
 *    `expectedRevision` / `requestId` / `afterSequence`，**不重发原需求**，
 *    因此服务端不会重新规划、不会重复调用工具。
 */
import type { AssistantResponse, AssistantResponseOrPause } from "./wpsPlanningTypes";
import {
  type PlanningRecoveryDescriptor,
  type TaskProgressEvent,
  parseTaskProgressEvent
} from "../../excel-addin/src/contracts";
import {
  API_BASE_URL,
  ApiRequestError,
  StreamingToolDeliveryUnavailableError
} from "../../excel-addin/src/api";
import { requiresStreamingToolDelivery } from "../../excel-addin/src/modelSelection";
import {
  MODEL_ERROR_CODE_PREFIX,
  applyPlanningEvent,
  newPlanningCursor,
  planningResumeDecision,
  reconcilePlanningCursor,
  type PlanningContinuationView,
  type PlanningCursor,
  type PlanningPauseType,
  type PlanningResumeDecision,
  type PlanningStatusView,
  type PlanningTaskStatus
} from "@excel-bro/core";

/** SSE 帧解析结果。`event` 缺省是 `message`，与 Excel 端 `parseSseFrame` 同口径。 */
export interface WpsSseFrame {
  event: string;
  data: string;
}

/** 与 Excel 端 `api.ts` 的 `parseSseFrame` 保持同一语义（含无 event 行的默认值）。 */
export function parseWpsSseFrame(frame: string): WpsSseFrame | null {
  let event = "message";
  const dataLines: string[] = [];
  for (const line of frame.split("\n")) {
    if (line.startsWith("event:")) {
      event = line.slice(6).trim();
    } else if (line.startsWith("data:")) {
      dataLines.push(line.slice(5).replace(/^ /, ""));
    }
  }
  if (dataLines.length === 0) return null;
  return { event, data: dataLines.join("\n") };
}

/** 合规的终态事件类型。未知事件一律**不吞**、按协议错误抛出（见 `consumeWpsPlanningStream`）。 */
export const WPS_KNOWN_STREAM_EVENTS = ["step", "progress", "paused", "result", "error"] as const;

/**
 * 消费 WPS 规划 SSE 流。
 *
 * **未知事件不静默吞掉。** 服务端协议若新增事件类型而本端未适配，直接抛
 * `WPS_PLANNING_UNKNOWN_EVENT` —— 表面成功（把未知事件当 answer 处理）比
 * 明确报错危险得多：用户会看到一个按钮齐全、实际什么都没做的卡片。
 */
export async function consumeWpsPlanningStream(
  response: Response,
  turnId: string,
  hooks: {
    onProgress?: (event: TaskProgressEvent) => void;
    signal?: AbortSignal;
    /**
     * 客户端**进入这条流之前**已确认消费到的事件序号（断线前的游标）。
     *
     * 传它是为了把共享层的两道闸接进**真实**消费链：
     *
     *  - **身份闸**：`step` 帧的 `operationId` 必须等于本次 `turnId`。不等即丢 ——
     *    这是"取消/换任务后旧事件的迟到帧不许污染新任务"的唯一守门点。
     *  - **去重闸**：`sequence` 必须严格大于已消费的最大值，重放/重复帧丢弃。
     *
     * 省略时按 0 起算（本端从未消费过这条任务的事件）。**绝不能**传一个比
     * 实际消费位置更大的值 —— 那会静默吞掉中间的事件。
     */
    afterSequence?: number;
  } = {}
): Promise<AssistantResponseOrPause> {
  if (!response.body) {
    throw new ApiRequestError(502, "SERVICE_ERROR", "本地 AI 服务没有返回数据流", true);
  }
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let result: AssistantResponseOrPause | null = null;
  /*
   * 游标从"客户端已确认位置"起算，并只由共享层 `applyPlanningEvent` 推进。
   * 本模块不自己维护 `legacySequence` 计数器：那样会绕过去重与身份闸。
   */
  let cursor = newPlanningCursor(turnId);
  if (hooks.afterSequence && hooks.afterSequence > 0) {
    cursor = { ...cursor, clientLastSequence: hooks.afterSequence, serverLastSequence: hooks.afterSequence };
  }

  const handleFrame = (frame: string): void => {
    const parsed = parseWpsSseFrame(frame);
    if (!parsed) return;
    let payload: unknown;
    try {
      payload = JSON.parse(parsed.data);
    } catch {
      throw new ApiRequestError(502, "WPS_PLANNING_BAD_FRAME", "规划流返回了无法解析的数据帧", false);
    }
    if (parsed.event === "step" || parsed.event === "progress") {
      const raw = (payload ?? {}) as Record<string, unknown>;
      /*
       * 事件身份取 `operationId`。缺失时退回本次 `turnId` —— 服务端所有
       * `step` 帧都带 `operationId`（`_planning_task_stream` 的 `on_event`），
       * 兜底只为兼容旧帧，不改变"身份必须是这条任务"这条判据。
       */
      const eventTurnId =
        typeof raw.operationId === "string" && raw.operationId.trim()
          ? raw.operationId
          : turnId;
      const sequence = Number(raw.sequence);
      const applied = applyPlanningEvent(
        cursor,
        {
          sequence: Number.isFinite(sequence) ? sequence : 0,
          stepId: typeof raw.stepId === "string" ? raw.stepId : "",
          phase: typeof raw.phase === "string" ? raw.phase : "planning",
          status: typeof raw.status === "string" ? raw.status : "progress",
          title: typeof raw.title === "string" ? raw.title : ""
        },
        eventTurnId
      );
      // 被拒的帧（身份不符 / 陈旧重复）**就地丢弃**：不推进游标、不上报进度。
      if (!applied.accepted) return;
      cursor = applied.cursor;
      const progress = parseTaskProgressEvent(payload, turnId, cursor.clientLastSequence);
      if (progress && hooks.onProgress) hooks.onProgress(progress);
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
    if (parsed.event === "paused" || parsed.event === "result") {
      result = payload as AssistantResponseOrPause;
      return;
    }
    // 未知事件：明确失败，绝不当作 answer 蒙混过去。
    throw new ApiRequestError(
      502,
      "WPS_PLANNING_UNKNOWN_EVENT",
      `规划流返回了本端尚不支持的事件类型：${parsed.event}`,
      false
    );
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
    throw new ApiRequestError(502, "SERVICE_ERROR", "本地 AI 服务未返回完整结果", true);
  }
  return result;
}

/** `GET /api/turn/{turnId}/status` 的响应视图。 */
export interface WpsPlanningTaskStatus {
  turnId: string;
  status: PlanningTaskStatus;
  revision: number;
  lastSequence: number;
  /**
   * **没有值被省略**时的完整响应。
   * 一旦服务端只给了 `recovery`（值被省略过），这里就是 `undefined` ——
   * 绝不能拿一个"矩阵全空"的记录冒充完整响应。
   */
  result?: AssistantResponse;
  /** 值被省略过时的版本化恢复描述（与 `result` 互斥）。 */
  recovery?: PlanningRecoveryDescriptor | null;
  error?: { status?: number; code?: string; message?: string; retryable?: boolean } | null;
  canContinue?: boolean;
  pauseType?: PlanningPauseType | null;
  reason?: string | null;
  /**
   * 「继续分析」的权威事实：能不能继续、为什么不能、停在哪个阶段。
   * 服务端用共享契约 + 执行账本算出，本端不得自行推断。
   */
  continuation?: PlanningContinuationView | null;
  /**
   * 原任务记录下来的来源指纹与来源表。
   *
   * 用于"读到了但不一样"的**本地**判定（服务端还会再判一次）。缺了它，WPS 只能
   * 等服务端回 409 才知道数据变过，用户在看到结果前多一次无效请求。
   */
  sourceFingerprint?: string | null;
  sourceFingerprintSheets?: string[];
}

/** `GET /api/turn/{turnId}/events` 的响应视图。 */
export interface WpsPlanningTaskEvents {
  turnId: string;
  status: PlanningTaskStatus;
  /** 这条任务此刻是否真的还在服务端跑（不是"账本里写着 running"）。 */
  active: boolean;
  lastSequence: number;
  events: TaskProgressEvent[];
}

/**
 * 补发**已经落盘**的规划事件（只读，不重新规划）。
 *
 * 服务端仍在跑时「继续」会被 409 挡住 —— 再发一次就是重复规划、重复读表。
 * 想接回**正在运行**的原任务，唯一正当的入口就是读它已保存的事件。
 */
export async function getWpsPlanningTaskEvents(
  turnId: string,
  afterSequence = 0,
  signal?: AbortSignal
): Promise<WpsPlanningTaskEvents> {
  const query = afterSequence > 0 ? `?afterSequence=${afterSequence}` : "";
  const response = await fetch(
    `${API_BASE_URL}/api/turn/${encodeURIComponent(turnId)}/events${query}`,
    { signal }
  );
  if (!response.ok) {
    throw await wpsPlanningRequestError(response, "读取原任务进度事件失败");
  }
  const value: unknown = await response.json();
  const raw = (value ?? {}) as Record<string, unknown>;
  if (!Array.isArray(raw.events)) {
    throw new ApiRequestError(
      502,
      "WPS_PLANNING_BAD_EVENT_PAGE",
      "原任务进度事件格式不正确",
      false
    );
  }
  const events: TaskProgressEvent[] = [];
  for (const item of raw.events) {
    const parsed = parseTaskProgressEvent(item);
    if (!parsed) {
      // 与流式通道同一条纪律：不认识的帧不静默吞掉。
      throw new ApiRequestError(
        502,
        "WPS_PLANNING_BAD_EVENT_PAGE",
        "原任务进度事件格式不正确",
        false
      );
    }
    events.push(parsed);
  }
  const status = String(raw.status ?? "");
  return {
    turnId,
    status: (["running", "waiting_user", "completed", "failed", "cancelled"].includes(status)
      ? status
      : "failed") as PlanningTaskStatus,
    active: raw.active === true,
    lastSequence: Number(raw.lastSequence ?? 0) || 0,
    events
  };
}

/** 查询规划任务的真实状态。**只查询，不创建任务**。 */
export async function getWpsPlanningTaskStatus(
  turnId: string,
  signal?: AbortSignal
): Promise<WpsPlanningTaskStatus> {
  const response = await fetch(
    `${API_BASE_URL}/api/turn/${encodeURIComponent(turnId)}/status`,
    { signal }
  );
  if (!response.ok) {
    throw await wpsPlanningRequestError(response, "读取规划任务状态失败");
  }
  const value: unknown = await response.json();
  return normalizeWpsPlanningStatus(turnId, value);
}

/**
 * 归并服务端状态到本端视图。
 *
 * 服务端可能只返回 `pause_payload` 的子集（`running` 分支），因此这里对
 * `status` 做白名单收窄，**不接受未知状态** —— 未知状态若被当作
 * `waiting_user`，界面会画出一个服务端并不认可的"继续"按钮。
 */
export function normalizeWpsPlanningStatus(
  turnId: string,
  value: unknown
): WpsPlanningTaskStatus {
  const raw = (value ?? {}) as Record<string, unknown>;
  const allowed: PlanningTaskStatus[] = [
    "running",
    "waiting_user",
    "interrupted",
    "completed",
    "failed",
    "cancelled"
  ];
  const status = allowed.includes(raw.status as PlanningTaskStatus)
    ? (raw.status as PlanningTaskStatus)
    : "failed";
  const revision = Number.isFinite(raw.revision) ? Math.max(1, Math.floor(Number(raw.revision))) : 1;
  const lastSequence = Number.isFinite(raw.lastSequence) ? Math.max(0, Math.floor(Number(raw.lastSequence))) : 0;
  const pauseType =
    raw.pauseType === "budget" ||
    raw.pauseType === "budget_exhausted" ||
    raw.pauseType === "model_error" ||
    raw.pauseType === "interrupted"
      ? raw.pauseType
      : null;
  const rawContinuation = (raw.continuation ?? null) as
    | Record<string, unknown>
    | null;
  const continuation: PlanningContinuationView | null =
    rawContinuation && typeof rawContinuation.available === "boolean"
      ? {
          available: rawContinuation.available,
          code:
            typeof rawContinuation.code === "string"
              ? rawContinuation.code
              : null,
          message:
            typeof rawContinuation.message === "string"
              ? rawContinuation.message
              : "",
          stage:
            rawContinuation.stage === "before_model" ||
            rawContinuation.stage === "after_tool"
              ? rawContinuation.stage
              : null,
          // 缺省按"需要核实"处理：没读懂服务端的意图时多读一次数据是安全方向，
          // 少读一次会让"数据已经变了"的现场直接放行。
          sourceVerificationRequired:
            rawContinuation.sourceVerificationRequired === false ? false : true
        }
      : null;
  return {
    turnId,
    status,
    revision,
    lastSequence,
    result: raw.result as AssistantResponse | undefined,
    recovery: (raw.recovery ?? null) as WpsPlanningTaskStatus["recovery"],
    error: (raw.error ?? null) as WpsPlanningTaskStatus["error"],
    canContinue: typeof raw.canContinue === "boolean" ? raw.canContinue : undefined,
    pauseType,
    reason: typeof raw.reason === "string" ? raw.reason : null,
    continuation,
    sourceFingerprint:
      typeof raw.sourceFingerprint === "string" ? raw.sourceFingerprint : null,
    sourceFingerprintSheets: Array.isArray(raw.sourceFingerprintSheets)
      ? raw.sourceFingerprintSheets.filter(
          (name): name is string => typeof name === "string" && name.trim() !== ""
        )
      : []
  };
}

/**
 * 把服务端状态映射到共享判定模块的保守视图。
 *
 * `hasCheckpoint` 的判据是**现场证据**，不是"服务端允许继续"：
 *
 *  - `pauseType`：服务端给出了暂停成因，说明它确实存了暂停现场；
 *  - `completed` + 记录：任务已跑完并把记录带回来了。这条记录**可能是**完整响应
 *    （`result`），也可能只是**版本化恢复描述**（`recovery`，值被省略过）。
 *    两者都算"服务端确实有一份记录"，但调用方必须自己分辨：只有 `result` 才能
 *    当完整响应回放，`recovery` 只说明结构在、值没留下。
 *
 * **刻意不把 `canContinue === true` 算作现场证据。** 这次是"能不能继续"的
 * 许可，不是"现场在不在"的事实；两者由服务端分开发两个字段，正是为了让
 * 客户端能分辨"我可以让你继续"与"我这儿确实有一份可续跑的现场"。
 * 把许可当现场，一旦服务端只发了许可（重放/半份状态），界面就会画出一个
 * 点下去只会重复规划的"继续"按钮 —— 而这恰恰是本模块要防的那种假成功。
 *
 * 同理**不拿**"服务端返回了 200"当现场存在的证据。
 *
 * 2026-10-19（A0）：服务端现在会给出 `continuation`——它由共享契约与执行账本
 * 算出，是"这份现场能不能继续"的**唯一权威**。有它就只信它；没有（旧载荷 /
 * 终态）才退回上面那套现场证据。`pauseType != null` 只是"暂停过"，一份损坏的、
 * 别的任务的、或已经写过原表的现场同样"暂停过"。
 */
export function planningStatusView(status: WpsPlanningTaskStatus): PlanningStatusView {
  const continuation = status.continuation ?? null;
  const hasCheckpoint = continuation
    ? continuation.available
    : status.pauseType != null ||
      (status.status === "completed" &&
        (status.result != null || status.recovery != null));
  return {
    status: status.status,
    canContinue: status.canContinue,
    pauseType: status.pauseType,
    reason: status.reason,
    errorCode: status.error?.code ?? null,
    errorMessage: status.error?.message ?? null,
    hasCheckpoint,
    continuation
  };
}

/** 由服务端状态推导续跑方案（按钮只从这里给）。 */
export function wpsPlanningResumeDecision(
  status: WpsPlanningTaskStatus
): PlanningResumeDecision {
  return planningResumeDecision(planningStatusView(status));
}

/**
 * 继续规划的请求体。**不含原需求** —— 那会让服务端重新规划。
 *
 * 现场证据是**必填**：省略、为空、读取失败都拿不到继续权。服务端不再接受
 * "没传就跳过比对"的请求，所以这里也不留可选位，缺证据应该在发请求之前停下。
 */
export interface WpsPlanningContinueBody {
  turnId: string;
  expectedRevision: number;
  requestId: string;
  afterSequence: number;
  modelId?: string;
  confirmModelSwitch?: true;
  /** 本轮重新读到的当前工作簿身份；读不到就不发请求。 */
  workbookIdentity: string;
  /** 当前已提交授权的成员，必须与服务端记录完全一致。 */
  authorizedSheetNames: string[];
  /** 来源核实的**三态**，不能合并成一个"有没有"。 */
  sourceVerification: "verified" | "not_collected" | "failed";
  /** 仅 `sourceVerification === "verified"` 时提供。 */
  sourceFingerprint?: string;
}

/**
 * 按服务端检查点继续规划。
 *
 * 只发送身份与游标；`confirmed: true` 是服务端要求的显式确认位。
 * 返回的仍是同一 `AssistantResponseOrPause` 联合类型 —— 与一次性通道共用出口。
 */
export async function continueWpsPlanning(
  body: WpsPlanningContinueBody,
  hooks: {
    onProgress?: (event: TaskProgressEvent) => void;
    signal?: AbortSignal;
    /** 客户端已确认消费到的事件序号；透传给消费链做去重与补发。 */
    clientAfterSequence?: number;
  } = {}
): Promise<AssistantResponseOrPause> {
  const response = await fetch(`${API_BASE_URL}/api/turn/continue/stream`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ...body, confirmed: true }),
    signal: hooks.signal
  });
  if (!response.ok || !response.body) {
    // 续跑只有流式通道；SDK 适配器拿不到流式就等于拿不到工具请求 —— 明确拒绝。
    if (response.status === 404 || response.status === 405 || response.ok) {
      refuseIfStreamingToolDeliveryUnavailable(body.modelId);
    }
    throw await wpsPlanningRequestError(response, "继续规划任务失败");
  }
  return consumeWpsPlanningStream(response, body.turnId, {
    onProgress: hooks.onProgress,
    signal: hooks.signal,
    /*
     * 补发窗口的起点由**客户端游标**决定，而不是 `body.afterSequence`。
     *
     * 两者的区别在断线场景里是致命的：`body.afterSequence` 是本端"请求服务端
     * 从哪补发"的诉求，而消费链的去重闸必须知道"本地**已经**展示到哪"，
     * 否则补发回来的前几条会被当成新事件重复渲染。
     */
    afterSequence: hooks.clientAfterSequence ?? body.afterSequence
  });
}

/** 发起流式规划。连不上时**不退回重发需求**，交给调用方查状态。 */
export async function streamWpsPlanning(
  request: Record<string, unknown>,
  turnId: string,
  hooks: {
    onProgress?: (event: TaskProgressEvent) => void;
    signal?: AbortSignal;
    /** 全新任务：本地从未消费过事件，默认从 0 起算。 */
    afterSequence?: number;
  } = {}
): Promise<AssistantResponseOrPause> {
  const response = await fetch(`${API_BASE_URL}/api/turn/stream`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request),
    signal: hooks.signal
  });
  if (!response.ok || !response.body) {
    /*
     * 流式通道不可用 ≠ 可以退回非流式。
     *
     * SDK 适配器（`planningReadPolicy: "none"`）取数的唯一通道，就是服务端经这条
     * SSE 投递 `bridgeTool` 指针；退回非流式后服务端仍会启动这一轮并等本机回执，
     * 于是干等到超时 —— "无人回执的等待"。这里明确拒绝，把原因说清楚。
     */
    if (response.status === 404 || response.status === 405 || response.ok) {
      refuseIfStreamingToolDeliveryUnavailable(request.modelId);
    }
    throw await wpsPlanningRequestError(response, "本地 AI 服务流式请求失败");
  }
  return consumeWpsPlanningStream(response, turnId, {
    onProgress: hooks.onProgress,
    signal: hooks.signal,
    afterSequence: hooks.afterSequence
  });
}

/**
 * 用一条真实状态响应校准游标。
 *
 * 抽出来是为了让"以服务端为准但不越过客户端游标"这条纪律只有一处实现。
 */
export function calibrateWpsCursor(
  cursor: PlanningCursor,
  status: WpsPlanningTaskStatus
): PlanningCursor {
  return reconcilePlanningCursor(cursor, {
    revision: status.revision,
    lastSequence: status.lastSequence
  });
}

/** 暂停类型是否为"模型连接故障"（可换模型接续）。 */
export function isModelErrorPause(status: WpsPlanningTaskStatus): boolean {
  return (
    status.pauseType === "model_error" ||
    (typeof status.error?.code === "string" &&
      status.error.code.startsWith(MODEL_ERROR_CODE_PREFIX))
  );
}

/**
 * 流式通道不可用时的**闸**：本轮请求需不需要"流式工具投递"。
 *
 * 需要（SDK 适配器）→ 抛与 Excel 端**同一个**错误类型和同一句话，绝不静默退化；
 * 不需要 → 原样返回，让调用方按原来的方式报错（旧后端/普通模型不受影响）。
 */
function refuseIfStreamingToolDeliveryUnavailable(modelId: unknown): void {
  if (typeof modelId === "string" && requiresStreamingToolDelivery(modelId)) {
    throw new StreamingToolDeliveryUnavailableError();
  }
}

/** 请求错误：与 Excel 端 `requestError` 同口径，保留服务端 code 供 UI 归因。 */
async function wpsPlanningRequestError(
  response: Response,
  fallbackMessage: string
): Promise<ApiRequestError> {
  let detail: { code?: string; message?: string; retryable?: boolean } = {};
  try {
    const body: unknown = await response.json();
    if (body && typeof body === "object") {
      const raw = body as Record<string, unknown>;
      const nested = raw.detail && typeof raw.detail === "object"
        ? (raw.detail as Record<string, unknown>)
        : raw;
      detail = {
        code: typeof nested.code === "string" ? nested.code : undefined,
        message: typeof nested.message === "string" ? nested.message : undefined,
        retryable: typeof nested.retryable === "boolean" ? nested.retryable : undefined
      };
    }
  } catch {
    // 响应体不是 JSON：退回纯状态码。
  }
  return new ApiRequestError(
    response.status,
    detail.code ?? "SERVICE_ERROR",
    detail.message ?? fallbackMessage,
    detail.retryable ?? response.status >= 500
  );
}
