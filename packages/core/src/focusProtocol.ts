/**
 * 专注窗口（focus window）的跨窗口协议。
 *
 * 两端共用**同一份**信封、校验与状态机：Excel 走 Office DialogApi，WPS 走
 * `Application.CreateWebDialog` / `ShowDialog` + 同源跨 WebView 通道，宿主差异只留在
 * 各端适配层。协议层不接触宿主对象、不发请求、不读工作表。
 *
 * 三条不可让渡的口径：
 *
 * 1. **打开请求 ≠ 窗口就绪。** `ShowDialog` 返回的 `true` 只代表宿主接受了请求，
 *    既不证明窗口已显示，也不证明子页面已加载。只有子窗口发出 `hello`、父窗格回以
 *    带会话标识的快照，才算握手完成。
 * 2. **消息必须可核对。** 每条消息带通道名、协议版本、会话标识、单调序号与工作簿身份；
 *    任一不符即拒收，并记下拒收原因（不是静默丢弃）。
 * 3. **失联必须说出来。** 父窗格刷新/关闭/切换工作簿/通道断开之后，子窗口不能继续把
 *    旧内容当成"当前现场"，必须明确进入 `stale` / `closed`。
 */

/** 通道名。写入每条消息，用来拒绝同源其它页面的广播。 */
export const FOCUS_PROTOCOL_CHANNEL = "excel-bro.focus";
export const FOCUS_PROTOCOL_VERSION = 1;

/** 父窗格心跳间隔；子窗口连续三个心跳周期收不到消息即判失联。 */
export const FOCUS_HEARTBEAT_INTERVAL_MS = 3_000;
export const FOCUS_SYNC_TTL_MS = FOCUS_HEARTBEAT_INTERVAL_MS * 3;
/** 子窗口 `hello` 重发间隔与整体握手超时。 */
export const FOCUS_HELLO_RETRY_MS = 400;
export const FOCUS_HANDSHAKE_TIMEOUT_MS = 10_000;

/**
 * 打开时快照缓存的键前缀。
 *
 * 只在**没有双向通道**（例如 Excel `DialogApi 1.2` 不可用）时兜底展示打开那一刻的内容，
 * 且键里带会话标识与过期时间：它**不是**当前现场的证据，界面必须按"快照"标注。
 * 绝不使用一份全局键，那会让上一个工作簿/会话的内容冒充这一个。
 */
export const FOCUS_CACHE_KEY_PREFIX = "excel-bro.focus.cache.v1.";
/** 打开时快照缓存的有效期：超出即视为过期，不再展示。 */
export const FOCUS_CACHE_TTL_MS = 30 * 60_000;

/** 单次快照的展示限额。专注窗口只做展示，不承担完整历史。 */
export const FOCUS_SNAPSHOT_LIMITS = {
  conversations: 30,
  messagesPerConversation: 100,
  resultRows: 200,
  tools: 200
} as const;

export type FocusCellValue = string | number | boolean | null;

export interface FocusPlanSummary {
  title: string;
  summary: string;
  steps: string[];
}

export interface FocusResultTable {
  title: string;
  headers: string[];
  rows: FocusCellValue[][];
}

export interface FocusMessage {
  id: string;
  role: "assistant" | "user" | "system";
  text: string;
  createdAt: string;
  plan?: FocusPlanSummary;
  result?: FocusResultTable;
}

export interface FocusConversation {
  id: string;
  title: string;
  updatedAt: string;
  messages: FocusMessage[];
}

export interface FocusTool {
  id: string;
  kind: "workflow" | "query";
  name: string;
  description: string;
  category: string;
  steps: number;
  stepLabels: string[];
  dsl?: string;
}

export interface FocusWorkbookRef {
  name: string;
  /** `wps-path:` / `excel-path:` 之类的稳定身份；读不到时 `null`（此时不校验工作簿）。 */
  identity: string | null;
}

export interface FocusSnapshot {
  initialView: "conversation" | "tools";
  activeConversationId: string;
  conversations: FocusConversation[];
  tools: FocusTool[];
  workbook: FocusWorkbookRef;
}

export type FocusMessageKind =
  /** 子窗口 → 父窗格：我已加载，请把当前内容发过来。 */
  | "hello"
  /** 父窗格 → 子窗口：当前会话/工具的展示快照。 */
  | "snapshot"
  /** 父窗格 → 子窗口：我还活着，现场未变。 */
  | "heartbeat"
  /** 子窗口 → 父窗格：用户要求关闭并返回任务窗格。 */
  | "close"
  /** 父窗格 → 子窗口：任务窗格即将关闭/刷新/切换工作簿。 */
  | "closing";

export interface FocusEnvelope {
  channel: typeof FOCUS_PROTOCOL_CHANNEL;
  version: number;
  kind: FocusMessageKind;
  sessionId: string;
  seq: number;
  sentAt: number;
  workbook?: FocusWorkbookRef;
  snapshot?: FocusSnapshot;
  reason?: string;
}

export type FocusRejectReason =
  | "not-json"
  | "not-object"
  | "wrong-channel"
  | "unsupported-version"
  | "unknown-kind"
  | "missing-session"
  | "session-mismatch"
  | "workbook-mismatch"
  | "stale-seq"
  | "expired"
  | "missing-snapshot";

export const FOCUS_REJECT_TEXT: Record<FocusRejectReason, string> = {
  "not-json": "消息不是合法的 JSON。",
  "not-object": "消息结构不完整。",
  "wrong-channel": "消息不属于专注窗口通道。",
  "unsupported-version": "消息协议版本不兼容。",
  "unknown-kind": "未知的消息类型。",
  "missing-session": "消息缺少窗口会话标识。",
  "session-mismatch": "消息属于另一个专注窗口会话。",
  "workbook-mismatch": "消息来自另一个工作簿。",
  "stale-seq": "消息顺序已过期。",
  expired: "消息已超过有效期。",
  "missing-snapshot": "快照缺少对话或工具内容。"
};

export interface FocusExpectation {
  sessionId: string;
  /** 期望的工作簿身份；`null` 表示"还没有身份可核对"，此时不校验。 */
  workbookIdentity?: string | null;
  /** 已接收的最大序号；新消息必须严格更大。 */
  minSeq?: number;
  now?: number;
  maxAgeMs?: number;
}

export type FocusReadResult =
  | { ok: true; message: FocusEnvelope }
  | { ok: false; reason: FocusRejectReason; detail: string };

const MESSAGE_KINDS: readonly FocusMessageKind[] = [
  "hello",
  "snapshot",
  "heartbeat",
  "close",
  "closing"
];

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function asString(value: unknown): string | null {
  return typeof value === "string" && value.trim() !== "" ? value : null;
}

function asFiniteNumber(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function clampArray<T>(value: unknown, limit: number): T[] {
  return Array.isArray(value) ? (value.slice(0, limit) as T[]) : [];
}

/**
 * 会话标识：32 位十六进制随机串。
 *
 * 它只是**窗口会话**的标识（哪个窗格开了哪个窗口），不是认证凭据；消息还要求工作簿
 * 身份与序号同时匹配。宿主没有 `crypto` 时返回 `null`，调用方必须如实报"无法创建
 * 窗口会话"，而不是退回到可预测的序号。
 */
export function createFocusSessionId(
  cryptoSource: { getRandomValues?: (bytes: Uint8Array) => Uint8Array } | null = typeof globalThis.crypto === "undefined"
    ? null
    : (globalThis.crypto as { getRandomValues?: (bytes: Uint8Array) => Uint8Array })
): string | null {
  const getRandomValues = cryptoSource?.getRandomValues;
  if (typeof getRandomValues !== "function") return null;
  const bytes = new Uint8Array(16);
  try {
    getRandomValues.call(cryptoSource, bytes);
  } catch {
    return null;
  }
  return Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("");
}

export function focusCacheKey(sessionId: string): string {
  return `${FOCUS_CACHE_KEY_PREFIX}${sessionId}`;
}

/** 解析结构，不做身份判断。身份/序号/时效由 `readFocusMessage` 判定。 */
export function parseFocusEnvelope(raw: unknown): FocusEnvelope | null {
  if (!isRecord(raw)) return null;
  if (raw.channel !== FOCUS_PROTOCOL_CHANNEL) return null;
  const version = asFiniteNumber(raw.version);
  const kind = asString(raw.kind);
  const sessionId = asString(raw.sessionId);
  const seq = asFiniteNumber(raw.seq);
  const sentAt = asFiniteNumber(raw.sentAt);
  if (version === null || kind === null || sessionId === null || seq === null || sentAt === null) {
    return null;
  }
  // 版本也在这里挡住：解析成功就必须是当前协议，否则调用方会把"看不懂的消息"当成能展示的内容。
  if (version !== FOCUS_PROTOCOL_VERSION) return null;
  if (!MESSAGE_KINDS.includes(kind as FocusMessageKind)) return null;
  const workbook = isRecord(raw.workbook)
    ? {
        name: typeof raw.workbook.name === "string" ? raw.workbook.name : "",
        identity: typeof raw.workbook.identity === "string" ? raw.workbook.identity : null
      }
    : undefined;
  const snapshot = raw.snapshot === undefined ? undefined : sanitizeFocusSnapshot(raw.snapshot);
  if (raw.snapshot !== undefined && !snapshot) return null;
  return {
    channel: FOCUS_PROTOCOL_CHANNEL,
    version,
    kind: kind as FocusMessageKind,
    sessionId,
    seq,
    sentAt,
    ...(workbook ? { workbook } : {}),
    ...(snapshot ? { snapshot } : {}),
    ...(typeof raw.reason === "string" ? { reason: raw.reason.slice(0, 200) } : {})
  };
}

/**
 * 把任意来源的快照收敛成可展示结构，并按限额截断。
 *
 * 只保留展示所需字段：对话文本、计划摘要、结果表头与前若干行、工具说明。凭据、连接
 * 元数据、工作表原始数据行都不在这个结构里 —— 上游构造快照时就不该放进来。
 */
export function sanitizeFocusSnapshot(raw: unknown): FocusSnapshot | null {
  if (!isRecord(raw)) return null;
  const conversationsRaw = clampArray<unknown>(raw.conversations, FOCUS_SNAPSHOT_LIMITS.conversations);
  const toolsRaw = clampArray<unknown>(raw.tools, FOCUS_SNAPSHOT_LIMITS.tools);
  if (!Array.isArray(raw.conversations) || !Array.isArray(raw.tools)) return null;

  const conversations: FocusConversation[] = [];
  for (const item of conversationsRaw) {
    if (!isRecord(item)) continue;
    const id = asString(item.id);
    if (!id) continue;
    conversations.push({
      id,
      title: typeof item.title === "string" ? item.title : "",
      updatedAt: typeof item.updatedAt === "string" ? item.updatedAt : "",
      messages: clampArray<unknown>(item.messages, FOCUS_SNAPSHOT_LIMITS.messagesPerConversation)
        .map((message) => sanitizeFocusMessage(message))
        .filter((message): message is FocusMessage => message !== null)
    });
  }

  const tools: FocusTool[] = [];
  for (const item of toolsRaw) {
    if (!isRecord(item)) continue;
    const id = asString(item.id);
    if (!id) continue;
    tools.push({
      id,
      kind: item.kind === "query" ? "query" : "workflow",
      name: typeof item.name === "string" ? item.name : "",
      description: typeof item.description === "string" ? item.description : "",
      category: typeof item.category === "string" ? item.category : "",
      steps: asFiniteNumber(item.steps) ?? 0,
      stepLabels: Array.isArray(item.stepLabels)
        ? item.stepLabels.filter((label): label is string => typeof label === "string")
        : [],
      ...(typeof item.dsl === "string" ? { dsl: item.dsl } : {})
    });
  }

  const workbookRaw = isRecord(raw.workbook) ? raw.workbook : null;
  return {
    initialView: raw.initialView === "tools" ? "tools" : "conversation",
    activeConversationId: typeof raw.activeConversationId === "string" ? raw.activeConversationId : "",
    conversations,
    tools,
    workbook: {
      name: workbookRaw && typeof workbookRaw.name === "string" ? workbookRaw.name : "",
      identity: workbookRaw && typeof workbookRaw.identity === "string" ? workbookRaw.identity : null
    }
  };
}

function sanitizeFocusMessage(raw: unknown): FocusMessage | null {
  if (!isRecord(raw)) return null;
  const id = asString(raw.id);
  if (!id) return null;
  const role = raw.role === "user" || raw.role === "system" ? raw.role : "assistant";
  const planRaw = isRecord(raw.plan) ? raw.plan : null;
  const resultRaw = isRecord(raw.result) ? raw.result : null;
  return {
    id,
    role,
    text: typeof raw.text === "string" ? raw.text : "",
    createdAt: typeof raw.createdAt === "string" ? raw.createdAt : "",
    ...(planRaw
      ? {
          plan: {
            title: typeof planRaw.title === "string" ? planRaw.title : "",
            summary: typeof planRaw.summary === "string" ? planRaw.summary : "",
            steps: Array.isArray(planRaw.steps)
              ? planRaw.steps.filter((step): step is string => typeof step === "string")
              : []
          }
        }
      : {}),
    ...(resultRaw
      ? {
          result: {
            title: typeof resultRaw.title === "string" ? resultRaw.title : "",
            headers: Array.isArray(resultRaw.headers)
              ? resultRaw.headers.filter((header): header is string => typeof header === "string")
              : [],
            rows: clampArray<unknown>(resultRaw.rows, FOCUS_SNAPSHOT_LIMITS.resultRows).map((row) =>
              Array.isArray(row)
                ? row.map((cell) =>
                    cell === null || typeof cell === "string" || typeof cell === "boolean" ||
                    (typeof cell === "number" && Number.isFinite(cell))
                      ? cell
                      : null
                  )
                : []
            )
          }
        }
      : {})
  };
}

/**
 * 收到一条原始消息时的**唯一**入口：结构 → 通道 → 版本 → 类型 → 会话 → 工作簿 → 序号 → 时效。
 *
 * 顺序有意义：先把"根本不是我们的消息"挡掉（同源其它页面的广播很常见），再核对身份，
 * 最后才判过期。任何一步失败都返回可展示的原因，调用方据此记拒收，**不得**当成成功。
 */
export function readFocusMessage(raw: unknown, expect: FocusExpectation): FocusReadResult {
  let parsed: unknown = raw;
  if (typeof raw === "string") {
    try {
      parsed = JSON.parse(raw);
    } catch {
      return reject("not-json");
    }
  }
  if (!isRecord(parsed)) return reject("not-object");
  if (parsed.channel !== FOCUS_PROTOCOL_CHANNEL) return reject("wrong-channel");
  const version = asFiniteNumber(parsed.version);
  if (version === null) return reject("not-object");
  if (version !== FOCUS_PROTOCOL_VERSION) return reject("unsupported-version");
  const kind = asString(parsed.kind);
  if (!kind || !MESSAGE_KINDS.includes(kind as FocusMessageKind)) return reject("unknown-kind");
  const sessionId = asString(parsed.sessionId);
  if (!sessionId) return reject("missing-session");
  if (sessionId !== expect.sessionId) return reject("session-mismatch");

  const message = parseFocusEnvelope(parsed);
  if (!message) return reject("not-object");

  const expectedIdentity = expect.workbookIdentity ?? null;
  const incomingIdentity = message.workbook?.identity ?? null;
  // 只有双方都有身份时才比较：宿主读不到身份（未保存工作簿）不等于"另一个工作簿"。
  if (expectedIdentity && incomingIdentity && expectedIdentity !== incomingIdentity) {
    return reject("workbook-mismatch");
  }

  if (message.kind === "snapshot") {
    if (!message.snapshot) return reject("missing-snapshot");
    if (message.snapshot.workbook.identity && expectedIdentity &&
        message.snapshot.workbook.identity !== expectedIdentity) {
      return reject("workbook-mismatch");
    }
  }

  if ((message.kind === "snapshot" || message.kind === "heartbeat") &&
      typeof expect.minSeq === "number" && message.seq <= expect.minSeq) {
    return reject("stale-seq");
  }

  const now = expect.now ?? Date.now();
  const maxAgeMs = expect.maxAgeMs ?? FOCUS_SYNC_TTL_MS;
  if (maxAgeMs > 0 && now - message.sentAt > maxAgeMs) return reject("expired");

  return { ok: true, message };
}

function reject(reason: FocusRejectReason): FocusReadResult {
  return { ok: false, reason, detail: FOCUS_REJECT_TEXT[reason] };
}

/**
 * 子窗口状态。
 *
 * `snapshot` 单独成一态：**没有双向通道**时只能展示打开那一刻的内容，它不是"当前现场"，
 * 界面上必须与 `ready` 区分，不能靠文案含糊过去。
 */
export type FocusChildStatus =
  | "connecting"
  | "ready"
  | "stale"
  | "snapshot"
  | "closed"
  | "failed";

export interface FocusChildState {
  status: FocusChildStatus;
  snapshot: FocusSnapshot | null;
  seq: number;
  lastSyncAt: number | null;
  /** 实际连通的通道名（宿主自报），未连通时 `null`。 */
  channel: string | null;
  reason: string | null;
}

export type FocusChildEvent =
  | { type: "transport"; channel: string }
  | { type: "transport-unavailable"; reason: string }
  | { type: "snapshot"; snapshot: FocusSnapshot; seq: number; at: number }
  | { type: "heartbeat"; seq: number; at: number }
  | { type: "closing"; reason: string }
  | { type: "tick"; at: number; ttlMs?: number }
  | { type: "handshake-timeout" }
  | { type: "cache-only"; snapshot: FocusSnapshot; reason: string };

export function initialFocusChildState(): FocusChildState {
  return { status: "connecting", snapshot: null, seq: -1, lastSyncAt: null, channel: null, reason: null };
}

export function reduceFocusChild(state: FocusChildState, event: FocusChildEvent): FocusChildState {
  // `closed` 是终态：父窗格已经明确说了"我关了/我换了工作簿"，之后任何内容都不再是现场。
  if (state.status === "closed" && event.type !== "tick") return state;

  switch (event.type) {
    case "transport":
      return { ...state, channel: event.channel };
    case "transport-unavailable":
      return { ...state, status: "failed", reason: event.reason };
    case "snapshot":
      // 序号不前进的快照一律不采纳：迟到或重复的同步不得把界面推回旧内容。
      if (event.seq <= state.seq) return state;
      return {
        ...state,
        status: "ready",
        snapshot: event.snapshot,
        seq: event.seq,
        lastSyncAt: event.at,
        reason: null
      };
    case "heartbeat":
      if (event.seq <= state.seq) return state;
      if (!state.snapshot) return { ...state, seq: event.seq, lastSyncAt: event.at };
      return {
        ...state,
        status: "ready",
        seq: event.seq,
        lastSyncAt: event.at,
        reason: null
      };
    case "closing":
      return {
        ...state,
        status: "closed",
        reason: event.reason
      };
    case "tick": {
      if (state.status !== "ready" || state.lastSyncAt === null) return state;
      const ttl = event.ttlMs ?? FOCUS_SYNC_TTL_MS;
      if (event.at - state.lastSyncAt <= ttl) return state;
      return {
        ...state,
        status: "stale",
        reason: "与任务窗格的连接已中断，内容可能已过期。"
      };
    }
    case "handshake-timeout":
      if (state.status !== "connecting") return state;
      return {
        ...state,
        status: "failed",
        reason: "任务窗格没有响应握手，无法确认这是当前工作簿的现场。"
      };
    case "cache-only":
      return {
        ...state,
        status: "snapshot",
        snapshot: event.snapshot,
        reason: event.reason
      };
    default:
      return state;
  }
}

/** 父窗格/子窗口共用的序号分配器。 */
export function nextFocusSeq(previous: number): number {
  return Number.isFinite(previous) ? previous + 1 : 1;
}

/** 构造一条消息信封。父窗格与子窗口都走这一个入口，避免两端各拼一份形状。 */
export function createFocusEnvelope(input: {
  kind: FocusMessageKind;
  sessionId: string;
  seq: number;
  sentAt: number;
  workbook?: FocusWorkbookRef;
  snapshot?: FocusSnapshot;
  reason?: string;
}): FocusEnvelope {
  return {
    channel: FOCUS_PROTOCOL_CHANNEL,
    version: FOCUS_PROTOCOL_VERSION,
    kind: input.kind,
    sessionId: input.sessionId,
    seq: input.seq,
    sentAt: input.sentAt,
    ...(input.workbook ? { workbook: input.workbook } : {}),
    ...(input.snapshot ? { snapshot: input.snapshot } : {}),
    ...(input.reason ? { reason: input.reason.slice(0, 200) } : {})
  };
}
