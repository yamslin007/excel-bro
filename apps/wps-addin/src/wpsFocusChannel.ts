/**
 * WPS 专注窗口的**跨窗口通道**。
 *
 * WPS 没有 Office 的 `dialog.messageChild`，父子窗口是同一个加载项源下的两个
 * WebView。可用的真实通道只有两类，按顺序尝试、如实自报名字：
 *
 * 1. `BroadcastChannel` —— 是否真的能跨两个 WPS WebView 投递，**须实机验证**；
 *    这里只负责"宿主给了就按协议用"，浏览器里能通不等于宿主里能通。
 * 2. 同源 `localStorage` 邮箱 —— 两个 WebView 共享同源存储时轮询收发。它只作
 *    **消息传输**，不是状态快照：每个键按会话隔离、带序号与发送时间，内容仍要走
 *    共享协议的完整校验；宿主不共享存储时自然连不上，界面如实报"通道不可用"。
 *
 * 无论哪条通道，消息都是共享协议信封（`@excel-bro/core/focusProtocol`）：通道名、
 * 版本、会话标识、序号、工作簿身份缺一不可，校验失败的消息一律拒收并记录原因。
 */

import {
  FOCUS_HANDSHAKE_TIMEOUT_MS,
  FOCUS_HEARTBEAT_INTERVAL_MS,
  FOCUS_HELLO_RETRY_MS,
  createFocusEnvelope,
  nextFocusSeq,
  readFocusMessage,
  type FocusChildEvent,
  type FocusRejectReason,
  type FocusSnapshot,
  type FocusWorkbookRef
} from "@excel-bro/core/focusProtocol";

export interface FocusTimers {
  setInterval: (handler: () => void, ms: number) => number;
  clearInterval: (id: number) => void;
  setTimeout: (handler: () => void, ms: number) => number;
  clearTimeout: (id: number) => void;
}

export const defaultFocusTimers: FocusTimers = {
  setInterval: (handler, ms) => window.setInterval(handler, ms),
  clearInterval: (id) => window.clearInterval(id),
  setTimeout: (handler, ms) => window.setTimeout(handler, ms),
  clearTimeout: (id) => window.clearTimeout(id)
};

export interface FocusTransport {
  /** 宿主自报的通道名，专注窗口底部如实展示。 */
  name: string;
  post(body: string): boolean;
  close(): void;
}

const MAILBOX_KEY_PREFIX = "excel-bro.focus.mail.v1.";
const MAILBOX_POLL_MS = 250;
const MAILBOX_KEEP = 50;

function mailboxKey(sessionId: string, role: "parent" | "child"): string {
  return `${MAILBOX_KEY_PREFIX}${sessionId}.${role}`;
}

function broadcastChannelName(sessionId: string): string {
  return `${MAILBOX_KEY_PREFIX}${sessionId}`;
}

/** `BroadcastChannel` 通道。宿主没有或构造抛错时返回 `null`。 */
export function createFocusBroadcastTransport(
  sessionId: string,
  onMessage: (body: string) => void
): FocusTransport | null {
  if (typeof BroadcastChannel !== "function") return null;
  let channel: BroadcastChannel;
  try {
    channel = new BroadcastChannel(broadcastChannelName(sessionId));
  } catch {
    return null;
  }
  channel.onmessage = (event: MessageEvent<unknown>) => {
    if (typeof event.data === "string") onMessage(event.data);
  };
  return {
    name: "BroadcastChannel",
    post(body) {
      try {
        channel.postMessage(body);
        return true;
      } catch {
        return false;
      }
    },
    close() {
      try {
        channel.close();
      } catch {
        // 关闭失败不影响会话结束。
      }
    }
  };
}

type MailboxEntry = { id: number; body: string; sentAt: number };

function readMailbox(key: string): MailboxEntry[] {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as { entries?: unknown };
    if (!Array.isArray(parsed?.entries)) return [];
    return parsed.entries.filter(
      (entry): entry is MailboxEntry =>
        typeof entry === "object" && entry !== null &&
        typeof (entry as MailboxEntry).id === "number" &&
        typeof (entry as MailboxEntry).body === "string" &&
        typeof (entry as MailboxEntry).sentAt === "number"
    );
  } catch {
    return [];
  }
}

function writeMailbox(key: string, entries: MailboxEntry[]): boolean {
  try {
    localStorage.setItem(key, JSON.stringify({ entries: entries.slice(-MAILBOX_KEEP) }));
    return true;
  } catch {
    return false;
  }
}

/**
 * 同源 `localStorage` 邮箱通道。它**只是传输**：两端仍会按共享协议校验每条消息，
 * 不存在"读了存储就算数"。
 *
 * 打开前先做一次真实写入探测 —— 宿主禁止本地存储时直接返回 `null`，
 * 不留一条"看着能连"的死通道。
 */
export function createFocusMailboxTransport(
  sessionId: string,
  role: "parent" | "child",
  onMessage: (body: string) => void,
  timers: FocusTimers = defaultFocusTimers
): FocusTransport | null {
  const ownKey = mailboxKey(sessionId, role);
  const peerKey = mailboxKey(sessionId, role === "parent" ? "child" : "parent");
  try {
    // 探测写：写不动就说明这条通道根本不存在。
    writeMailbox(ownKey, []);
    localStorage.getItem(ownKey);
  } catch {
    return null;
  }
  if (!writeMailbox(ownKey, [])) return null;

  let nextId = 1;
  let lastSeenId = 0;
  const poll = timers.setInterval(() => {
    const entries = readMailbox(peerKey);
    for (const entry of entries) {
      if (entry.id <= lastSeenId) continue;
      lastSeenId = entry.id;
      onMessage(entry.body);
    }
  }, MAILBOX_POLL_MS);

  return {
    name: "localStorage 邮箱",
    post(body) {
      const entries = readMailbox(ownKey);
      const id = Math.max(nextId, ...entries.map((entry) => entry.id), 0) + 1;
      nextId = id + 1;
      return writeMailbox(ownKey, [...entries, { id, body, sentAt: Date.now() }]);
    },
    close() {
      timers.clearInterval(poll);
      try {
        localStorage.removeItem(ownKey);
      } catch {
        // 清不掉也只是留下本会话的孤立键，不影响正确性。
      }
    }
  };
}

/** 按可用性顺序建立通道；两条都没有返回 `null`（如实报"通道不可用"）。 */
export function createFocusTransport(
  sessionId: string,
  role: "parent" | "child",
  onMessage: (body: string) => void,
  timers: FocusTimers = defaultFocusTimers
): FocusTransport | null {
  return (
    createFocusBroadcastTransport(sessionId, onMessage) ??
    createFocusMailboxTransport(sessionId, role, onMessage, timers)
  );
}

/* ── 父窗格会话 ── */

export type WpsFocusParentStatus =
  | { phase: "waiting"; transport: string }
  | { phase: "ready"; transport: string }
  | { phase: "handshake-timeout"; transport: string | null }
  | { phase: "transport-unavailable"; reason: string }
  | { phase: "child-closed" };

export interface WpsFocusParentSession {
  readonly transport: string | null;
  /** 立即把当前快照发一次（窗口已打开后再点菜单、或会话内容变化时）。 */
  sync(): boolean;
  /** 告知子窗口本侧即将关闭/切换。 */
  sendClosing(reason: string): void;
  stop(): void;
}

export interface WpsFocusParentSessionOptions {
  sessionId: string;
  workbook: FocusWorkbookRef;
  /** 取最新快照的函数：内容以任务窗格为准，发送那一刻现取，不缓存旧值。 */
  snapshot(): FocusSnapshot;
  onStatus(status: WpsFocusParentStatus): void;
  onRejected?(reason: FocusRejectReason, detail: string): void;
  now?: () => number;
  timers?: FocusTimers;
  handshakeTimeoutMs?: number;
}

export function startWpsFocusParentSession(
  options: WpsFocusParentSessionOptions
): WpsFocusParentSession {
  const now = options.now ?? (() => Date.now());
  const timers = options.timers ?? defaultFocusTimers;
  let seq = 0;
  let ready = false;
  let stopped = false;
  let heartbeat: number | null = null;
  let handshake: number | null = null;

  const post = (kind: "snapshot" | "heartbeat" | "closing", extra: {
    snapshot?: FocusSnapshot;
    reason?: string;
  } = {}): boolean => {
    if (!transport || stopped) return false;
    seq = nextFocusSeq(seq);
    return transport.post(
      JSON.stringify(
        createFocusEnvelope({
          kind,
          sessionId: options.sessionId,
          seq,
          sentAt: now(),
          workbook: options.workbook,
          snapshot: extra.snapshot,
          reason: extra.reason
        })
      )
    );
  };

  const markReady = () => {
    if (ready || stopped) return;
    ready = true;
    if (heartbeat === null) {
      heartbeat = timers.setInterval(() => {
        post("heartbeat");
      }, FOCUS_HEARTBEAT_INTERVAL_MS);
    }
    if (transport) options.onStatus({ phase: "ready", transport: transport.name });
  };

  const onMessage = (body: string) => {
    if (stopped) return;
    const result = readFocusMessage(body, { sessionId: options.sessionId, now: now() });
    if (!result.ok) {
      options.onRejected?.(result.reason, result.detail);
      return;
    }
    if (result.message.kind === "hello") {
      markReady();
      // 每次 hello 都用**当前**快照回应：子窗口重试 hello 不会拿到旧内容。
      post("snapshot", { snapshot: options.snapshot() });
      return;
    }
    if (result.message.kind === "close") {
      options.onStatus({ phase: "child-closed" });
    }
  };

  const transport = createFocusTransport(options.sessionId, "parent", onMessage, timers);
  if (!transport) {
    options.onStatus({
      phase: "transport-unavailable",
      reason: "当前环境没有可用的跨窗口通道（BroadcastChannel 与同源存储都不可用）。"
    });
    return {
      transport: null,
      sync: () => false,
      sendClosing: () => undefined,
      stop: () => {
        stopped = true;
      }
    };
  }

  options.onStatus({ phase: "waiting", transport: transport.name });
  handshake = timers.setTimeout(() => {
    handshake = null;
    if (!ready && !stopped) {
      options.onStatus({ phase: "handshake-timeout", transport: transport.name });
    }
  }, options.handshakeTimeoutMs ?? FOCUS_HANDSHAKE_TIMEOUT_MS);

  return {
    get transport() {
      return transport.name;
    },
    sync() {
      return post("snapshot", { snapshot: options.snapshot() });
    },
    sendClosing(reason: string) {
      post("closing", { reason });
    },
    stop() {
      if (stopped) return;
      stopped = true;
      if (heartbeat !== null) timers.clearInterval(heartbeat);
      if (handshake !== null) timers.clearTimeout(handshake);
      transport.close();
    }
  };
}

/* ── 子窗口会话 ── */

export interface WpsFocusChildSession {
  readonly transport: string | null;
  /** 发送 close 并停止通道；窗口能否真正关闭由调用方的 `window.close()` 决定。 */
  sendClose(): boolean;
  stop(): void;
}

export interface WpsFocusChildSessionOptions {
  sessionId: string;
  onEvent(event: FocusChildEvent): void;
  onRejected?(reason: FocusRejectReason, detail: string): void;
  now?: () => number;
  timers?: FocusTimers;
  handshakeTimeoutMs?: number;
}

export function startWpsFocusChildSession(
  options: WpsFocusChildSessionOptions
): WpsFocusChildSession {
  const now = options.now ?? (() => Date.now());
  const timers = options.timers ?? defaultFocusTimers;
  let seq = 0;
  let minSeq = -1;
  let workbookIdentity: string | null = null;
  let ready = false;
  let stopped = false;
  let hello: number | null = null;
  let handshake: number | null = null;
  let tick: number | null = null;

  const post = (kind: "hello" | "close"): boolean => {
    if (!transport || stopped) return false;
    seq = nextFocusSeq(seq);
    return transport.post(
      JSON.stringify(
        createFocusEnvelope({ kind, sessionId: options.sessionId, seq, sentAt: now() })
      )
    );
  };

  const markReady = () => {
    if (ready || stopped) return;
    ready = true;
    if (hello !== null) {
      timers.clearInterval(hello);
      hello = null;
    }
    if (handshake !== null) {
      timers.clearTimeout(handshake);
      handshake = null;
    }
  };

  const onMessage = (body: string) => {
    if (stopped) return;
    const result = readFocusMessage(body, {
      sessionId: options.sessionId,
      workbookIdentity,
      minSeq,
      now: now()
    });
    if (!result.ok) {
      options.onRejected?.(result.reason, result.detail);
      return;
    }
    const message = result.message;
    if (message.kind === "snapshot" && message.snapshot) {
      // 第一条通过校验的快照即锁定工作簿身份，之后的消息都按它核对。
      if (!workbookIdentity && message.snapshot.workbook.identity) {
        workbookIdentity = message.snapshot.workbook.identity;
      }
      minSeq = message.seq;
      markReady();
      options.onEvent({
        type: "snapshot",
        snapshot: message.snapshot,
        seq: message.seq,
        at: now()
      });
      return;
    }
    if (message.kind === "heartbeat") {
      minSeq = message.seq;
      options.onEvent({ type: "heartbeat", seq: message.seq, at: now() });
      return;
    }
    if (message.kind === "closing") {
      options.onEvent({
        type: "closing",
        reason: message.reason || "任务窗格已关闭或切换了工作簿。"
      });
    }
  };

  const transport = createFocusTransport(options.sessionId, "child", onMessage, timers);
  if (!transport) {
    options.onEvent({
      type: "transport-unavailable",
      reason: "当前环境没有可用的跨窗口通道（BroadcastChannel 与同源存储都不可用）。"
    });
    return {
      transport: null,
      sendClose: () => false,
      stop: () => {
        stopped = true;
      }
    };
  }

  options.onEvent({ type: "transport", channel: transport.name });
  post("hello");
  hello = timers.setInterval(() => {
    if (!ready) post("hello");
  }, FOCUS_HELLO_RETRY_MS);
  handshake = timers.setTimeout(() => {
    handshake = null;
    if (!ready && !stopped) options.onEvent({ type: "handshake-timeout" });
  }, options.handshakeTimeoutMs ?? FOCUS_HANDSHAKE_TIMEOUT_MS);
  // 失联检测：心跳停够久就把界面标成"内容可能已过期"，不继续冒充当前现场。
  tick = timers.setInterval(() => {
    options.onEvent({ type: "tick", at: now() });
  }, 1_000);

  return {
    get transport() {
      return transport.name;
    },
    sendClose() {
      return post("close");
    },
    stop() {
      if (stopped) return;
      stopped = true;
      if (hello !== null) timers.clearInterval(hello);
      if (handshake !== null) timers.clearTimeout(handshake);
      if (tick !== null) timers.clearInterval(tick);
      transport.close();
    }
  };
}
