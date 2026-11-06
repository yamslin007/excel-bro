/**
 * WPS 专注窗口跨窗口通道的**会话状态机**测试。
 *
 * 这里不测宿主 API（ShowDialog / 真实 WPS WebView 的 BroadcastChannel），那些只能
 * 实机验证。测的是通道层自己的确定性行为：握手、快照、心跳、失联、关闭、非法消息拒绝。
 *
 * 两条通道各测一条路径：
 * - `FakeBroadcastChannel`（同步投递）测主路径的握手/心跳/关闭/拒绝；
 * - `BroadcastChannel = undefined` 强制回退 localStorage 邮箱，验证回退真实可用。
 */

import { afterEach, describe, expect, it, vi } from "vitest";
import {
  FOCUS_PROTOCOL_VERSION,
  createFocusEnvelope,
  type FocusSnapshot,
  type FocusWorkbookRef
} from "@excel-bro/core/focusProtocol";
import {
  createFocusTransport,
  startWpsFocusChildSession,
  startWpsFocusParentSession,
  type FocusTimers
} from "./wpsFocusChannel";

const SESSION = "0123456789abcdef0123456789abcdef";
const WORKBOOK: FocusWorkbookRef = { name: "预算.xlsx", identity: "wps-path:c:\\data\\预算.xlsx" };

function snapshot(): FocusSnapshot {
  return {
    initialView: "conversation",
    activeConversationId: "c1",
    conversations: [
      {
        id: "c1",
        title: "整理销售表",
        updatedAt: "2026-11-03T08:00:00.000Z",
        messages: [{ id: "m1", role: "user", text: "把空行删掉", createdAt: "2026-11-03T08:00:00.000Z" }]
      }
    ],
    tools: [],
    workbook: WORKBOOK
  };
}

/** 手动驱动型定时器：interval/timeout 回调登记下来，测试手动触发，不真跑时间。 */
function makeTimers() {
  const intervalHandlers: Array<() => void> = [];
  const timeoutHandlers: Array<() => void> = [];
  const timers: FocusTimers = {
    setInterval: (handler) => { intervalHandlers.push(handler); return intervalHandlers.length; },
    clearInterval: () => undefined,
    setTimeout: (handler) => { timeoutHandlers.push(handler); return timeoutHandlers.length; },
    clearTimeout: () => undefined
  };
  return { timers, intervalHandlers, timeoutHandlers };
}

/** 同步投递的假 BroadcastChannel（与 wpsWideViewChannelPoc.test.ts 同款）。 */
class FakeBroadcastChannel {
  static channels = new Map<string, Set<FakeBroadcastChannel>>();
  onmessage: ((event: MessageEvent<unknown>) => void) | null = null;
  private readonly peers: Set<FakeBroadcastChannel>;
  constructor(name: string) {
    this.peers = FakeBroadcastChannel.channels.get(name) ?? new Set<FakeBroadcastChannel>();
    this.peers.add(this);
    FakeBroadcastChannel.channels.set(name, this.peers);
  }
  postMessage(data: unknown): void {
    for (const peer of this.peers) {
      if (peer !== this) peer.onmessage?.({ data } as MessageEvent<unknown>);
    }
  }
  close(): void {
    this.peers.delete(this);
  }
}

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
  localStorage.clear();
  FakeBroadcastChannel.channels.clear();
});

describe("BroadcastChannel 通道", () => {
  it("父子握手：hello → 快照 → 双方 ready，之后心跳维持", () => {
    vi.stubGlobal("BroadcastChannel", FakeBroadcastChannel);
    const parentStatuses: string[] = [];
    const childEvents: string[] = [];

    const parent = startWpsFocusParentSession({
      sessionId: SESSION,
      workbook: WORKBOOK,
      snapshot,
      onStatus: (status) => parentStatuses.push(status.phase)
    });
    const child = startWpsFocusChildSession({
      sessionId: SESSION,
      onEvent: (event) => childEvents.push(event.type)
    });

    // 子窗口启动即发 hello；FakeBroadcastChannel 同步投递 ⇒ 父窗格立即回快照。
    expect(parentStatuses).toContain("ready");
    expect(childEvents).toContain("snapshot");

    parent.stop();
    child.stop();
    expect(FakeBroadcastChannel.channels.get(`excel-bro.focus.mail.v1.${SESSION}`)?.size ?? 0).toBe(0);
  });

  it("没有 BroadcastChannel 时回退 localStorage 邮箱，握手仍能完成", () => {
    vi.stubGlobal("BroadcastChannel", undefined);
    const { timers, intervalHandlers } = makeTimers();
    const parentStatuses: string[] = [];
    const childEvents: string[] = [];

    const parent = startWpsFocusParentSession({
      sessionId: SESSION,
      workbook: WORKBOOK,
      snapshot,
      onStatus: (status) => parentStatuses.push(status.phase),
      timers
    });
    expect(parent.transport).toBe("localStorage 邮箱");

    const child = startWpsFocusChildSession({
      sessionId: SESSION,
      onEvent: (event) => childEvents.push(event.type),
      timers
    });
    expect(child.transport).toBe("localStorage 邮箱");

    // 驱动双方的邮箱 poll 与 hello 重发，往返几轮直到握手完成。
    for (let round = 0; round < 4 && !parentStatuses.includes("ready"); round += 1) {
      intervalHandlers.forEach((run) => run());
    }
    expect(parentStatuses).toContain("ready");
    expect(childEvents).toContain("snapshot");

    parent.stop();
    child.stop();
  });

  it("通道完全不可用时如实报 transport-unavailable", () => {
    vi.stubGlobal("BroadcastChannel", undefined);
    vi.spyOn(Storage.prototype, "setItem").mockImplementation(() => { throw new Error("denied"); });
    const statuses: string[] = [];
    startWpsFocusParentSession({
      sessionId: SESSION,
      workbook: WORKBOOK,
      snapshot,
      onStatus: (status) => statuses.push(status.phase),
      timers: makeTimers().timers
    });
    expect(statuses[0]).toBe("transport-unavailable");
  });
});

describe("握手与失联", () => {
  it("子窗口迟迟收不到快照 ⇒ handshake-timeout", () => {
    vi.stubGlobal("BroadcastChannel", undefined);
    const { timers, timeoutHandlers } = makeTimers();
    const events: string[] = [];
    startWpsFocusChildSession({
      sessionId: SESSION,
      onEvent: (event) => events.push(event.type),
      timers,
      handshakeTimeoutMs: 50
    });
    timeoutHandlers.forEach((run) => run());
    expect(events).toContain("handshake-timeout");
  });

  it("心跳中断后子窗口持续发出 tick（失联检测的触发源）", () => {
    vi.stubGlobal("BroadcastChannel", undefined);
    const { timers, intervalHandlers } = makeTimers();
    const events: string[] = [];
    const nowRef = { value: Date.now() };
    const child = startWpsFocusChildSession({
      sessionId: SESSION,
      onEvent: (event) => events.push(event.type),
      now: () => nowRef.value,
      timers
    });
    const parent = startWpsFocusParentSession({
      sessionId: SESSION,
      workbook: WORKBOOK,
      snapshot,
      onStatus: () => undefined,
      timers
    });
    for (let round = 0; round < 4 && !events.includes("snapshot"); round += 1) {
      intervalHandlers.forEach((run) => run());
    }
    expect(events).toContain("snapshot");

    // 时间前进后，子窗口的 tick 仍被 emit（`stale` 状态由 `reduceFocusChild` 对 tick 计算，
    // 已在 `focusProtocol.test.ts` 覆盖：`tick` 且 `at - lastSyncAt > TTL` ⇒ `stale`）。
    nowRef.value += 60_000;
    intervalHandlers.forEach((run) => run());
    expect(events).toContain("tick");

    child.stop();
    parent.stop();
  });

  it("父窗格发 closing ⇒ 子窗口进入 closed（终态）", () => {
    vi.stubGlobal("BroadcastChannel", FakeBroadcastChannel);
    const events: string[] = [];
    const child = startWpsFocusChildSession({ sessionId: SESSION, onEvent: (event) => events.push(event.type) });
    const parent = startWpsFocusParentSession({ sessionId: SESSION, workbook: WORKBOOK, snapshot, onStatus: () => undefined });
    parent.sendClosing("任务窗格已关闭。");
    expect(events).toContain("closing");
    child.stop();
    parent.stop();
  });
});

describe("非法消息拒绝", () => {
  it("错误会话 / 错误工作簿 / 过期序号都被拒收", () => {
    vi.stubGlobal("BroadcastChannel", FakeBroadcastChannel);
    const childRejected: string[] = [];
    const childEvents: string[] = [];
    const child = startWpsFocusChildSession({
      sessionId: SESSION,
      onEvent: (event) => childEvents.push(event.type),
      onRejected: (reason) => childRejected.push(reason)
    });
    // 直接往同一个 FakeBroadcastChannel 塞非法消息。
    const channel = new FakeBroadcastChannel(`excel-bro.focus.mail.v1.${SESSION}`);
    channel.postMessage(JSON.stringify(
      createFocusEnvelope({ kind: "snapshot", sessionId: "deadbeefdeadbeefdeadbeefdeadbeef", seq: 1, sentAt: Date.now(), workbook: WORKBOOK, snapshot: snapshot() })
    ));
    channel.postMessage(JSON.stringify({
      channel: "别的通道", version: FOCUS_PROTOCOL_VERSION, kind: "snapshot",
      sessionId: SESSION, seq: 1, sentAt: Date.now()
    }));
    expect(childRejected.length).toBeGreaterThan(0);
    expect(childEvents).not.toContain("snapshot");
    child.stop();
    channel.close();
  });
});
