import { afterEach, describe, expect, it, vi } from "vitest";
import {
  createProbeNonce,
  startChildChannelProbe,
  startParentChannelProbe,
  type ChannelProbeStatus
} from "./wpsWideViewChannelPoc";

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
  vi.useRealTimers();
  vi.unstubAllGlobals();
  FakeBroadcastChannel.channels.clear();
});

describe("WPS 宽窗口无数据通信探针", () => {
  it("只生成随机校验码，不从工作簿或本地存储取内容", () => {
    expect(createProbeNonce()).toMatch(/^[0-9a-f]{32}$/);
  });

  it("双向挑战应答成功才报告连通", () => {
    vi.stubGlobal("BroadcastChannel", FakeBroadcastChannel);
    const parent: ChannelProbeStatus[] = [];
    const child: ChannelProbeStatus[] = [];
    const stopParent = startParentChannelProbe("test-nonce", (status) => parent.push(status));
    const stopChild = startChildChannelProbe("test-nonce", (status) => child.push(status));

    expect(parent).toEqual(["waiting", "connected"]);
    expect(child).toEqual(["waiting", "connected"]);
    stopParent();
    stopChild();
    expect(FakeBroadcastChannel.channels.get("excel-bro-wide-poc-test-nonce")?.size).toBe(0);
  });

  it("不同校验码不能互相证明连通，超时如实报告", () => {
    vi.useFakeTimers();
    vi.stubGlobal("BroadcastChannel", FakeBroadcastChannel);
    const parent: ChannelProbeStatus[] = [];
    const child: ChannelProbeStatus[] = [];
    startParentChannelProbe("first", (status) => parent.push(status));
    startChildChannelProbe("second", (status) => child.push(status));
    vi.advanceTimersByTime(15000);
    expect(parent).toEqual(["waiting", "timeout"]);
    expect(child).toEqual(["waiting", "timeout"]);
  });

  it("WebView 没有 BroadcastChannel 时不得宣称成功", () => {
    vi.stubGlobal("BroadcastChannel", undefined);
    const statuses: ChannelProbeStatus[] = [];
    startParentChannelProbe("nonce", (status) => statuses.push(status));
    expect(statuses).toEqual(["unsupported"]);
  });
});
