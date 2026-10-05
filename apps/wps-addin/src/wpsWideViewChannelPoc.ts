/**
 * WPS 宽窗口跨 WebView 通信探针。只交换随机 nonce，不传单元格或聊天内容。
 * BroadcastChannel 是否跨 WPS 的两个浏览器实例可用，必须由真机判定。
 */

export type ChannelProbeStatus = "waiting" | "connected" | "unsupported" | "timeout";

type ProbeMessage = {
  nonce: string;
  type: "child-ready" | "parent-ack" | "child-confirm";
};

const TIMEOUT_MS = 15000;

function channelName(nonce: string): string {
  return `excel-bro-wide-poc-${nonce}`;
}

function isProbeMessage(value: unknown, nonce: string): value is ProbeMessage {
  if (!value || typeof value !== "object") return false;
  const message = value as Partial<ProbeMessage>;
  return message.nonce === nonce && (
    message.type === "child-ready" ||
    message.type === "parent-ack" ||
    message.type === "child-confirm"
  );
}

export function createProbeNonce(): string | null {
  if (typeof globalThis.crypto?.getRandomValues !== "function") return null;
  const bytes = new Uint8Array(16);
  globalThis.crypto.getRandomValues(bytes);
  return Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("");
}

export function startParentChannelProbe(
  nonce: string,
  onStatus: (status: ChannelProbeStatus) => void
): () => void {
  if (typeof BroadcastChannel !== "function") {
    onStatus("unsupported");
    return () => undefined;
  }
  let channel: BroadcastChannel;
  try {
    channel = new BroadcastChannel(channelName(nonce));
  } catch {
    onStatus("unsupported");
    return () => undefined;
  }
  let finished = false;
  const stop = () => {
    if (finished) return;
    finished = true;
    clearTimeout(timer);
    channel.close();
  };
  const timer = setTimeout(() => {
    onStatus("timeout");
    stop();
  }, TIMEOUT_MS);
  channel.onmessage = (event: MessageEvent<unknown>) => {
    if (!isProbeMessage(event.data, nonce)) return;
    if (event.data.type === "child-ready") {
      channel.postMessage({ type: "parent-ack", nonce } satisfies ProbeMessage);
    } else if (event.data.type === "child-confirm") {
      onStatus("connected");
      stop();
    }
  };
  onStatus("waiting");
  return stop;
}

export function startChildChannelProbe(
  nonce: string,
  onStatus: (status: ChannelProbeStatus) => void
): () => void {
  if (typeof BroadcastChannel !== "function") {
    onStatus("unsupported");
    return () => undefined;
  }
  let channel: BroadcastChannel;
  try {
    channel = new BroadcastChannel(channelName(nonce));
  } catch {
    onStatus("unsupported");
    return () => undefined;
  }
  let finished = false;
  const stop = () => {
    if (finished) return;
    finished = true;
    clearInterval(retry);
    clearTimeout(timer);
    channel.close();
  };
  const announce = () => channel.postMessage({ type: "child-ready", nonce } satisfies ProbeMessage);
  const retry = setInterval(announce, 250);
  const timer = setTimeout(() => {
    onStatus("timeout");
    stop();
  }, TIMEOUT_MS);
  channel.onmessage = (event: MessageEvent<unknown>) => {
    if (!isProbeMessage(event.data, nonce) || event.data.type !== "parent-ack") return;
    channel.postMessage({ type: "child-confirm", nonce } satisfies ProbeMessage);
    onStatus("connected");
    stop();
  };
  onStatus("waiting");
  announce();
  return stop;
}
