import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  FOCUS_HANDSHAKE_TIMEOUT_MS,
  FOCUS_HELLO_RETRY_MS,
  FOCUS_PROTOCOL_VERSION,
  createFocusEnvelope,
  focusCacheKey,
  initialFocusChildState,
  nextFocusSeq,
  readFocusMessage,
  reduceFocusChild,
  sanitizeFocusSnapshot,
  type FocusChildState,
  type FocusSnapshot
} from "@excel-bro/core/focusProtocol";
import FocusWorkspaceView from "./components/FocusWorkspaceView";

/**
 * Excel 专注窗口的**宿主适配层**。
 *
 * 展示交给共享的 `FocusWorkspaceView`；这里只做 Office DialogApi 的三件事：
 * 发 `hello` 握手、按共享协议校验父窗格消息、关闭对话框。任何"打开成功"的结论
 * 都只在收到通过校验的快照之后才成立。
 *
 * 打开时快照缓存（`focusCacheKey(sessionId)`）只在双向通道不可用、握手失败之后作为
 * **明确的快照**展示：它按会话隔离、带版本与过期时间，绝不冒充当前现场。
 */

const OFFICE_CHANNEL_LABEL = "Office DialogApi";

function sessionFromLocation(): string {
  try {
    return new URLSearchParams(window.location.search).get("session") ?? "";
  } catch {
    return "";
  }
}

function readCachedSnapshot(sessionId: string): FocusSnapshot | null {
  try {
    const raw = localStorage.getItem(focusCacheKey(sessionId));
    if (!raw) return null;
    const parsed = JSON.parse(raw) as {
      version?: number;
      sessionId?: string;
      expiresAt?: number;
      snapshot?: unknown;
    };
    if (parsed?.version !== FOCUS_PROTOCOL_VERSION || parsed.sessionId !== sessionId) return null;
    if (typeof parsed.expiresAt !== "number" || parsed.expiresAt < Date.now()) {
      localStorage.removeItem(focusCacheKey(sessionId));
      return null;
    }
    return sanitizeFocusSnapshot(parsed.snapshot);
  } catch {
    return null;
  }
}

function closeContainer() {
  try {
    if (Office.context?.ui?.closeContainer) {
      Office.context.ui.closeContainer();
      return;
    }
  } catch {
    // 老版本没有 closeContainer，退回 window.close()。
  }
  window.close();
}

export default function FocusWorkspace() {
  const sessionId = useMemo(sessionFromLocation, []);
  const [state, setState] = useState<FocusChildState>(() =>
    sessionId
      ? initialFocusChildState()
      : {
          ...initialFocusChildState(),
          status: "failed",
          reason: "缺少窗口会话标识，无法确认这是当前工作簿的现场。"
        }
  );
  const stateRef = useRef(state);
  stateRef.current = state;
  const seqRef = useRef(0);
  // 已收到的父窗格消息序号（独立于自己发消息的 `seqRef`）。父/子各自维护序号空间，
  // 用 `seqRef` 当 `minSeq` 会把父窗格第一条 seq=1 的快照误判成 stale-seq（见 2026-11-03 修复）。
  const receivedSeqRef = useRef(-1);
  const identityRef = useRef<string | null>(null);

  const send = useCallback(
    (kind: "hello" | "close") => {
      try {
        if (!Office.context?.ui?.messageParent) return false;
        const seq = nextFocusSeq(seqRef.current);
        Office.context.ui.messageParent(
          JSON.stringify(
            createFocusEnvelope({ kind, sessionId, seq, sentAt: Date.now() })
          )
        );
        seqRef.current = seq;
        return true;
      } catch {
        return false;
      }
    },
    [sessionId]
  );

  useEffect(() => {
    if (!sessionId) return;
    let disposed = false;
    let hello: number | undefined;
    let handshake: number | undefined;
    let tick: number | undefined;

    const receive = (event: Office.DialogParentMessageReceivedEventArgs) => {
      if (disposed) return;
      const result = readFocusMessage(event.message, {
        sessionId,
        workbookIdentity: identityRef.current,
        minSeq: receivedSeqRef.current,
        now: Date.now()
      });
      if (!result.ok) return;
      const message = result.message;
      if (message.kind === "snapshot" && message.snapshot) {
        // 第一条通过校验的快照即锁定工作簿身份，之后的快照/心跳都按它核对。
        if (!identityRef.current && message.snapshot.workbook.identity) {
          identityRef.current = message.snapshot.workbook.identity;
        }
        receivedSeqRef.current = message.seq;
        if (hello !== undefined) {
          window.clearInterval(hello);
          hello = undefined;
        }
        if (handshake !== undefined) {
          window.clearTimeout(handshake);
          handshake = undefined;
        }
        setState((current) =>
          reduceFocusChild(current, {
            type: "snapshot",
            snapshot: message.snapshot!,
            seq: message.seq,
            at: Date.now()
          })
        );
        return;
      }
      if (message.kind === "heartbeat") {
        receivedSeqRef.current = message.seq;
        setState((current) =>
          reduceFocusChild(current, { type: "heartbeat", seq: message.seq, at: Date.now() })
        );
        return;
      }
      if (message.kind === "closing") {
        setState((current) =>
          reduceFocusChild(current, {
            type: "closing",
            reason: message.reason || "任务窗格已关闭或切换了工作簿。"
          })
        );
      }
    };

    const startTicking = () => {
      tick = window.setInterval(() => {
        if (disposed) return;
        setState((current) => reduceFocusChild(current, { type: "tick", at: Date.now() }));
      }, 1_000);
    };

    const failWithoutChannel = (reason: string) => {
      setState((current) => reduceFocusChild(current, { type: "transport-unavailable", reason }));
    };

    const begin = () => {
      if (disposed) return;
      if (!Office.context?.ui?.addHandlerAsync || !Office.context?.ui?.messageParent) {
        failWithoutChannel("当前窗口没有 Office 对话框通道，无法与任务窗格同步。");
        return;
      }
      Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        receive,
        (result) => {
          if (disposed) return;
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            failWithoutChannel("无法监听任务窗格消息，专注窗口未建立同步。");
            return;
          }
          setState((current) =>
            reduceFocusChild(current, { type: "transport", channel: OFFICE_CHANNEL_LABEL })
          );
          send("hello");
          hello = window.setInterval(() => send("hello"), FOCUS_HELLO_RETRY_MS);
          handshake = window.setTimeout(() => {
            if (disposed) return;
            if (hello !== undefined) {
              window.clearInterval(hello);
              hello = undefined;
            }
            if (stateRef.current.status !== "connecting") return;
            // 双向通道不可用（DialogApi 1.2 以下）时，父窗格只能留下会话隔离的打开时快照。
            const cached = readCachedSnapshot(sessionId);
            setState((current) =>
              cached
                ? reduceFocusChild(current, {
                    type: "cache-only",
                    snapshot: cached,
                    reason: "双向通道不可用，只显示打开窗口时的快照。"
                  })
                : reduceFocusChild(current, { type: "handshake-timeout" })
            );
          }, FOCUS_HANDSHAKE_TIMEOUT_MS);
          startTicking();
        }
      );
    };

    if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
      Office.onReady(() => begin());
    } else {
      failWithoutChannel("当前窗口没有 Office 运行环境。");
    }

    return () => {
      disposed = true;
      if (hello !== undefined) window.clearInterval(hello);
      if (handshake !== undefined) window.clearTimeout(handshake);
      if (tick !== undefined) window.clearInterval(tick);
      // 对话框关闭时宿主会一并回收消息处理器（老版本没有可移除的 API）；
      // `disposed` 与计时器清理已保证回收之后不再处理任何消息。
    };
  }, [send, sessionId]);

  const handleClose = useCallback(() => {
    send("close");
    closeContainer();
  }, [send]);

  return (
    <FocusWorkspaceView
      snapshot={state.snapshot}
      status={state.status}
      statusDetail={state.reason}
      channelLabel={state.channel}
      onClose={handleClose}
    />
  );
}
