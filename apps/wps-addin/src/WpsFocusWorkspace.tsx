import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  initialFocusChildState,
  reduceFocusChild,
  type FocusChildState
} from "@excel-bro/core/focusProtocol";
import FocusWorkspaceView from "../../excel-addin/src/components/FocusWorkspaceView";
import {
  startWpsFocusChildSession,
  type WpsFocusChildSession
} from "./wpsFocusChannel";

/**
 * WPS 专注窗口**页面**（宿主适配层）。
 *
 * 它只通过会话标识加入父窗格发起的窗口会话，并按共享协议校验每一条消息；
 * 不访问 `window.Application`、不读工作表、不发请求。握手没完成就永远是
 * "正在连接/未连接"，不会展示任何可能过期的内容。
 */

function sessionFromLocation(): string {
  try {
    return new URLSearchParams(window.location.search).get("session") ?? "";
  } catch {
    return "";
  }
}

export default function WpsFocusWorkspace() {
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
  const sessionRef = useRef<WpsFocusChildSession | null>(null);

  useEffect(() => {
    if (!sessionId) return;
    const session = startWpsFocusChildSession({
      sessionId,
      onEvent: (event) => setState((current) => reduceFocusChild(current, event))
    });
    sessionRef.current = session;
    return () => {
      session.stop();
      sessionRef.current = null;
    };
  }, [sessionId]);

  const handleClose = useCallback(() => {
    sessionRef.current?.sendClose();
    // 句柄在父窗格手上时父窗格会关；这里尽力自行关闭，关不掉就停在"已结束"。
    try {
      window.close();
    } catch {
      // 宿主不允许脚本关窗时，由用户点窗口自带关闭按钮。
    }
  }, []);

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
