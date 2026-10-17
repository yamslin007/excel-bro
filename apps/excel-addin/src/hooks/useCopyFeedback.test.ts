// @vitest-environment jsdom
import { StrictMode, act as reactAct, createElement, useRef, useEffect } from "react";
import { createRoot } from "react-dom/client";
import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { renderHook, act } from "@testing-library/react";
import { copyTextToClipboard } from "../utils";
import { useCopyFeedback } from "./useCopyFeedback";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT =
  true;

/**
 * 在**真实 StrictMode 根**下挂载 `useCopyFeedback`。
 *
 * 不能用 `@testing-library/react` 的 `renderHook` + wrapper：实测那样挂载
 * **不会**触发 React 的重放（实测日志比真实根少一次 `setup`），于是"挂载期开关"
 * 这类缺陷在测试里永远看不见。`main.tsx` 用的是真实 `createRoot` + `StrictMode`，
 * 所以判据必须在同样的挂载方式下取证。
 */
function mountUnderStrictMode(options: {
  onMessageCopyError?: () => void;
  onFormulaCopyError?: (messageId: string) => void;
} = {}) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  // 挂载期重放的"现场证据"：只记录本组件实例真实经历过的 setup / cleanup 序列。
  const lifecycle: string[] = [];
  const latest: { current: ReturnType<typeof useCopyFeedback> | null } = {
    current: null
  };

  function Harness() {
    const mountedRef = useRef(true);
    useEffect(() => {
      lifecycle.push("setup");
      return () => {
        lifecycle.push("cleanup");
        mountedRef.current = false;
      };
    }, []);
    latest.current = useCopyFeedback(options);
    return createElement(
      "button",
      {
        type: "button",
        className: "probe-copy",
        onClick: () => void latest.current?.copyMessageText("message-1", "hello")
      },
      "copy"
    );
  }

  reactAct(() => {
    root.render(createElement(StrictMode, null, createElement(Harness)));
  });

  return {
    container,
    lifecycle,
    latest,
    copyButton: () => container.querySelector<HTMLButtonElement>(".probe-copy"),
    unmount: () => {
      reactAct(() => root.unmount());
      container.remove();
    }
  };
}

vi.mock("../utils", () => ({
  copyTextToClipboard: vi.fn()
}));

const copyTextToClipboardMock = vi.mocked(copyTextToClipboard);

describe("useCopyFeedback", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.clearAllMocks();
    copyTextToClipboardMock.mockResolvedValue(true);
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it("should initialize with no copied feedback", () => {
    const { result } = renderHook(() => useCopyFeedback({}));

    expect(result.current.copiedMessageId).toBeNull();
    expect(result.current.copiedFunctionPreviewId).toBeNull();
  });

  it("should do nothing when message text is empty", async () => {
    const { result } = renderHook(() => useCopyFeedback({}));

    await act(async () => {
      await result.current.copyMessageText("message-1", "   ");
    });

    expect(copyTextToClipboardMock).not.toHaveBeenCalled();
    expect(result.current.copiedMessageId).toBeNull();
  });

  it("should show and clear message copy feedback", async () => {
    const { result } = renderHook(() => useCopyFeedback({}));

    await act(async () => {
      await result.current.copyMessageText("message-1", "hello");
    });

    expect(copyTextToClipboardMock).toHaveBeenCalledWith("hello");
    expect(result.current.copiedMessageId).toBe("message-1");

    act(() => {
      vi.advanceTimersByTime(1600);
    });

    expect(result.current.copiedMessageId).toBeNull();
  });

  it("should report message copy failure", async () => {
    copyTextToClipboardMock.mockResolvedValue(false);
    const onMessageCopyError = vi.fn();
    const { result } = renderHook(() =>
      useCopyFeedback({ onMessageCopyError })
    );

    await act(async () => {
      await result.current.copyMessageText("message-1", "hello");
    });

    expect(onMessageCopyError).toHaveBeenCalledTimes(1);
    expect(result.current.copiedMessageId).toBeNull();
  });

  it("should show and clear formula copy feedback", async () => {
    const { result } = renderHook(() => useCopyFeedback({}));

    await act(async () => {
      await result.current.copyFunctionFormula("message-2", "=SUM(A1:A2)");
    });

    expect(copyTextToClipboardMock).toHaveBeenCalledWith("=SUM(A1:A2)");
    expect(result.current.copiedFunctionPreviewId).toBe("message-2");

    act(() => {
      vi.advanceTimersByTime(1500);
    });

    expect(result.current.copiedFunctionPreviewId).toBeNull();
  });

  it("should report formula copy failure", async () => {
    copyTextToClipboardMock.mockResolvedValue(false);
    const onFormulaCopyError = vi.fn();
    const { result } = renderHook(() =>
      useCopyFeedback({ onFormulaCopyError })
    );

    await act(async () => {
      await result.current.copyFunctionFormula("message-2", "=SUM(A1:A2)");
    });

    expect(onFormulaCopyError).toHaveBeenCalledWith("message-2");
    expect(result.current.copiedFunctionPreviewId).toBeNull();
  });

  // 回归：Excel 端「点复制按钮要出现『已复制』」这一步在 StrictMode 下曾整条失效。
  // 判据是**反馈状态真的被置上**（hook 是"要不要写状态"的唯一管辖者），
  // 而不是"回调被调了几次"——次数在 StrictMode 下本来就是双份，写次数是错的正确性标准。
  it("should still publish the copied feedback when StrictMode replays effects", async () => {
    const view = mountUnderStrictMode();
    try {
      // 前置闸：必须真的发生过重放，否则这条用例测不到它要测的机制。
      expect(view.lifecycle.filter((step) => step === "setup")).toHaveLength(2);
      expect(view.lifecycle).toEqual(["setup", "cleanup", "setup"]);

      await act(async () => {
        view.copyButton()?.click();
      });

      expect(copyTextToClipboardMock).toHaveBeenCalledWith("hello");
      expect(view.latest.current?.copiedMessageId).toBe("message-1");
    } finally {
      view.unmount();
    }
  });

  it("should still clear the copied feedback after the feedback window under StrictMode", async () => {
    const view = mountUnderStrictMode();
    try {
      await act(async () => {
        view.copyButton()?.click();
      });
      expect(view.latest.current?.copiedMessageId).toBe("message-1");

      act(() => {
        vi.advanceTimersByTime(1600);
      });

      expect(view.latest.current?.copiedMessageId).toBeNull();
    } finally {
      view.unmount();
    }
  });
});
