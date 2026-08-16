// @vitest-environment jsdom
import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { renderHook, act } from "@testing-library/react";
import { copyTextToClipboard } from "../utils";
import { useCopyFeedback } from "./useCopyFeedback";

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
});
