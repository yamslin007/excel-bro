// @vitest-environment jsdom
import { act, renderHook } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import type { UIEvent } from "react";
import {
  isNearScrollBottom,
  useMessageAutoFollow
} from "./useMessageAutoFollow";

function scrollElement(): HTMLElement {
  const element = document.createElement("section");
  Object.defineProperties(element, {
    scrollHeight: { configurable: true, value: 1000 },
    clientHeight: { configurable: true, value: 300 }
  });
  element.scrollTop = 700;
  return element;
}

/** 造一次滚动事件；`currentTarget` 就是被滚动的容器本身。 */
function scrollEvent(element: HTMLElement, scrollTop: number): UIEvent<HTMLElement> {
  element.scrollTop = scrollTop;
  return { currentTarget: element } as UIEvent<HTMLElement>;
}

describe("useMessageAutoFollow", () => {
  it("eases between positions, ignores new progress mid-flight and yields to the wheel", () => {
    const callbacks = new Map<number, FrameRequestCallback>();
    let serial = 0;
    vi.stubGlobal("matchMedia", () => ({ matches: false }));
    vi.stubGlobal("requestAnimationFrame", (callback: FrameRequestCallback) => { callbacks.set(++serial, callback); return serial; });
    vi.stubGlobal("cancelAnimationFrame", (id: number) => callbacks.delete(id));
    const now = vi.spyOn(performance, "now").mockReturnValue(0);
    const { result, rerender, unmount } = renderHook(({ signal }) => useMessageAutoFollow(signal, "chat"),
      { initialProps: { signal: "1" } });
    try {
      const element = scrollElement();
      const row = document.createElement("article"); row.dataset.messageId = "old";
      row.getBoundingClientRect = () => ({ top: -684 } as DOMRect);
      element.append(row); result.current.containerRef.current = element;
      act(() => result.current.scrollToMessage("old"));
      expect(element.scrollTop).toBe(700);
      act(() => { const callback = callbacks.get(1)!; callbacks.delete(1); callback(160); });
      expect(element.scrollTop).toBe(350);
      rerender({ signal: "new progress" });
      expect(element.scrollTop).toBe(350);
      act(() => element.dispatchEvent(new Event("wheel")));
      expect(callbacks.size).toBe(0);
      expect(element.style.scrollBehavior).toBe("");
      rerender({ signal: "more progress" });
      expect(element.scrollTop).toBe(350);
    } finally { unmount(); now.mockRestore(); vi.unstubAllGlobals(); }
  });
  it("navigation holds even inside the usual bottom threshold, until explicitly returning", () => {
    const { result, rerender } = renderHook(({ signal }) => useMessageAutoFollow(signal, "chat"),
      { initialProps: { signal: "1" } });
    const element = scrollElement();
    const row = document.createElement("article"); row.dataset.messageId = "old";
    row.getBoundingClientRect = () => ({ top: -84 } as DOMRect);
    element.append(row); result.current.containerRef.current = element;
    act(() => result.current.scrollToMessage("old"));
    expect(element.scrollTop).toBe(600);
    act(() => result.current.handleScroll(scrollEvent(element, 600)));
    rerender({ signal: "new progress" });
    expect(element.scrollTop).toBe(600);
    expect(result.current.awayFromBottom).toBe(true);
    act(() => result.current.scrollToLatest());
    expect(element.scrollTop).toBe(1000);
    expect(result.current.awayFromBottom).toBe(false);
  });
  it("recognizes the configured near-bottom area", () => {
    expect(
      isNearScrollBottom({ scrollHeight: 1000, scrollTop: 590, clientHeight: 300 })
    ).toBe(true);
    expect(
      isNearScrollBottom({ scrollHeight: 1000, scrollTop: 400, clientHeight: 300 })
    ).toBe(false);
  });

  it("keeps the stream pinned to the bottom while the user stays there", () => {
    const { result, rerender } = renderHook(
      ({ signal }) => useMessageAutoFollow(signal, "chat-1"),
      { initialProps: { signal: "1" } }
    );
    const element = scrollElement();
    result.current.containerRef.current = element;

    rerender({ signal: "2" });

    expect(element.scrollTop).toBe(1000);
    expect(result.current.awayFromBottom).toBe(false);
  });

  it("shows the way back as soon as the user scrolls up, even with no new content", () => {
    const { result, rerender } = renderHook(
      ({ signal }) => useMessageAutoFollow(signal, "chat-1"),
      { initialProps: { signal: "1" } }
    );
    const element = scrollElement();
    result.current.containerRef.current = element;

    act(() => result.current.handleScroll(scrollEvent(element, 200)));

    // 这就是用户报的那条：往上滑的**那一刻**就该出现，不看有没有新内容、
    // 也不等下一次重渲染。旧口径要等到 signal 变化才置真，于是"只往上翻历史"
    // 时界面上什么都不出现。
    expect(result.current.awayFromBottom).toBe(true);

    // 之后即便来了新内容，也不许把位置抢回底部；浮标留在原地。
    rerender({ signal: "2" });
    expect(element.scrollTop).toBe(200);
    expect(result.current.awayFromBottom).toBe(true);

    act(() => result.current.scrollToLatest());
    expect(element.scrollTop).toBe(1000);
    expect(result.current.awayFromBottom).toBe(false);
  });

  it.each(["new-chat", "history-chat"])(
    "resets the way back when switching to %s, even without a new signal",
    (conversationId) => {
      const { result, rerender } = renderHook(
        ({ signal, conversationId: id }) => useMessageAutoFollow(signal, id),
        { initialProps: { signal: "1", conversationId: "chat-1" } }
      );
      const element = scrollElement();
      result.current.containerRef.current = element;

      act(() => result.current.handleScroll(scrollEvent(element, 200)));
      expect(result.current.awayFromBottom).toBe(true);

      rerender({ signal: "2", conversationId: "chat-1" });
      expect(result.current.awayFromBottom).toBe(true);

      rerender({ signal: "2", conversationId });
      expect(element.scrollTop).toBe(1000);
      expect(result.current.awayFromBottom).toBe(false);

      rerender({ signal: "3", conversationId });
      expect(result.current.awayFromBottom).toBe(false);
      rerender({ signal: "3", conversationId: "chat-1" });
      expect(result.current.awayFromBottom).toBe(false);
    }
  );

  it("does not offer the way back in a conversation shorter than the viewport", () => {
    const { result, rerender } = renderHook(
      ({ signal, conversationId }) => useMessageAutoFollow(signal, conversationId),
      { initialProps: { signal: "1", conversationId: "chat-1" } }
    );
    const element = scrollElement();
    result.current.containerRef.current = element;

    act(() => result.current.handleScroll(scrollEvent(element, 200)));
    expect(result.current.awayFromBottom).toBe(true);

    // 切到一个短会话：内容连一屏都占不满，没有"下面"可去。
    Object.defineProperty(element, "scrollHeight", { configurable: true, value: 150 });
    element.scrollTop = 0;
    rerender({ signal: "welcome", conversationId: "new-chat" });
    expect(result.current.awayFromBottom).toBe(false);

    // 同一个会话里答案变长：跟随语义把它钉回底部，不亮浮标。
    Object.defineProperty(element, "scrollHeight", { configurable: true, value: 1000 });
    rerender({ signal: "first-answer", conversationId: "new-chat" });
    expect(element.scrollTop).toBe(1000);
    expect(result.current.awayFromBottom).toBe(false);
  });

  it("hides the way back again when the content shrinks to fit the viewport", () => {
    const { result, rerender } = renderHook(
      ({ signal }) => useMessageAutoFollow(signal, "chat-1"),
      { initialProps: { signal: "1" } }
    );
    const element = scrollElement();
    result.current.containerRef.current = element;

    act(() => result.current.handleScroll(scrollEvent(element, 200)));
    rerender({ signal: "2" });
    expect(result.current.awayFromBottom).toBe(true);

    Object.defineProperty(element, "scrollHeight", { configurable: true, value: 150 });
    element.scrollTop = 0;
    rerender({ signal: "collapsed" });
    expect(result.current.awayFromBottom).toBe(false);
  });
});
