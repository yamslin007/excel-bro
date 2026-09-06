// @vitest-environment jsdom
import { act, renderHook } from "@testing-library/react";
import { describe, expect, it } from "vitest";
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

describe("useMessageAutoFollow", () => {
  it("recognizes the configured near-bottom area", () => {
    expect(
      isNearScrollBottom({ scrollHeight: 1000, scrollTop: 590, clientHeight: 300 })
    ).toBe(true);
    expect(
      isNearScrollBottom({ scrollHeight: 1000, scrollTop: 400, clientHeight: 300 })
    ).toBe(false);
  });

  it("keeps new progress visible while the user remains near the bottom", () => {
    const { result, rerender } = renderHook(
      ({ signal }) => useMessageAutoFollow(signal),
      { initialProps: { signal: "1" } }
    );
    const element = scrollElement();
    result.current.containerRef.current = element;

    rerender({ signal: "2" });

    expect(element.scrollTop).toBe(1000);
    expect(result.current.hasUnseenProgress).toBe(false);
  });

  it("does not steal the scroll position after the user scrolls up", () => {
    const { result, rerender } = renderHook(
      ({ signal }) => useMessageAutoFollow(signal),
      { initialProps: { signal: "1" } }
    );
    const element = scrollElement();
    result.current.containerRef.current = element;
    element.scrollTop = 200;
    act(() =>
      result.current.handleScroll({
        currentTarget: element
      } as UIEvent<HTMLElement>)
    );

    rerender({ signal: "2" });

    expect(element.scrollTop).toBe(200);
    expect(result.current.hasUnseenProgress).toBe(true);

    act(() => result.current.scrollToLatest());
    expect(element.scrollTop).toBe(1000);
    expect(result.current.hasUnseenProgress).toBe(false);
  });
});
