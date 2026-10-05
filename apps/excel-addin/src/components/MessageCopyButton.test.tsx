// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import MessageCopyButton from "./MessageCopyButton";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const mounted: Array<() => void> = [];

afterEach(() => {
  while (mounted.length) mounted.pop()?.();
});

describe("MessageCopyButton", () => {
  it("shows the same icon-only copy control and accessible label in both hosts", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    const onClick = vi.fn();
    mounted.push(() => {
      act(() => root.unmount());
      container.remove();
    });

    await act(async () => root.render(<MessageCopyButton copied={false} onClick={onClick} />));
    const button = container.querySelector<HTMLButtonElement>(".message-copy-button");
    expect(button?.getAttribute("aria-label")).toBe("复制这条消息");
    expect(button?.title).toBe("复制这条消息");
    expect(button?.querySelector("svg rect")).not.toBeNull();
    expect(button?.textContent).toBe("");

    await act(async () => button?.click());
    expect(onClick).toHaveBeenCalledTimes(1);

    await act(async () => root.render(<MessageCopyButton copied onClick={onClick} />));
    expect(button?.getAttribute("aria-label")).toBe("消息已复制");
    expect(button?.title).toBe("已复制");
    expect(button?.querySelector("svg path")).not.toBeNull();
    expect(button?.textContent).toBe("已复制");
  });
});
