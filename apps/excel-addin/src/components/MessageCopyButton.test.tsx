// @vitest-environment jsdom
import { StrictMode, act, createElement, useEffect, useRef } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import MessageCopyButton from "./MessageCopyButton";
import { useCopyFeedback } from "../hooks/useCopyFeedback";

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

  /**
   * 用户报的缺陷："点击 Excel 里的复制按钮，要出现『已复制』的字样"。
   *
   * 这条判据读的是**界面上真的渲染出来的文本**，与 `copied` 这个 prop 是否被传对无关：
   * 它把真实 `useCopyFeedback` 接在真实按钮上，在**真实 StrictMode 根**下（`main.tsx`
   * 就是这么挂 `<App/>` 的）点按钮。挂载期重放会让只写不读的"挂载期开关"永久失效，
   * 于是"点复制有反应"这件事只有在这种挂载方式下才测得到。
   */
  it("shows 已复制 after a real click when mounted under StrictMode", async () => {
    const writeText = vi.fn().mockResolvedValue(undefined);
    Object.defineProperty(navigator, "clipboard", {
      value: { writeText },
      configurable: true
    });

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    const lifecycle: string[] = [];
    const latest: { current: ReturnType<typeof useCopyFeedback> | null } = {
      current: null
    };
    mounted.push(() => {
      act(() => root.unmount());
      container.remove();
    });

    function Harness() {
      const mountedRef = useRef(true);
      useEffect(() => {
        lifecycle.push("setup");
        return () => {
          lifecycle.push("cleanup");
          mountedRef.current = false;
        };
      }, []);
      latest.current = useCopyFeedback({});
      return createElement(MessageCopyButton, {
        copied: latest.current.copiedMessageId === "message-1",
        onClick: () => void latest.current?.copyMessageText("message-1", "B2:E5 加粗")
      });
    }

    await act(async () => {
      root.render(createElement(StrictMode, null, createElement(Harness)));
    });

    // 前置闸：重放必须真的发生过，否则这条用例测不到它要测的机制（会假绿）。
    expect(lifecycle).toEqual(["setup", "cleanup", "setup"]);

    const button = container.querySelector<HTMLButtonElement>(".message-copy-button");
    expect(button?.textContent).toBe("");

    await act(async () => {
      button?.click();
    });

    expect(writeText).toHaveBeenCalledWith("B2:E5 加粗");
    const clicked = container.querySelector<HTMLButtonElement>(".message-copy-button");
    expect(clicked?.textContent).toBe("已复制");
    expect(clicked?.className).toContain("copied");
    expect(clicked?.title).toBe("已复制");
  });
});
