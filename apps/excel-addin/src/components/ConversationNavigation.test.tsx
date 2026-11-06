// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import ConversationNavigation from "./ConversationNavigation";

afterEach(cleanup);
describe("ConversationNavigation", () => {
  it("indexes only user turns, previews text and navigates by stable identity", () => {
    const onNavigate = vi.fn();
    render(<ConversationNavigation streamRef={{ current: null }} onNavigate={onNavigate} messages={[
      { id: "u1", role: "user", text: "解释移动平均" },
      { id: "a1", role: "assistant", text: "过程与回复" },
      { id: "u2", role: "user", text: "" }
    ]} />);
    expect(screen.getAllByRole("button")).toHaveLength(2);
    fireEvent.focus(screen.getByRole("button", { name: "第 1 轮：解释移动平均" }));
    expect(screen.getByRole("navigation").dataset.expanded).toBe("true");
    expect(onNavigate).not.toHaveBeenCalled();
    fireEvent.click(screen.getByRole("button", { name: "第 2 轮：图片或附件消息" }));
    expect(onNavigate).toHaveBeenCalledExactlyOnceWith("u2");
    expect(screen.getByRole("navigation").dataset.expanded).toBe("true");
    fireEvent.scroll(screen.getByRole("navigation").firstElementChild!);
    expect(onNavigate).toHaveBeenCalledTimes(1);
    fireEvent.mouseLeave(screen.getByRole("navigation"));
    expect(screen.getByRole("navigation").dataset.expanded).toBe("false");
    fireEvent.mouseEnter(screen.getByRole("navigation"));
    fireEvent.keyDown(screen.getByRole("navigation"), { key: "Escape" });
    expect(screen.getByRole("navigation").dataset.expanded).toBe("false");
  });
  it("tracks the visible turn and releases listeners on conversation removal", () => {
    const stream = document.createElement("main");
    let offset = 0;
    ["first", "second"].forEach((id, index) => {
      const row = document.createElement("article"); row.dataset.messageId = id;
      row.getBoundingClientRect = () => ({ top: index * 400 - offset } as DOMRect);
      stream.append(row);
    });
    const ref = { current: stream };
    const { rerender } = render(<ConversationNavigation streamRef={ref} onNavigate={vi.fn()} messages={[
      { id: "first", role: "user", text: "第一题" }, { id: "second", role: "user", text: "第二题" }
    ]} />);
    expect(screen.getByRole("button", { name: /第一题/ }).getAttribute("aria-current")).toBe("location");
    offset = 410; fireEvent.scroll(stream);
    expect(screen.getByRole("button", { name: /第二题/ }).getAttribute("aria-current")).toBe("location");
    rerender(<ConversationNavigation streamRef={ref} onNavigate={vi.fn()} messages={[]} />);
    expect(screen.queryByRole("navigation")).toBeNull();
    fireEvent.scroll(stream);
    expect(screen.queryByRole("button")).toBeNull();
  });
});
