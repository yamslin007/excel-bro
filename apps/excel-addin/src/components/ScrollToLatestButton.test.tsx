// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import ScrollToLatestButton from "./ScrollToLatestButton";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const mounted: Array<() => void> = [];

afterEach(() => {
  while (mounted.length) mounted.pop()?.();
});

async function mountButton(running: boolean, onClick: () => void) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  mounted.push(() => {
    act(() => root.unmount());
    container.remove();
  });
  await act(async () => root.render(<ScrollToLatestButton running={running} onClick={onClick} />));
  const button = container.querySelector<HTMLButtonElement>(".new-progress-button");
  expect(button, "浮标按钮没有渲染出来").not.toBeNull();
  return { container, root, button: button! };
}

describe("ScrollToLatestButton", () => {
  it("runs the three-dot indicator while EB is working", async () => {
    const onClick = vi.fn();
    const { button } = await mountButton(true, onClick);

    // ① 状态标在 DOM 上：宿主接线（busy → running）可以被测试直接读到。
    expect(button.dataset.mode).toBe("running");
    // ② 运行中必须是**三点**，而不是锚。
    expect(button.querySelectorAll(".progress-dots > span")).toHaveLength(3);
    expect(button.querySelector(".progress-anchor-icon")).toBeNull();
    // ③ 无障碍名说的是"执行进度"，与运行态一致。
    expect(button.getAttribute("aria-label")).toBe("查看最新执行进度");
    expect(button.title).toBe("查看最新进度");

    await act(async () => button.click());
    expect(onClick).toHaveBeenCalledTimes(1);
  });

  it("swaps the three dots for a theme-coloured anchor once EB stops", async () => {
    const onClick = vi.fn();
    const { button } = await mountButton(false, onClick);

    expect(button.dataset.mode).toBe("idle");
    // 停下后**不再有**三点：留着它会被读成"还在跑"。
    expect(button.querySelector(".progress-dots")).toBeNull();

    const anchor = button.querySelector<SVGSVGElement>(".progress-anchor-icon");
    expect(anchor, "停下后应当渲染锚图标").not.toBeNull();
    // 实心剪影锚的几何：锚眼环（一条 path，外圆减内圆）、横杆（rect）、
    // 两端球（两个 circle）、锚杆 + 底部尖角、锚臂。缺任何一段都不再是锚。
    expect(anchor!.querySelectorAll("circle")).toHaveLength(2);
    expect(anchor!.querySelectorAll("rect")).toHaveLength(1);
    expect(anchor!.querySelectorAll("path")).toHaveLength(3);
    // 锚眼环必须带 evenodd：这是"中空"的唯一来源，漏了它内孔会被填实、锚眼糊死。
    expect(anchor!.querySelector("path")!.getAttribute("fill-rule")).toBe("evenodd");
    // 图标不承担语义（按钮自己有 aria-label），保持 aria-hidden。
    expect(anchor!.getAttribute("aria-hidden")).toBe("true");

    expect(button.getAttribute("aria-label")).toBe("回到底部查看最新内容");
    expect(button.title).toBe("回到底部");

    await act(async () => button.click());
    expect(onClick).toHaveBeenCalledTimes(1);
  });

  it("switches glyphs in place when the run finishes without remounting the control", async () => {
    const onClick = vi.fn();
    const { container, root, button } = await mountButton(true, onClick);
    const runningGlyph = button.querySelector(".progress-dots");
    expect(runningGlyph).not.toBeNull();

    await act(async () => root.render(<ScrollToLatestButton running={false} onClick={onClick} />));

    // 同一个按钮节点：状态切换不该让浮标"消失再出现"（那会在面板里闪一下）。
    const after = container.querySelector<HTMLButtonElement>(".new-progress-button")!;
    expect(after).toBe(button);
    expect(after.dataset.mode).toBe("idle");
    expect(after.querySelector(".progress-anchor-icon")).not.toBeNull();
    expect(after.querySelector(".progress-dots")).toBeNull();
  });
});
