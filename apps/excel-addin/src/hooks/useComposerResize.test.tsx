// @vitest-environment jsdom
import { useRef, useState } from "react";
import { act, fireEvent, render } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  COMPOSER_AUTO_GROW_MAX_HEIGHT,
  COMPOSER_DRAG_MAX_HEIGHT,
  COMPOSER_KEYBOARD_STEP,
  COMPOSER_MIN_HEIGHT,
  COMPOSER_RESIZE_LABEL,
  maximumComposerHeight,
  useComposerResize
} from "./useComposerResize";

type HandleProps = ReturnType<typeof useComposerResize>["handleProps"];
type PointerEventArg = Parameters<HandleProps["onPointerDown"]>[0];
type KeyboardEventArg = Parameters<HandleProps["onKeyDown"]>[0];

/**
 * 这段逻辑原先内联在 `App.tsx` 里、没有测试；WPS 侧照猫画虎写成原生
 * `resize: vertical`，结果在真机 WebView 里拖不动，超出部分还被继承来的
 * `overflow-y: hidden` 切掉。抽成共享 hook 之后行为终于可以被钉住：
 *
 * - **超过上限必须允许滚动**（`overflow-y: auto`），而不是把文字切掉 —— 本次修复点；
 * - 拖到底／拖到顶都要被 clamp，不能把输入框拖出屏幕或压到看不见；
 * - 双击恢复"跟着内容长高"；键盘 ↑/↓ 等价于拖动；
 * - 多指／错配的 pointerId 不得互相串扰。
 *
 * jsdom 不计算布局，所以这里给 textarea 显式定义 `scrollHeight` 与
 * `getBoundingClientRect()` —— 这两个值正是真宿主会提供的输入。
 */
function metrics(
  input: HTMLTextAreaElement,
  values: { scrollHeight: number; height?: number }
) {
  Object.defineProperty(input, "scrollHeight", {
    configurable: true,
    get: () => values.scrollHeight
  });
  Object.defineProperty(input, "getBoundingClientRect", {
    configurable: true,
    value: () => ({
      height: values.height ?? 0,
      width: 0,
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      x: 0,
      y: 0,
      toJSON: () => ({})
    })
  });
}

/** `{...handleProps}` 之后监听器读不回来，所以 harness 顺手把它交出来。 */
let capturedHandleProps: HandleProps | null = null;

function Harness({
  minHeight,
  autoGrowMaxHeight
}: {
  minHeight?: number;
  autoGrowMaxHeight?: number;
} = {}) {
  const inputRef = useRef<HTMLTextAreaElement | null>(null);
  const [height, setHeight] = useState<number | null>(null);
  const [tick, setTick] = useState(0);
  const { handleProps } = useComposerResize({
    inputRef,
    height,
    setHeight,
    // `value` 只是"重新量一次"的触发源；用递增的数字最省事。
    value: String(tick),
    minHeight,
    autoGrowMaxHeight
  });
  capturedHandleProps = handleProps;
  return (
    <div>
      <textarea ref={inputRef} aria-label="composer" />
      <div data-testid="handle" {...handleProps} />
      <span data-testid="height">{height === null ? "auto" : height}</span>
      <button type="button" onClick={() => setTick((current) => current + 1)}>
        sync
      </button>
    </div>
  );
}

/** 手搓指针事件：精确控制 pointerId / clientY，也便于断言指针捕获的调用。 */
function pointer(pointerId: number, clientY: number) {
  const setPointerCapture = vi.fn();
  const releasePointerCapture = vi.fn();
  const hasPointerCapture = vi.fn(() => true);
  const event = {
    pointerId,
    clientY,
    preventDefault: vi.fn(),
    currentTarget: { setPointerCapture, releasePointerCapture, hasPointerCapture }
  } as unknown as PointerEventArg;
  return { event, setPointerCapture, releasePointerCapture, hasPointerCapture };
}

function press(key: string) {
  return { key, preventDefault: vi.fn() } as unknown as KeyboardEventArg;
}

function setup(options: { minHeight?: number; autoGrowMaxHeight?: number } = {}) {
  const view = render(<Harness {...options} />);
  const input = view.getByLabelText("composer") as HTMLTextAreaElement;
  const handle = () => {
    if (!capturedHandleProps) throw new Error("把手 props 尚未捕获");
    return capturedHandleProps;
  };
  return {
    view,
    input,
    handle,
    height: () => view.getByTestId("height").textContent,
    /** 改完 metrics 后点一下 sync，让 hook 的副作用重跑一次。 */
    sync: () => fireEvent.click(view.getByText("sync"))
  };
}

/** 拖动一次：从 clientY `from` 拖到 `to`。返回按下与松开两个事件的探针。 */
function drag(id: number, from: number, to: number) {
  const handle = capturedHandleProps;
  if (!handle) throw new Error("把手 props 尚未捕获");
  const down = pointer(id, from);
  const up = pointer(id, to);
  act(() => {
    handle.onPointerDown(down.event);
    handle.onPointerMove(pointer(id, to).event);
    handle.onPointerUp(up.event);
  });
  return { down, up };
}

afterEach(() => {
  document.body.innerHTML = "";
  capturedHandleProps = null;
});

describe("输入框上拉：自动长高 + 超过上限可滚动", () => {
  it("内容不多时跟着内容长高，且不需要滚动", () => {
    const { input, sync } = setup();
    metrics(input, { scrollHeight: 80 });

    act(() => {
      sync();
    });

    expect(input.style.height).toBe("80px");
    expect(input.style.overflowY).toBe("hidden");
  });

  it("内容超过上限时**停在**上限并改为可滚动 —— 不能把多余的行切掉", () => {
    const { input, sync } = setup();
    metrics(input, { scrollHeight: COMPOSER_AUTO_GROW_MAX_HEIGHT + 200 });

    act(() => {
      sync();
    });

    expect(input.style.height).toBe(`${COMPOSER_AUTO_GROW_MAX_HEIGHT}px`);
    // 这条就是 WPS 侧文字被截断的直接修复点：hidden 会让超出部分够不着。
    expect(input.style.overflowY).toBe("auto");
  });

  it("拖出来的固定高度：装得下不滚动，装不下必须能滚动", () => {
    const { input, height, sync } = setup();
    metrics(input, { scrollHeight: 120, height: 60 });

    // clientY 300 → 100 = 向上 200px；起点高度 60 → 260。
    drag(1, 300, 100);
    expect(height()).toBe("260");
    act(() => {
      sync();
    });
    expect(input.style.height).toBe("260px");
    expect(input.style.overflowY).toBe("hidden");

    // 内容长到 400，超过 260 → 必须能滚动查看。
    metrics(input, { scrollHeight: 400, height: 60 });
    act(() => {
      sync();
    });
    expect(input.style.overflowY).toBe("auto");
  });

  it("拖动被 clamp 在 [最小高度, 视口上限] 之间", () => {
    const { height } = setup();
    const ceiling = maximumComposerHeight(window.innerHeight);

    drag(1, 100, 5100);
    expect(height()).toBe(String(COMPOSER_MIN_HEIGHT));

    drag(2, 5100, 100);
    expect(height()).toBe(String(ceiling));
    expect(ceiling).toBeLessThanOrEqual(COMPOSER_DRAG_MAX_HEIGHT);
  });

  it("minHeight 由调用方给出时，拖动下限跟着它走", () => {
    const { height } = setup({ minHeight: 56 });
    drag(1, 100, 5100);
    expect(height()).toBe("56");
  });

  it("pointerId 对不上的移动事件不得改变高度（防多指串扰）", () => {
    const { input, handle, height } = setup();
    metrics(input, { scrollHeight: 80, height: 44 });
    act(() => {
      handle().onPointerDown(pointer(7, 300).event);
      handle().onPointerMove(pointer(8, 100).event);
    });
    expect(height()).toBe("auto");
  });

  it("松开后释放指针捕获，且再移动不再改高度", () => {
    const { input, handle, height } = setup();
    metrics(input, { scrollHeight: 80, height: COMPOSER_MIN_HEIGHT });

    const down = pointer(3, 300);
    const up = pointer(3, 250);
    act(() => {
      handle().onPointerDown(down.event);
      handle().onPointerMove(pointer(3, 250).event);
    });
    // 44 + (300-250) = 94。
    expect(height()).toBe("94");
    expect(down.setPointerCapture).toHaveBeenCalledWith(3);

    act(() => {
      handle().onPointerUp(up.event);
      handle().onPointerMove(pointer(3, 100).event);
    });
    expect(height()).toBe("94");
    expect(up.releasePointerCapture).toHaveBeenCalledWith(3);
  });

  it("双击回到自动高度", () => {
    const { height, handle } = setup();
    drag(1, 300, 200);
    expect(height()).not.toBe("auto");
    act(() => {
      handle().onDoubleClick();
    });
    expect(height()).toBe("auto");
  });

  it("键盘 ↑/↓ 按步长微调，且同样被 clamp", () => {
    const { input, handle, height } = setup();
    metrics(input, { scrollHeight: 80, height: 80 });

    act(() => {
      handle().onKeyDown(press("ArrowUp"));
    });
    expect(height()).toBe(String(80 + COMPOSER_KEYBOARD_STEP));

    act(() => {
      handle().onKeyDown(press("ArrowDown"));
    });
    expect(height()).toBe("80");
  });

  it("无关按键不拦截默认行为", () => {
    const { handle } = setup();
    const event = press("Enter");
    act(() => {
      handle().onKeyDown(event);
    });
    expect(event.preventDefault).not.toHaveBeenCalled();
  });

  it("还没有 input 时拖动不崩溃，也不进入拖动会话", () => {
    const setHeight = vi.fn();
    let props: HandleProps | null = null;
    function Probe() {
      props = useComposerResize({
        inputRef: { current: null },
        height: null,
        setHeight,
        value: ""
      }).handleProps;
      return null;
    }
    render(<Probe />);
    expect(props).not.toBeNull();
    act(() => {
      (props as unknown as HandleProps).onPointerDown(pointer(1, 200).event);
      (props as unknown as HandleProps).onPointerMove(pointer(1, 100).event);
    });
    expect(setHeight).not.toHaveBeenCalled();
  });

  it("把手的无障碍说明来自共享常量", () => {
    const { view } = setup();
    const element = view.getByTestId("handle");
    expect(element.getAttribute("role")).toBe("separator");
    expect(element.getAttribute("aria-label")).toBe(COMPOSER_RESIZE_LABEL);
    expect(element.getAttribute("aria-orientation")).toBe("horizontal");
    expect(element.getAttribute("tabindex")).toBe("0");
  });
});
