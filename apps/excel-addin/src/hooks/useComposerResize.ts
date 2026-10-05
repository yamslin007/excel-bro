import {
  useLayoutEffect,
  useRef,
  type KeyboardEvent as ReactKeyboardEvent,
  type PointerEvent as ReactPointerEvent,
  type RefObject
} from "react";

/**
 * 输入框的「上拉」行为 —— 拖动上方把手调整高度，内容变多时自动长高，
 * 长到上限之后**可以滚动查看**（而不是把多出来的行切掉、也够不着）。
 *
 * 为什么抽成共享 hook：Excel 侧原本把这段逻辑内联在 `App.tsx` 里，WPS 侧于是
 * 自己另写了一套 —— 用了浏览器原生的 `resize: vertical`。在 WPS 的 WebView 里
 * 那个原生把手既拖不动，又因为从共享 CSS 继承来的 `overflow-y: hidden` 看不到
 * 超出高度的那几行，文字被生生切掉。**一种行为只能有一份实现**，否则两端必然
 * 各自漂移，而"漂移"在这里的表现就是"一端能用、另一端不能"。
 *
 * 类名 `.composer-input-region` / `.composer-resize-handle` 定义在两端共用的
 * `apps/excel-addin/src/styles.css`，所以两端挂同一套样式。
 */

/** 内容少时的最小高度。与共享 CSS 里 `.composer textarea` 的 `min-height` 一致。 */
export const COMPOSER_MIN_HEIGHT = 44;
/** 自动长高的上限；再高就不长高了，改由滚动查看。 */
export const COMPOSER_AUTO_GROW_MAX_HEIGHT = 156;
/** 手动拖动的上限：px 与视口比例取小者。 */
export const COMPOSER_DRAG_MAX_HEIGHT = 360;
export const COMPOSER_DRAG_MAX_VIEWPORT_FRACTION = 0.45;
/** ↑/↓ 微调步长。 */
export const COMPOSER_KEYBOARD_STEP = 20;

/**
 * 把手的无障碍说明。两端引用同一句 —— 各自写一份就会漂移，而这句话是键盘与
 * 读屏用户理解"这个把手能干什么"的唯一入口。
 */
export const COMPOSER_RESIZE_LABEL = "上下拖动调整输入框高度，双击恢复自动高度";

/** 拖动上限随视口走：小窗口里不让输入框把整屏吃掉。 */
export function maximumComposerHeight(viewportHeight: number): number {
  return Math.max(
    COMPOSER_MIN_HEIGHT,
    Math.min(
      COMPOSER_DRAG_MAX_HEIGHT,
      Math.floor(viewportHeight * COMPOSER_DRAG_MAX_VIEWPORT_FRACTION)
    )
  );
}

type ResizeSession = {
  pointerId: number;
  startY: number;
  startHeight: number;
};

export type ComposerResizeOptions = {
  inputRef: RefObject<HTMLTextAreaElement | null>;
  /**
   * `null` = 自动高度（跟着内容长高）；数字 = 用户拖动或键盘定下的固定高度。
   * 状态由调用方持有，两端各自决定放在哪（Excel 放在 `useUIState`）。
   */
  height: number | null;
  setHeight: (next: number | null) => void;
  /**
   * **只**用来触发"重新量一次内容高度"。内容变了就要重算，hook 不读它的值。
   */
  value: string;
  /** 与调用方 CSS 的 `min-height` 保持一致，否则拖到底和实际渲染高度会对不上。 */
  minHeight?: number;
  autoGrowMaxHeight?: number;
};

/**
 * 返回挂到把手元素上的 props。高度本身由副作用写到 textarea 的 `style` 上，
 * 调用方不需要自己量 `scrollHeight`。
 */
export function useComposerResize(options: ComposerResizeOptions) {
  const { inputRef, height, setHeight, value } = options;
  const minHeight = options.minHeight ?? COMPOSER_MIN_HEIGHT;
  const autoGrowMaxHeight = options.autoGrowMaxHeight ?? COMPOSER_AUTO_GROW_MAX_HEIGHT;
  const sessionRef = useRef<ResizeSession | null>(null);

  useLayoutEffect(() => {
    const input = inputRef.current;
    if (!input) return;
    const viewportHeight = globalThis.innerHeight ?? 0;
    if (height !== null) {
      input.style.height = `${height}px`;
      /*
       * 关键的一行：固定高度下，内容超出就**允许滚动**。
       * 写成 `hidden` 就会把多出来的行切掉 —— 那正是 WPS 侧文字被截断的原因。
       */
      input.style.overflowY = input.scrollHeight > height ? "auto" : "hidden";
      return;
    }
    input.style.height = "auto";
    const nextHeight = Math.min(input.scrollHeight, autoGrowMaxHeight);
    input.style.height = `${nextHeight}px`;
    input.style.overflowY = input.scrollHeight > autoGrowMaxHeight ? "auto" : "hidden";
    void viewportHeight;
  }, [value, height, inputRef, minHeight, autoGrowMaxHeight]);

  const clamp = (next: number) =>
    Math.max(
      minHeight,
      Math.min(
        maximumComposerHeight(globalThis.innerHeight ?? 0),
        Math.round(next)
      )
    );

  const finishPointer = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (sessionRef.current?.pointerId !== event.pointerId) return;
    sessionRef.current = null;
    if (event.currentTarget.hasPointerCapture?.(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId);
    }
  };

  const handleProps = {
    role: "separator" as const,
    "aria-orientation": "horizontal" as const,
    "aria-label": COMPOSER_RESIZE_LABEL,
    tabIndex: 0,
    onPointerDown(event: ReactPointerEvent<HTMLDivElement>) {
      const input = inputRef.current;
      if (!input) return;
      event.preventDefault();
      event.currentTarget.setPointerCapture?.(event.pointerId);
      sessionRef.current = {
        pointerId: event.pointerId,
        startY: event.clientY,
        startHeight: input.getBoundingClientRect().height
      };
    },
    onPointerMove(event: ReactPointerEvent<HTMLDivElement>) {
      const session = sessionRef.current;
      if (!session || session.pointerId !== event.pointerId) return;
      // 往上拖 = 变高，所以位移取反。
      setHeight(clamp(session.startHeight + session.startY - event.clientY));
    },
    onPointerUp: finishPointer,
    onPointerCancel: finishPointer,
    onDoubleClick() {
      // 恢复自动高度：回到"跟着内容长高"。
      setHeight(null);
    },
    onKeyDown(event: ReactKeyboardEvent<HTMLDivElement>) {
      if (event.key !== "ArrowUp" && event.key !== "ArrowDown") return;
      event.preventDefault();
      const current =
        height ?? inputRef.current?.getBoundingClientRect().height ?? minHeight;
      setHeight(
        clamp(current + (event.key === "ArrowUp" ? COMPOSER_KEYBOARD_STEP : -COMPOSER_KEYBOARD_STEP))
      );
    }
  };

  return { handleProps };
}
