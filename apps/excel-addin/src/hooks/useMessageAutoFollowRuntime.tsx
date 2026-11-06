import {
  followingAfterScroll,
  isNearScrollBottom,
  shouldFollowAfterLayout
} from "@excel-bro/core/scrollFollow";
import {
  useCallback,
  useLayoutEffect,
  useRef,
  useState,
  type UIEvent
} from "react";

// Keep this custom hook in a .tsx module so Vite generates a Fast Refresh
// signature. Otherwise adding a primitive hook can reuse incompatible state.

/**
 * 判据本身住在共享层（`@excel-bro/core/scrollFollow`）：执行过程框内部的滚动
 * 读的是**同一份**，否则"贴底跟随 / 上翻保持"在两个容器里会各长一套。
 * 这里只做转发，保持原有导出名不变。
 */
export { isNearScrollBottom };

/**
 * 消息流的自动跟随 + 「回到底部」的可见性。
 *
 * `awayFromBottom` 的判据是**位置本身**：用户滚离底部（下面还有超过阈值的
 * 内容）就是 true，滚回底部就是 false。
 *
 * 它曾经只在"滚离底部**并且**又来了一次新内容"时才为 true（那时叫
 * `hasUnseenProgress`）。那个口径意味着：用户往上翻看历史、期间没有新内容，
 * 界面上**什么都不会出现** —— 想回到最新那条只能手动滚回去。用户要的是
 * "向上滑就出现一个回到底部的入口"，所以改成看位置；字形再由调用方按
 * "这一轮是否在跑"决定（见 `ScrollToLatestButton`）。
 *
 * 这里只负责"该不该显示"；显示成三点还是锚由调用方决定，所以名字不再
 * 绑定任何一种字形或语义。
 */
export function useMessageAutoFollow(signal: string, conversationId: string) {
  const containerRef = useRef<HTMLElement | null>(null);
  const followingRef = useRef(true);
  const navigationHoldRef = useRef(false);
  const cancelNavigationRef = useRef<(() => void) | null>(null);
  const conversationRef = useRef(conversationId);
  const [awayFromBottom, setAwayFromBottom] = useState(false);

  const handleScroll = useCallback((event: UIEvent<HTMLElement>) => {
    if (cancelNavigationRef.current) return;
    if (navigationHoldRef.current && isNearScrollBottom(event.currentTarget, 1)) {
      navigationHoldRef.current = false;
    }
    const following = !navigationHoldRef.current && followingAfterScroll(event.currentTarget);
    followingRef.current = following;
    // 滚出去就放出来、滚回来就收回 —— **不要求**期间有新内容。
    setAwayFromBottom(!following);
  }, []);

  const scrollToLatest = useCallback(() => {
    const container = containerRef.current;
    if (!container) return;
    cancelNavigationRef.current?.();
    followingRef.current = true;
    navigationHoldRef.current = false;
    container.scrollTop = container.scrollHeight;
    setAwayFromBottom(false);
  }, []);

  const scrollToMessage = useCallback((id: string) => {
    const container = containerRef.current;
    if (!container) return;
    const row = Array.from(container.querySelectorAll<HTMLElement>("[data-message-id]"))
      .find((element) => element.dataset.messageId === id);
    if (!row) return;
    cancelNavigationRef.current?.();
    followingRef.current = false;
    navigationHoldRef.current = true;
    const start = container.scrollTop;
    const target = Math.max(0, Math.min(container.scrollHeight - container.clientHeight,
      start + row.getBoundingClientRect().top - container.getBoundingClientRect().top - 16));
    const behavior = container.style.scrollBehavior;
    container.style.scrollBehavior = "auto";
    // Respect reduced motion; older hosts without matchMedia use immediate navigation.
    if (!window.matchMedia || window.matchMedia("(prefers-reduced-motion: reduce)").matches) {
      container.scrollTop = target;
      container.style.scrollBehavior = behavior;
      setAwayFromBottom(!isNearScrollBottom(container, 1));
      return;
    }
    let frame = 0;
    const began = performance.now();
    const finish = () => {
      cancelAnimationFrame(frame);
      container.style.scrollBehavior = behavior;
      cancelNavigationRef.current = null;
      for (const name of ["wheel", "touchstart", "pointerdown", "keydown"]) container.removeEventListener(name, finish);
      setAwayFromBottom(!isNearScrollBottom(container, 1));
    };
    cancelNavigationRef.current = finish;
    for (const name of ["wheel", "touchstart", "pointerdown", "keydown"]) container.addEventListener(name, finish, { passive: true });
    const tick = (now: number) => {
      const t = Math.min(1, (now - began) / 320);
      const eased = t < .5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2;
      container.scrollTop = start + (target - start) * eased;
      if (t < 1) frame = requestAnimationFrame(tick);
      else finish();
    };
    setAwayFromBottom(true);
    frame = requestAnimationFrame(tick);
  }, []);

  useLayoutEffect(() => () => { cancelNavigationRef.current?.(); }, [conversationId]);

  useLayoutEffect(() => {
    // Scroll intent belongs to the displayed conversation, not to the pane.
    if (conversationRef.current !== conversationId) {
      conversationRef.current = conversationId;
      followingRef.current = true;
      navigationHoldRef.current = false;
      setAwayFromBottom(false);
    }
    const container = containerRef.current;
    if (!container) return;
    if (cancelNavigationRef.current) return;
    if (navigationHoldRef.current && !isNearScrollBottom(container, 1)) {
      setAwayFromBottom(true);
      return;
    }
    navigationHoldRef.current = false;
    // 意图优先、位置兜底：这一句就是"用户上翻后新内容不许抢滚动"的落点。
    if (shouldFollowAfterLayout({ following: followingRef.current, metrics: container })) {
      followingRef.current = true;
      container.scrollTop = container.scrollHeight;
      setAwayFromBottom(false);
      return;
    }
    // 不看 `signal` 也能到这：`signal` 只用来触发这次重算，结论取自**当前**位置。
    setAwayFromBottom(true);
  }, [signal, conversationId]);

  return {
    containerRef,
    awayFromBottom,
    handleScroll,
    scrollToLatest,
    scrollToMessage
  };
}
