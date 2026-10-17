import {
  useCallback,
  useLayoutEffect,
  useRef,
  useState,
  type UIEvent
} from "react";

// Keep this custom hook in a .tsx module so Vite generates a Fast Refresh
// signature. Otherwise adding a primitive hook can reuse incompatible state.
const DEFAULT_BOTTOM_THRESHOLD = 120;

interface ScrollMetrics {
  scrollHeight: number;
  scrollTop: number;
  clientHeight: number;
}

export function isNearScrollBottom(
  metrics: ScrollMetrics,
  threshold = DEFAULT_BOTTOM_THRESHOLD
): boolean {
  return (
    Math.max(
      0,
      metrics.scrollHeight - metrics.scrollTop - metrics.clientHeight
    ) <= threshold
  );
}

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
  const conversationRef = useRef(conversationId);
  const [awayFromBottom, setAwayFromBottom] = useState(false);

  const handleScroll = useCallback((event: UIEvent<HTMLElement>) => {
    const following = isNearScrollBottom(event.currentTarget);
    followingRef.current = following;
    // 滚出去就放出来、滚回来就收回 —— **不要求**期间有新内容。
    setAwayFromBottom(!following);
  }, []);

  const scrollToLatest = useCallback(() => {
    const container = containerRef.current;
    if (!container) return;
    followingRef.current = true;
    container.scrollTop = container.scrollHeight;
    setAwayFromBottom(false);
  }, []);

  useLayoutEffect(() => {
    // Scroll intent belongs to the displayed conversation, not to the pane.
    if (conversationRef.current !== conversationId) {
      conversationRef.current = conversationId;
      followingRef.current = true;
      setAwayFromBottom(false);
    }
    const container = containerRef.current;
    if (!container) return;
    if (followingRef.current || isNearScrollBottom(container)) {
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
    scrollToLatest
  };
}
