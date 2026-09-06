import {
  useCallback,
  useLayoutEffect,
  useRef,
  useState,
  type UIEvent
} from "react";

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

export function useMessageAutoFollow(signal: string) {
  const containerRef = useRef<HTMLElement | null>(null);
  const followingRef = useRef(true);
  const [hasUnseenProgress, setHasUnseenProgress] = useState(false);

  const handleScroll = useCallback((event: UIEvent<HTMLElement>) => {
    const following = isNearScrollBottom(event.currentTarget);
    followingRef.current = following;
    if (following) setHasUnseenProgress(false);
  }, []);

  const scrollToLatest = useCallback(() => {
    const container = containerRef.current;
    if (!container) return;
    followingRef.current = true;
    container.scrollTop = container.scrollHeight;
    setHasUnseenProgress(false);
  }, []);

  useLayoutEffect(() => {
    const container = containerRef.current;
    if (!container) return;
    if (followingRef.current) {
      container.scrollTop = container.scrollHeight;
      setHasUnseenProgress(false);
      return;
    }
    setHasUnseenProgress(true);
  }, [signal]);

  return {
    containerRef,
    hasUnseenProgress,
    handleScroll,
    scrollToLatest
  };
}
