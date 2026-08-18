import { useCallback, useRef } from "react";

export function useLongPress(
  onTrigger: () => void,
  duration = 2000
) {
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const glowTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const activeRef = useRef(false);

  const clear = useCallback(() => {
    if (timerRef.current) {
      clearTimeout(timerRef.current);
      timerRef.current = null;
    }
    if (glowTimerRef.current) {
      clearTimeout(glowTimerRef.current);
      glowTimerRef.current = null;
    }
    activeRef.current = false;
    document.querySelector(".app-logo")?.classList.remove("glow");
  }, []);

  const start = useCallback(() => {
    if (activeRef.current) return;
    activeRef.current = true;
    const logo = document.querySelector(".app-logo");

    glowTimerRef.current = setTimeout(() => {
      logo?.classList.add("glow");
    }, 1000);

    timerRef.current = setTimeout(() => {
      clear();
      logo?.classList.remove("glow");
      logo?.classList.add("trigger-success");
      onTrigger();
      window.setTimeout(() => {
        logo?.classList.remove("trigger-success");
      }, 400);
    }, duration);
  }, [clear, duration, onTrigger]);

  return {
    onMouseDown: start,
    onMouseUp: clear,
    onMouseLeave: clear,
    onTouchStart: start,
    onTouchEnd: clear,
    onTouchCancel: clear
  };
}
