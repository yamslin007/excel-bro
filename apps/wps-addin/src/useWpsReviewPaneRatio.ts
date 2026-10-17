import { useEffect, useState } from 'react';
import { REVIEW_PANE_LIMITS as limits, ReviewPaneSizing } from '@excel-bro/core/reviewPaneSizing';
import type { WpsNativeComparison } from './wpsNativeComparison';
import { getWpsApplication } from './wpsHost';
import { createWpsReviewPaneAdapter } from './wpsReviewPaneSizing';
import { readReviewPaneLease, refreshReviewPaneLease, REVIEW_PANE_LEASE_KEY } from './wpsReviewPaneLease';

export function useWpsReviewPaneRatio(review: WpsNativeComparison | null, rearm = 0) {
  const [needsResize, setNeedsResize] = useState(false);
  useEffect(() => {
    setNeedsResize(false);
    if (!review) return;
    let disposed = false, stopped = false, manual = false;
    let timer: ReturnType<typeof setTimeout> | undefined;
    // UI-only ownership, not an execution credential; older WebViews lack randomUUID.
    const owner = globalThis.crypto?.randomUUID?.() ?? `pane-${Date.now()}-${Math.random().toString(16).slice(2)}`;
    const sizing = new ReviewPaneSizing(window.innerWidth);
    let adapter: ReturnType<typeof createWpsReviewPaneAdapter> | undefined;
    let app: ReturnType<typeof getWpsApplication> | undefined;
    const schedule = (delay: number) => {
      clearTimeout(timer);
      if (!disposed && !stopped) timer = setTimeout(probe, delay);
    };
    function probe() {
      if (disposed || stopped) return;
      try {
        if (document.visibilityState === 'hidden' && !adapter) return;
        app ??= getWpsApplication();
        // Missing capabilities never trigger creation of a pane or writes to a sheet.
        adapter ??= createWpsReviewPaneAdapter(app, review!);
        if (getWpsApplication().Hwnd !== app.Hwnd) throw new Error('Application changed');
        adapter.checkPane();
        refreshReviewPaneLease(app, adapter.paneId, owner);
        if (manual || document.visibilityState === 'hidden') return;
        const { hostWidth } = adapter.measure();
        const decision = sizing.observe(hostWidth, window.innerWidth);
        manual = decision.manual;
        setNeedsResize(decision.needsResize);
        if (decision.requestWidth !== undefined) adapter.resize(decision.requestWidth, window.innerWidth);
      } catch {
        stopped = true;
        if (!disposed) setNeedsResize(true);
      } finally { schedule(limits.reviewPaneMeasureIntervalMs); }
    }
    const resized = () => schedule(limits.reviewPaneResizeDebounceMs);
    const visible = () => { if (document.visibilityState === 'visible') resized(); };
    window.addEventListener('resize', resized);
    window.addEventListener('focus', resized);
    document.addEventListener('visibilitychange', visible);
    probe();
    return () => {
      disposed = true; clearTimeout(timer);
      window.removeEventListener('resize', resized);
      window.removeEventListener('focus', resized);
      document.removeEventListener('visibilitychange', visible);
      if (app && adapter && readReviewPaneLease(app, true)?.owner === owner) {
        try {
          // 只恢复**本次流程自动改过**的宽度。用户中途手动拖过的宽度是他自己的选择，
          // 退出审阅时不得强行覆盖（那会把用户的布局当成我们的临时布局抹掉）。
          if (!manual && getWpsApplication().Hwnd === app.Hwnd) adapter.restore();
        } catch { /* Never resize a replaced/hidden pane. */ }
        try {
          app.PluginStorage.removeItem?.(REVIEW_PANE_LEASE_KEY);
          if (readReviewPaneLease(app, true)?.owner === owner) app.PluginStorage.setItem(REVIEW_PANE_LEASE_KEY, 'null');
        } catch { /* TTL expires a stale marker. */ }
      }
    };
  }, [review, rearm]);
  return needsResize;
}
