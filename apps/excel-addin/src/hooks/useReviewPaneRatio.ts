import { useEffect, useState } from 'react';
import { REVIEW_PANE_LIMITS as limits, ReviewPaneSizing } from '@excel-bro/core/reviewPaneSizing';
export { reviewPaneTargetWidth } from '@excel-bro/core/reviewPaneSizing';
import { measureLocalExcelWindow } from '../excelWindowLayout';

export type ReviewPaneScope = { workbookIdentity: string; original: { name: string }; copySheetName: string };

/** Only this live review's host is measured; never use iframe/screen width as host width. */
export function useReviewPaneRatio(scope?: ReviewPaneScope) {
  const [needsResize, setNeedsResize] = useState(false);
  const identity = scope?.workbookIdentity, sheet = scope?.original.name, nonce = scope?.copySheetName;
  useEffect(() => {
    setNeedsResize(false);
    if (!identity || !sheet || !nonce) return;
    let disposed = false, busy = false, manual = false, stopped = false;
    const sizing = new ReviewPaneSizing(window.innerWidth);
    let timer: ReturnType<typeof setTimeout> | undefined;
    const schedule = (delay: number) => {
      clearTimeout(timer);
      if (!disposed && !manual && !stopped) timer = setTimeout(() => void probe(), delay);
    };
    async function probe() {
      if (disposed || busy || manual || stopped) return;
      if (document.visibilityState === 'hidden') { schedule(limits.reviewPaneMeasureIntervalMs); return; }
      busy = true;
      try {
        const { clientWidthCssPx } = await measureLocalExcelWindow(identity!, sheet!, nonce!);
        if (disposed) return;
        const decision = sizing.observe(clientWidthCssPx, window.innerWidth);
        manual = decision.manual;
        setNeedsResize(decision.needsResize);
        if (decision.requestWidth !== undefined) Office.extensionLifeCycle.taskpane.setWidth(decision.requestWidth);
      } catch {
        if (!disposed) { stopped = true; setNeedsResize(true); }
      } finally {
        busy = false;
        schedule(limits.reviewPaneMeasureIntervalMs);
      }
    }
    const resized = () => schedule(limits.reviewPaneResizeDebounceMs);
    const visible = () => { if (document.visibilityState === 'visible') resized(); };
    window.addEventListener('resize', resized);
    window.addEventListener('focus', resized);
    document.addEventListener('visibilitychange', visible);
    if (typeof Office !== 'undefined') Office.onReady(() => {
      if (disposed) return;
      try {
        if (!Office.context.requirements.isSetSupported('TaskPaneApi', '1.1')) {
          stopped = true; setNeedsResize(true); return;
        }
        void probe();
      } catch { stopped = true; setNeedsResize(true); }
    });
    else setNeedsResize(true);
    return () => {
      disposed = true; clearTimeout(timer);
      window.removeEventListener('resize', resized);
      window.removeEventListener('focus', resized);
      document.removeEventListener('visibilitychange', visible);
    };
  }, [identity, sheet, nonce]);
  return needsResize;
}
