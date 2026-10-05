import { useEffect, useState } from 'react';
import capabilities from '../../../../config/capabilities.json';
import { measureLocalExcelWindow } from '../excelWindowLayout';

const limits = capabilities.excelWindowLayout;
export type ReviewPaneScope = { workbookIdentity: string; original: { name: string }; copySheetName: string };

export function reviewPaneTargetWidth(clientWidth: number): number {
  // Office on Windows caps the native pane at half the host client width.
  return Math.min(Math.floor(clientWidth / 2), Math.round(Math.max(limits.reviewTaskPaneMinWidthCssPx,
    clientWidth * limits.reviewTaskPaneWidthFraction)));
}

/** Only this live review's host is measured; never use iframe/screen width as host width. */
export function useReviewPaneRatio(scope?: ReviewPaneScope) {
  const [needsResize, setNeedsResize] = useState(false);
  const identity = scope?.workbookIdentity, sheet = scope?.original.name, nonce = scope?.copySheetName;
  useEffect(() => {
    setNeedsResize(false);
    if (!identity || !sheet || !nonce) return;
    let disposed = false, busy = false, manual = false, stopped = false;
    let hostWidth: number | undefined, lastPaneWidth = window.innerWidth;
    let expectedWidth: number | undefined;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const tolerance = limits.taskPaneWidthToleranceCssPx;
    const close = (a: number, b: number) => Math.abs(a - b) <= tolerance;
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
        const changedHost = hostWidth === undefined || !close(clientWidthCssPx, hostWidth);
        const paneWidth = window.innerWidth;
        // Native host resize and native pane resize are different events. Only a
        // pane-only change not acknowledging our request grants manual override.
        if (!changedHost && !close(paneWidth, lastPaneWidth) &&
          (expectedWidth === undefined || !close(paneWidth, expectedWidth))) {
          manual = true; setNeedsResize(false); return;
        }
        if (expectedWidth !== undefined && close(paneWidth, expectedWidth)) expectedWidth = undefined;
        if (changedHost) hostWidth = clientWidthCssPx;
        lastPaneWidth = paneWidth;
        const target = reviewPaneTargetWidth(clientWidthCssPx);
        setNeedsResize(!close(paneWidth, target)); // setWidth returning is not evidence.
        if (changedHost && !close(paneWidth, target)) {
          expectedWidth = target;
          Office.extensionLifeCycle.taskpane.setWidth(target);
        }
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
