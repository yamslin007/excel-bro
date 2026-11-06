import { useEffect, useState, useRef } from 'react';
import { REVIEW_PANE_LIMITS as limits, ReviewPaneSizing } from '@excel-bro/core/reviewPaneSizing';
export { reviewPaneTargetWidth, reviewLayoutPlan } from '@excel-bro/core/reviewPaneSizing';
import { measureLocalExcelWindow } from '../excelWindowLayout';

export type ReviewPaneScope = { workbookIdentity: string; original: { name: string }; copySheetName: string };

/**
 * 只在**本次审阅**内测量并自适应 AI 窗格宽度；绝不拿 iframe/屏幕宽度充当宿主宽度。
 *
 * 参照是原表所在显示器的**工作区宽度**（`workAreaWidthCssPx`，由本地桥只读测量），
 * 与 WPS 端 `Application.Width` 同口径 —— 这样同一份 `reviewAiWidthFraction`
 * 在两端都表示"占整个可用宽度的比例"，不会变成"占右侧窗口的比例"。
 *
 * `rearm` 变化时（用户主动点「重新排列」）重建状态机：清掉"用户已手动调整"标记，
 * 重新套用默认比例。消息更新、流式输出、回读等普通重渲染**不会**触发重排。
 */
export function useReviewPaneRatio(scope?: ReviewPaneScope, rearm = 0) {
  const currentScope = useRef(scope);
  currentScope.current = scope;
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
        const { workAreaWidthCssPx } = await measureLocalExcelWindow(identity!, currentScope.current!.original.name, nonce!);
        if (disposed) return;
        const decision = sizing.observe(workAreaWidthCssPx, window.innerWidth);
        manual = decision.manual;
        setNeedsResize(decision.needsResize);
        // TaskPaneApi 1.1 的 setWidth 可能被宿主静默忽略：只有实际宽度测量能清除指引。
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
  }, [identity, nonce, rearm]);
  return needsResize;
}
