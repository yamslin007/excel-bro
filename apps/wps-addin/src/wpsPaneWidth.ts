import type { WpsApplication } from './wpsHost';
import { TASK_PANE_STORAGE_KEY, readReviewPaneLease } from './wpsReviewPaneLease';

/** User-initiated geometry only; no workbook reads, writes or new task panes. */
export function widenWpsPane(app: WpsApplication, viewportWidth: number): string {
  const id = app.PluginStorage.getItem(TASK_PANE_STORAGE_KEY);
  if (typeof id !== 'string' && typeof id !== 'number') return '无法找到当前任务窗格，宽度未调整。';
  const pane = app.GetTaskPane(id);
  if (!pane || String(pane.ID) !== String(id) || !pane.Visible || pane.DockPosition !== 2 ||
      typeof pane.Width !== 'number' || !Number.isFinite(pane.Width) || pane.Width <= 0 ||
      !Number.isFinite(viewportWidth) || viewportWidth <= 0 ||
      typeof app.Width !== 'number' || !Number.isFinite(app.Width) || app.Width <= 0) {
    return '无法核实当前窗格宽度，宽度未调整。';
  }
  if (readReviewPaneLease(app)) return '双视图正在管理窗格宽度，请先退出双视图后调整。';
  // WPS application width is points; task-pane units are calibrated from the current viewport.
  const hostCssWidth = app.Width * 96 / 72;
  const maxCssWidth = Math.floor(hostCssWidth * 0.7);
  const target = Math.min(viewportWidth + 160, maxCssWidth);
  if (target <= viewportWidth + 2) return '窗格已达到当前窗口的宽度上限。';
  const previous = pane.Width;
  try {
    pane.Width = Math.round(target * previous / viewportWidth);
    if (typeof pane.Width !== 'number' || !Number.isFinite(pane.Width) || pane.Width <= previous) {
      return '宿主未确认宽度调整，请使用窗格边界拖动。';
    }
    return '已请求加宽对话窗格。';
  } catch {
    return '宿主未能调整窗格宽度，请使用窗格边界拖动。';
  }
}
