import capabilities from '../../../config/capabilities.json';

export const REVIEW_PANE_LIMITS = capabilities.reviewPaneSizing;

/**
 * 并排审阅的**共享布局计算**（Excel / WPS 同一份）。
 *
 * 唯一的输入是**整个可用工作区宽度**（CSS 逻辑像素）。两端都必须传这个口径：
 *  - Excel：本地桥 `measure` 返回的显示器工作区宽度（原表所在显示器，扣除任务栏）；
 *  - WPS：`Application.Width`（包含两张表的主窗口，非任一表格子窗口）。
 *
 * 绝不能用"右侧原表窗口客户区宽度"或任务窗格 iframe 宽度来算 —— 那样同一份配置在两端
 * 会代表不同含义，AI 窗格要么过窄、要么把原表挤没（这就是此前的真实缺陷）。
 */
export type ReviewLayoutPlan = {
  /** 整个可用工作区宽度。 */
  workAreaWidth: number;
  /** 左侧审阅副本窗口宽度。 */
  copyWidth: number;
  /** 右侧原表窗口宽度（含 AI 窗格）。 */
  originalWidth: number;
  /** AI 任务窗格目标宽度。 */
  aiPaneWidth: number;
  /** 右侧原表**可见表格**宽度 = 原表窗口宽 − AI 窗格宽。 */
  originalSheetWidth: number;
  /** 空间不足以同时满足最小 AI 宽与最小表格宽时为 true（需要用户手动调整）。 */
  degraded: boolean;
};

function assertWorkArea(workAreaWidth: number): void {
  if (!Number.isFinite(workAreaWidth) || workAreaWidth <= 0 || workAreaWidth > REVIEW_PANE_LIMITS.maxHostWidthCssPx) {
    throw new Error('Host width is unavailable');
  }
}

/**
 * AI 任务窗格在给定**工作区宽度**下的目标宽度。
 *
 * 目标 = clamp(工作区 × reviewAiWidthFraction, 最小 AI 宽, 原表窗口的一半)：
 * 最后一项既满足 Office 对任务窗格的半窗口上限，也保证表格至少留一半原表窗口。
 */
export function reviewPaneTargetWidth(workAreaWidth: number): number {
  assertWorkArea(workAreaWidth);
  const half = Math.floor(workAreaWidth / 2);
  const wanted = Math.round(workAreaWidth * REVIEW_PANE_LIMITS.reviewAiWidthFraction);
  return Math.min(half, Math.max(REVIEW_PANE_LIMITS.reviewTaskPaneMinWidthCssPx, wanted));
}

/** 三栏初始布局（副本 / 原表+AI / AI 内部占比）按工作区宽度一次算清。 */
export function reviewLayoutPlan(workAreaWidth: number): ReviewLayoutPlan {
  assertWorkArea(workAreaWidth);
  const copyWidth = Math.round(workAreaWidth * REVIEW_PANE_LIMITS.reviewCopyWidthFraction);
  const originalWidth = workAreaWidth - copyWidth;
  if (copyWidth <= 0 || originalWidth <= 0) throw new Error('Host width is unavailable');
  const wanted = Math.round(workAreaWidth * REVIEW_PANE_LIMITS.reviewAiWidthFraction);
  // AI 窗格：不小于最小 AI 宽，也不吃掉原表窗口的最小表格宽，且不超过原表窗口的一半。
  const aiPaneWidth = Math.min(
    Math.max(REVIEW_PANE_LIMITS.reviewTaskPaneMinWidthCssPx, wanted),
    Math.floor(originalWidth / 2)
  );
  const originalSheetWidth = originalWidth - aiPaneWidth;
  return {
    workAreaWidth,
    copyWidth,
    originalWidth,
    aiPaneWidth,
    originalSheetWidth,
    degraded:
      originalSheetWidth < REVIEW_PANE_LIMITS.reviewSheetMinWidthCssPx ||
      aiPaneWidth < REVIEW_PANE_LIMITS.reviewTaskPaneMinWidthCssPx
  };
}

/** Per-review state, shared by both hosts. A width request is not an acknowledgement. */
export class ReviewPaneSizing {
  private hostWidth?: number;
  private expectedWidth?: number;
  private lastPaneWidth: number;
  private manual = false;
  constructor(paneWidth: number) { this.lastPaneWidth = paneWidth; }
  observe(hostWidth: number, paneWidth: number) {
    const target = reviewPaneTargetWidth(hostWidth);
    if (!Number.isFinite(paneWidth) || paneWidth <= 0) throw new Error('Pane width is unavailable');
    const close = (a: number, b: number) => Math.abs(a - b) <= REVIEW_PANE_LIMITS.taskPaneWidthToleranceCssPx;
    const changedHost = this.hostWidth === undefined || !close(hostWidth, this.hostWidth);
    if (!changedHost && !close(paneWidth, this.lastPaneWidth) &&
        (this.expectedWidth === undefined || !close(paneWidth, this.expectedWidth))) this.manual = true;
    if (this.expectedWidth !== undefined && close(paneWidth, this.expectedWidth)) this.expectedWidth = undefined;
    if (changedHost) this.hostWidth = hostWidth;
    this.lastPaneWidth = paneWidth;
    const requestWidth = !this.manual && changedHost && !close(paneWidth, target) ? target : undefined;
    if (requestWidth !== undefined) this.expectedWidth = requestWidth;
    return { target, manual: this.manual, needsResize: !this.manual && !close(paneWidth, target), requestWidth };
  }
}
