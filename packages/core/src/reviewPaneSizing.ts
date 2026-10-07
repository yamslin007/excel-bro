import capabilities from '../../../config/capabilities.json';

export const REVIEW_PANE_LIMITS = capabilities.reviewPaneSizing;
export function reviewPaneTargetWidth(hostWidth: number): number {
  if (!Number.isFinite(hostWidth) || hostWidth <= 0 || hostWidth > REVIEW_PANE_LIMITS.maxHostWidthCssPx) {
    throw new Error('Host width is unavailable');
  }
  // Product bound leaves at least half the host for sheets; also meets Office's native cap.
  return Math.min(Math.floor(hostWidth / 2), Math.round(Math.max(
    REVIEW_PANE_LIMITS.reviewTaskPaneMinWidthCssPx, hostWidth * REVIEW_PANE_LIMITS.reviewTaskPaneWidthFraction)));
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
