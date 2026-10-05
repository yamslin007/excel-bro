import type { ChangeReviewStatus, ChangeReviewSummary } from "@excel-bro/core/changeReview";

// 只保留定位、计数与真实回读状态；不渲染网页前后网格。
export const CHANGE_REVIEW_STATUS_TEXT: Record<ChangeReviewStatus, string> = {
  ready: "预计，尚未写入",
  applied: "已写入并回读",
  executed_unverified: "已写入，但独立验收未全部通过",
  readback_failed: "回读失败，未能核对",
  readback_mismatch: "回读与预期不一致",
  not_executed: "未执行"
};

export function ChangeReviewUndoNote({ summaries }: { summaries: readonly ChangeReviewSummary[] }) {
  if (!summaries.length) return null;
  const changed = summaries.reduce((total, item) => total + item.changedCount, 0);
  const total = summaries.reduce((sum, item) => sum + item.totalCount, 0);
  const where = summaries.map((item) => `${item.sheet}!${item.range}`).join("、");
  return (
    <p className="change-review-undo-note" role="status">
      本次曾在 {where} 改动 {changed} / {total} 个单元格；现已撤销，逐格明细已按要求释放。
    </p>
  );
}
