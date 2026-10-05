/**
 * 聊天侧栏里的变更摘要卡：只显示定位、计数与真实回读状态。
 * 原生副本负责审阅，不再提供网页对比入口。
 *
 * 摘要本身来自 `changeReviewSummary`，**不含任何单元格内容**，
 * 因此可以安全地渲染在聊天里（也允许进活动日志）。
 */

import {
  canClaimWriteSuccess,
  changeReviewSummary,
  type RangeChangeReview
} from "@excel-bro/core/changeReview";
import { CHANGE_REVIEW_STATUS_TEXT } from "./ChangeReviewPresentation";

export default function ChangeReviewSummaryCard({
  reviews
}: {
  reviews: readonly RangeChangeReview[];
}) {
  if (!reviews.length) return null;
  const summaries = reviews.map(changeReviewSummary);
  const changed = summaries.reduce((total, item) => total + item.changedCount, 0);
  const total = summaries.reduce((sum, item) => sum + item.totalCount, 0);
  const allSuccess = reviews.every((review) => canClaimWriteSuccess(review.status));
  const where = summaries.map((item) => `${item.sheet}!${item.range}`).join("、");

  return (
    <section className="change-review-summary" aria-label="变更摘要">
      <div className="change-review-summary-head">
        <strong>{allSuccess ? "变更摘要" : "变更摘要（含未核对区域）"}</strong>
        <span className="change-review-summary-count">
          改动 {changed} / {total} 个单元格
        </span>
      </div>
      <p className="change-review-summary-where">区域：{where}</p>
      <ul className="change-review-summary-statuses">
        {summaries.map((item) => (
          <li key={`${item.sheet}!${item.range}`}>
            {item.sheet}!{item.range}：{CHANGE_REVIEW_STATUS_TEXT[item.status]}
            {item.truncated && `（共 ${item.totalCount} 格，明细有截断）`}
          </li>
        ))}
      </ul>
    </section>
  );
}
