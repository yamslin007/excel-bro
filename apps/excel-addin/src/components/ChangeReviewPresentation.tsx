import {
  changeReviewUndoFacts,
  type ChangeReviewStatus,
  type ChangeReviewSummary
} from "@excel-bro/core/changeReview";

// 只保留定位、计数与真实回读状态；不渲染网页前后网格。
export const CHANGE_REVIEW_STATUS_TEXT: Record<ChangeReviewStatus, string> = {
  ready: "预计，尚未写入",
  applied: "已写入并回读",
  executed_unverified: "已写入，但独立验收未全部通过",
  readback_failed: "回读失败，未能核对",
  readback_mismatch: "回读与预期不一致",
  not_executed: "未执行"
};

/**
 * 撤销之后的说明。
 *
 * 它**只**说三件可以核实的事：涉及哪些定位、这次执行自己确认改了几格、
 * 撤销已经完成。三件都不能含糊：
 *
 * 1. **必须标阶段。** 这份计数停在产生恢复点的那次执行（中断那一刻）；续跑不会
 *    重新生成审阅，撤销覆盖的却是整个恢复点。写成"本次曾在…改动 2 / 3 个单元格"
 *    会被读成"整次执行只改了 2 格 / 撤销只覆盖 2 格"—— 两件都是假话。
 * 2. **不许把"回读失败"算成改动。** 写后快照缺失时双栏只能拿空矩阵当写后，
 *    "写前非空 → 写后空"看起来就像一次改动；那是没核对上，不是改过了。
 * 3. **不许为了好看凑整。** 没有 `applied` 就如实说 0，不写 3/3。
 *    逐格明细已经释放，这里也不重建任何单元格内容。
 */
export function ChangeReviewUndoNote({ summaries }: { summaries: readonly ChangeReviewSummary[] }) {
  if (!summaries.length) return null;
  const facts = changeReviewUndoFacts(summaries);
  const head = `${facts.interrupted ? "这次执行中断时" : "这次执行"}涉及 ${facts.where} 共 ${facts.totalCells} 个单元格`;
  const details = [`经回读确认改动 ${facts.verifiedCells} 个`];
  if (facts.unverifiedCells > 0) details.push(`写入但未通过回读核对 ${facts.unverifiedCells} 个`);
  if (facts.unattemptedCells > 0) details.push(`未执行 ${facts.unattemptedCells} 个`);
  return (
    <p className="change-review-undo-note" role="status">
      {head}：{details.join("、")}。现已撤销，主表已按恢复点撤回到写前值；逐格明细已按要求释放。
    </p>
  );
}
