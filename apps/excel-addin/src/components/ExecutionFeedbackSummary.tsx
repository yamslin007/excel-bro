import {
  compactOperationFeedback,
  type OperationFeedbackRecord
} from "@excel-bro/core/operationFeedback";

/**
 * 挂在**对应执行记录下面**的紧凑操作反馈（Excel / WPS 共用）。
 *
 * "已撤销 / 已放弃并恢复 / 已确认变更"与"展示副本未清理"是两件事：
 * 前者说工作簿的业务结论，后者只说对照用的展示副本还没关掉。副本失败按真实
 * 身份去重后给出数量，逐份原因收在 `<details>` 里，既不假装全部清理完成，
 * 也不把恢复结果说成失败。
 */
export function ExecutionFeedbackSummary({
  records
}: {
  records: readonly OperationFeedbackRecord[];
}) {
  const digest = compactOperationFeedback(records);
  const cleanup = digest.records.find((record) => record.kind === "copy_cleanup_failed");
  const issues = cleanup?.issues ?? [];
  if (digest.records.length === 0) return null;
  return (
    <section className="execution-feedback" aria-label="本次执行的操作反馈"
      data-feedback-count={digest.records.length} data-copy-issue-count={issues.length}>
      <p className="execution-feedback-line">{digest.text}</p>
      {issues.length > 0 && (
        <details className="execution-feedback-details">
          <summary>查看 {issues.length} 份未清理副本的原因</summary>
          <ul>
            {issues.map((issue) => (
              <li key={issue.identity}>
                <strong>{issue.label ?? issue.identity}</strong>
                <span>{issue.reason}</span>
              </li>
            ))}
          </ul>
        </details>
      )}
    </section>
  );
}

export default ExecutionFeedbackSummary;
