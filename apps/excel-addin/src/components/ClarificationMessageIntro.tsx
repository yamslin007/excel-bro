/** Shared clarification typography and state wording; host actions remain outside. */
export function ClarificationMessageIntro({ status, summary, question, reason, scopeLabel }: {
  status: string; summary: string; question: string; reason: string; scopeLabel: string;
}) {
  const label = status === "pending" ? "需要你确认需求" : status === "resolving" ? "正在理解你的补充"
    : status === "resolved" ? "需求已确认" : status === "invalidated" ? "数据范围已变化"
    : status === "unconfirmed" ? "需要核对数据范围" : status === "continued" ? "已重新授权并继续"
    : status === "cancelled" ? "需求已取消" : "需要核对需求状态";
  return <>
    {status !== "pending" && <div className="clarification-heading"><span>{label}</span></div>}
    <p className="clarification-question">{question}</p>
    <span className="clarification-scope">{scopeLabel}</span>
    {(summary || reason) && <details className="clarification-context">
      <summary>查看问题说明</summary>
      {summary && <p className="clarification-summary">{summary}</p>}
      {reason && <p className="clarification-reason">{reason}</p>}
    </details>}
  </>;
}
