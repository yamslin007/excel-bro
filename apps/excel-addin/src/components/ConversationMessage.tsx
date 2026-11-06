import type { ReactNode } from "react";

/** Shared conversation layout. Hosts supply facts and actions, not separate message shells. */
export function ConversationMessage({ id, role, groupStart = false, author = false, badge,
  executionResult = false, className = "", actions, feedback, children }: {
  id?: string; role: string; groupStart?: boolean; author?: boolean; badge?: string | null;
  executionResult?: boolean; className?: string; actions?: ReactNode;
  /** 归属到这条执行记录的操作反馈；由宿主按 executionId 精确匹配后传入。 */
  feedback?: ReactNode; children: ReactNode;
}) {
  return <article className={`message-row conversation-message ${role} ${className}`.trim()}
    data-message-id={id} data-group={groupStart ? "start" : "continue"}>
    <div className={`message-content conversation-reading${executionResult ? " has-execution-result" : ""}`}>
      {author && <span className={`message-author${badge ? " has-source" : " sr-only"}`}>
        <span className="sr-only">Excel Bro</span>{badge && <em>{badge}</em>}
      </span>}
      {children}
      {feedback}
    </div>
    {actions && <div className="message-copy-row conversation-actions">{actions}</div>}
  </article>;
}
