import type { ReactNode } from "react";

/** Both hosts provide verified facts; this component only presents the record. */
export function RecoveryTerminalRecord({ kind, title, summary, children, className = "", operationId }: {
  kind: string; title: string; summary: string | null; children: ReactNode; className?: string; operationId?: string;
}) {
  return <details className={`recovery-terminal-record is-terminal ${className}`.trim()}
    aria-label="已结束的恢复记录" data-recovery-terminal={kind} data-operation-id={operationId}>
    <summary><strong>{title}</strong><span className="recovery-terminal-summary">{summary}</span></summary>
    <div className="recovery-terminal-details">{children}</div>
  </details>;
}
