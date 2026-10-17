import type { ReactNode } from "react";
import "./executionResult.css";

/** Presentation only: each host supplies its own verified outcome and recovery controls. */
export function ExecutionResultCard({
  tone, title, summary, className = "", children
}: {
  tone: "verified" | "failed" | "executed_unverified" | "unverified" | "partial" | "undone";
  title: string;
  summary: ReactNode;
  className?: string;
  children?: ReactNode;
}) {
  const symbol = tone === "verified" ? "✓" : tone === "undone" ? "↶" : "!";
  return (
    <section className={`verification-card execution-result-card ${tone} ${className}`.trim()}
      aria-label="写入执行结果">
      <div className="execution-result-heading">
        <span className="execution-result-icon" aria-hidden="true">{symbol}</span>
        <strong>{title}</strong>
      </div>
      <div className="execution-result-summary">{summary}</div>
      {children}
    </section>
  );
}

export function ExecutionResultDetails({ children, open = false }: { children: ReactNode; open?: boolean }) {
  return <details className="execution-result-details" open={open}>
    <summary>查看执行明细</summary>
    {children}
  </details>;
}
