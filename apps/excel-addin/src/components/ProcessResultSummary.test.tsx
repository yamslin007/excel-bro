// @vitest-environment jsdom
import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { ProcessResultSummary } from "./ProcessResultSummary";
import { TaskProgressTimeline } from "./TaskProgressTimeline";
import type { ActivityLog } from "../hooks/useActivityProgress";

afterEach(() => cleanup());

const stepLog = (): ActivityLog => ({
  operationId: "op-1",
  totalMs: 12_000,
  outcome: "completed",
  steps: [{ id: "s1", label: "写入汇总表", elapsedMs: 12_000, status: "completed" }]
});

describe("ProcessResultSummary", () => {
  it("进行中不宣布结论", () => {
    const { container } = render(<ProcessResultSummary input={{ live: true, phase: "executing" }} />);
    expect(container.firstChild).toBeNull();
  });

  it("正常结束给出改了哪里 / 验证结果 / 下一步", () => {
    render(<ProcessResultSummary input={{
      live: false,
      phase: "reviewing",
      reviews: [{ sheet: "汇总", range: "B2:D4", status: "applied", rows: Array.from({ length: 9 }, (_, i) => ({ changed: true, key: String(i) })) }],
      verification: { status: "verified", total: 9, passed: 9, unverified: 0 }
    }} />);
    expect(screen.getByText("改了：汇总!B2:D4（9 格）")).toBeTruthy();
    expect(screen.getByText("9/9 项符合预期")).toBeTruthy();
    expect(screen.getByText("在主表核对结果后点「确认变更」或「放弃变更」。")).toBeTruthy();
  });

  it("已撤销：只说恢复事实，不再复述旧的验证结论", () => {
    render(<ProcessResultSummary input={{
      live: false,
      phase: "undone",
      reviews: [{ sheet: "汇总", range: "B2", status: "applied", rows: [{ changed: true }] }],
      verification: { status: "verified", total: 1, passed: 1, unverified: 0 }
    }} />);
    expect(screen.getByText(/还原/)).toBeTruthy();
    expect(screen.queryByText(/项符合预期/)).toBeNull();
  });

  it("**在过程框之外**：过程框收起后这三行仍然可见", () => {
    const { container } = render(<div>
      <TaskProgressTimeline log={stepLog()} />
      <ProcessResultSummary input={{
        live: false,
        phase: "reviewing",
        verification: { status: "verified", total: 1, passed: 1, unverified: 0 }
      }} />
    </div>);
    // 过程框已收起（步骤行不在 DOM 里）。
    expect(container.querySelector(".activity-steps")).toBeNull();
    // 摘要仍在，而且不是过程框的后代 —— 它不会被收起带走。
    const summary = container.querySelector(".process-result-summary")!;
    expect(summary).not.toBeNull();
    expect(summary.closest(".task-progress-timeline")).toBeNull();
    expect(summary.textContent).toContain("1/1 项符合预期");
  });
});
