// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import type { ActivityLog } from "../hooks/useActivityProgress";
import { TaskProgressTimeline } from "./TaskProgressTimeline";

afterEach(() => cleanup());

function sampleLog(): ActivityLog {
  return {
    operationId: "operation-timeline-test",
    totalMs: 10800,
    steps: [
      {
        id: "planning.tool.1.1",
        label: "已读取数据范围",
        elapsedMs: 420,
        phase: "planning",
        status: "completed",
        toolName: "read_range",
        actionType: "splitGroupAggregate",
        target: { file: "报表.xlsx", sheet: "sheet-dc1", range: "A1:W811" },
        facts: [
          { label: "结果", value: "811 行 × 23 列", style: "text" }
        ],
        trace: {
          id: "planning.tool.1.1:evidence",
          kind: "evidence",
          summary: "只读取已授权范围",
          basis: "工作表授权"
        }
      }
    ]
  };
}

describe("TaskProgressTimeline", () => {
  it("is expanded by default and shows tools, actions, targets and facts", () => {
    render(<TaskProgressTimeline log={sampleLog()} />);

    expect(
      screen.getByRole("button", { name: "收起" }).getAttribute("aria-expanded")
    ).toBe("true");
    expect(screen.getByText("read_range")).toBeTruthy();
    expect(screen.getByText("splitGroupAggregate")).toBeTruthy();
    expect(
      screen.getByText("报表.xlsx › sheet-dc1!A1:W811")
    ).toBeTruthy();
    expect(screen.getByText("811 行 × 23 列")).toBeTruthy();
    expect(screen.getByText("判断：只读取已授权范围")).toBeTruthy();
    expect(
      screen.queryByRole("button", { name: /收起执行过程/ })
    ).toBeNull();
  });

  it("only collapses after the user chooses to collapse it", () => {
    render(<TaskProgressTimeline log={sampleLog()} />);

    fireEvent.click(screen.getByRole("button", { name: "收起" }));

    expect(
      screen.getByRole("button", { name: "展开" }).getAttribute("aria-expanded")
    ).toBe("false");
    expect(screen.queryByText("read_range")).toBeNull();
  });

  it("shows a deterministic fallback when all optional details are absent", () => {
    const log: ActivityLog = {
      totalMs: 50,
      steps: [
        {
          id: "minimal",
          label: "基础步骤",
          elapsedMs: 50,
          phase: "planning",
          status: "progress"
        }
      ]
    };

    render(<TaskProgressTimeline log={log} live />);

    expect(screen.getByText("基础步骤")).toBeTruthy();
    expect(screen.getAllByText("进行中").length).toBeGreaterThan(0);
  });

  it("labels a completed intent analysis as waiting for confirmation", () => {
    const log = sampleLog();
    log.outcome = "awaiting_input";

    render(<TaskProgressTimeline log={log} />);

    expect(screen.getByText(/需求分析过程 · 1 步/)).toBeTruthy();
    expect(screen.getByText("等待确认")).toBeTruthy();
    expect(screen.queryByText("已完成")).toBeNull();
  });

  it("adds a bottom collapse control when the expanded content is long", () => {
    const originalDescriptor = Object.getOwnPropertyDescriptor(
      HTMLElement.prototype,
      "scrollHeight"
    );
    Object.defineProperty(HTMLElement.prototype, "scrollHeight", {
      configurable: true,
      get() {
        return this.classList.contains("task-progress-content-body") ? 600 : 0;
      }
    });

    try {
      render(<TaskProgressTimeline log={sampleLog()} />);

      fireEvent.click(
        screen.getByRole("button", { name: /收起执行过程/ })
      );

      expect(
        screen.getByRole("button", { name: "展开" }).getAttribute(
          "aria-expanded"
        )
      ).toBe("false");
      expect(screen.queryByText("read_range")).toBeNull();
    } finally {
      if (originalDescriptor) {
        Object.defineProperty(
          HTMLElement.prototype,
          "scrollHeight",
          originalDescriptor
        );
      } else {
        Reflect.deleteProperty(HTMLElement.prototype, "scrollHeight");
      }
    }
  });
});
