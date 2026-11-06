// @vitest-environment jsdom
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";
import { readFileSync } from "node:fs";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { ActivityLog } from "../hooks/useActivityProgress";
import { TaskProgressTimeline } from "./TaskProgressTimeline";

afterEach(() => cleanup());

it("普通聊天展开后只显示完成说明，不重复阶段和过时的准备文案", () => {
  const log: ActivityLog = {
    operationId: "chat", outcome: "completed", totalMs: 2100,
    steps: [{ id: "message-intent", label: "理解消息", phase: "intent", status: "completed",
      elapsedMs: 2100, note: "已理解并直接回复；未读取工作表",
      detail: "先判断是否需要工作簿上下文；尚未读取字段或数据。" }]
  };
  const view = render(<TaskProgressTimeline log={log} />);
  fireEvent.click(screen.getByRole("button", { name: "查看过程" }));
  expect(screen.getByText("已理解并直接回复；未读取工作表")).toBeTruthy();
  expect(screen.queryByText("理解消息")).toBeNull();
  expect(screen.queryByText(log.steps[0].detail!)).toBeNull();
  expect(view.container.querySelector(".activity-step-badges")).toBeNull();
  view.rerender(<TaskProgressTimeline log={{ ...log, outcome: "failed" }} />);
  expect(screen.getByText(log.steps[0].detail!)).toBeTruthy();
});

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
  it("keeps an explicit question compact without hiding pause or recovery controls", () => {
    const log = { ...sampleLog(), outcome: "awaiting_input" as const };
    const view = render(<TaskProgressTimeline log={log} awaitingAnswer />);
    expect(screen.getByRole("button", { name: "查看过程" }).getAttribute("aria-expanded")).toBe("false");
    expect(view.container.querySelector('.task-progress-heading strong')?.textContent).toContain("等待你的回答");
    fireEvent.click(screen.getByRole("button", { name: "查看过程" }));
    expect(screen.getByText("read_range")).toBeTruthy();
    view.rerender(<TaskProgressTimeline log={log} awaitingAnswer />);
    expect(screen.getByRole("button", { name: "收起" }).getAttribute("aria-expanded")).toBe("true");
  });
  it("freezes the entire stopped segment even while backend status is unknown", () => {
    vi.useFakeTimers();
    try {
      vi.setSystemTime(2000);
      const { container } = render(<TaskProgressTimeline log={{ ...sampleLog(),
        timelineStartedAt: 1000, totalMs: 1000, excludedKind: "waiting",
        excludedSince: 2000, outcome: "awaiting_input" }} />);
      const before = container.textContent;
      act(() => vi.advanceTimersByTime(10000));
      expect(container.textContent).toBe(before);
      expect(vi.getTimerCount()).toBe(0);
    } finally { vi.useRealTimers(); }
  });

  it("keeps a live task running after its latest step has completed", () => {
    const { container } = render(<TaskProgressTimeline log={sampleLog()} live />);
    expect(container.querySelector(".task-progress-timeline")?.getAttribute("data-status")).toBe("progress");
    expect(container.querySelector(".task-progress-overall")?.textContent).toBe("进行中");
    expect(container.querySelector(".activity-steps")).toBeNull();
    fireEvent.click(screen.getByRole("button", { name: "查看过程" }));
    expect(container.querySelector(".activity-steps > li")?.getAttribute("data-status")).toBe("completed");
  });

  it("shows pause and progress-wait time separately from foreground time", () => {
    const log = { ...sampleLog(), totalMs: 10000, pausedMs: 2000, waitingMs: 3000, outcome: "completed" as const };
    render(<TaskProgressTimeline log={log} />);
    /*
     * 正常结束后过程框收起成一行：**阶段结论 + 前台耗时**（5.0s = 10s − 2s 暂停 − 3s 等待）。
     * 计时分账没有变，只是标题从"执行过程 · N 步 · …"换成结论。
     */
    expect(screen.getByText("已完成 · 5.0s")).toBeTruthy();
    // 分账明细留在折叠区**外**：收起不该把"暂停/等待各多久"一起藏起来。
    expect(screen.getByText("暂停 2.0s · 等待进度 3.0s")).toBeTruthy();

    fireEvent.click(screen.getByRole("button", { name: "查看过程" }));
    expect(screen.getByText("执行过程 · 1 步 · 5.0s")).toBeTruthy();
  });

  it("keeps early continuation compact inside the original timeline", () => {
    const resume = vi.fn();
    const { container } = render(<TaskProgressTimeline log={sampleLog()} onContinue={resume} />);
    expect(screen.getByText("已暂停")).toBeTruthy();
    const button = screen.getByRole("button", { name: "继续分析" });
    expect(button.closest(".task-progress-actions")).not.toBeNull();
    expect(container.querySelector(".folder-resume-card, .pause-details")).toBeNull();
    expect(screen.queryByText(/尚未生成规划检查点/)).toBeNull();
    fireEvent.click(button);
    expect(resume).toHaveBeenCalledOnce();
  });

  it("labels the non-submitting interruption action as pause", () => {
    const pause = vi.fn();
    render(<TaskProgressTimeline log={sampleLog()} live onStop={pause} />);
    const button = screen.getByRole("button", { name: "暂停" });
    expect(button.getAttribute("type")).toBe("button");
    expect(screen.queryByRole("button", { name: "停止" })).toBeNull();
    fireEvent.click(button);
    expect(pause).toHaveBeenCalledOnce();
  });

  it("folds completed ordinary chat after its live phase, but keeps details expandable", () => {
    const log = { ...sampleLog(), defaultCollapsed: true, outcome: "completed" as const };
    const view = render(<TaskProgressTimeline log={log} live />);
    expect(screen.queryByText("read_range")).toBeNull();
    view.rerender(<TaskProgressTimeline log={log} />);
    expect(screen.getByRole("button", { name: "查看过程" }).getAttribute("aria-expanded")).toBe("false");
    expect(screen.queryByText("read_range")).toBeNull();
    fireEvent.click(screen.getByRole("button", { name: "查看过程" }));
    view.rerender(<TaskProgressTimeline log={{ ...log, totalMs: 12000 }} />);
    expect(screen.getByText("read_range")).toBeTruthy();
  });

  it.each(["failed", "awaiting_input", "warning", "interrupted", "cancelled"] as const)("does not fold %s even with a chat presentation flag", (outcome) => {
    render(<TaskProgressTimeline log={{ ...sampleLog(), outcome, defaultCollapsed: true }} />);
    expect(screen.getByRole("button", { name: "收起" }).getAttribute("aria-expanded")).toBe("true");
    expect(screen.getByText("read_range")).toBeTruthy();
  });

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
      screen.getByRole("button", { name: "查看过程" }).getAttribute("aria-expanded")
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
    fireEvent.click(screen.getByRole("button", { name: "查看过程" }));

    expect(screen.getByText("基础步骤")).toBeTruthy();
    expect(screen.getAllByText("进行中").length).toBeGreaterThan(0);
  });

  it("labels a completed intent analysis as waiting for confirmation", () => {
    const log = sampleLog();
    log.outcome = "awaiting_input";

    render(<TaskProgressTimeline log={log} />);

    expect(screen.getByText(/需求分析过程 · 1 步/)).toBeTruthy();
    expect(screen.getByText("等待下一步操作")).toBeTruthy();
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
        screen.getByRole("button", { name: "查看过程" }).getAttribute(
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


it("distinguishes superseded segments and interrupted steps from real warnings", () => {
  const log: ActivityLog = { totalMs: 90, outcome: "awaiting_input", closedByFinish: ["interrupted"], steps: [
    { id: "interrupted", label: "规划阶段", status: "warning", elapsedMs: 90 },
    { id: "real-warning", label: "数据异常", status: "warning", elapsedMs: 10 }
  ] };
  render(<TaskProgressTimeline log={log} segment={{ continuesTask: true, onShowLater: () => {} }} />);
  expect(screen.getByText("本段已结束 · 已接续")).toBeTruthy();
  expect(screen.queryByText("等待下一步操作")).toBeNull();
  expect(screen.getByLabelText("本段已中断").textContent).toBe("Ⅱ");
  expect(screen.getByLabelText("需要注意").textContent).toBe("!");
  expect(screen.getByText(/接续过程 · 2 步/)).toBeTruthy();
});

it("identifies the model stage without reordering its completed tool calls", () => {
  const log: ActivityLog = { totalMs: 1000, steps: [
    { id: "planning.model", label: "正在生成操作方案", status: "progress", elapsedMs: 1000 },
    { id: "tool-1", label: "查询完成", status: "completed", elapsedMs: 30 }
  ] };
  const view = render(<TaskProgressTimeline log={log} live />);
  fireEvent.click(screen.getByRole("button", { name: "查看过程" }));
  expect(screen.getByText("正在生成操作方案（整体阶段，包含工具调用）")).toBeTruthy();
  expect([...view.container.querySelectorAll(".activity-step-label")].map(el => el.textContent)).toEqual([
    "正在生成操作方案（整体阶段，包含工具调用）", "查询完成"
  ]);
});

it("only describes a write preview with completed preflight evidence", () => {
  render(<TaskProgressTimeline log={{ totalMs: 10, outcome: "awaiting_input", steps: [
    { id: "preflight-plan", label: "计划校验", elapsedMs: 10, status: "completed" }
  ] }} />);
  expect(screen.getByText("预览已生成 · 待确认写入")).toBeTruthy();
});

describe("限高与内部滚动（A 批）", () => {
  function scrollableBody(container: HTMLElement): HTMLElement {
    const body = container.querySelector<HTMLElement>(".task-progress-content-body");
    expect(body).not.toBeNull();
    return body!;
  }

  it("进行中默认 160px，靠内部滚动而不是把输入框挤出窗格", () => {
    const { container } = render(<TaskProgressTimeline log={sampleLog()} live />);
    fireEvent.click(screen.getByRole("button", { name: "查看过程" }));
    const body = scrollableBody(container);
    expect(body.style.getPropertyValue("--process-panel-max-height")).toBe("160px");
    // 限高必须真的由 CSS 执行，不能只写在变量里没人读。
    expect(readFileSync("src/styles.css", "utf8")).toContain(
      "max-height: min(var(--process-panel-max-height, 560px), 46vh)"
    );
    expect(body.getAttribute("role")).toBe("group");
    // 可键盘滚动：没有 tabIndex 的滚动容器键盘用户根本进不去。
    expect(body.getAttribute("tabindex")).toBe("0");
  });

  it("用户要求看更多时才放宽到 420px，并能还原", () => {
    const originalDescriptor = Object.getOwnPropertyDescriptor(
      HTMLElement.prototype,
      "scrollHeight"
    );
    Object.defineProperty(HTMLElement.prototype, "scrollHeight", {
      configurable: true,
      get() {
        return this.classList.contains("task-progress-content-body") ? 900 : 0;
      }
    });
    try {
      const { container } = render(<TaskProgressTimeline log={sampleLog()} live />);
      fireEvent.click(screen.getByRole("button", { name: "查看过程" }));
      const body = scrollableBody(container);
      fireEvent.click(screen.getByRole("button", { name: "多显示几行" }));
      expect(body.style.getPropertyValue("--process-panel-max-height")).toBe("420px");
      fireEvent.click(screen.getByRole("button", { name: "还原高度" }));
      expect(body.style.getPropertyValue("--process-panel-max-height")).toBe("160px");
    } finally {
      if (originalDescriptor) {
        Object.defineProperty(HTMLElement.prototype, "scrollHeight", originalDescriptor);
      } else {
        Reflect.deleteProperty(HTMLElement.prototype, "scrollHeight");
      }
    }
  });

  it("正常结束后收起成一行阶段结论 + 耗时，且不把阅读位置顶走", () => {
    const scrollIntoView = vi.fn();
    (Element.prototype as unknown as { scrollIntoView: () => void }).scrollIntoView = scrollIntoView;
    try {
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
        render(<TaskProgressTimeline log={{ ...sampleLog(), outcome: "completed" }} live />);
        fireEvent.click(screen.getByRole("button", { name: "查看过程" }));
        fireEvent.click(screen.getByRole("button", { name: "收起" }));
        expect(screen.getByText("已完成 · 10.8s")).toBeTruthy();
        // 关键：收起**不许**替用户决定"现在该看哪"。
        expect(scrollIntoView).not.toHaveBeenCalled();
      } finally {
        if (originalDescriptor) {
          Object.defineProperty(HTMLElement.prototype, "scrollHeight", originalDescriptor);
        } else {
          Reflect.deleteProperty(HTMLElement.prototype, "scrollHeight");
        }
      }
    } finally {
      Reflect.deleteProperty(Element.prototype, "scrollIntoView");
    }
  });

  it("用户上翻内部日志后给出回到最新步骤的入口，贴底时不出现", () => {
    const { container } = render(<TaskProgressTimeline log={sampleLog()} live />);
    fireEvent.click(screen.getByRole("button", { name: "查看过程" }));
    const body = scrollableBody(container);
    Object.defineProperties(body, {
      scrollHeight: { configurable: true, value: 900 },
      clientHeight: { configurable: true, value: 160 }
    });
    body.scrollTop = 700;
    expect(screen.queryByRole("button", { name: /最新步骤/ })).toBeNull();

    body.scrollTop = 100;
    fireEvent.scroll(body);
    const jump = screen.getByRole("button", { name: /最新步骤/ });
    // 还在跑 ⇒ 说"有新进展"；停下之后同一个位置说"回到底部"。
    expect(jump.textContent).toBe("有新进展");

    fireEvent.click(jump);
    expect(body.scrollTop).toBe(900);
    expect(screen.queryByRole("button", { name: /最新步骤/ })).toBeNull();
  });

  it("等待回答时状态栏说的是「等待你的回答」，而不是通用的等待下一步", () => {
    render(<TaskProgressTimeline log={{ ...sampleLog(), outcome: "awaiting_input" }} awaitingAnswer />);
    expect(screen.getByText("等待你的回答")).toBeTruthy();
    expect(screen.queryByText("等待下一步操作")).toBeNull();
  });
});
