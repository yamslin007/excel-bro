// @vitest-environment jsdom
import { act, renderHook } from "@testing-library/react";
import { createElement, StrictMode, type ReactNode } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { TaskProgressEvent } from "../contracts";
import { resumeActivityLog, useActivityProgress } from "./useActivityProgress";

function progressEvent(
  overrides: Partial<TaskProgressEvent> = {}
): TaskProgressEvent {
  return {
    schemaVersion: 1,
    operationId: "execution-progress-1",
    sequence: 1,
    phase: "executing",
    stepId: "execution.action.1",
    status: "started",
    title: "正在执行第 1 步",
    timestamp: "2026-09-02T00:00:00Z",
    ...overrides
  };
}

describe("useActivityProgress", () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it("完成回执重放不改耗时，暂停等待经持久化后继续累计", () => {
    const persisted = vi.fn();
    const { result } = renderHook(() => useActivityProgress({ onPersistLog: persisted }));
    vi.setSystemTime(1000);
    act(() => result.current.startActivity("步骤", "", { operationId: "execution-progress-1", stepId: "execution.action.1" }));
    vi.setSystemTime(1500);
    act(() => result.current.applyProgressEvent(progressEvent({ status: "completed" })));
    vi.setSystemTime(9000);
    act(() => result.current.applyProgressEvent(progressEvent({ sequence: 2, status: "completed" })));
    expect(result.current.activity?.steps[0].elapsedMs).toBe(500);
    act(() => result.current.applyProgressEvent(progressEvent({ sequence: 3, stepId: "next" })));
    vi.setSystemTime(10000);
    act(() => result.current.completeActivity("warning", "暂停接收", "awaiting_input"));
    const log = persisted.mock.calls[0][0];
    expect(log.excludedKind).toBe("waiting");
    vi.setSystemTime(20000);
    act(() => result.current.resumeActivity("继续", "", { operationId: "execution-progress-1", previousLog: log }));
    expect(result.current.activity?.waitingMs).toBe(10000);
    expect(result.current.activity?.steps[1].elapsedMs).toBe(1000);
    vi.setSystemTime(20500);
    act(() => result.current.applyProgressEvent(progressEvent({ sequence: 4, stepId: "next", status: "completed" })));
    expect(result.current.activity?.steps[1].elapsedMs).toBe(1500);
  });

  it("starts with one visible in-progress step", () => {
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog: vi.fn() })
    );

    act(() => result.current.startActivity("标题", "明细"));

    expect(result.current.activity).toMatchObject({
      title: "标题",
      detail: "明细",
      connectionState: "connected",
      completed: []
    });
    expect(result.current.activity?.steps).toEqual([
      expect.objectContaining({
        id: "local",
        label: "标题",
        status: "started",
        detail: "明细"
      })
    ]);
  });

  it("updates a legacy step in place and appends the next active step", () => {
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog: vi.fn() })
    );

    act(() => result.current.startActivity("开始", "第一步"));
    act(() =>
      result.current.advanceActivity(
        "正在确认需求",
        "下一步",
        "已识别字段",
        { note: "识别到 23 个字段", detail: "只使用已授权工作表" }
      )
    );

    expect(result.current.activity?.steps).toHaveLength(2);
    expect(result.current.activity?.steps[0]).toMatchObject({
      label: "已识别字段",
      status: "completed",
      note: "识别到 23 个字段"
    });
    expect(result.current.activity?.steps[1]).toMatchObject({
      label: "正在确认需求",
      status: "started"
    });
  });

  it("completes a terminal decision without appending a duplicate step", () => {
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog: vi.fn() })
    );

    act(() => result.current.startActivity("正在确认需求", "等待模型判断"));
    act(() =>
      result.current.advanceActivity(
        "模型判断需要先澄清需求",
        "发现两个互斥维度",
        "模型判断需要先澄清需求",
        { note: "需要确认建表维度" }
      )
    );

    expect(result.current.activity?.steps).toHaveLength(1);
    expect(result.current.activity?.steps[0]).toMatchObject({
      label: "模型判断需要先澄清需求",
      status: "completed",
      note: "需要确认建表维度"
    });
  });

  it("persists an awaiting-input outcome to the stable message key", () => {
    const onPersistLog = vi.fn();
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog })
    );

    act(() =>
      result.current.startActivity("正在确认需求", "等待模型判断", {
        persistenceKey: "activity-message-1"
      })
    );
    act(() =>
      result.current.completeCurrentStep("模型判断需要先澄清需求", {
        note: "需要用户确认"
      })
    );
    act(() =>
      result.current.completeActivity(
        "completed",
        undefined,
        "awaiting_input"
      )
    );

    expect(onPersistLog).toHaveBeenCalledWith(
      expect.objectContaining({
        outcome: "awaiting_input",
        steps: [
          expect.objectContaining({
            label: "模型判断需要先澄清需求",
            status: "completed",
            note: "需要用户确认"
          })
        ]
      }),
      "activity-message-1"
    );
  });

  it("persists a completed activity only once in StrictMode", () => {
    const onPersistLog = vi.fn();
    const wrapper = ({ children }: { children: ReactNode }) =>
      createElement(StrictMode, null, children);
    const { result } = renderHook(
      () => useActivityProgress({ onPersistLog }),
      { wrapper }
    );

    act(() =>
      result.current.startActivity("正在分析", "等待模型判断", {
        persistenceKey: "activity-message-strict"
      })
    );
    act(() => result.current.completeActivity());

    expect(onPersistLog).toHaveBeenCalledTimes(1);
    expect(onPersistLog).toHaveBeenCalledWith(
      expect.objectContaining({
        steps: [expect.objectContaining({ label: "正在分析" })]
      }),
      "activity-message-strict"
    );
  });

  it("keeps the stable message key after activity steps advance", () => {
    const onPersistLog = vi.fn();
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog })
    );

    act(() =>
      result.current.startActivity("正在识别字段", "", {
        persistenceKey: "activity-message-advanced"
      })
    );
    act(() =>
      result.current.advanceActivity(
        "模型判断需要先澄清需求",
        "",
        "已识别字段"
      )
    );
    act(() =>
      result.current.completeActivity(
        "completed",
        undefined,
        "awaiting_input"
      )
    );

    expect(onPersistLog).toHaveBeenCalledTimes(1);
    expect(onPersistLog).toHaveBeenCalledWith(
      expect.objectContaining({ outcome: "awaiting_input" }),
      "activity-message-advanced"
    );
  });

  it("inserts started events immediately and updates the same step in place", () => {
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog: vi.fn() })
    );
    act(() =>
      result.current.startActivity("开始", "", {
        operationId: "execution-progress-1",
        stepId: "execution.preflight"
      })
    );

    act(() =>
      result.current.applyProgressEvent(
        progressEvent({
          toolName: "read_range",
          facts: [{ label: "范围", value: "A1:B3", style: "code" }],
          target: { sheet: "Sheet1", range: "A1:B3" }
        })
      )
    );
    act(() =>
      result.current.applyProgressEvent(
        progressEvent({
          sequence: 2,
          status: "progress",
          title: "正在执行第 1 步",
          actionType: "writeValues",
          current: 1,
          total: 2,
          unit: "actions"
        })
      )
    );

    const actionSteps = result.current.activity?.steps.filter(
      (step) => step.id === "execution.action.1"
    );
    expect(actionSteps).toHaveLength(1);
    expect(actionSteps?.[0]).toMatchObject({
      status: "progress",
      toolName: "read_range",
      actionType: "writeValues",
      target: { sheet: "Sheet1", range: "A1:B3" },
      facts: [{ label: "范围", value: "A1:B3", style: "code" }]
    });
  });

  it("ignores duplicate sequences and events from another operation", () => {
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog: vi.fn() })
    );
    act(() =>
      result.current.startActivity("当前任务", "", {
        operationId: "execution-progress-1"
      })
    );
    act(() => result.current.applyProgressEvent(progressEvent()));
    act(() =>
      result.current.applyProgressEvent(
        progressEvent({ title: "不应覆盖的新标题" })
      )
    );
    act(() =>
      result.current.applyProgressEvent(
        progressEvent({
          operationId: "execution-other",
          sequence: 2,
          title: "错误任务"
        })
      )
    );

    expect(result.current.activity?.title).toBe("正在执行第 1 步");
    expect(
      result.current.activity?.steps.some((step) => step.label === "错误任务")
    ).toBe(false);
  });

  it("associates trace metadata with its step and persists the full timeline", () => {
    const onPersistLog = vi.fn();
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog })
    );
    act(() =>
      result.current.startActivity("开始", "", {
        operationId: "execution-progress-1",
        stepId: "execution.action.1"
      })
    );
    act(() =>
      result.current.applyProgressEvent(
        progressEvent({
          sequence: 2,
          status: "completed",
          title: "第 1 步已完成",
          summary: "已写入目标表",
          trace: {
            kind: "evidence",
            summary: "已重新读取目标区域",
            basis: "本地执行器验收"
          }
        })
      )
    );
    act(() => result.current.completeActivity());

    expect(onPersistLog).toHaveBeenCalledWith(
      expect.objectContaining({
        operationId: "execution-progress-1",
        steps: [
          expect.objectContaining({
            id: "execution.action.1",
            status: "completed",
            trace: expect.objectContaining({ kind: "evidence" })
          })
        ]
      })
    );
  });

  it("marks an interrupted connection without reporting a false failure", () => {
    const onPersistLog = vi.fn();
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog })
    );
    act(() => result.current.startActivity("正在规划", "等待服务端"));
    act(() =>
      result.current.markActivityInterrupted("连接中断，执行结果待确认。")
    );

    expect(result.current.activity).toMatchObject({
      status: "started",
      connectionState: "interrupted",
      connectionMessage: "连接中断，执行结果待确认。"
    });
    expect(result.current.activity?.steps[0].status).toBe("started");

    act(() => result.current.completeActivity());
    expect(onPersistLog).toHaveBeenCalledWith(
      expect.objectContaining({
        connectionState: "interrupted",
        steps: [expect.objectContaining({ status: "warning" })]
      })
    );
  });

  it("keeps a transient disconnect non-terminal after reconnecting", () => {
    const onPersistLog = vi.fn();
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog })
    );
    act(() => result.current.startActivity("正在执行", "等待事件"));
    act(() => result.current.markActivityReconnecting("正在恢复进度。"));
    act(() => result.current.markActivityInterrupted("连接中断。"));
    act(() => result.current.markActivityConnected());
    act(() => result.current.completeActivity());

    expect(onPersistLog).toHaveBeenCalledWith(
      expect.objectContaining({
        connectionState: "connected",
        steps: [expect.objectContaining({ status: "completed" })]
      })
    );
  });

  it("resumes the same persisted timeline and only accepts newer events", () => {
    const onPersistLog = vi.fn();
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog })
    );
    act(() =>
      result.current.startActivity("正在规划", "读取字段", {
        persistenceKey: "activity-message-resume",
        operationId: "turn-resume-1",
        stepId: "planning.fields"
      })
    );
    act(() =>
      result.current.applyProgressEvent(
        progressEvent({
          operationId: "turn-resume-1",
          sequence: 3,
          stepId: "planning.fields",
          status: "completed",
          title: "字段读取完成"
        })
      )
    );
    act(() => result.current.completeActivity("warning", "连接中断", "interrupted"));
    const previousLog = onPersistLog.mock.calls[0][0];

    act(() =>
      result.current.resumeActivity("正在从断点继续", "保留已完成步骤", {
        persistenceKey: "activity-message-resume",
        operationId: "turn-resume-1",
        stepId: "planning.resume.request-1",
        initialSequence: 3,
        previousLog
      })
    );
    act(() =>
      result.current.applyProgressEvent(
        progressEvent({
          operationId: "turn-resume-1",
          sequence: 3,
          stepId: "planning.duplicate",
          title: "不应重复显示"
        })
      )
    );
    act(() =>
      result.current.applyProgressEvent(
        progressEvent({
          operationId: "turn-resume-1",
          sequence: 4,
          stepId: "planning.next",
          title: "继续生成方案"
        })
      )
    );

    expect(result.current.activity).toMatchObject({
      persistenceKey: "activity-message-resume",
      operationId: "turn-resume-1",
      sequence: 4
    });
    expect(result.current.activity?.steps).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          id: "planning.fields",
          label: "字段读取完成"
        }),
        expect.objectContaining({ id: "planning.next", label: "继续生成方案" })
      ])
    );
    expect(result.current.activity?.steps).toHaveLength(2);
    expect(
      result.current.activity?.steps.some(
        (step) => step.label === "不应重复显示"
      )
    ).toBe(false);
  });

  it.each([1, 2, 4, 7, 10])("resumes step %s in place, preserves completed rows and excludes paused time", (pausedAt) => {
    const { result } = renderHook(() => useActivityProgress({ onPersistLog: vi.fn() }));
    const steps = Array.from({ length: pausedAt }, (_, index) => ({
      id: `step-${index + 1}`, label: `步骤${index + 1}`, elapsedMs: 100,
      status: index === pausedAt - 1 ? "warning" as const : "completed" as const
    }));
    act(() => result.current.resumeActivity("继续", "核对后继续", {
      operationId: "task-seven", persistenceKey: "original-card", initialSequence: 7,
      previousLog: { operationId: "task-seven", sequence: 7, outcome: "awaiting_input", steps, totalMs: 700 }
    }));
    expect(result.current.activity?.steps).toHaveLength(pausedAt);
    expect(result.current.activity?.steps.slice(0, pausedAt - 1)).toEqual(steps.slice(0, pausedAt - 1));
    expect(result.current.activity?.steps[pausedAt - 1]).toMatchObject({ id: `step-${pausedAt}`, status: "progress", elapsedMs: 100 });
    act(() => result.current.applyProgressEvent(progressEvent({ operationId: "task-seven", sequence: 8,
      stepId: `step-${pausedAt}`, title: "当前步骤继续", status: "completed" })));
    expect(result.current.activity?.steps).toHaveLength(pausedAt);
    expect(result.current.activity?.steps[pausedAt - 1].elapsedMs).toBe(100);
    act(() => result.current.applyProgressEvent(progressEvent({ operationId: "task-seven", sequence: 9,
      stepId: `step-${pausedAt + 1}`, title: "下一步骤", status: "started" })));
    expect(result.current.activity?.steps.map((step) => step.id)).toEqual([...steps.map((step) => step.id), `step-${pausedAt + 1}`]);
  });

  it("persists an unfinished local step with the supplied final status", () => {
    const onPersistLog = vi.fn();
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog })
    );
    act(() => result.current.startActivity("开始", "明细"));
    act(() => result.current.completeActivity("cancelled", "用户已取消"));

    expect(onPersistLog).toHaveBeenCalledWith(
      expect.objectContaining({
        steps: [
          expect.objectContaining({
            label: "开始",
            status: "cancelled",
            note: "用户已取消"
          })
        ]
      })
    );
  });

  it("updates elapsed seconds and clears the timer on completion", () => {
    const clearIntervalSpy = vi.spyOn(window, "clearInterval");
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog: vi.fn() })
    );
    act(() => result.current.startActivity("开始", "明细"));
    act(() => {
      vi.advanceTimersByTime(3000);
    });

    expect(result.current.activitySeconds).toBe(3);

    act(() => result.current.completeActivity());
    expect(result.current.activity).toBeNull();
    expect(result.current.activitySeconds).toBe(0);
    expect(clearIntervalSpy).toHaveBeenCalled();
    clearIntervalSpy.mockRestore();
  });

  it("hands preparation to remote planning without leaving an earlier running step", () => {
    const onPersistLog = vi.fn();
    const { result } = renderHook(() => useActivityProgress({ onPersistLog }));
    act(() => result.current.startActivity("正在读取字段", ""));
    act(() => result.current.advanceActivity("正在理解需求", "", "已识别字段"));
    act(() => result.current.completeCurrentStep("需求已确认"));
    act(() => result.current.advanceActivity("正在准备规划请求", "重新核对范围"));
    act(() => vi.advanceTimersByTime(100));
    act(() => result.current.completeCurrentStep("已准备规划请求"));
    act(() => result.current.applyProgressEvent(progressEvent({
      phase: "planning", stepId: "planning.understand", status: "completed"
    })));
    act(() => result.current.applyProgressEvent(progressEvent({
      sequence: 2, phase: "planning", stepId: "planning.model.1"
    })));
    expect(result.current.activity?.steps.slice(0, -1).every(
      (step) => step.status === "completed"
    )).toBe(true);
    expect(result.current.activity?.steps[2]).toMatchObject({ elapsedMs: 100 });
    act(() => result.current.completeActivity("completed", undefined, "awaiting_input"));
    expect(onPersistLog.mock.calls[0][0].steps.every(
      (step: { status: string }) => step.status === "completed"
    )).toBe(true);
  });

  it("does not persist orphan running steps or pretend they completed", () => {
    const onPersistLog = vi.fn();
    const { result } = renderHook(() => useActivityProgress({ onPersistLog }));
    act(() => result.current.startActivity("缺少结束事件的步骤", ""));
    act(() => result.current.applyProgressEvent(progressEvent()));
    act(() => result.current.completeActivity());
    const log = onPersistLog.mock.calls[0][0];
    expect(log.outcome).toBe("warning");
    expect(log.steps[0]).toMatchObject({ status: "warning" });
    expect(log.steps[0].note).toContain("未收到结束确认");
    expect(log.steps.every((step: { status: string }) =>
      !["started", "progress"].includes(step.status))).toBe(true);
  });

  /*
   * 收尾的**两档结论**必须分开（2026-10-21）。
   *
   * 活动级 `outcome` 说的是"这一轮怎么停下的"，步骤状态说的是"这一步做完了没有"。
   * 两者混成一个的直接后果：用户确认完计划、这一轮带着"待确认"收尾时，明明做完
   * 的那一步被写成"未收到结束确认"；或者反过来 —— 被暂停打断的那一步被写成
   * 已完成，恢复后第一条进度被"已结束不许重激活"丢掉，界面再也不动。
   */
  it("带着待确认结果收尾时这一步就是做完了，恢复后也不被重新打开", () => {
    const onPersistLog = vi.fn();
    const { result } = renderHook(() => useActivityProgress({ onPersistLog }));
    act(() => result.current.startActivity("规划", "正在规划", {
      stepId: "step-5", operationId: "turn-1", phase: "planning"
    }));
    act(() => result.current.applyProgressEvent(progressEvent({
      operationId: "turn-1", phase: "planning", stepId: "step-5", status: "progress"
    })));
    act(() => result.current.completeActivity("completed", "计划待确认", "awaiting_input"));

    const log = onPersistLog.mock.calls[0][0];
    // 活动等用户确认，但这一步自己确实做完了。
    expect(log.outcome).toBe("awaiting_input");
    expect(log.steps[0].status).toBe("completed");
    // 不是"被收尾关掉的"，因此恢复时不会被重新打开。
    expect(log.closedByFinish).toEqual([]);
    expect(resumeActivityLog(log, Date.now() + 5_000).steps[0].status).toBe("completed");
  });

  it("暂停收尾时那一步没有结论：落盘 closedByFinish，恢复后重新可推进", () => {
    const onPersistLog = vi.fn();
    const { result } = renderHook(() => useActivityProgress({ onPersistLog }));
    act(() => result.current.startActivity("规划", "正在规划", {
      stepId: "step-7", operationId: "turn-1", phase: "planning"
    }));
    act(() => result.current.applyProgressEvent(progressEvent({
      operationId: "turn-1", phase: "planning", stepId: "step-7", status: "progress"
    })));
    act(() => result.current.completeActivity("warning", "已暂停", "awaiting_input"));

    const log = onPersistLog.mock.calls[0][0];
    expect(log.steps[0].status).toBe("warning");
    // 关键：这句不是"你做完了"，而是"我们没等到结论"。
    expect(log.closedByFinish).toEqual(["step-7"]);
    expect(resumeActivityLog(log, Date.now() + 5_000).steps[0].status).toBe("progress");
  });

  describe("Python 代码工件", () => {
    const codeEvent = (overrides: Record<string, unknown> = {}) => ({
      codeId: "code_py_0123456789ab",
      attemptId: "py_0123456789ab",
      stage: "validating" as const,
      version: 1,
      chars: 812,
      complete: true,
      ...overrides
    });

    it("同一步骤的生命周期事件合并成一个工件，不新建步骤", () => {
      const { result } = renderHook(() => useActivityProgress({ onPersistLog: vi.fn() }));
      act(() =>
        result.current.startActivity("规划", "正在规划", {
          stepId: "python.analysis:py_0123456789ab",
          operationId: "turn-1",
          phase: "planning"
        })
      );
      act(() =>
        result.current.applyProgressEvent(
          progressEvent({
            operationId: "turn-1",
            phase: "planning",
            stepId: "python.analysis:py_0123456789ab",
            status: "started",
            pythonCode: codeEvent()
          })
        )
      );
      act(() =>
        result.current.applyProgressEvent(
          progressEvent({
            operationId: "turn-1",
            phase: "planning",
            stepId: "python.analysis:py_0123456789ab",
            status: "progress",
            sequence: 2,
            pythonCode: codeEvent({ stage: "running", executedVersion: 1 })
          })
        )
      );

      const steps = result.current.activity?.steps ?? [];
      expect(steps).toHaveLength(1);
      expect(steps[0].pythonCode?.stage).toBe("running");
      // 正文没有随事件下发，但工件记得执行过哪一版。
      expect(steps[0].pythonCode?.code).toBe("");
      expect(steps[0].pythonCode?.executedVersion).toBe(1);
    });

    it("重新生成是新的尝试：旧工件被替换，不把两段代码粘在一起", () => {
      const { result } = renderHook(() => useActivityProgress({ onPersistLog: vi.fn() }));
      act(() =>
        result.current.startActivity("规划", "正在规划", {
          stepId: "python.analysis:py_first",
          operationId: "turn-1",
          phase: "planning"
        })
      );
      act(() =>
        result.current.applyProgressEvent(
          progressEvent({
            operationId: "turn-1",
            phase: "planning",
            stepId: "python.analysis:py_first",
            status: "started",
            pythonCode: codeEvent({ code: "x = 1" })
          })
        )
      );
      act(() =>
        result.current.applyProgressEvent(
          progressEvent({
            operationId: "turn-1",
            phase: "planning",
            stepId: "python.analysis:py_second",
            status: "started",
            sequence: 2,
            pythonCode: {
              ...codeEvent(),
              attemptId: "py_second",
              codeId: "code_py_1111111111ab",
              code: "y = 2"
            }
          })
        )
      );

      const steps = result.current.activity?.steps ?? [];
      const latest = steps[steps.length - 1].pythonCode;
      expect(latest?.code).toBe("y = 2");
      expect(latest?.code).not.toContain("x = 1");
    });
  });
});
