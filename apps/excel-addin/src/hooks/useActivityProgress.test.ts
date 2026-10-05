// @vitest-environment jsdom
import { act, renderHook } from "@testing-library/react";
import { createElement, StrictMode, type ReactNode } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { TaskProgressEvent } from "../contracts";
import { useActivityProgress } from "./useActivityProgress";

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
        expect.objectContaining({ id: "planning.resume.request-1" }),
        expect.objectContaining({ id: "planning.next", label: "继续生成方案" })
      ])
    );
    expect(
      result.current.activity?.steps.some(
        (step) => step.label === "不应重复显示"
      )
    ).toBe(false);
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
});
