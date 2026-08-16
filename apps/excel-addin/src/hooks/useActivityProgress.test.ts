// @vitest-environment jsdom
import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { renderHook, act } from "@testing-library/react";
import { useActivityProgress } from "./useActivityProgress";

describe("useActivityProgress", () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it("should initialize with no activity and zero seconds", () => {
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog: vi.fn() })
    );

    expect(result.current.activity).toBeNull();
    expect(result.current.activitySeconds).toBe(0);
  });

  it("should start an activity with empty completed steps", () => {
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog: vi.fn() })
    );

    act(() => result.current.startActivity("标题", "明细"));

    expect(result.current.activity).toMatchObject({
      title: "标题",
      detail: "明细",
      completed: []
    });
    expect(result.current.activity?.startedAt).toBe(
      result.current.activity?.lastStepAt
    );
    expect(result.current.activitySeconds).toBe(0);
  });

  it("should advance activity title and detail without adding a step", () => {
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog: vi.fn() })
    );

    act(() => result.current.startActivity("开始", "第一步"));
    act(() =>
      result.current.advanceActivity("进行中", "下一步")
    );

    expect(result.current.activity?.title).toBe("进行中");
    expect(result.current.activity?.detail).toBe("下一步");
    expect(result.current.activity?.completed).toHaveLength(0);
  });

  it("should append a completed step and update lastStepAt", () => {
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog: vi.fn() })
    );

    act(() => result.current.startActivity("开始", "第一步"));
    act(() =>
      result.current.advanceActivity(
        "已完成一步",
        "继续",
        "已识别字段",
        { note: "note", detail: "detail" }
      )
    );

    expect(result.current.activity?.completed).toHaveLength(1);
    expect(result.current.activity?.completed[0]).toMatchObject({
      label: "已识别字段",
      note: "note",
      detail: "detail"
    });
    expect(result.current.activity?.lastStepAt).toBeGreaterThanOrEqual(
      result.current.activity?.startedAt ?? 0
    );
  });

  it("should update only the activity detail", () => {
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog: vi.fn() })
    );

    act(() => result.current.startActivity("开始", "旧明细"));
    act(() => result.current.updateActivityDetail("新明细"));

    expect(result.current.activity).toMatchObject({
      title: "开始",
      detail: "新明细"
    });
  });

  it("should persist completed steps and clear activity on completion", () => {
    const onPersistLog = vi.fn();
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog })
    );

    act(() => result.current.startActivity("开始", "明细"));
    act(() =>
      result.current.advanceActivity("进行中", "明细", "完成一步")
    );
    act(() => result.current.completeActivity());

    expect(onPersistLog).toHaveBeenCalledTimes(1);
    expect(onPersistLog.mock.calls[0][0].steps).toHaveLength(1);
    expect(result.current.activity).toBeNull();
    expect(result.current.activitySeconds).toBe(0);
  });

  it("should clear activity without persisting when there are no steps", () => {
    const onPersistLog = vi.fn();
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog })
    );

    act(() => result.current.startActivity("开始", "明细"));
    act(() => result.current.completeActivity());

    expect(onPersistLog).not.toHaveBeenCalled();
    expect(result.current.activity).toBeNull();
    expect(result.current.activitySeconds).toBe(0);
  });

  it("should increment elapsed seconds on an interval", () => {
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog: vi.fn() })
    );

    act(() => result.current.startActivity("开始", "明细"));
    expect(result.current.activitySeconds).toBe(0);

    act(() => {
      vi.advanceTimersByTime(3000);
    });

    expect(result.current.activitySeconds).toBe(3);
  });

  it("should clear the interval when activity is completed", () => {
    const clearIntervalSpy = vi.spyOn(window, "clearInterval");
    const { result } = renderHook(() =>
      useActivityProgress({ onPersistLog: vi.fn() })
    );

    act(() => result.current.startActivity("开始", "明细"));
    act(() => result.current.completeActivity());

    expect(clearIntervalSpy).toHaveBeenCalled();
    clearIntervalSpy.mockRestore();
  });
});
