// @vitest-environment jsdom
import { describe, it, expect, vi, beforeEach } from "vitest";
import { renderHook, act } from "@testing-library/react";
import type { ExecutionUndoSnapshot } from "../contracts";
import { isRunningInExcel, undoExecution } from "../excel";
import { useUndoSnapshot } from "./useUndoSnapshot";

vi.mock("../excel", () => ({
  isRunningInExcel: vi.fn(),
  undoExecution: vi.fn()
}));

const isRunningInExcelMock = vi.mocked(isRunningInExcel);
const undoExecutionMock = vi.mocked(undoExecution);

const snapshot = makeSnapshot("plan-1");

function makeSnapshot(
  planId: string,
  executionId?: string
): ExecutionUndoSnapshot {
  return {
    planId,
    executionId,
    capturedAt: "2026-08-22T00:00:00.000Z",
    ranges: [],
    createdWorksheets: [],
    deletedWorksheets: [],
    skippedWorksheets: []
  };
}

describe("useUndoSnapshot", () => {
  it("pauses undo for target confirmation before any host read/write or failure message", async () => {
    const beforeOperation = vi.fn(async () => false);
    const onMessage = vi.fn();
    const onStatusChange = vi.fn();
    const { result } = renderHook(() => useUndoSnapshot({ isBusy: false, beforeOperation, onMessage, onStatusChange }));
    act(() => result.current.pushUndoSnapshot(snapshot));
    await act(async () => { await result.current.undoLastExecution(); });
    expect(undoExecutionMock).not.toHaveBeenCalled();
    expect(onMessage).not.toHaveBeenCalled(); expect(onStatusChange).not.toHaveBeenCalled();
    expect(result.current.lastUndoSnapshot).toBe(snapshot);
    beforeOperation.mockResolvedValue(true);
    await act(async () => { await result.current.undoLastExecution(); });
    expect(undoExecutionMock).toHaveBeenCalledOnce();
  });
  beforeEach(() => {
    vi.clearAllMocks();
    isRunningInExcelMock.mockReturnValue(true);
    undoExecutionMock.mockResolvedValue(undefined);
  });

  it("should initialize with no undo snapshot", () => {
    const { result } = renderHook(() =>
      useUndoSnapshot({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage: vi.fn()
      })
    );

    expect(result.current.lastUndoSnapshot).toBeNull();
  });

  it("should set and clear an undo snapshot", () => {
    const { result } = renderHook(() =>
      useUndoSnapshot({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage: vi.fn()
      })
    );

    act(() => result.current.pushUndoSnapshot(snapshot));
    expect(result.current.lastUndoSnapshot).toBe(snapshot);

    act(() => result.current.clearUndoSnapshot());
    expect(result.current.lastUndoSnapshot).toBeNull();
  });

  it("should do nothing when there is no snapshot", async () => {
    const onStatusChange = vi.fn();
    const onMessage = vi.fn();
    const { result } = renderHook(() =>
      useUndoSnapshot({
        isBusy: false,
        onStatusChange,
        onMessage
      })
    );

    await act(async () => {
      await result.current.undoLastExecution();
    });

    expect(undoExecutionMock).not.toHaveBeenCalled();
    expect(onStatusChange).not.toHaveBeenCalled();
    expect(onMessage).not.toHaveBeenCalled();
  });

  it("should do nothing while busy or outside Excel", async () => {
    const onStatusChange = vi.fn();
    const { result } = renderHook(() =>
      useUndoSnapshot({
        isBusy: true,
        onStatusChange,
        onMessage: vi.fn()
      })
    );

    act(() => result.current.pushUndoSnapshot(snapshot));
    await act(async () => {
      await result.current.undoLastExecution();
    });

    expect(undoExecutionMock).not.toHaveBeenCalled();

    isRunningInExcelMock.mockReturnValue(false);
    await act(async () => {
      await result.current.undoLastExecution();
    });

    expect(undoExecutionMock).not.toHaveBeenCalled();
  });

  it("should execute undo, clear snapshot, notify and refresh", async () => {
    const onStatusChange = vi.fn();
    const onMessage = vi.fn();
    const onAfterUndo = vi.fn().mockResolvedValue(undefined);
    const { result } = renderHook(() =>
      useUndoSnapshot({
        isBusy: false,
        onStatusChange,
        onMessage,
        onAfterUndo
      })
    );

    act(() => result.current.pushUndoSnapshot(snapshot));

    await act(async () => {
      await result.current.undoLastExecution();
    });

    expect(undoExecutionMock).toHaveBeenCalledWith(snapshot);
    expect(result.current.lastUndoSnapshot).toBeNull();
    expect(onStatusChange).toHaveBeenNthCalledWith(1, "executing");
    expect(onStatusChange).toHaveBeenLastCalledWith("idle");
    expect(onMessage).toHaveBeenCalledWith(
      "已撤销上一次执行中记录的单元格值、公式和格式。"
    );
    expect(onAfterUndo).toHaveBeenCalledTimes(1);
  });

  it("should notify on undo failure and restore idle status", async () => {
    undoExecutionMock.mockRejectedValue(new Error("undo failed"));
    const onStatusChange = vi.fn();
    const onMessage = vi.fn();
    const { result } = renderHook(() =>
      useUndoSnapshot({
        isBusy: false,
        onStatusChange,
        onMessage
      })
    );

    act(() => result.current.pushUndoSnapshot(snapshot));

    await act(async () => {
      await result.current.undoLastExecution();
    });

    expect(onMessage).toHaveBeenCalledWith("undo failed");
    expect(onStatusChange).toHaveBeenLastCalledWith("idle");
    expect(result.current.lastUndoSnapshot).toBe(snapshot);
  });

  it("keeps multiple snapshots with the newest on top", () => {
    const { result } = renderHook(() =>
      useUndoSnapshot({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage: vi.fn()
      })
    );

    const first = makeSnapshot("plan-1");
    const second = makeSnapshot("plan-2");
    act(() => result.current.pushUndoSnapshot(first));
    act(() => result.current.pushUndoSnapshot(second));

    expect(result.current.undoStack).toEqual([first, second]);
    expect(result.current.lastUndoSnapshot).toBe(second);
  });

  it("cascades undo from the newest back to the requested plan", async () => {
    const onMessage = vi.fn();
    const { result } = renderHook(() =>
      useUndoSnapshot({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage
      })
    );

    const first = makeSnapshot("plan-1");
    const second = makeSnapshot("plan-2");
    const third = makeSnapshot("plan-3");
    act(() => result.current.pushUndoSnapshot(first));
    act(() => result.current.pushUndoSnapshot(second));
    act(() => result.current.pushUndoSnapshot(third));

    await act(async () => {
      await result.current.undoExecutionTo("plan-2");
    });

    // 逆序撤销：先 plan-3 再 plan-2，plan-1 保留。
    expect(undoExecutionMock.mock.calls.map((call) => call[0])).toEqual([
      third,
      second
    ]);
    expect(result.current.undoStack).toEqual([first]);
    expect(onMessage).toHaveBeenCalledWith(
      "已撤销最近 2 次执行中记录的单元格值、公式和格式。"
    );
  });

  it("keeps the failing snapshot on the stack for retry", async () => {
    undoExecutionMock.mockRejectedValueOnce(new Error("undo failed"));
    const { result } = renderHook(() =>
      useUndoSnapshot({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage: vi.fn()
      })
    );

    const first = makeSnapshot("plan-1");
    const second = makeSnapshot("plan-2");
    act(() => result.current.pushUndoSnapshot(first));
    act(() => result.current.pushUndoSnapshot(second));

    await act(async () => {
      await result.current.undoExecutionTo("plan-1");
    });

    // plan-2 撤销即失败，两条都留在栈上。
    expect(result.current.undoStack).toEqual([first, second]);
  });

  it("uses executionId to distinguish repeated executions of one plan", async () => {
    const { result } = renderHook(() =>
      useUndoSnapshot({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage: vi.fn()
      })
    );
    const first = makeSnapshot("same-plan", "execution-1");
    const second = makeSnapshot("same-plan", "execution-2");
    act(() => {
      result.current.pushUndoSnapshot(first);
      result.current.pushUndoSnapshot(second);
    });

    await act(async () => {
      await result.current.undoExecutionTo("execution-2");
    });

    expect(undoExecutionMock).toHaveBeenCalledWith(second);
    expect(result.current.undoStack).toEqual([first]);
  });

  it("onRestored fires only for snapshots that actually succeeded", async () => {
    const onRestored = vi.fn();
    const { result } = renderHook(() =>
      useUndoSnapshot({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage: vi.fn(),
        onRestored
      })
    );
    const first = makeSnapshot("plan-1", "execution-1");
    const second = makeSnapshot("plan-2", "execution-2");
    const third = makeSnapshot("plan-3", "execution-3");
    act(() => {
      result.current.pushUndoSnapshot(first);
      result.current.pushUndoSnapshot(second);
      result.current.pushUndoSnapshot(third);
    });

    // 逆序撤销 third → second → first；让 second（较早、较后执行）失败。
    undoExecutionMock
      .mockResolvedValueOnce(undefined) // third 成功
      .mockRejectedValueOnce(new Error("undo failed")); // second 失败
    await act(async () => {
      await result.current.undoExecutionTo("plan-1");
    });

    // 只有真正成功的 third 触发 onRestored；失败的 second 与未执行的 first 都不触发。
    expect(onRestored).toHaveBeenCalledTimes(1);
    expect(onRestored).toHaveBeenCalledWith(third);
  });

  it("onRestored fires once per successfully undone snapshot, in reverse order", async () => {
    const onRestored = vi.fn();
    const { result } = renderHook(() =>
      useUndoSnapshot({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage: vi.fn(),
        onRestored
      })
    );
    const first = makeSnapshot("plan-1", "execution-1");
    const second = makeSnapshot("plan-2", "execution-2");
    act(() => {
      result.current.pushUndoSnapshot(first);
      result.current.pushUndoSnapshot(second);
    });

    await act(async () => {
      await result.current.undoExecutionTo("plan-1");
    });

    // 逆序：second 先成功、first 后成功。
    expect(onRestored.mock.calls.map((call) => call[0])).toEqual([second, first]);
  });

  it("onRestored is not fired on conflict/failure (no partial 'undone' marking)", async () => {
    const onRestored = vi.fn();
    undoExecutionMock.mockRejectedValue(new Error("在本次撤销将要写回的字段上又发生了变化，已停止撤销"));
    const { result } = renderHook(() =>
      useUndoSnapshot({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage: vi.fn(),
        onRestored
      })
    );
    act(() => result.current.pushUndoSnapshot(snapshot));

    await act(async () => {
      await result.current.undoLastExecution();
    });

    expect(onRestored).not.toHaveBeenCalled();
    expect(result.current.lastUndoSnapshot).toBe(snapshot);
  });

  it("clearUndoSnapshot does not fire onRestored (clearing the stack is not an undo)", async () => {
    const onRestored = vi.fn();
    const { result } = renderHook(() =>
      useUndoSnapshot({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage: vi.fn(),
        onRestored
      })
    );
    act(() => result.current.pushUndoSnapshot(snapshot));
    act(() => result.current.clearUndoSnapshot());

    expect(onRestored).not.toHaveBeenCalled();
  });
});
