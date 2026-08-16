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

const snapshot = { ranges: [] } as unknown as ExecutionUndoSnapshot;

describe("useUndoSnapshot", () => {
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

    act(() => result.current.setLastUndoSnapshot(snapshot));
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

    act(() => result.current.setLastUndoSnapshot(snapshot));
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

    act(() => result.current.setLastUndoSnapshot(snapshot));

    await act(async () => {
      await result.current.undoLastExecution();
    });

    expect(undoExecutionMock).toHaveBeenCalledWith(snapshot);
    expect(result.current.lastUndoSnapshot).toBeNull();
    expect(onStatusChange).toHaveBeenNthCalledWith(1, "executing");
    expect(onStatusChange).toHaveBeenLastCalledWith("idle");
    expect(onMessage).toHaveBeenCalledWith(
      "已撤销上一次执行中记录的单元格值、公式和常用格式。"
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

    act(() => result.current.setLastUndoSnapshot(snapshot));

    await act(async () => {
      await result.current.undoLastExecution();
    });

    expect(onMessage).toHaveBeenCalledWith("undo failed");
    expect(onStatusChange).toHaveBeenLastCalledWith("idle");
    expect(result.current.lastUndoSnapshot).toBe(snapshot);
  });
});
