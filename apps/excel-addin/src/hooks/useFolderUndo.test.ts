// @vitest-environment jsdom
import { act, renderHook } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { undoFolderExecution } from "../api";
import { useFolderUndo, type FolderUndoEntry } from "./useFolderUndo";

vi.mock("../api", () => ({
  undoFolderExecution: vi.fn()
}));

const undoFolderExecutionMock = vi.mocked(undoFolderExecution);

function entry(planId: string, undoId = `undo-${planId}`): FolderUndoEntry {
  return {
    planId,
    undoId,
    sessionId: "folder-session"
  };
}

describe("useFolderUndo", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    undoFolderExecutionMock.mockResolvedValue({
      filesRestored: ["scores.xlsx"],
      executionsUndone: 1,
      remainingUndoCount: 0
    });
  });

  it("registers and restores a folder execution even when verification is unrelated", async () => {
    const onMessage = vi.fn();
    const onAfterUndo = vi.fn();
    const { result } = renderHook(() =>
      useFolderUndo({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage,
        onAfterUndo
      })
    );
    const undo = entry("failed-verification-plan");

    act(() => result.current.pushFolderUndo(undo));
    await act(async () => {
      await result.current.undoFolderExecutionTo(undo.planId);
    });

    expect(undoFolderExecutionMock).toHaveBeenCalledWith(
      undo.sessionId,
      undo.undoId
    );
    expect(result.current.folderUndoStack).toEqual([]);
    expect(onMessage).toHaveBeenCalledWith(
      "已撤销上一次文件夹执行，并恢复到执行前状态：scores.xlsx。"
    );
    expect(onAfterUndo).toHaveBeenCalledTimes(1);
  });

  it("keeps the recovery entry when the file changed after execution", async () => {
    undoFolderExecutionMock.mockRejectedValue(
      new Error("执行后文件又发生了变化")
    );
    const onMessage = vi.fn();
    const { result } = renderHook(() =>
      useFolderUndo({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage
      })
    );
    const undo = entry("changed-plan");

    act(() => result.current.pushFolderUndo(undo));
    await act(async () => {
      await result.current.undoFolderExecutionTo(undo.planId);
    });

    expect(result.current.folderUndoStack).toEqual([undo]);
    expect(onMessage).toHaveBeenCalledWith("执行后文件又发生了变化");
  });

  it("uses a server recovery reference even when the local stack was reloaded", async () => {
    const onAfterUndo = vi.fn();
    const { result } = renderHook(() =>
      useFolderUndo({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage: vi.fn(),
        onAfterUndo
      })
    );

    await act(async () => {
      await result.current.undoFolderExecutionTo("reloaded-plan", {
        sessionId: "folder-session",
        undoId: "server-undo-id"
      });
    });

    expect(undoFolderExecutionMock).toHaveBeenCalledWith(
      "folder-session",
      "server-undo-id"
    );
    expect(onAfterUndo).toHaveBeenCalledTimes(1);
  });

  it("does not report a completed restore as failed when only refresh fails", async () => {
    const onMessage = vi.fn();
    const { result } = renderHook(() =>
      useFolderUndo({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage,
        onAfterUndo: vi.fn().mockRejectedValue(new Error("刷新失败"))
      })
    );
    const undo = entry("refresh-failure-plan");

    act(() => result.current.pushFolderUndo(undo));
    await act(async () => {
      await result.current.undoFolderExecutionTo(undo.planId);
    });

    expect(result.current.folderUndoStack).toEqual([]);
    expect(onMessage).toHaveBeenLastCalledWith(
      "撤销已经完成，但文件夹列表刷新失败：刷新失败"
    );
  });

  it("uses one server request when undoing an older execution", async () => {
    undoFolderExecutionMock.mockResolvedValue({
      filesRestored: ["scores.xlsx"],
      executionsUndone: 2,
      remainingUndoCount: 1
    });
    const { result } = renderHook(() =>
      useFolderUndo({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage: vi.fn()
      })
    );
    const first = entry("plan-1");
    const second = entry("plan-2");
    const third = entry("plan-3");
    act(() => {
      result.current.pushFolderUndo(first);
      result.current.pushFolderUndo(second);
      result.current.pushFolderUndo(third);
    });

    await act(async () => {
      await result.current.undoFolderExecutionTo("plan-2");
    });

    expect(undoFolderExecutionMock).toHaveBeenCalledWith(
      second.sessionId,
      second.undoId
    );
    expect(result.current.folderUndoStack).toEqual([first]);
  });

  it("drops stale recovery entries when the folder session changes", () => {
    const { result } = renderHook(() =>
      useFolderUndo({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage: vi.fn()
      })
    );
    const oldEntry = entry("old-plan");
    const newEntry = {
      ...entry("new-plan"),
      sessionId: "rebuilt-session"
    };

    act(() => {
      result.current.pushFolderUndo(oldEntry);
      result.current.pushFolderUndo(newEntry);
    });

    expect(result.current.folderUndoStack).toEqual([newEntry]);
  });

  it("reconciles local entries with the server after managed deletion", () => {
    const { result } = renderHook(() =>
      useFolderUndo({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage: vi.fn()
      })
    );
    const first = entry("same-plan", "undo-1");
    const second = entry("same-plan", "undo-2");
    act(() => {
      result.current.pushFolderUndo(first);
      result.current.pushFolderUndo(second);
      result.current.reconcileFolderUndoIds(["undo-2"]);
    });

    expect(result.current.folderUndoStack).toEqual([second]);
  });

  it("uses undoId to distinguish repeated executions of one plan", async () => {
    const { result } = renderHook(() =>
      useFolderUndo({
        isBusy: false,
        onStatusChange: vi.fn(),
        onMessage: vi.fn()
      })
    );
    const first = entry("same-plan", "undo-1");
    const second = entry("same-plan", "undo-2");
    act(() => {
      result.current.pushFolderUndo(first);
      result.current.pushFolderUndo(second);
    });

    await act(async () => {
      await result.current.undoFolderExecutionTo("undo-2");
    });

    expect(undoFolderExecutionMock).toHaveBeenCalledWith(
      second.sessionId,
      second.undoId
    );
    expect(result.current.folderUndoStack).toEqual([first]);
  });
});
