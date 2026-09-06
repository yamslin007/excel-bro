// @vitest-environment jsdom
import { act, renderHook, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  cleanupFolderBackups,
  collectLegacyFolderBackups,
  deleteFolderBackup,
  exportFolderBackup,
  listFolderBackups,
  openFolderBackupDirectory,
  setFolderBackupPinned
} from "../api";
import { useFolderBackups } from "./useFolderBackups";

vi.mock("../api", () => ({
  cleanupFolderBackups: vi.fn(),
  collectLegacyFolderBackups: vi.fn(),
  deleteFolderBackup: vi.fn(),
  exportFolderBackup: vi.fn(),
  listFolderBackups: vi.fn(),
  openFolderBackupDirectory: vi.fn(),
  setFolderBackupPinned: vi.fn()
}));

const listMock = vi.mocked(listFolderBackups);
const cleanupMock = vi.mocked(cleanupFolderBackups);
const deleteMock = vi.mocked(deleteFolderBackup);
const exportMock = vi.mocked(exportFolderBackup);
const collectMock = vi.mocked(collectLegacyFolderBackups);
const openMock = vi.mocked(openFolderBackupDirectory);
const pinMock = vi.mocked(setFolderBackupPinned);

const emptyResult = {
  folderPath: "C:\\work",
  backupDirectory: "C:\\work\\.excel-bro-backups",
  openDirectoryToken: null,
  backupManagementToken: null,
  backups: [],
  legacyCandidates: [],
  warnings: []
};

const sessionTarget = {
  kind: "session" as const,
  sessionId: "session-1",
  folderPath: "C:\\work"
};

describe("useFolderBackups", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    listMock.mockResolvedValue(emptyResult);
    cleanupMock.mockResolvedValue({
      deletedExecutions: 0,
      deletedFiles: 0,
      bytesFreed: 0,
      skippedActive: 0,
      failures: [],
      legacyCandidateCount: 0
    });
    deleteMock.mockResolvedValue({
      deletedUndoId: "a".repeat(32),
      discardedUndo: true,
      filesDeleted: 1,
      bytesFreed: 1024,
      remainingUndoCount: 0,
      invalidatedUndoIds: [],
      remainingUndoIds: [],
      cleanupPending: false,
      cleanupWarnings: []
    });
    exportMock.mockResolvedValue({
      operationId: "export-test-operation",
      exportDirectory: "C:\\work\\.excel-bro-backups\\exports\\copy",
      files: ["copy\\scores.xlsx"]
    });
    openMock.mockResolvedValue({
      opened: true,
      directory: "C:\\work\\.excel-bro-backups"
    });
    pinMock.mockResolvedValue({ undoId: "a".repeat(32), pinned: true });
    collectMock.mockResolvedValue({ collectedFiles: [], failures: [] });
  });

  it("loads backups for the active work directory session", async () => {
    const { result } = renderHook(() => useFolderBackups(sessionTarget));

    await waitFor(() => expect(result.current.loading).toBe(false));

    expect(listMock).toHaveBeenCalledWith({ sessionId: "session-1" });
    expect(result.current.data).toEqual(emptyResult);
    expect(result.current.readOnly).toBe(false);
  });

  it("cleans managed backups and refreshes the list", async () => {
    const { result } = renderHook(() => useFolderBackups(sessionTarget));
    await waitFor(() => expect(result.current.loading).toBe(false));

    await act(async () => {
      await result.current.cleanup();
    });

    expect(cleanupMock).toHaveBeenCalledWith({ sessionId: "session-1" });
    expect(listMock).toHaveBeenCalledTimes(2);
    expect(result.current.notice).toBe("没有需要清理的旧备份。");
  });

  it("opens the backend-registered backup directory token", async () => {
    listMock.mockResolvedValue({
      ...emptyResult,
      openDirectoryToken: "open-token"
    });
    const { result } = renderHook(() => useFolderBackups(sessionTarget));
    await waitFor(() => expect(result.current.loading).toBe(false));

    await act(async () => {
      await result.current.openBackupDirectory();
    });

    expect(openMock).toHaveBeenCalledWith("open-token");
    expect(result.current.notice).toContain("文件资源管理器");
  });

  it("deletes a managed backup, reports discarded undo and refreshes", async () => {
    listMock.mockResolvedValue({
      ...emptyResult,
      backupManagementToken: "manage-token"
    });
    const { result } = renderHook(() => useFolderBackups(sessionTarget));
    await waitFor(() => expect(result.current.loading).toBe(false));

    let deleted: Awaited<ReturnType<typeof result.current.deleteBackup>> = null;
    await act(async () => {
      deleted = await result.current.deleteBackup("a".repeat(32));
    });

    expect(
      (deleted as Awaited<ReturnType<typeof result.current.deleteBackup>>)
        ?.deletedUndoId
    ).toBe("a".repeat(32));
    expect(deleteMock).toHaveBeenCalledWith(
      "manage-token",
      "a".repeat(32),
      undefined,
      undefined
    );
    expect(listMock).toHaveBeenCalledTimes(2);
    expect(result.current.notice).toContain("放弃本次执行的撤销能力");
  });

  it("updates the long-term retention flag through the management token", async () => {
    listMock.mockResolvedValue({
      ...emptyResult,
      backupManagementToken: "manage-token"
    });
    const { result } = renderHook(() => useFolderBackups(sessionTarget));
    await waitFor(() => expect(result.current.loading).toBe(false));

    await act(async () => {
      await result.current.setBackupPinned("a".repeat(32), true);
    });

    expect(pinMock).toHaveBeenCalledWith("manage-token", "a".repeat(32), true);
    expect(result.current.notice).toContain("长期保留");
    expect(listMock).toHaveBeenCalledTimes(2);
  });

  it("reuses the same export operation id after an ambiguous failure", async () => {
    listMock.mockResolvedValue({
      ...emptyResult,
      backupManagementToken: "manage-token"
    });
    exportMock
      .mockRejectedValueOnce(new Error("连接中断"))
      .mockResolvedValueOnce({
        operationId: "export-retry-operation",
        exportDirectory: "C:\\work\\.excel-bro-backups\\exports\\copy",
        files: ["copy\\scores.xlsx"]
      });
    const { result } = renderHook(() => useFolderBackups(sessionTarget));
    await waitFor(() => expect(result.current.loading).toBe(false));

    await act(async () => {
      await result.current.exportBackup("a".repeat(32));
    });
    await act(async () => {
      await result.current.exportBackup("a".repeat(32));
    });

    const firstOperationId = exportMock.mock.calls[0][2];
    const secondOperationId = exportMock.mock.calls[1][2];
    expect(firstOperationId).toBeTruthy();
    expect(secondOperationId).toBe(firstOperationId);
    expect(result.current.notice).toContain("备份副本已导出");
  });

  it("按文件夹路径加载时使用目录令牌管理备份", async () => {
    // 工作簿模式下只有路径没有数据会话；受管仓库操作复用目录令牌，
    // 不需要扩大工作表数据范围。
    listMock.mockResolvedValue({
      ...emptyResult,
      openDirectoryToken: "manage-token",
      backupManagementToken: "manage-token"
    });
    const { result } = renderHook(() =>
      useFolderBackups({ kind: "path", folderPath: "C:\\work" })
    );
    await waitFor(() => expect(result.current.loading).toBe(false));

    expect(listMock).toHaveBeenCalledWith({ folderPath: "C:\\work" });
    expect(result.current.readOnly).toBe(true);

    await act(async () => {
      await result.current.cleanup();
    });

    expect(cleanupMock).toHaveBeenCalledWith({
      backupManagementToken: "manage-token"
    });
    expect(result.current.error).toBe("");

    await act(async () => {
      await result.current.deleteBackup("a".repeat(32));
    });

    expect(deleteMock).toHaveBeenCalledWith(
      "manage-token",
      "a".repeat(32),
      undefined,
      undefined
    );
    expect(result.current.error).toBe("");
  });

  it("目标不可用时不发请求", async () => {
    const { result } = renderHook(() =>
      useFolderBackups({ kind: "unavailable", reason: "remote" })
    );
    await waitFor(() => expect(result.current.loading).toBe(false));

    expect(listMock).not.toHaveBeenCalled();
    expect(result.current.data).toBeNull();
  });
});
