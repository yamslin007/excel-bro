// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { useFolderBackups } from "../hooks/useFolderBackups";
import type { FolderBackupTarget } from "../hooks/useFolderBackups";
import { FolderBackupManager } from "./FolderBackupManager";

vi.mock("../hooks/useFolderBackups", () => ({
  useFolderBackups: vi.fn()
}));

const useFolderBackupsMock = vi.mocked(useFolderBackups);

type BackupsState = ReturnType<typeof useFolderBackups>;

function backupsState(overrides: Partial<BackupsState> = {}): BackupsState {
  return {
    data: {
      folderPath: "C:\\work",
      backupDirectory: "C:\\work\\.excel-bro-backups",
      openDirectoryToken: "open-token",
      backupManagementToken: "manage-token",
      backups: [
        {
          revision: 3,
          undoId: "a".repeat(32),
          undoSessionId: "undo-session",
          planId: "plan-1",
          planTitle: "整理报表",
          createdAt: "2026-09-02T10:30:00+08:00",
          status: "ready",
          undoAvailability: "eligible",
          deleteInvalidatesUndoIds: [],
          pinned: false,
          protectedReasons: ["当前可撤销链需要"],
          canUndo: true,
          fileCount: 1,
          totalBytes: 1024,
          files: [
            {
              originalRelativePath: "华南\\报表.xlsx",
              backupRelativePath: "华南/报表.xlsx",
              recoveryAction: "restore",
              sizeBytes: 1024
            }
          ]
        }
      ],
      legacyCandidates: [],
      warnings: []
    },
    loading: false,
    working: false,
    error: "",
    notice: "",
    readOnly: false,
    reload: vi.fn().mockResolvedValue(undefined),
    openBackupDirectory: vi.fn().mockResolvedValue(
      "C:\\work\\.excel-bro-backups"
    ),
    exportBackup: vi.fn().mockResolvedValue("C:\\work\\exports"),
    deleteBackup: vi.fn().mockResolvedValue({
      deletedUndoId: "a".repeat(32),
      discardedUndo: true,
      filesDeleted: 1,
      bytesFreed: 1024,
      remainingUndoCount: 0,
      invalidatedUndoIds: [],
      remainingUndoIds: [],
      cleanupPending: false,
      cleanupWarnings: []
    }),
    setBackupPinned: vi.fn().mockResolvedValue(true),
    cleanup: vi.fn().mockResolvedValue(undefined),
    collectLegacy: vi.fn().mockResolvedValue(undefined),
    ...overrides
  };
}

const sessionTarget: FolderBackupTarget = {
  kind: "session",
  sessionId: "current-session",
  folderPath: "C:\\work"
};

describe("FolderBackupManager", () => {
  beforeEach(() => {
    useFolderBackupsMock.mockReturnValue(backupsState());
  });

  // vite.config.ts 没开 test.globals，@testing-library 的自动清理不会注册，
  // 上一条用例渲染的 DOM 会残留在 document.body 里污染后续 screen 查询。
  afterEach(() => {
    cleanup();
  });

  it("shows the fixed work directory entry and restores with the server session", async () => {
    const onUndo = vi.fn().mockResolvedValue(undefined);
    render(
      <FolderBackupManager
        target={sessionTarget}
        onClose={vi.fn()}
        onUndo={onUndo}
      />
    );

    expect(screen.getByText("C:\\work\\.excel-bro-backups")).toBeTruthy();
    expect(screen.getByText("整理报表")).toBeTruthy();
    fireEvent.click(screen.getByRole("button", { name: "撤销本次执行" }));

    expect(onUndo).toHaveBeenCalledWith(
      "plan-1",
      "a".repeat(32),
      "undo-session"
    );
  });

  it("summarizes session undo and separates backup storage categories", () => {
    const state = backupsState();
    useFolderBackupsMock.mockReturnValue({
      ...state,
      data: {
        ...state.data!,
        storageSummary: {
          automaticBytes: 1024,
          pinnedBytes: 2048,
          abnormalBytes: 4096,
          exportedBytes: 8192,
          legacyBytes: 16384,
          pendingCleanupBytes: 512,
          totalBytes: 32256,
          softLimitBytes: 1073741824,
          complete: true,
          calculatedAt: "2026-09-06T00:00:00.000Z"
        }
      }
    });

    render(
      <FolderBackupManager
        target={sessionTarget}
        onClose={vi.fn()}
        onUndo={vi.fn().mockResolvedValue(undefined)}
        workbookUndoCount={2}
        folderUndoCount={1}
      />
    );

    expect(screen.getByLabelText("当前工作簿 2 条")).toBeTruthy();
    expect(screen.getByLabelText("文件夹执行 1 条")).toBeTruthy();
    expect(screen.getByText("空间占用")).toBeTruthy();
    expect(screen.getByLabelText("自动恢复点 1.0 KB")).toBeTruthy();
    expect(screen.getByLabelText("长期保留 2.0 KB")).toBeTruthy();
    expect(screen.getByLabelText("异常资料 4.0 KB")).toBeTruthy();
  });

  it("opens the registered managed backup folder and keeps a copy-path fallback", () => {
    const state = backupsState();
    useFolderBackupsMock.mockReturnValue(state);
    render(
      <FolderBackupManager
        target={sessionTarget}
        onClose={vi.fn()}
        onUndo={vi.fn().mockResolvedValue(undefined)}
      />
    );

    fireEvent.click(screen.getByRole("button", { name: "打开文件夹" }));

    expect(state.openBackupDirectory).toHaveBeenCalledTimes(1);
    expect(screen.getByRole("button", { name: "复制路径" })).toBeTruthy();
  });

  it("confirms before deleting an active backup and explains undo loss", async () => {
    const state = backupsState();
    const onBackupDeleted = vi.fn();
    useFolderBackupsMock.mockReturnValue(state);
    render(
      <FolderBackupManager
        target={sessionTarget}
        onClose={vi.fn()}
        onUndo={vi.fn().mockResolvedValue(undefined)}
        onBackupDeleted={onBackupDeleted}
      />
    );

    fireEvent.click(screen.getByRole("button", { name: "删除并放弃撤销" }));

    expect(screen.getByText("删除并放弃撤销？")).toBeTruthy();
    expect(screen.getByText(/不会改变当前工作文件/)).toBeTruthy();
    expect(state.deleteBackup).not.toHaveBeenCalled();

    fireEvent.click(screen.getByRole("button", { name: "确认删除" }));
    await waitFor(() =>
      expect(state.deleteBackup).toHaveBeenCalledWith("a".repeat(32), 3, [])
    );
    expect(onBackupDeleted).toHaveBeenCalledWith([]);
  });

  it("explains a new-file-only recovery point without showing a fake backup size", () => {
    const base = backupsState();
    const backup = base.data!.backups[0];
    useFolderBackupsMock.mockReturnValue({
      ...base,
      data: {
        ...base.data!,
        backups: [
          {
            ...backup,
            fileCount: 1,
            totalBytes: 0,
            files: [
              {
                originalRelativePath: "Excel Bro 结果.xlsx",
                backupRelativePath: null,
                recoveryAction: "remove",
                sizeBytes: 0
              }
            ]
          }
        ]
      }
    });

    render(
      <FolderBackupManager
        target={sessionTarget}
        onClose={vi.fn()}
        onUndo={vi.fn().mockResolvedValue(undefined)}
      />
    );

    expect(screen.getByText("1 个新建文件撤销记录")).toBeTruthy();
    expect(screen.getByText(/无需保存原文件副本/)).toBeTruthy();
    expect(screen.queryByText("1 个文件 · 0 B")).toBeNull();
    expect(
      (screen.getByRole("button", { name: "取出副本" }) as HTMLButtonElement)
        .disabled
    ).toBe(true);
  });

  it("offers verified file extraction for a recovery point needing review", () => {
    const base = backupsState();
    const backup = base.data!.backups[0];
    useFolderBackupsMock.mockReturnValue({
      ...base,
      data: {
        ...base.data!,
        backups: [
          {
            ...backup,
            status: "needs_attention",
            statusReason: "部分文件状态未知",
            canUndo: false,
            undoAvailability: "view_only"
          }
        ]
      }
    });

    render(
      <FolderBackupManager
        target={sessionTarget}
        onClose={vi.fn()}
        onUndo={vi.fn().mockResolvedValue(undefined)}
      />
    );

    expect(screen.getByRole("button", { name: "取出可用副本" })).toBeTruthy();
    expect(
      (screen.getByRole("button", { name: "删除备份" }) as HTMLButtonElement)
        .disabled
    ).toBe(true);
  });

  it("can cancel an inline backup deletion confirmation", () => {
    const state = backupsState();
    useFolderBackupsMock.mockReturnValue(state);
    render(
      <FolderBackupManager
        target={sessionTarget}
        onClose={vi.fn()}
        onUndo={vi.fn().mockResolvedValue(undefined)}
      />
    );

    fireEvent.click(screen.getByRole("button", { name: "删除并放弃撤销" }));
    fireEvent.click(screen.getByRole("button", { name: "取消" }));

    expect(screen.queryByRole("button", { name: "确认删除" })).toBeNull();
    expect(state.deleteBackup).not.toHaveBeenCalled();
  });

  it("does not offer to open a backup directory that does not exist", () => {
    const state = backupsState();
    useFolderBackupsMock.mockReturnValue({
      ...state,
      data: {
        ...state.data!,
        openDirectoryToken: null,
        backups: []
      }
    });
    render(
      <FolderBackupManager
        target={sessionTarget}
        onClose={vi.fn()}
        onUndo={vi.fn().mockResolvedValue(undefined)}
      />
    );

    for (const name of ["打开文件夹", "复制路径"]) {
      expect(
        (screen.getByRole("button", { name }) as HTMLButtonElement).disabled
      ).toBe(true);
    }
    expect(state.openBackupDirectory).not.toHaveBeenCalled();
  });

  it("点击归拢旧版备份直接触发 collectLegacy", () => {
    // 加载项运行在 Office 的 iframe 中，window.confirm 会被宿主屏蔽并静默返回
    // 假值，按钮表现为点击无反应。这里断言点击直接调用归拢，防止再退回到
    // confirm 包裹的写法。
    const base = backupsState();
    const collectLegacy = vi.fn().mockResolvedValue(undefined);
    useFolderBackupsMock.mockReturnValue({
      ...base,
      data: {
        ...base.data!,
        legacyCandidates: [
          {
            relativePath: "旧备份\\报表.xlsx.bak",
            sizeBytes: 2048,
            modifiedAt: "2026-08-20T09:00:00+08:00"
          }
        ]
      },
      collectLegacy
    });

    render(
      <FolderBackupManager
        target={sessionTarget}
        onClose={vi.fn()}
        onUndo={vi.fn().mockResolvedValue(undefined)}
      />
    );

    expect(screen.getByText("发现 1 个旧版散落备份")).toBeTruthy();
    fireEvent.click(screen.getByRole("button", { name: "归拢旧版备份" }));

    expect(collectLegacy).toHaveBeenCalledTimes(1);
  });

  it("当前工作簿入口通过目录令牌显示完整备份管理操作", () => {
    useFolderBackupsMock.mockReturnValue(backupsState({ readOnly: true }));

    render(
      <FolderBackupManager
        target={{ kind: "path", folderPath: "C:\\work" }}
        onClose={vi.fn()}
        onUndo={vi.fn().mockResolvedValue(undefined)}
      />
    );

    expect(screen.getByRole("button", { name: "清理旧备份" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "取出副本" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "保留此副本" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "删除并放弃撤销" })).toBeTruthy();
  });

  it("marks an automatic recovery point for long-term retention", () => {
    const state = backupsState();
    useFolderBackupsMock.mockReturnValue(state);
    render(
      <FolderBackupManager
        target={sessionTarget}
        onClose={vi.fn()}
        onUndo={vi.fn().mockResolvedValue(undefined)}
      />
    );

    fireEvent.click(screen.getByRole("button", { name: "保留此副本" }));

    expect(state.setBackupPinned).toHaveBeenCalledWith("a".repeat(32), true);
  });

  it("云端工作簿显示不可用空态", () => {
    useFolderBackupsMock.mockReturnValue(
      backupsState({ data: null, loading: false })
    );

    render(
      <FolderBackupManager
        target={{ kind: "unavailable", reason: "remote" }}
        onClose={vi.fn()}
        onUndo={vi.fn().mockResolvedValue(undefined)}
      />
    );

    expect(screen.getByText("云端工作簿暂不支持")).toBeTruthy();
  });
});
