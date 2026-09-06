import { useCallback, useEffect, useRef, useState } from "react";
import {
  cleanupFolderBackups,
  collectLegacyFolderBackups,
  deleteFolderBackup,
  exportFolderBackup,
  listFolderBackups,
  openFolderBackupDirectory,
  setFolderBackupPinned
} from "../api";
import type { FolderBackupsResult } from "../contracts";

/**
 * 备份查看的定位方式。
 *
 * sessionId：文件夹模式下已有会话，包含数据处理与撤销能力。
 * folderPath：工作簿模式下由文件位置推导；受管仓库操作使用短期目录令牌，
 * 不扩大工作表数据处理范围。
 * unavailable：工作簿未保存或在云端，无法定位备份目录。
 */
export type FolderBackupTarget =
  | { kind: "session"; sessionId: string; folderPath: string }
  | { kind: "path"; folderPath: string }
  | { kind: "unavailable"; reason: "unsaved" | "remote" };

export function useFolderBackups(target: FolderBackupTarget) {
  const [data, setData] = useState<FolderBackupsResult | null>(null);
  const [loading, setLoading] = useState(true);
  const [working, setWorking] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const pendingExportOperationIds = useRef(new Map<string, string>());

  const sessionId = target.kind === "session" ? target.sessionId : "";
  const folderPath = target.kind === "unavailable" ? "" : target.folderPath;
  const readOnly = target.kind === "path";

  const reload = useCallback(async () => {
    if (!sessionId && !folderPath) {
      setData(null);
      setLoading(false);
      return;
    }
    setLoading(true);
    setError("");
    try {
      setData(
        await listFolderBackups(
          sessionId ? { sessionId } : { folderPath }
        )
      );
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "读取备份失败");
    } finally {
      setLoading(false);
    }
  }, [folderPath, sessionId]);

  useEffect(() => {
    void reload();
  }, [reload]);

  useEffect(() => {
    pendingExportOperationIds.current.clear();
  }, [folderPath, sessionId]);

  const openBackupDirectory = useCallback(async () => {
    const directoryToken = data?.openDirectoryToken;
    if (!directoryToken) {
      setError("打开凭证不存在，请刷新备份管理后重试。");
      return null;
    }
    setWorking(true);
    setError("");
    try {
      const result = await openFolderBackupDirectory(directoryToken);
      setNotice("已在文件资源管理器中打开统一备份位置。");
      return result.directory;
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "打开备份文件夹失败");
      return null;
    } finally {
      setWorking(false);
    }
  }, [data?.openDirectoryToken]);

  const exportBackup = useCallback(
    async (undoId: string) => {
      const managementToken = data?.backupManagementToken;
      if (!sessionId && !managementToken) {
        setError("备份管理凭证不存在，请刷新后重试。");
        return null;
      }
      setWorking(true);
      setError("");
      const operationId =
        pendingExportOperationIds.current.get(undoId) ??
        `export-${Date.now().toString(36)}-${Math.random().toString(16).slice(2)}`;
      pendingExportOperationIds.current.set(undoId, operationId);
      try {
        const result = await exportFolderBackup(
          sessionId ? { sessionId } : { backupManagementToken: managementToken! },
          undoId,
          operationId
        );
        pendingExportOperationIds.current.delete(undoId);
        setNotice(
          `${result.complete === false ? "已取出可用副本" : "备份副本已导出"}到：${
            result.exportDirectory
          }${
            result.failures?.length
              ? `；${result.failures.length} 项未通过完整性检查或需要人工核对。`
              : ""
          }`
        );
        return result.exportDirectory;
      } catch (reason) {
        setError(reason instanceof Error ? reason.message : "导出备份失败");
        return null;
      } finally {
        setWorking(false);
      }
    },
    [data?.backupManagementToken, sessionId]
  );

  const deleteBackup = useCallback(
    async (
      undoId: string,
      expectedRevision?: number,
      expectedInvalidatedUndoIds?: string[]
    ) => {
      const managementToken = data?.backupManagementToken;
      if (!managementToken) {
        setError("备份管理凭证不存在，请刷新后重试。");
        return null;
      }
      setWorking(true);
      setError("");
      try {
        const result = await deleteFolderBackup(
          managementToken,
          undoId,
          expectedRevision,
          expectedInvalidatedUndoIds
        );
        setNotice(
          `${
            result.discardedUndo
              ? "备份已删除，并已放弃本次执行的撤销能力。"
              : "备份已删除。"
          }${
            result.invalidatedUndoIds.length > 0
              ? ` 另有 ${result.invalidatedUndoIds.length} 条较早撤销记录因依赖中断而失效。`
              : ""
          }${result.cleanupPending ? " 内部删除暂存将在后续安全清理。" : ""}`
        );
        await reload();
        return result;
      } catch (reason) {
        setError(reason instanceof Error ? reason.message : "删除备份失败");
        return null;
      } finally {
        setWorking(false);
      }
    },
    [data?.backupManagementToken, reload]
  );

  const setBackupPinned = useCallback(
    async (undoId: string, pinned: boolean) => {
      const managementToken = data?.backupManagementToken;
      if (!managementToken) {
        setError("备份管理凭证不存在，请刷新后重试。");
        return false;
      }
      setWorking(true);
      setError("");
      try {
        await setFolderBackupPinned(managementToken, undoId, pinned);
        setNotice(pinned ? "这份副本已设为长期保留。" : "已取消长期保留，将重新按自动清理规则管理。");
        await reload();
        return true;
      } catch (reason) {
        setError(reason instanceof Error ? reason.message : "更新保留状态失败");
        return false;
      } finally {
        setWorking(false);
      }
    },
    [data?.backupManagementToken, reload]
  );

  const cleanup = useCallback(async () => {
    const managementToken = data?.backupManagementToken;
    if (!sessionId && !managementToken) {
      setError("备份管理凭证不存在，请刷新后重试。");
      return;
    }
    setWorking(true);
    setError("");
    try {
      const result = await cleanupFolderBackups(
        sessionId ? { sessionId } : { backupManagementToken: managementToken! }
      );
      setNotice(
        `${
          result.deletedExecutions > 0
            ? `已清理 ${result.deletedExecutions} 份旧备份，释放 ${result.bytesFreed} 字节。`
            : "没有需要清理的旧备份。"
        }${
          (result.protectedExecutions ?? 0) > 0
            ? ` ${result.protectedExecutions} 份因仍可撤销、属于最新副本、主动保留或需要核对而保留。`
            : ""
        }${
          (result.overLimitBytes ?? 0) > 0
            ? ` 受保护和保留期内资料仍超过软上限 ${result.overLimitBytes} 字节，未强制删除。`
            : ""
        }${
          result.failures.length > 0
            ? ` ${result.failures.length} 项因安全检查或文件占用被保留。`
            : ""
        }`
      );
      await reload();
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "清理旧备份失败");
    } finally {
      setWorking(false);
    }
  }, [data?.backupManagementToken, reload, sessionId]);

  const collectLegacy = useCallback(async () => {
    const managementToken = data?.backupManagementToken;
    if (!sessionId && !managementToken) {
      setError("备份管理凭证不存在，请刷新后重试。");
      return;
    }
    setWorking(true);
    setError("");
    try {
      const result = await collectLegacyFolderBackups(
        sessionId ? { sessionId } : { backupManagementToken: managementToken! }
      );
      setNotice(
        `已归拢 ${result.collectedFiles.length} 个旧版备份。${
          result.failures.length > 0
            ? ` ${result.failures.length} 项未移动。`
            : ""
        }`
      );
      await reload();
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "归拢旧版备份失败");
    } finally {
      setWorking(false);
    }
  }, [data?.backupManagementToken, reload, sessionId]);

  return {
    data,
    loading,
    working,
    error,
    notice,
    readOnly,
    reload,
    openBackupDirectory,
    exportBackup,
    deleteBackup,
    setBackupPinned,
    cleanup,
    collectLegacy
  };
}
