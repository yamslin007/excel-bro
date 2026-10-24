import { useState, type ReactNode } from "react";
import { useFolderBackups, type FolderBackupTarget } from "../hooks/useFolderBackups";

interface FolderBackupManagerProps {
  target: FolderBackupTarget;
  onClose: () => void;
  onUndo: (planId: string, undoId: string, undoSessionId: string) => Promise<void>;
  onBackupDeleted?: (remainingUndoIds: string[]) => void;
  workbookUndoCount?: number;
  folderUndoCount?: number;
  /**
   * 工作表删除恢复点区段。
   *
   * 由调用方注入而不是在这里直接取：这一层只负责"备份与恢复"的版式，
   * 检查点的归属与查询留在检查点自己那一份实现里。
   */
  checkpointSection?: ReactNode;
}

function formatBytes(value: number): string {
  if (value < 1024) return `${value} B`;
  if (value < 1024 * 1024) return `${(value / 1024).toFixed(1)} KB`;
  if (value < 1024 * 1024 * 1024) {
    return `${(value / (1024 * 1024)).toFixed(1)} MB`;
  }
  return `${(value / (1024 * 1024 * 1024)).toFixed(2)} GB`;
}

function formatDate(value: string): string {
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? value : parsed.toLocaleString("zh-CN");
}

function backupStatusLabel(status: string, canUndo: boolean): string {
  if (status === "needs_attention") return "需要核对";
  if (status === "creating" || status === "preparing") return "正在准备";
  if (status === "prepared" || status === "applying") return "正在写入";
  if (status === "recovering") return "正在恢复";
  if (status === "rolled_back") return "已安全回滚";
  if (status === "restored") return "已撤销·历史副本";
  return canUndo ? "可撤销" : "历史副本";
}

function isBackupBusy(status: string): boolean {
  return ["creating", "preparing", "prepared", "applying", "recovering"].includes(
    status
  );
}

function canDeleteNormally(status: string): boolean {
  return ["ready", "restored", "rolled_back"].includes(status);
}

function canRecoverFromManagedRecord(status: string): boolean {
  return [
    "preparing",
    "prepared",
    "applying",
    "ready",
    "recovering",
    "needs_attention"
  ].includes(status);
}

export function FolderBackupManager({
  target,
  onClose,
  onUndo,
  onBackupDeleted,
  workbookUndoCount = 0,
  folderUndoCount = 0,
  checkpointSection
}: FolderBackupManagerProps) {
  const {
    data,
    managedDirectory,
    loading,
    working,
    error,
    notice,
    readOnly,
    reload,
    openBackupDirectory,
    exportBackup,
    recoverBackup,
    deleteBackup,
    setBackupPinned,
    cleanup,
    collectLegacy
  } = useFolderBackups(target);
  const [confirmingDeleteId, setConfirmingDeleteId] = useState<string | null>(null);
  const [confirmingRecoveryId, setConfirmingRecoveryId] = useState<string | null>(null);

  const copyPath = async (path: string) => {
    try {
      await navigator.clipboard.writeText(path);
    } catch {
      // 完整路径始终显示在界面上，剪贴板不可用时仍可手动复制。
    }
  };

  return (
    <aside className="tool-drawer folder-backup-drawer" aria-label="备份与恢复">
      <div className="tool-drawer-header">
        <div>
          <strong>备份与恢复</strong>
          <span>
            {target.kind === "unavailable"
              ? target.reason === "unsaved"
                ? "工作簿尚未保存到本地"
                : "云端工作簿"
              : data?.folderPath ?? "正在读取当前工作目录"}
          </span>
        </div>
        <button onClick={onClose} aria-label="关闭备份与恢复">
          ×
        </button>
      </div>
      <div className="folder-backup-content">
        <section className="folder-backup-session-summary">
          <div>
            <strong>本次会话可撤销</strong>
            <span>加载项关闭或刷新后，当前工作簿的内存撤销可能失效。</span>
          </div>
          <div className="folder-backup-session-counts">
            <span aria-label={`当前工作簿 ${workbookUndoCount} 条`}>
              当前工作簿 <b>{workbookUndoCount}</b> 条
            </span>
            <span aria-label={`文件夹执行 ${folderUndoCount} 条`}>
              文件夹执行 <b>{folderUndoCount}</b> 条
            </span>
          </div>
          {workbookUndoCount + folderUndoCount > 0 && (
            <small>请在对应的执行结果消息中撤销；这里只汇总当前会话能力。</small>
          )}
        </section>
        {target.kind === "unavailable" && (
          <div className="tool-empty folder-backup-empty">
            <i aria-hidden="true">↶</i>
            <strong>
              {target.reason === "unsaved"
                ? "工作簿尚未保存到本地"
                : "云端工作簿暂不支持"}
            </strong>
            <span>
              {target.reason === "unsaved"
                ? "Excel Bro 的备份存储在工作簿所在的文件夹下。保存工作簿后重新打开此面板即可查看。"
                : "OneDrive 和 SharePoint 上的工作簿暂不支持文件夹备份。如需使用备份功能，请下载到本地磁盘后操作。"}
            </span>
          </div>
        )}
        {target.kind !== "unavailable" && loading && (
          <div className="folder-backup-state">正在读取备份…</div>
        )}
        {target.kind !== "unavailable" && !loading && error && (
          <div className="folder-backup-feedback error" role="alert">
            {error}
          </div>
        )}
        {target.kind !== "unavailable" && !loading && data && (
          <>
            <section className="folder-backup-location">
              <div>
                <span>恢复点存储位置</span>
                <strong title={data.backupDirectory}>{data.backupDirectory}</strong>
              </div>
              <div className="folder-backup-location-actions">
                <button
                  type="button"
                  className="primary-button"
                  disabled={working || !data.openDirectoryToken}
                  onClick={() => void openBackupDirectory()}
                  title={
                    data.openDirectoryToken
                      ? "在文件资源管理器中打开统一备份位置"
                      : "当前没有可打开的备份目录"
                  }
                >
                  打开文件夹
                </button>
                <button
                  type="button"
                  className="secondary-button"
                  disabled={!data.openDirectoryToken}
                  onClick={() => void copyPath(data.backupDirectory)}
                  title={
                    data.openDirectoryToken
                      ? "复制统一备份位置"
                      : "当前没有可复制的备份目录"
                  }
                >
                  复制路径
                </button>
              </div>
            </section>

            <div className="folder-backup-toolbar">
              <span>{data.backups.length} 个恢复点</span>
              <div>
                <button
                  type="button"
                  className="secondary-button"
                  disabled={working}
                  onClick={() => void reload()}
                >
                  刷新
                </button>
                <button
                  type="button"
                  className="secondary-button"
                  disabled={working || !data.backupManagementToken}
                  onClick={() => void cleanup()}
                >
                  清理旧备份
                </button>
              </div>
            </div>

            {data.storageSummary && (
              <section className="folder-backup-storage-summary">
                <header>
                  <strong>
                    {data.storageSummary.complete ? "空间占用" : "已知空间占用"}
                  </strong>
                  <span>{formatBytes(data.storageSummary.totalBytes)}</span>
                </header>
                <div>
                  <span aria-label={`自动恢复点 ${formatBytes(data.storageSummary.automaticBytes)}`}>自动恢复点 <b>{formatBytes(data.storageSummary.automaticBytes)}</b></span>
                  <span aria-label={`长期保留 ${formatBytes(data.storageSummary.pinnedBytes)}`}>长期保留 <b>{formatBytes(data.storageSummary.pinnedBytes)}</b></span>
                  <span aria-label={`异常资料 ${formatBytes(data.storageSummary.abnormalBytes)}`}>异常资料 <b>{formatBytes(data.storageSummary.abnormalBytes)}</b></span>
                  <span>已取出副本 <b>{formatBytes(data.storageSummary.exportedBytes)}</b></span>
                  <span>旧版副本 <b>{formatBytes(data.storageSummary.legacyBytes)}</b></span>
                  <span>待清理资料 <b>{formatBytes(data.storageSummary.pendingCleanupBytes)}</b></span>
                </div>
                <small>
                  自动恢复资料软上限 {formatBytes(data.storageSummary.softLimitBytes)}；
                  期内、最新、主动保留和异常资料不会为满足上限而强制删除。
                </small>
              </section>
            )}

            {readOnly && data.backups.length > 0 && (
              <div className="folder-backup-feedback" role="status">
                当前工作簿入口可以管理这个已登记目录内的恢复点；恢复前会再次核对文件状态，
                不会扩大工作表数据范围。
              </div>
            )}

            {notice && (
              <div className="folder-backup-feedback" role="status">
                {notice}
              </div>
            )}
            {data.warnings.map((warning) => (
              <div className="folder-backup-feedback warning" key={warning}>
                {warning}
              </div>
            ))}

            {data.backups.length === 0 ? (
              <div className="tool-empty folder-backup-empty">
                <i aria-hidden="true">↶</i>
                <strong>当前没有恢复点</strong>
                <span>执行产生实际文件变化后，恢复点会集中显示在这里。</span>
              </div>
            ) : (
              <div className="folder-backup-list">
                {data.backups.map((backup) => {
                  const confirmingDelete = confirmingDeleteId === backup.undoId;
                  const confirmingRecovery =
                    confirmingRecoveryId === backup.undoId;
                  const restoreFileCount = backup.files.filter(
                    (file) => file.recoveryAction === "restore"
                  ).length;
                  const createdFileCount = backup.files.filter(
                    (file) => file.recoveryAction === "remove"
                  ).length;
                  const backupBusy = isBackupBusy(backup.status);
                  const ordinaryDeleteAllowed = canDeleteNormally(backup.status);
                  const managedRecoveryAllowed =
                    !backup.canUndo &&
                    managedDirectory?.status === "active" &&
                    backup.revision !== undefined &&
                    canRecoverFromManagedRecord(backup.status);
                  return (
                  <article className="folder-backup-card" key={backup.undoId}>
                    <header>
                      <div>
                        <strong>{backup.planTitle}</strong>
                        <span>{formatDate(backup.createdAt)}</span>
                      </div>
                      <em className={backup.canUndo ? "available" : "saved"}>
                        {backupStatusLabel(backup.status, backup.canUndo)}
                      </em>
                    </header>
                    <p className="folder-backup-recovery-summary">
                      {restoreFileCount > 0 && (
                        <span>
                          {restoreFileCount} 个原文件副本 · {formatBytes(backup.totalBytes)}
                        </span>
                      )}
                      {createdFileCount > 0 && (
                        <span>{createdFileCount} 个新建文件撤销记录</span>
                      )}
                    </p>
                    {createdFileCount > 0 && restoreFileCount === 0 && (
                      <div className="folder-backup-new-files-note">
                        本次执行只新建了文件；撤销时会移除这些文件，无需保存原文件副本。
                      </div>
                    )}
                    {backup.statusReason && (
                      <div
                        className={`folder-backup-feedback ${
                          backup.status === "needs_attention" ? "warning" : ""
                        }`}
                        role={backup.status === "needs_attention" ? "alert" : "status"}
                      >
                        {backup.statusReason}
                      </div>
                    )}
                    <details>
                      <summary>查看涉及文件</summary>
                      <ul>
                        {backup.files.map((file) => (
                          <li key={file.originalRelativePath}>
                            <em>
                              {file.recoveryAction === "restore"
                                ? "恢复原文件"
                                : "撤销时移除"}
                            </em>
                            {file.originalRelativePath}
                          </li>
                        ))}
                      </ul>
                    </details>
                    {confirmingRecovery ? (
                      <div
                        className="folder-backup-delete-confirm folder-backup-recovery-confirm"
                        role="alert"
                      >
                        <div>
                          <strong>恢复到本次执行前？</strong>
                          <span>
                            将恢复 {restoreFileCount} 个原文件
                            {createdFileCount > 0
                              ? `，并移除本次执行新建的 ${createdFileCount} 个文件`
                              : ""}
                            。执行前会重新核对文件指纹；发现后续人工修改时会安全停止，不会覆盖。
                          </span>
                        </div>
                        <div className="folder-backup-delete-confirm-actions">
                          <button
                            type="button"
                            className="secondary-button"
                            disabled={working}
                            onClick={() => setConfirmingRecoveryId(null)}
                          >
                            取消
                          </button>
                          <button
                            type="button"
                            className="primary-button"
                            disabled={working}
                            onClick={async () => {
                              const result = await recoverBackup(
                                backup.undoId,
                                backup.revision
                              );
                              if (result) setConfirmingRecoveryId(null);
                            }}
                          >
                            确认恢复
                          </button>
                        </div>
                      </div>
                    ) : confirmingDelete ? (
                      <div className="folder-backup-delete-confirm" role="alert">
                        <div>
                          <strong>
                            {backup.canUndo ? "删除并放弃撤销？" : "确认删除这份备份？"}
                          </strong>
                          <span>
                            {backup.canUndo
                              ? `删除后无法再撤销本次执行${
                                  backup.deleteInvalidatesUndoIds.length > 0
                                    ? `，并会使 ${backup.deleteInvalidatesUndoIds.length} 条较早撤销记录失效`
                                    : ""
                                }，但不会改变当前工作文件。`
                              : "这份历史副本将被永久删除，当前工作文件不会受影响。"}
                          </span>
                        </div>
                        <div className="folder-backup-delete-confirm-actions">
                          <button
                            type="button"
                            className="secondary-button"
                            disabled={working}
                            onClick={() => setConfirmingDeleteId(null)}
                          >
                            取消
                          </button>
                          <button
                            type="button"
                            className="danger-button"
                            disabled={working}
                            onClick={async () => {
                              const result = await deleteBackup(
                                backup.undoId,
                                backup.revision,
                                backup.deleteInvalidatesUndoIds
                              );
                              if (result) {
                                onBackupDeleted?.(result.remainingUndoIds);
                                setConfirmingDeleteId(null);
                              }
                            }}
                          >
                            确认删除
                          </button>
                        </div>
                      </div>
                    ) : (
                    <div className="folder-backup-card-actions">
                      {backup.canUndo && backup.undoSessionId && (
                        <button
                          type="button"
                          className="primary-button"
                          disabled={working}
                          onClick={async () => {
                            await onUndo(
                              backup.planId,
                              backup.undoId,
                              backup.undoSessionId!
                            );
                            await reload();
                          }}
                        >
                          撤销本次执行
                        </button>
                      )}
                      {managedRecoveryAllowed && (
                        <button
                          type="button"
                          className="primary-button"
                          disabled={working}
                          onClick={() => {
                            setConfirmingDeleteId(null);
                            setConfirmingRecoveryId(backup.undoId);
                          }}
                        >
                          {backup.status === "ready"
                            ? "恢复到执行前"
                            : "核对并恢复到执行前"}
                        </button>
                      )}
                      <button
                        type="button"
                        className="secondary-button"
                        disabled={
                          working ||
                          backup.totalBytes === 0 ||
                          backupBusy ||
                          !data.backupManagementToken
                        }
                        onClick={() => void exportBackup(backup.undoId)}
                      >
                        {backup.status === "needs_attention" ? "取出可用副本" : "取出副本"}
                      </button>
                      <button
                        type="button"
                        className="secondary-button"
                        disabled={working || backupBusy || !data.backupManagementToken}
                        onClick={() => void setBackupPinned(backup.undoId, !backup.pinned)}
                      >
                        {backup.pinned ? "取消保留" : "保留此副本"}
                      </button>
                      <button
                        type="button"
                        className="folder-backup-delete-trigger"
                        disabled={
                          working ||
                          !ordinaryDeleteAllowed ||
                          !data.backupManagementToken
                        }
                        title={
                          !ordinaryDeleteAllowed
                            ? backupBusy
                              ? "这份恢复资料仍在使用中，暂时不能删除"
                              : "这份恢复资料需要先核对，不能作为普通历史副本删除"
                            : !data.backupManagementToken
                              ? "备份管理凭证不存在，请刷新后重试"
                              : undefined
                        }
                        onClick={() => setConfirmingDeleteId(backup.undoId)}
                      >
                        {backup.canUndo ? "删除并放弃撤销" : "删除备份"}
                      </button>
                    </div>
                    )}
                  </article>
                  );
                })}
              </div>
            )}

            {data.legacyCandidates.length > 0 && (
              <section className="folder-legacy-backups">
                <div>
                  <strong>发现 {data.legacyCandidates.length} 个旧版散落备份</strong>
                  <span>只会归拢到当前工作目录的 legacy 文件夹，不会删除。</span>
                </div>
                <button
                  type="button"
                  className="secondary-button"
                  disabled={working || !data.backupManagementToken}
                  onClick={() => void collectLegacy()}
                >
                  归拢旧版备份
                </button>
              </section>
            )}
          </>
        )}
        {checkpointSection}
      </div>
    </aside>
  );
}
