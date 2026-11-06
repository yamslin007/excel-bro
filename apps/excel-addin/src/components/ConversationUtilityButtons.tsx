type Props = {
  disabled: boolean;
  backupCount: number;
  backupTitle?: string;
  onBackup: () => void;
  onHistory: () => void;
};

/** The recovery icon opens the existing backup/recovery surface; it never undoes directly. */
export function ConversationUtilityButtons({ disabled, backupCount, backupTitle, onBackup, onHistory }: Props) {
  return <>
    <button className="scope-utility-button folder-backup-entry" type="button" disabled={disabled}
      title={backupTitle || "备份与恢复"} aria-label="备份与恢复" onClick={onBackup}>
      <svg className="recovery-turn-icon" viewBox="0 0 24 24" aria-hidden="true" focusable="false">
        <path d="M9 3 1 10l8 7v-5c7 0 10.5 2.5 13 9 0-10-4-15-13-15V3Z" />
      </svg>
      {backupCount > 0 && <b>{backupCount}</b>}
    </button>
    <button className="scope-utility-button history-entry" type="button" disabled={disabled}
      title="历史记录" aria-label="历史对话" onClick={onHistory}>
      <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
        <circle cx="12" cy="12" r="8.5" /><path d="M12 7v5l3.5 2" />
      </svg>
    </button>
  </>;
}
