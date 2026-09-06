interface WorkDirectoryGuideProps {
  busy?: boolean;
  onDismiss: () => void;
  onChoose: () => void;
}

interface WorkDirectoryGuideVisibility {
  ready: boolean;
  previousStepComplete: boolean;
  folderConfigured: boolean;
  dismissed: boolean;
  blocked: boolean;
}

export function shouldShowWorkDirectoryGuide({
  ready,
  previousStepComplete,
  folderConfigured,
  dismissed,
  blocked
}: WorkDirectoryGuideVisibility): boolean {
  return (
    ready &&
    previousStepComplete &&
    !folderConfigured &&
    !dismissed &&
    !blocked
  );
}

/** 首次产品介绍中的可跳过工作目录设置卡片。 */
export function WorkDirectoryGuide({
  busy = false,
  onDismiss,
  onChoose
}: WorkDirectoryGuideProps) {
  return (
    <section
      className="first-work-directory-guide"
      aria-label="首次工作目录设置引导"
    >
      <div className="first-work-directory-guide-icon" aria-hidden="true">
        <span>⌂</span>
      </div>
      <div className="first-work-directory-guide-copy">
        <strong>设置常用工作目录</strong>
        <span>
          选择你平时存放 Excel 文件的文件夹。批量处理时无需重复选择；只有你随后勾选的工作表会被授权。
        </span>
        <small>
          可以选择已有文件夹，也可以在选择窗口中新建；恢复点只会在发生写入后创建。
        </small>
      </div>
      <div className="first-work-directory-guide-actions">
        <button type="button" onClick={onDismiss} disabled={busy}>
          以后再设置
        </button>
        <button
          type="button"
          className="primary"
          onClick={onChoose}
          disabled={busy}
        >
          选择工作目录
        </button>
      </div>
    </section>
  );
}
