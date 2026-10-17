import type { ToolEligibility } from "../storage";

/**
 * 「保存为工具」确认对话框（双端共用）。
 *
 * 这是把 Excel 任务窗格里内联的固化确认流程抽出来的共用呈现层：名称、用途说明、
 * 固化检查（固定内容 / 覆盖清除等风险项）+ 对应确认勾选。资格判断本身不在这里，
 * 由 `analyzeToolEligibility`（shared storage）产出 `eligibility` 传入。
 *
 * 安全约束由本组件与调用方共同保证：
 *  - 所有确认项默认不勾选（初始值由调用方 `beginSaveCandidate` 重置）；
 *  - 不可保存（`eligibility.blocked`）的计划，勾选框根本不会出现，确认按钮恒禁用，
 *    不能靠勾选绕过；
 *  - 需要确认的风险项未勾选时，确认按钮禁用；
 *  - `blockedReason` 非空（候选在打开后失效：计划已变 / 执行已撤销 / 资格不再成立）
 *    时确认按钮恒禁用，调用方必须在确认时重新检查。
 */
export interface SaveToolDialogProps {
  toolName: string;
  toolDescription: string;
  eligibility: ToolEligibility | null;
  approveFixedContent: boolean;
  approveDestructive: boolean;
  onToolNameChange: (value: string) => void;
  onToolDescriptionChange: (value: string) => void;
  onApproveFixedContentChange: (value: boolean) => void;
  onApproveDestructiveChange: (value: boolean) => void;
  onCancel: () => void;
  onConfirm: () => void;
  /** 候选在打开对话框后失效时的阻止原因；非空则确认按钮恒禁用。 */
  blockedReason?: string | null;
  title?: string;
  subtitle?: string;
  note?: string;
}

export function SaveToolDialog({
  toolName,
  toolDescription,
  eligibility,
  approveFixedContent,
  approveDestructive,
  onToolNameChange,
  onToolDescriptionChange,
  onApproveFixedContentChange,
  onApproveDestructiveChange,
  onCancel,
  onConfirm,
  blockedReason = null,
  title = "保存为工具",
  subtitle = "验证后可从「我的工具」重复使用",
  note = "只保存白名单操作，不会保存 API Key，也不会执行系统命令。"
}: SaveToolDialogProps) {
  const required = eligibility?.requiredApprovals ?? [];
  const confirmDisabled =
    !toolName.trim() ||
    Boolean(blockedReason) ||
    !eligibility ||
    Boolean(eligibility.blocked) ||
    (required.includes("fixedContent") && !approveFixedContent) ||
    (required.includes("destructive") && !approveDestructive);

  return (
    <div className="tool-dialog-backdrop" role="presentation">
      <section
        className="tool-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={title}
      >
        <div className="tool-dialog-title">
          <div>
            <strong>{title}</strong>
            <span>{subtitle}</span>
          </div>
          <button type="button" onClick={onCancel} aria-label="关闭">
            ×
          </button>
        </div>
        <label>
          <span>工具名称</span>
          <input
            value={toolName}
            maxLength={60}
            autoFocus
            onChange={(event) => onToolNameChange(event.target.value)}
          />
        </label>
        <label>
          <span>用途说明</span>
          <textarea
            value={toolDescription}
            maxLength={240}
            onChange={(event) => onToolDescriptionChange(event.target.value)}
          />
        </label>
        <div className="tool-dialog-note">{note}</div>
        {blockedReason && (
          <p role="alert" className="tool-save-blocked-reason">
            {blockedReason}
          </p>
        )}
        {eligibility && eligibility.issues.length > 0 && (
          <div className="tool-save-review">
            <strong>固化检查</strong>
            {eligibility.issues.map((issue) => (
              <div
                className={`tool-save-issue ${issue.severity}`}
                key={issue.code}
              >
                <span>{issue.message}</span>
                {issue.approval === "fixedContent" && (
                  <label>
                    <input
                      type="checkbox"
                      checked={approveFixedContent}
                      onChange={(event) =>
                        onApproveFixedContentChange(event.target.checked)
                      }
                    />
                    我确认以后仍写入这些固定内容
                  </label>
                )}
                {issue.approval === "destructive" && (
                  <label>
                    <input
                      type="checkbox"
                      checked={approveDestructive}
                      onChange={(event) =>
                        onApproveDestructiveChange(event.target.checked)
                      }
                    />
                    我确认此工具包含覆盖或删除操作
                  </label>
                )}
              </div>
            ))}
          </div>
        )}
        <div className="tool-dialog-actions">
          <button type="button" className="secondary-button" onClick={onCancel}>
            取消
          </button>
          <button
            type="button"
            className="primary-button"
            disabled={confirmDisabled}
            onClick={onConfirm}
          >
            保存到我的工具
          </button>
        </div>
      </section>
    </div>
  );
}
