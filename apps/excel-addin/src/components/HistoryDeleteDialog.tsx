export const HISTORY_MENU_DESCRIPTION = "查看并打开已保存的对话";

export function HistoryDeleteDialog({ title, onCancel, onConfirm, disabled = false }: {
  title: string;
  onCancel: () => void;
  onConfirm: () => void;
  disabled?: boolean;
}) {
  return <div className="tool-dialog-backdrop" role="presentation">
    <section className="tool-dialog history-delete-dialog" role="dialog" aria-modal="true" aria-label="删除历史对话">
      <div className="tool-dialog-title">
        <div><strong>删除历史对话？</strong><span>此操作无法撤销</span></div>
        <button type="button" onClick={onCancel} aria-label="关闭">×</button>
      </div>
      <p>确定删除「{title}」吗？仅删除聊天记录，不撤销执行，也不删除已保存的工具、执行记录、恢复材料或未完成任务。</p>
      <div className="tool-dialog-actions">
        <button type="button" className="secondary-button" onClick={onCancel}>取消</button>
        <button type="button" className="danger-button" disabled={disabled} onClick={onConfirm}>确认删除</button>
      </div>
    </section>
  </div>;
}
