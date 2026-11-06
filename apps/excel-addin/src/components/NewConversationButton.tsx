type Props = {
  disabled: boolean;
  onClick: () => void;
};

/** 会话入口与本次数据范围同排；两端保留各自既有的新对话处理器。 */
export function NewConversationButton({ disabled, onClick }: Props) {
  return (
    <button
      className="scope-new-chat new-chat-entry"
      type="button"
      disabled={disabled}
      onClick={onClick}
      title="新对话"
      aria-label="新对话"
    >
      <span aria-hidden="true">＋</span>
      <span>新对话</span>
    </button>
  );
}
