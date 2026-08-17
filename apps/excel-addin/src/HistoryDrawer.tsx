// 历史对话抽屉：按更新时间排序展示会话，可打开 / 新建 / 删除。
// 纯展示组件，JSX 与 formatConversationTime 从 App.tsx 逐字搬移。
import type { Dispatch, SetStateAction } from "react";
import type { ChatHistoryState } from "./types/chat";

interface HistoryDrawerProps {
  historyOpen: boolean;
  chatHistory: ChatHistoryState;
  busy: boolean;
  setHistoryOpen: Dispatch<SetStateAction<boolean>>;
  openConversation: (conversationId: string) => void;
  deleteConversation: (conversationId: string) => void;
  newChat: () => void;
}

function formatConversationTime(value: string): string {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return new Intl.DateTimeFormat("zh-CN", {
    month: "numeric",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  }).format(date);
}

export default function HistoryDrawer({
  historyOpen,
  chatHistory,
  busy,
  setHistoryOpen,
  openConversation,
  deleteConversation,
  newChat
}: HistoryDrawerProps) {
  if (!historyOpen) return null;
  return (
    <aside className="tool-drawer history-drawer" aria-label="历史对话">
      <div className="tool-drawer-header">
        <div>
          <strong>历史对话</strong>
          <span>新建对话不会覆盖之前的记录</span>
        </div>
        <button onClick={() => setHistoryOpen(false)} aria-label="关闭">
          ×
        </button>
      </div>
      <div className="history-list">
        {[...chatHistory.conversations]
          .sort((left, right) =>
            right.updatedAt.localeCompare(left.updatedAt)
          )
          .map((conversation) => {
            const active =
              conversation.id === chatHistory.activeConversationId;
            const questionCount = conversation.messages.filter(
              (message) => message.role === "user"
            ).length;
            return (
              <article
                className={`history-item${active ? " active" : ""}`}
                key={conversation.id}
              >
                <button
                  className="history-open-button"
                  onClick={() => openConversation(conversation.id)}
                  disabled={busy}
                >
                  <span className="history-title-row">
                    <strong>{conversation.title}</strong>
                    {active && <em>当前</em>}
                  </span>
                  <span>
                    {formatConversationTime(conversation.updatedAt)}
                    {questionCount > 0
                      ? ` · ${questionCount} 条提问`
                      : " · 尚未提问"}
                  </span>
                </button>
                <button
                  className="history-delete-button"
                  onClick={() => deleteConversation(conversation.id)}
                  disabled={busy}
                  title={`删除「${conversation.title}」`}
                  aria-label={`删除「${conversation.title}」`}
                >
                  删除
                </button>
              </article>
            );
          })}
      </div>
      <button
        className="history-new-button"
        onClick={newChat}
        disabled={busy}
      >
        ＋ 新建对话
      </button>
    </aside>
  );
}
