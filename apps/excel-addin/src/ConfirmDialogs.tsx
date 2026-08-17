// 确认类对话框：删除工具 / 删除历史对话 / 保存为工具。
// 纯展示组件，JSX 从 App.tsx 逐字搬移；状态与回调由 props 传入。
import type { Dispatch, SetStateAction } from "react";
import type { AnalysisPlan } from "./contracts";
import type { ToolEligibility } from "./storage";
import type { ChatHistoryState } from "./types/chat";
import type { PendingToolDeletion } from "./hooks/useToolManagement";

interface ConfirmDialogsProps {
  pendingToolDeletion: PendingToolDeletion | null;
  setPendingToolDeletion: Dispatch<SetStateAction<PendingToolDeletion | null>>;
  confirmToolDeletion: () => void;
  pendingDeleteConversationId: string | null;
  setPendingDeleteConversationId: Dispatch<SetStateAction<string | null>>;
  confirmDeleteConversation: () => void;
  chatHistory: ChatHistoryState;
  saveCandidate: AnalysisPlan | null;
  setSaveCandidate: Dispatch<SetStateAction<AnalysisPlan | null>>;
  toolName: string;
  setToolName: Dispatch<SetStateAction<string>>;
  toolDescription: string;
  setToolDescription: Dispatch<SetStateAction<string>>;
  saveEligibility: ToolEligibility | null;
  approveFixedContent: boolean;
  setApproveFixedContent: Dispatch<SetStateAction<boolean>>;
  approveDestructive: boolean;
  setApproveDestructive: Dispatch<SetStateAction<boolean>>;
  confirmSaveTool: () => void;
}

export default function ConfirmDialogs({
  pendingToolDeletion,
  setPendingToolDeletion,
  confirmToolDeletion,
  pendingDeleteConversationId,
  setPendingDeleteConversationId,
  confirmDeleteConversation,
  chatHistory,
  saveCandidate,
  setSaveCandidate,
  toolName,
  setToolName,
  toolDescription,
  setToolDescription,
  saveEligibility,
  approveFixedContent,
  setApproveFixedContent,
  approveDestructive,
  setApproveDestructive,
  confirmSaveTool
}: ConfirmDialogsProps) {
  return (
    <>
      {pendingToolDeletion && (
        <div className="tool-dialog-backdrop" role="presentation">
          <section
            className="tool-dialog history-delete-dialog"
            role="dialog"
            aria-modal="true"
            aria-label="删除保存的工具"
          >
            <div className="tool-dialog-title">
              <div>
                <strong>删除这个工具？</strong>
                <span>删除后无法恢复，但不会影响工作簿中的数据</span>
              </div>
              <button
                onClick={() => setPendingToolDeletion(null)}
                aria-label="关闭"
              >
                ×
              </button>
            </div>
            <p>确定从工具箱删除「{pendingToolDeletion.name}」吗？</p>
            <div className="tool-dialog-actions">
              <button
                className="secondary-button"
                onClick={() => setPendingToolDeletion(null)}
              >
                取消
              </button>
              <button
                className="danger-button"
                onClick={confirmToolDeletion}
              >
                确认删除
              </button>
            </div>
          </section>
        </div>
      )}

      {pendingDeleteConversationId &&
        chatHistory.conversations
          .filter(
            (conversation) =>
              conversation.id === pendingDeleteConversationId
          )
          .map((conversation) => (
            <div
              className="tool-dialog-backdrop"
              role="presentation"
              key={conversation.id}
            >
              <section
                className="tool-dialog history-delete-dialog"
                role="dialog"
                aria-modal="true"
                aria-label="删除历史对话"
              >
                <div className="tool-dialog-title">
                  <div>
                    <strong>删除历史对话？</strong>
                    <span>此操作无法撤销</span>
                  </div>
                  <button
                    onClick={() => setPendingDeleteConversationId(null)}
                    aria-label="关闭"
                  >
                    ×
                  </button>
                </div>
                <p>
                  确定删除「{conversation.title}」吗？删除后不会影响已保存的工具。
                </p>
                <div className="tool-dialog-actions">
                  <button
                    className="secondary-button"
                    onClick={() => setPendingDeleteConversationId(null)}
                  >
                    取消
                  </button>
                  <button
                    className="danger-button"
                    onClick={confirmDeleteConversation}
                  >
                    确认删除
                  </button>
                </div>
              </section>
            </div>
          ))}

      {saveCandidate && (
        <div className="tool-dialog-backdrop" role="presentation">
          <section
            className="tool-dialog"
            role="dialog"
            aria-modal="true"
            aria-label="保存为工具"
          >
            <div className="tool-dialog-title">
              <div>
                <strong>保存为工具</strong>
                <span>验证后可从“我的工具”重复使用</span>
              </div>
              <button onClick={() => setSaveCandidate(null)}>×</button>
            </div>
            <label>
              <span>工具名称</span>
              <input
                value={toolName}
                maxLength={60}
                autoFocus
                onChange={(event) => setToolName(event.target.value)}
              />
            </label>
            <label>
              <span>用途说明</span>
              <textarea
                value={toolDescription}
                maxLength={240}
                onChange={(event) => setToolDescription(event.target.value)}
              />
            </label>
            <div className="tool-dialog-note">
              只保存白名单 Excel 操作，不会保存 API Key，也不会执行系统命令。
            </div>
            {saveEligibility && saveEligibility.issues.length > 0 && (
              <div className="tool-save-review">
                <strong>固化检查</strong>
                {saveEligibility.issues.map((issue) => (
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
                            setApproveFixedContent(event.target.checked)
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
                            setApproveDestructive(event.target.checked)
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
              <button
                className="secondary-button"
                onClick={() => setSaveCandidate(null)}
              >
                取消
              </button>
              <button
                className="primary-button"
                disabled={
                  !toolName.trim() ||
                  Boolean(saveEligibility?.blocked) ||
                  Boolean(
                    saveEligibility?.requiredApprovals.includes(
                      "fixedContent"
                    ) && !approveFixedContent
                  ) ||
                  Boolean(
                    saveEligibility?.requiredApprovals.includes(
                      "destructive"
                    ) && !approveDestructive
                  )
                }
                onClick={confirmSaveTool}
              >
                保存到我的工具
              </button>
            </div>
          </section>
        </div>
      )}
    </>
  );
}
