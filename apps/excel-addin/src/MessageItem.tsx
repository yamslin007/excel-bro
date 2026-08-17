// 消息流：消息列表 + 运行中活动卡片 + 滚动锚点。
// 纯展示组件，JSX 从 App.tsx 逐字搬移；状态与回调由 props 传入。
import type { Dispatch, RefObject, SetStateAction } from "react";
import type {
  AnalysisPlan,
  ExecutionUndoSnapshot,
  IntentOption
} from "./contracts";
import { actionLabel, formatGenerateMs, formatStepElapsed } from "./format";
import type { ActivityProgress } from "./hooks/useActivityProgress";
import type {
  ChatMessage,
  FunctionPreview,
  Status
} from "./types/chat";

interface MessageItemProps {
  messages: ChatMessage[];
  messageEndRef: RefObject<HTMLDivElement | null>;
  busy: boolean;
  activity: ActivityProgress | null;
  activitySeconds: number;
  status: Status;
  stopTurn: () => void;
  verifiedPlanIds: Set<string>;
  clarificationDrafts: Record<string, string>;
  setClarificationDrafts: Dispatch<SetStateAction<Record<string, string>>>;
  copiedMessageId: string | null;
  copiedFunctionPreviewId: string | null;
  recomputeReusedMessage: (message: ChatMessage) => void;
  saveQueryFromMessage: (message: ChatMessage) => void;
  resolveClarification: (
    message: ChatMessage,
    resolution: string,
    label: string
  ) => Promise<void>;
  editClarificationScope: (message: ChatMessage) => void;
  cancelClarification: (message: ChatMessage) => void;
  markFunctionPreview: (
    messageId: string,
    patch: Partial<FunctionPreview>
  ) => void;
  pickFunctionTarget: (message: ChatMessage) => Promise<void>;
  cancelFunctionPreview: (message: ChatMessage) => void;
  confirmFunctionTarget: (message: ChatMessage) => Promise<void>;
  copyFunctionFormula: (message: ChatMessage) => Promise<void>;
  applyFunctionPreview: (message: ChatMessage) => Promise<void>;
  beginSaveTool: (plan: AnalysisPlan) => void;
  runPlan: (plan: AnalysisPlan) => Promise<void>;
  copyMessage: (message: ChatMessage) => Promise<void>;
  lastUndoSnapshot: ExecutionUndoSnapshot | null;
  undoLastExecution: () => Promise<void>;
  replyToMessage: (message: ChatMessage) => void;
}

export default function MessageItem({
  messages,
  messageEndRef,
  busy,
  activity,
  activitySeconds,
  status,
  stopTurn,
  verifiedPlanIds,
  clarificationDrafts,
  setClarificationDrafts,
  copiedMessageId,
  copiedFunctionPreviewId,
  recomputeReusedMessage,
  saveQueryFromMessage,
  resolveClarification,
  editClarificationScope,
  cancelClarification,
  markFunctionPreview,
  pickFunctionTarget,
  cancelFunctionPreview,
  confirmFunctionTarget,
  copyFunctionFormula,
  applyFunctionPreview,
  beginSaveTool,
  runPlan,
  copyMessage,
  lastUndoSnapshot,
  undoLastExecution,
  replyToMessage
}: MessageItemProps) {
  return (
    <section className="message-stream" aria-live="polite">
      {messages.map((message) => (
        <article
          className={`message-row ${message.role}`}
          key={message.id}
        >
          <div className="message-content">
            {message.role === "assistant" && (
              <span className="message-author">
                Excel Bro
                {message.provider && (
                  <em>
                    {message.clarification
                      ? "需求确认"
                      : message.provider === "model"
                        ? "Agent"
                        : "基础模式"}
                  </em>
                )}
              </span>
            )}
            {message.reused && (
              <div className="reuse-badge">
                <span>♻ 复用上次结果（数据未变化）</span>
                <button
                  type="button"
                  className="link-button"
                  disabled={busy}
                  onClick={() => recomputeReusedMessage(message)}
                >
                  仍要重新计算
                </button>
              </div>
            )}
            {message.text && <p className="message-text">{message.text}</p>}
            {message.intentMemory?.toolRequest && message.resultContext && (
              <button
                type="button"
                className="secondary-button"
                disabled={busy}
                onClick={() => saveQueryFromMessage(message)}
              >
                保存为固化查询
              </button>
            )}
            {message.attachmentNames &&
              message.attachmentNames.length > 0 && (
                <div className="message-attachments">
                  {message.attachmentNames.map((name, index) => (
                    <span key={`${message.id}-attachment-${index}`}>
                      ▧ {name}
                    </span>
                  ))}
                </div>
              )}
            {message.clarification && (
              <section
                className={`clarification-card ${message.clarification.status}`}
                aria-label="确认需求"
              >
                <div className="clarification-heading">
                  <div>
                    <span>执行前确认</span>
                    <strong>
                      {message.clarification.status === "pending"
                        ? "我需要确认你的需求"
                        : message.clarification.status === "resolving"
                          ? "正在理解你的补充"
                        : message.clarification.status === "resolved"
                          ? "需求已确认"
                          : message.clarification.status === "invalidated"
                            ? "数据范围已变化"
                            : "需求已取消"}
                    </strong>
                  </div>
                  <em>{message.clarification.scopeLabel}</em>
                </div>
                <p className="clarification-summary">
                  {message.clarification.summary}
                </p>
                <strong className="clarification-question">
                  {message.clarification.question}
                </strong>
                <span className="clarification-reason">
                  {message.clarification.reason}
                </span>

                {message.clarification.status === "pending" ? (
                  <>
                    <div className="clarification-options">
                      {message.clarification.options.map(
                        (option: IntentOption) => (
                          <button
                            key={option.id}
                            disabled={busy}
                            onClick={() => {
                              if (option.action === "editScope") {
                                editClarificationScope(message);
                                return;
                              }
                              void resolveClarification(
                                message,
                                option.resolution,
                                option.label
                              );
                            }}
                          >
                            <strong>{option.label}</strong>
                            <span>{option.description}</span>
                          </button>
                        )
                      )}
                    </div>
                    <span className="clarification-custom-hint">
                      都不符合？请直接写下你的想法。
                    </span>
                    <div className="clarification-custom-input">
                      <textarea
                        value={
                          clarificationDrafts[
                            message.clarification.id
                          ] ?? ""
                        }
                        disabled={busy}
                        rows={2}
                        placeholder="例如：按所有已选工作表汇总后，再比较每个分类的整体占比"
                        onChange={(event) =>
                          setClarificationDrafts((current) => ({
                            ...current,
                            [message.clarification!.id]:
                              event.target.value
                          }))
                        }
                      />
                      <button
                        disabled={
                          busy ||
                          !(
                            clarificationDrafts[
                              message.clarification.id
                            ] ?? ""
                          ).trim()
                        }
                        onClick={() => {
                          const customAnswer = (
                            clarificationDrafts[
                              message.clarification!.id
                            ] ?? ""
                          ).trim();
                          void resolveClarification(
                            message,
                            customAnswer,
                            customAnswer
                          );
                        }}
                      >
                        提交补充
                      </button>
                    </div>
                    <div className="clarification-actions">
                      <button
                        disabled={busy}
                        onClick={() => editClarificationScope(message)}
                      >
                        修改数据范围
                      </button>
                      <button
                        disabled={busy}
                        onClick={() =>
                          void resolveClarification(
                            message,
                            "按用户原始需求继续；若仍需自行判断，请在结果中明确列出采用的假设。",
                            "按原话继续"
                          )
                        }
                      >
                        按原话继续
                      </button>
                      <button
                        disabled={busy}
                        onClick={() => cancelClarification(message)}
                      >
                        取消
                      </button>
                    </div>
                  </>
                ) : (
                  <div className="clarification-outcome">
                    {message.clarification.status === "resolving"
                      ? "正在结合你的回答重新判断，不会执行任何写入。"
                      : message.clarification.status === "resolved"
                      ? `已选择：${message.clarification.resolvedLabel}`
                      : message.clarification.status === "invalidated"
                        ? "请检查数据范围后重新发送需求。"
                        : "没有进行后续分析或写入。"}
                  </div>
                )}
              </section>
            )}
            {message.activityLog &&
              message.activityLog.steps.length > 0 && (
                <details className="activity-log-card">
                  <summary>
                    执行过程 · {message.activityLog.steps.length} 步 ·{" "}
                    {formatStepElapsed(message.activityLog.totalMs)}
                  </summary>
                  <ul className="activity-steps">
                    {message.activityLog.steps.map((step, index) => (
                      <li key={`${step.label}-${index}`}>
                        <span
                          className="activity-step-check"
                          aria-hidden="true"
                        >
                          ✓
                        </span>
                        <div className="activity-step-body">
                          <span className="activity-step-label">
                            {step.label}
                          </span>
                          {step.note && (
                            <span className="activity-step-note">
                              {step.note}
                            </span>
                          )}
                          {step.detail && (
                            <details className="activity-step-detail">
                              <summary>查看详情</summary>
                              <pre>{step.detail}</pre>
                            </details>
                          )}
                        </div>
                        <span className="activity-step-time">
                          {formatStepElapsed(step.elapsedMs)}
                        </span>
                      </li>
                    ))}
                  </ul>
                </details>
              )}
            {message.verification && (
              <div
                className={`verification-card ${message.verification.status}`}
              >
                <strong>
                  {message.verification.status === "verified"
                    ? "验证通过"
                    : message.verification.status === "executed_unverified"
                      ? "已执行，部分操作未独立验证"
                      : "执行完成，但验证未通过"}
                </strong>
                <span>
                  {
                    message.verification.checks.filter((check) => check.passed)
                      .length
                  }
                  /{message.verification.checks.length} 项符合预期
                  {message.verification.unverifiedActions.length > 0
                    ? `；${message.verification.unverifiedActions.length} 步缺少独立验收`
                    : ""}
                </span>
                {message.verification.status === "failed" && (
                  <ul>
                    {message.verification.checks
                      .filter((check) => !check.passed)
                      .map((check, index) => (
                        <li key={`${message.id}-verification-${index}`}>
                          {check.message}
                        </li>
                      ))}
                  </ul>
                )}
                {message.verification.status ===
                  "executed_unverified" && (
                  <ul>
                    {message.verification.unverifiedActions.map(
                      (action) => (
                        <li
                          key={`${message.id}-unverified-${action.index}`}
                        >
                          {action.message}
                        </li>
                      )
                    )}
                  </ul>
                )}
              </div>
            )}
            {message.executedPlanId &&
              lastUndoSnapshot?.planId === message.executedPlanId && (
                <div className="undo-row">
                  <button
                    type="button"
                    className="secondary-button"
                    disabled={busy}
                    onClick={() => void undoLastExecution()}
                  >
                    ↶ 撤销本次执行
                  </button>
                  <span className="undo-row-hint">
                    只保留最近一次执行的撤销数据
                  </span>
                </div>
              )}
            {message.functionPreview &&
              message.functionPreview.phase === "target" && (
                <div className="inline-plan">
                  <div className="inline-plan-title">
                    <div>
                      <span>写入单元格内</span>
                      <strong>{message.functionPreview.sheet}</strong>
                    </div>
                  </div>
                  <p className="function-preview-explain">
                    确定公式要写到哪个单元格或区域，再生成。已为你预填建议位置，可直接改，或先在表里点选目标格、再点「拾取当前选区」。
                  </p>
                  <div className="function-write-target">
                    <input
                      className="function-write-target-input"
                      type="text"
                      value={message.functionPreview.writeTarget}
                      disabled={busy || message.functionPreview.pickingTarget}
                      placeholder="如 E2 或 E2:E20"
                      spellCheck={false}
                      onChange={(event) =>
                        markFunctionPreview(message.id, {
                          writeTarget: event.target.value,
                          targetError: undefined
                        })
                      }
                    />
                    <button
                      className="secondary-button"
                      disabled={busy}
                      onClick={() => void pickFunctionTarget(message)}
                    >
                      {message.functionPreview.pickingTarget
                        ? "读取中…"
                        : "拾取当前选区"}
                    </button>
                  </div>
                  <div className="inline-notes function-pick-hint">
                    拾取方式：先在工作表里点一下（或框选）目标单元格，再点上面的「拾取当前选区」，会自动填入那一刻选中的位置。
                  </div>
                  {message.functionPreview.targetError && (
                    <div className="function-write-target-error">
                      {message.functionPreview.targetError}
                    </div>
                  )}
                  <div className="inline-plan-buttons">
                    <button
                      className="secondary-button"
                      disabled={busy}
                      onClick={() => cancelFunctionPreview(message)}
                    >
                      取消
                    </button>
                    <button
                      className="primary-button"
                      disabled={busy}
                      onClick={() => void confirmFunctionTarget(message)}
                    >
                      {status === "planning" ? "正在生成…" : "确定并生成"}
                    </button>
                  </div>
                </div>
              )}
            {message.functionPreview &&
              message.functionPreview.phase === "preview" && (
                <div className="inline-plan">
                  <div className="inline-plan-title">
                    <div>
                      <span>公式预览</span>
                      <strong>{message.functionPreview.writeTarget}</strong>
                    </div>
                    {message.functionPreview.generateMs !== undefined && (
                      <span className="function-preview-timing">
                        生成耗时 {formatGenerateMs(message.functionPreview.generateMs)}
                      </span>
                    )}
                  </div>

                  {!message.functionPreview.applied &&
                    !message.functionPreview.cancelled && (
                      <div className="function-preview-versions">
                        <button
                          className={
                            message.functionPreview.version === "compat"
                              ? "version-tab active"
                              : "version-tab"
                          }
                          onClick={() =>
                            markFunctionPreview(message.id, {
                              version: "compat"
                            })
                          }
                        >
                          兼容版 · 2016/2019
                        </button>
                        <button
                          className={
                            message.functionPreview.version === "modern"
                              ? "version-tab active"
                              : "version-tab"
                          }
                          onClick={() =>
                            markFunctionPreview(message.id, {
                              version: "modern"
                            })
                          }
                        >
                          现代版 · 365/2021
                        </button>
                      </div>
                    )}

                  <pre className="function-preview-formula">
                    {message.functionPreview.version === "modern"
                      ? message.functionPreview.modernFormula
                      : message.functionPreview.compatFormula}
                  </pre>

                  <div className="function-preview-trial">
                    <span>首格试算</span>
                    <strong>
                      {message.functionPreview.version === "modern"
                        ? message.functionPreview.modernResult
                        : message.functionPreview.compatResult}
                    </strong>
                  </div>

                  {(message.functionPreview.version === "modern"
                    ? message.functionPreview.modernExplanation
                    : message.functionPreview.compatExplanation) && (
                    <p className="function-preview-explain">
                      {message.functionPreview.version === "modern"
                        ? message.functionPreview.modernExplanation
                        : message.functionPreview.compatExplanation}
                    </p>
                  )}

                  {message.functionPreview.applied ? (
                    <div className="inline-notes">
                      已写入 {message.functionPreview.appliedTarget}。
                    </div>
                  ) : message.functionPreview.cancelled ? (
                    <div className="inline-notes">已取消。</div>
                  ) : (
                    <>
                      <div className="function-write-target">
                        <input
                          className="function-write-target-input"
                          type="text"
                          value={message.functionPreview.writeTarget}
                          disabled={
                            busy || message.functionPreview.pickingTarget
                          }
                          placeholder="如 E2 或 E2:E20"
                          spellCheck={false}
                          onChange={(event) =>
                            markFunctionPreview(message.id, {
                              writeTarget: event.target.value,
                              targetError: undefined
                            })
                          }
                        />
                        <button
                          className="secondary-button"
                          disabled={busy}
                          onClick={() => void pickFunctionTarget(message)}
                        >
                          {message.functionPreview.pickingTarget
                            ? "读取中…"
                            : "拾取当前选区"}
                        </button>
                      </div>
                      {message.functionPreview.targetError && (
                        <div className="function-write-target-error">
                          {message.functionPreview.targetError}
                        </div>
                      )}
                      <div className="inline-plan-buttons">
                        <button
                          className="secondary-button"
                          disabled={busy}
                          onClick={() => void copyFunctionFormula(message)}
                        >
                          {copiedFunctionPreviewId === message.id
                            ? "已复制"
                            : "复制公式"}
                        </button>
                        <button
                          className="secondary-button"
                          disabled={busy}
                          onClick={() => cancelFunctionPreview(message)}
                        >
                          取消
                        </button>
                        <button
                          className="primary-button"
                          disabled={busy}
                          onClick={() => void applyFunctionPreview(message)}
                        >
                          {status === "executing" ? "正在写入…" : "确认写入"}
                        </button>
                      </div>
                    </>
                  )}
                </div>
              )}
            {message.plan && (
              <div className="inline-plan">
                <div className="inline-plan-title">
                  <div>
                    <span>准备就绪</span>
                    <strong>{message.plan.title}</strong>
                  </div>
                </div>

                <div className="inline-plan-buttons">
                  <button
                    className="secondary-button"
                    disabled={busy || !verifiedPlanIds.has(message.plan.id)}
                    title={
                      verifiedPlanIds.has(message.plan.id)
                        ? "保存为可重复使用的个人工具"
                        : "请先执行，并通过结果验证"
                    }
                    onClick={() => beginSaveTool(message.plan!)}
                  >
                    {verifiedPlanIds.has(message.plan.id)
                      ? "保存为工具"
                      : "验证后保存"}
                  </button>
                  <button
                    className="primary-button"
                    disabled={busy}
                    onClick={() => void runPlan(message.plan!)}
                  >
                    {status === "executing" ? "正在执行…" : "执行到 Excel"}
                  </button>
                </div>

                <details className="plan-details">
                  <summary>
                    查看执行细节
                    <span>{message.plan.actions.length} 步</span>
                  </summary>

                  {message.plan.assumptions.length > 0 && (
                    <div className="inline-notes">
                      <strong>采用的信息</strong>
                      <ul>
                        {message.plan.assumptions.map((item) => (
                          <li key={item}>{item}</li>
                        ))}
                      </ul>
                    </div>
                  )}

                  {message.plan.warnings.length > 0 && (
                    <div className="inline-notes warning">
                      <strong>补充说明</strong>
                      <ul>
                        {message.plan.warnings.map((item) => (
                          <li key={item}>{item}</li>
                        ))}
                      </ul>
                    </div>
                  )}

                  <ol className="inline-actions">
                    {message.plan.actions.map((action, index) => (
                      <li key={`${message.id}-${action.type}-${index}`}>
                        <span>{index + 1}</span>
                        {actionLabel(action)}
                      </li>
                    ))}
                  </ol>
                </details>
              </div>
            )}
            {message.text && (
              <div className="message-copy-row">
                {message.role === "assistant" && message.resultContext && (
                  <button
                    className="message-copy-button"
                    type="button"
                    disabled={busy}
                    title="针对这条结果解释或纠错"
                    aria-label="回复这条结果"
                    onClick={() => replyToMessage(message)}
                  >
                    <svg
                      viewBox="0 0 24 24"
                      aria-hidden="true"
                      focusable="false"
                    >
                      <path d="M9 14 4 9l5-5" />
                      <path d="M4 9h10a6 6 0 0 1 0 12h-3" />
                    </svg>
                    <span>回复</span>
                  </button>
                )}
                <button
                  className={`message-copy-button ${
                    copiedMessageId === message.id ? "copied" : ""
                  }`}
                  type="button"
                  title={
                    copiedMessageId === message.id
                      ? "已复制"
                      : "复制这条消息"
                  }
                  aria-label={
                    copiedMessageId === message.id
                      ? "消息已复制"
                      : "复制这条消息"
                  }
                  onClick={() => void copyMessage(message)}
                >
                  {copiedMessageId === message.id ? (
                    <>
                      <svg
                        viewBox="0 0 24 24"
                        aria-hidden="true"
                        focusable="false"
                      >
                        <path d="m5 12 4 4L19 6" />
                      </svg>
                      <span>已复制</span>
                    </>
                  ) : (
                    <svg
                      viewBox="0 0 24 24"
                      aria-hidden="true"
                      focusable="false"
                    >
                      <rect x="8" y="8" width="11" height="11" rx="2" />
                      <path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h2" />
                    </svg>
                  )}
                </button>
              </div>
            )}
          </div>
        </article>
      ))}

      {(status === "planning" || status === "tooling") && (
        <article className="message-row assistant">
          <div className="activity-card" aria-live="polite">
            <div className="activity-heading">
              <span className="activity-pulse" aria-hidden="true">
                <i />
                <i />
                <i />
              </span>
              <strong>
                {activity?.title ??
                  (status === "tooling"
                    ? "正在本地读取并计算"
                    : "正在理解工作簿")}
              </strong>
              <time>{activitySeconds} 秒</time>
            </div>
            {activity?.detail && <p>{activity.detail}</p>}
            {busy && (
              <button
                type="button"
                className="activity-stop-button"
                onClick={stopTurn}
              >
                停止
              </button>
            )}
            {activity && activity.completed.length > 0 && (
              <ul className="activity-steps">
                {activity.completed.map((step, index) => (
                  <li key={`${step.label}-${index}`}>
                    <span className="activity-step-check" aria-hidden="true">
                      ✓
                    </span>
                    <div className="activity-step-body">
                      <span className="activity-step-label">{step.label}</span>
                      {step.note && (
                        <span className="activity-step-note">{step.note}</span>
                      )}
                      {step.detail && (
                        <details className="activity-step-detail">
                          <summary>查看详情</summary>
                          <pre>{step.detail}</pre>
                        </details>
                      )}
                    </div>
                    <span className="activity-step-time">
                      {formatStepElapsed(step.elapsedMs)}
                    </span>
                  </li>
                ))}
              </ul>
            )}
          </div>
        </article>
      )}
      <div ref={messageEndRef} />
    </section>
  );
}
