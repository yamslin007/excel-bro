import { useEffect, useMemo, useState } from "react";
import type {
  FocusChildStatus,
  FocusSnapshot
} from "@excel-bro/core/focusProtocol";
import { brandLogoUrl } from "../brandLogo";

/**
 * 专注窗口的**展示层**，两端共用（Excel DialogApi / WPS WebDialog）。
 *
 * 它只接收"已经校验过的快照 + 子窗口状态"，不认识 Office 全局对象、不认识 WPS
 * `Application`，也不发任何请求 —— 打开、通信、关闭的宿主差异全在各自的适配层。
 *
 * 状态与内容**分开表达**：`stale` / `closed` / `snapshot` 时仍然保留已有内容（用户
 * 还能看），但顶部必须明确标注"这不是当前现场"。把失联画成正常界面，等于让用户
 * 对着一份旧内容做判断。
 */
export interface FocusWorkspaceViewProps {
  snapshot: FocusSnapshot | null;
  status: FocusChildStatus;
  statusDetail: string | null;
  /** 实际连通的通道名（宿主自报）；未连通时 `null`。 */
  channelLabel: string | null;
  onClose: () => void;
}

const STATUS_BADGE: Record<FocusChildStatus, string> = {
  connecting: "正在连接",
  ready: "已同步",
  stale: "已失联",
  snapshot: "仅快照",
  closed: "已结束",
  failed: "未连接"
};

const STATUS_DETAIL: Record<FocusChildStatus, string> = {
  connecting: "正在与任务窗格握手，确认这是当前工作簿的现场…",
  ready: "",
  stale: "任务窗格的心跳已中断，下面的内容可能已过期；请回到任务窗格确认后再操作。",
  snapshot: "当前只能显示打开窗口那一刻的快照，后续变化不会同步到这里。",
  closed: "任务窗格已关闭、刷新或切换了工作簿，此窗口不再同步。",
  failed: "无法确认这是当前工作簿的现场。"
};

function timeLabel(value: string): string {
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "";
  return parsed.toLocaleString("zh-CN");
}

export default function FocusWorkspaceView({
  snapshot,
  status,
  statusDetail,
  channelLabel,
  onClose
}: FocusWorkspaceViewProps) {
  const [view, setView] = useState<"conversation" | "tools">(
    snapshot?.initialView ?? "conversation"
  );
  const [conversationId, setConversationId] = useState(
    snapshot?.activeConversationId ?? ""
  );
  const [toolId, setToolId] = useState(snapshot?.tools[0]?.id ?? "");
  const [toolDetailMode, setToolDetailMode] = useState<"standard" | "expert">("standard");

  /*
   * 只做"选中项仍然存在"的校正，不跟随每次同步重置视图。
   * 心跳/新消息会不断送来新快照，如果每次都把视图拉回 initialView，用户正在看的
   * 那条对话会被反复顶掉 —— 那正是"同步"看起来像"刷新"的原因。
   */
  useEffect(() => {
    if (!snapshot) return;
    setConversationId((current) =>
      snapshot.conversations.some((item) => item.id === current)
        ? current
        : snapshot.activeConversationId || snapshot.conversations[0]?.id || ""
    );
    setToolId((current) =>
      snapshot.tools.some((item) => item.id === current)
        ? current
        : snapshot.tools[0]?.id ?? ""
    );
  }, [snapshot]);

  const conversation = useMemo(
    () =>
      snapshot?.conversations.find((item) => item.id === conversationId) ??
      snapshot?.conversations[0],
    [conversationId, snapshot]
  );
  const tool = useMemo(
    () => snapshot?.tools.find((item) => item.id === toolId) ?? snapshot?.tools[0],
    [snapshot, toolId]
  );

  if (!snapshot) {
    return (
      <main className="focus-loading" data-focus-status={status}>
        <img className="focus-brand-mark" src={brandLogoUrl} alt="" />
        <strong>
          {status === "failed" ? "专注窗口未能连接" : "正在准备专注窗口"}
        </strong>
        <span>{statusDetail ?? STATUS_DETAIL.connecting}</span>
        {status === "failed" && (
          <button type="button" className="focus-close" onClick={onClose}>
            关闭窗口
          </button>
        )}
      </main>
    );
  }

  const detail = statusDetail ?? STATUS_DETAIL[status];

  return (
    <main className="focus-shell" data-focus-status={status}>
      <header className="focus-header">
        <img className="focus-brand-mark" src={brandLogoUrl} alt="" />
        <div className="focus-brand-copy">
          <strong>Excel Bro</strong>
          <span>
            {snapshot.workbook.name || "当前工作簿"} · 专注查看
          </span>
        </div>
        <span className={`focus-status-badge tone-${status}`}>
          {STATUS_BADGE[status]}
        </span>
        <nav aria-label="专注窗口导航">
          <button
            className={view === "conversation" ? "active" : ""}
            onClick={() => setView("conversation")}
          >
            对话
          </button>
          <button
            className={view === "tools" ? "active" : ""}
            onClick={() => setView("tools")}
          >
            工具
          </button>
        </nav>
        <button className="focus-close" onClick={onClose}>
          返回任务窗格
        </button>
      </header>

      {status !== "ready" && detail && (
        <div className={`focus-status-banner tone-${status}`} role="status">
          <strong>{STATUS_BADGE[status]}</strong>
          <span>{detail}</span>
          {status === "connecting" ? null : (
            <span className="focus-status-hint">
              这里只做展示：写入、确认与撤销仍在任务窗格完成。
            </span>
          )}
        </div>
      )}

      {view === "conversation" ? (
        <div className="focus-layout">
          <aside className="focus-sidebar">
            <span className="focus-eyebrow">历史对话</span>
            {snapshot.conversations.map((item) => (
              <button
                key={item.id}
                className={item.id === conversation?.id ? "selected" : ""}
                onClick={() => setConversationId(item.id)}
              >
                <strong>{item.title}</strong>
                <span>{timeLabel(item.updatedAt)}</span>
              </button>
            ))}
          </aside>
          <section className="focus-content">
            <div className="focus-page-title">
              <span className="focus-eyebrow">当前对话</span>
              <h1>{conversation?.title ?? "暂无对话"}</h1>
            </div>
            <div className="focus-message-list">
              {conversation?.messages.map((message) => (
                <article
                  key={message.id}
                  className={`focus-message role-${message.role}`}
                >
                  <small>
                    {message.role === "user"
                      ? "你"
                      : message.role === "assistant"
                        ? "Excel Bro"
                        : "系统"}
                  </small>
                  {message.text && <p>{message.text}</p>}
                  {message.plan && (
                    <section className="focus-plan">
                      <strong>{message.plan.title}</strong>
                      <p>{message.plan.summary}</p>
                      <ol>
                        {message.plan.steps.map((step, index) => (
                          <li key={`${message.id}-${index}`}>{step}</li>
                        ))}
                      </ol>
                    </section>
                  )}
                  {message.result && (
                    <section className="focus-result">
                      <strong>{message.result.title}</strong>
                      <div>
                        <table>
                          <thead>
                            <tr>
                              {message.result.headers.map((header, index) => (
                                <th key={`${message.id}-h-${index}`}>{header}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {message.result.rows.map((row, rowIndex) => (
                              <tr key={rowIndex}>
                                {row.map((value, columnIndex) => (
                                  <td key={columnIndex}>{String(value ?? "")}</td>
                                ))}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </section>
                  )}
                </article>
              ))}
              {conversation && conversation.messages.length === 0 && (
                <div className="focus-empty-detail">这段对话还没有内容。</div>
              )}
            </div>
          </section>
        </div>
      ) : (
        <div className="focus-layout">
          <aside className="focus-sidebar">
            <span className="focus-eyebrow">我的工具</span>
            {snapshot.tools.map((item) => (
              <button
                key={item.id}
                className={item.id === tool?.id ? "selected" : ""}
                onClick={() => setToolId(item.id)}
              >
                <strong>{item.name}</strong>
                <span>{item.category}</span>
              </button>
            ))}
          </aside>
          <section className="focus-content">
            {tool ? (
              <>
                <div className="focus-page-title">
                  <span className="focus-eyebrow">{tool.category}</span>
                  <h1>{tool.name}</h1>
                  <p>{tool.description}</p>
                </div>
                <div
                  className="focus-tool-switch"
                  role="tablist"
                  aria-label="工具说明视图"
                >
                  <button
                    role="tab"
                    aria-selected={toolDetailMode === "standard"}
                    className={toolDetailMode === "standard" ? "active" : ""}
                    onClick={() => setToolDetailMode("standard")}
                  >
                    普通视图
                  </button>
                  <button
                    role="tab"
                    aria-selected={toolDetailMode === "expert"}
                    className={toolDetailMode === "expert" ? "active" : ""}
                    onClick={() => setToolDetailMode("expert")}
                  >
                    专家视图
                  </button>
                </div>
                {toolDetailMode === "standard" ? (
                  <>
                    <div className="focus-tool-facts">
                      <div>
                        <span>类型</span>
                        <strong>
                          {tool.kind === "workflow" ? "受控工作流" : "本地查询"}
                        </strong>
                      </div>
                      <div>
                        <span>处理步骤</span>
                        <strong>{tool.steps}</strong>
                      </div>
                      <div>
                        <span>执行位置</span>
                        <strong>返回任务窗格确认</strong>
                      </div>
                    </div>
                    {tool.stepLabels.length > 0 ? (
                      <div className="focus-tool-steps">
                        <span className="focus-eyebrow">完整处理步骤</span>
                        <ol>
                          {tool.stepLabels.map((label, index) => (
                            <li key={`${tool.id}-${index}`}>{label}</li>
                          ))}
                        </ol>
                      </div>
                    ) : (
                      <div className="focus-empty-detail">
                        此工具在本地确定性运行，一步返回结果。
                      </div>
                    )}
                  </>
                ) : tool.dsl ? (
                  <pre className="focus-dsl">
                    <code>{tool.dsl}</code>
                  </pre>
                ) : (
                  <div className="focus-empty-detail">
                    此工具在本地确定性运行，不包含专家计划。
                  </div>
                )}
              </>
            ) : (
              <div className="focus-empty-detail">暂无已保存工具。</div>
            )}
          </section>
        </div>
      )}

      <footer className="focus-footer">
        <span>
          {channelLabel ? `同步通道：${channelLabel}` : "同步通道：未建立"}
        </span>
        <span>只读展示 · 不执行任务、不修改工作簿</span>
      </footer>
    </main>
  );
}
