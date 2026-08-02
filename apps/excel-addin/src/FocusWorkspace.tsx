import { useEffect, useMemo, useState } from "react";
import {
  FOCUS_PAYLOAD_STORAGE_KEY,
  type FocusPayload
} from "./focusState";

function storedPayload(): FocusPayload | null {
  try {
    const value = JSON.parse(
      localStorage.getItem(FOCUS_PAYLOAD_STORAGE_KEY) ?? "null"
    ) as FocusPayload | null;
    return value?.type === "focus-state" ? value : null;
  } catch {
    return null;
  }
}

function requestClose() {
  try {
    if (Office.context?.ui?.messageParent) {
      Office.context.ui.messageParent(
        JSON.stringify({ type: "focus-close" })
      );
      return;
    }
    window.close();
  } catch {
    window.close();
  }
}

export default function FocusWorkspace() {
  const initialPayload = useMemo(storedPayload, []);
  const [payload, setPayload] = useState<FocusPayload | null>(initialPayload);
  const [view, setView] = useState<"conversation" | "tools">(
    initialPayload?.initialView ?? "conversation"
  );
  const [conversationId, setConversationId] = useState(
    initialPayload?.activeConversationId ?? ""
  );
  const [toolId, setToolId] = useState(initialPayload?.tools[0]?.id ?? "");
  const [toolDetailMode, setToolDetailMode] = useState<"standard" | "expert">(
    "standard"
  );

  useEffect(() => {
    const receive = (event: Office.DialogParentMessageReceivedEventArgs) => {
      try {
        const next = JSON.parse(event.message) as FocusPayload;
        if (next.type !== "focus-state") return;
        setPayload(next);
        setView(next.initialView);
        setConversationId(next.activeConversationId);
        setToolId(next.tools[0]?.id ?? "");
      } catch {
        // Keep the waiting screen when a malformed message is received.
      }
    };
    Office.onReady(() => {
      if (!Office.context?.ui?.addHandlerAsync) return;
      Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        receive,
        () => {
          Office.context.ui.messageParent(
            JSON.stringify({ type: "focus-ready" })
          );
        }
      );
    });
  }, []);

  const conversation = useMemo(
    () =>
      payload?.conversations.find((item) => item.id === conversationId) ??
      payload?.conversations[0],
    [conversationId, payload]
  );
  const tool = useMemo(
    () =>
      payload?.tools.find((item) => item.id === toolId) ?? payload?.tools[0],
    [payload, toolId]
  );

  if (!payload) {
    return (
      <main className="focus-loading">
        <div className="focus-brand-mark">EB</div>
        <strong>正在准备专注窗口</strong>
        <span>正在从任务窗格同步对话与工具说明…</span>
      </main>
    );
  }

  return (
    <main className="focus-shell">
      <header className="focus-header">
        <div className="focus-brand-mark">EB</div>
        <div className="focus-brand-copy">
          <strong>Excel Bro</strong>
          <span>{payload.workbookName || "当前工作簿"} · 专注查看</span>
        </div>
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
        <button className="focus-close" onClick={requestClose}>
          返回任务窗格
        </button>
      </header>

      {view === "conversation" ? (
        <div className="focus-layout">
          <aside className="focus-sidebar">
            <span className="focus-eyebrow">历史对话</span>
            {payload.conversations.map((item) => (
              <button
                key={item.id}
                className={item.id === conversation?.id ? "selected" : ""}
                onClick={() => setConversationId(item.id)}
              >
                <strong>{item.title}</strong>
                <span>{new Date(item.updatedAt).toLocaleString("zh-CN")}</span>
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
                              {message.result.headers.map((header) => (
                                <th key={header}>{header}</th>
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
            </div>
          </section>
        </div>
      ) : (
        <div className="focus-layout">
          <aside className="focus-sidebar">
            <span className="focus-eyebrow">我的工具</span>
            {payload.tools.map((item) => (
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
                    此工具在本地确定性运行，不包含专家脚本。
                  </div>
                )}
              </>
            ) : (
              <div className="focus-empty-detail">暂无已保存工具。</div>
            )}
          </section>
        </div>
      )}
    </main>
  );
}
