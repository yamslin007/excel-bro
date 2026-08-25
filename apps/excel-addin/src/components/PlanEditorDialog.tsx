import {
  useCallback,
  useEffect,
  useRef,
  useState,
  type PointerEvent as ReactPointerEvent
} from "react";
import type { AnalysisPlan, WorkbookSnapshot } from "../contracts";
import { assertEditablePlan } from "../contracts";
import { chatModifyPlan, validatePlan } from "../api";
import { updateToolPlanTemplate, type SavedTool } from "../storage";
import {
  isToolSchemaFingerprint,
  toolSchemaFingerprintForSnapshot
} from "../excel";
import type { PlanEditorTarget } from "../hooks/usePlanEditor";
import { applyThemePalette } from "../hooks/useTheme";
import { loadThemeSettings } from "../utils/imageStorage";
import { JsonCodeEditor } from "./JsonCodeEditor";

// 消息协议
interface PlanEditorInitMessage {
  type: "plan-editor-init";
  payload: {
    plan: AnalysisPlan;
    target: PlanEditorTarget;
    workbook: WorkbookSnapshot;
  };
}

interface PlanEditorReadyMessage {
  type: "plan-editor-ready";
}

interface PlanEditorSaveMessage {
  type: "plan-editor-save";
  payload: {
    plan: AnalysisPlan;
    target: PlanEditorTarget;
    updatedTools?: SavedTool[];
    /** 保存按钮触发时由父窗口负责关闭对话框 */
    closeAfterSave?: boolean;
  };
}

interface PlanEditorCloseMessage {
  type: "plan-editor-close";
}

type DialogPhase = "loading" | "editing" | "error" | "chatting" | "validating";

interface ChatMessage {
  role: "user" | "assistant";
  content: string;
  modifiedPlan?: AnalysisPlan;
}

const STORAGE_KEY = "EB_PLAN_EDITOR_PAYLOAD";

function getStoredPayload(): PlanEditorInitMessage["payload"] | null {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as PlanEditorInitMessage;
    return parsed.type === "plan-editor-init" ? parsed.payload : null;
  } catch {
    return null;
  }
}

function sendMessageToParent(
  message: PlanEditorReadyMessage | PlanEditorSaveMessage | PlanEditorCloseMessage
) {
  try {
    if (Office.context?.ui?.messageParent) {
      Office.context.ui.messageParent(JSON.stringify(message));
      return;
    }
  } catch {
    // 降级：如果不在 Dialog 环境中运行，静默失败
  }
}

export default function PlanEditorDialog() {
  const [phase, setPhase] = useState<DialogPhase>("loading");
  const [draft, setDraft] = useState("");
  const [errorMessage, setErrorMessage] = useState("");
  const [payload, setPayload] = useState<PlanEditorInitMessage["payload"] | null>(
    null
  );
  const [chatMessages, setChatMessages] = useState<ChatMessage[]>([]);
  const [chatInput, setChatInput] = useState("");
  const [chatPanelCollapsed, setChatPanelCollapsed] = useState(false);
  const [chatWidth, setChatWidth] = useState(400);
  const [confirmDiscardOpen, setConfirmDiscardOpen] = useState(false);
  const [showSavedToast, setShowSavedToast] = useState(false);
  const abortRef = useRef<AbortController | null>(null);
  const toastTimerRef = useRef<number | null>(null);
  const splitRef = useRef<HTMLDivElement | null>(null);
  const messagesRef = useRef<HTMLDivElement | null>(null);
  const originalDraftRef = useRef<string | null>(null);

  const applyPayload = useCallback((next: PlanEditorInitMessage["payload"]) => {
    const formatted = JSON.stringify(next.plan, null, 2);
    setPayload(next);
    setDraft(formatted);
    originalDraftRef.current = formatted;
    setPhase("editing");
    setErrorMessage("");
  }, []);

  // 组件卸载时清理提示定时器
  useEffect(() => {
    return () => {
      if (toastTimerRef.current !== null) {
        window.clearTimeout(toastTimerRef.current);
      }
    };
  }, []);

  // 独立窗口不会自动继承任务窗格的主题预设，这里读取并应用调色板
  useEffect(() => {
    let active = true;
    void (async () => {
      try {
        const settings = await loadThemeSettings();
        if (active && settings) applyThemePalette(settings.preset);
      } catch {
        // 读取失败时保持默认主题
      }
    })();
    return () => {
      active = false;
    };
  }, []);

  // Office.js 初始化 + 消息监听
  useEffect(() => {
    Office.onReady(() => {
      // 尝试从 localStorage 读取初始数据（兜底）
      const stored = getStoredPayload();
      if (stored) applyPayload(stored);

      // 监听父窗口消息
      if (Office.context?.ui?.addHandlerAsync) {
        Office.context.ui.addHandlerAsync(
          Office.EventType.DialogParentMessageReceived,
          (event: Office.DialogParentMessageReceivedEventArgs) => {
            if (!("message" in event)) return;
            try {
              const message = JSON.parse(event.message) as PlanEditorInitMessage;
              if (message.type === "plan-editor-init") {
                applyPayload(message.payload);
              }
            } catch {
              // 忽略非法消息
            }
          },
          () => {
            // 注册成功后发送就绪信号
            sendMessageToParent({ type: "plan-editor-ready" });
          }
        );
      } else {
        // 如果不支持消息通信，直接发送就绪信号
        sendMessageToParent({ type: "plan-editor-ready" });
      }
    });
  }, [applyPayload]);

  const handleDraftChange = useCallback((text: string) => {
    setDraft(text);
    setErrorMessage("");
  }, []);

  const isDirty =
    originalDraftRef.current !== null &&
    draft !== originalDraftRef.current;

  // 对话历史变化时自动滚动到底部
  useEffect(() => {
    const container = messagesRef.current;
    if (container) container.scrollTop = container.scrollHeight;
  }, [chatMessages, phase]);

  const closeWindow = useCallback(() => {
    sendMessageToParent({ type: "plan-editor-close" });
    window.close();
  }, []);

  const handleCancel = useCallback(() => {
    if (originalDraftRef.current !== null && draft !== originalDraftRef.current) {
      setConfirmDiscardOpen(true);
      return;
    }
    closeWindow();
  }, [closeWindow, draft]);

  const handleChatSubmit = useCallback(async () => {
    if (!payload || !chatInput.trim()) return;

    const userMessage: ChatMessage = {
      role: "user",
      content: chatInput.trim()
    };
    setChatMessages((previous) => [...previous, userMessage]);
    setChatInput("");
    setPhase("chatting");
    setErrorMessage("");

    abortRef.current?.abort();
    const controller = new AbortController();
    abortRef.current = controller;

    try {
      let currentPlan: AnalysisPlan;
      try {
        currentPlan = JSON.parse(draft);
      } catch {
        setPhase("error");
        setErrorMessage("当前 JSON 格式有误，无法使用 AI 修改");
        return;
      }

      const response = await chatModifyPlan(
        {
          plan: currentPlan,
          message: userMessage.content
        },
        controller.signal
      );

      const assistantMessage: ChatMessage = {
        role: "assistant",
        content: response.explanation,
        modifiedPlan: response.modifiedPlan
      };
      setChatMessages((previous) => [...previous, assistantMessage]);
      setPhase("editing");
    } catch (reason) {
      if (controller.signal.aborted) return;
      setPhase("error");
      setErrorMessage(
        reason instanceof Error ? reason.message : "AI 修改失败，请重试"
      );
    }
  }, [payload, chatInput, draft]);

  const handleApplyPlan = useCallback((plan: AnalysisPlan) => {
    setDraft(JSON.stringify(plan, null, 2));
    setErrorMessage("");
    setPhase("editing");
  }, []);

  const handleSave = useCallback(async (closeAfterSave: boolean) => {
    if (!payload) return;

    // 第一道门：本地 JSON 解析 + 白名单结构校验，给即时反馈
    let parsed: unknown;
    try {
      parsed = JSON.parse(draft);
    } catch {
      setPhase("error");
      setErrorMessage("JSON 格式有误，请检查括号和引号是否匹配");
      return;
    }

    try {
      assertEditablePlan(parsed);
    } catch (reason) {
      setPhase("error");
      setErrorMessage(
        reason instanceof Error ? reason.message : "计划结构不合法"
      );
      return;
    }

    // 第二道门：服务端 Pydantic 规范化。逐动作校验合法性并重算验收条件
    // （acceptanceCriteria），返回的才是最终要保存的版本 —— 前端白名单只
    // 保证「字段在不在」，值是否合法、验收条件怎么算都只有后端知道。
    setPhase("validating");
    setErrorMessage("");
    abortRef.current?.abort();
    const controller = new AbortController();
    abortRef.current = controller;

    let normalized: AnalysisPlan;
    try {
      normalized = await validatePlan(parsed as AnalysisPlan);
    } catch (reason) {
      if (controller.signal.aborted) return;
      setPhase("error");
      setErrorMessage(
        reason instanceof Error
          ? reason.message
          : "服务端校验失败，请检查计划内容"
      );
      return;
    }

    // 固化工具：把规范化后的计划写回 storage
    let updatedTools: SavedTool[] | undefined;
    if (payload.target.kind === "savedTool") {
      try {
        updatedTools = updateToolPlanTemplate(
          payload.target.toolId,
          normalized
        );
      } catch (reason) {
        setPhase("error");
        setErrorMessage(
          reason instanceof Error ? reason.message : "写回工具失败"
        );
        return;
      }
    }

    // 指纹重算（只针对 schema 指纹）
    if (
      payload.workbook &&
      normalized.sourceFingerprint &&
      isToolSchemaFingerprint(normalized.sourceFingerprint)
    ) {
      normalized = {
        ...normalized,
        sourceFingerprint: toolSchemaFingerprintForSnapshot(
          normalized,
          payload.workbook
        )
      };
    }

    // 发送保存成功消息
    sendMessageToParent({
      type: "plan-editor-save",
      payload: {
        plan: normalized,
        target: payload.target,
        updatedTools,
        closeAfterSave
      }
    });

    if (closeAfterSave) {
      // 保存按钮：保存成功后直接关闭窗口
      window.close();
      return;
    }

    // Ctrl+S：不关闭窗口，把编辑器内容更新为最终保存的计划，并提示保存成功
    const savedText = JSON.stringify(normalized, null, 2);
    setDraft(savedText);
    originalDraftRef.current = savedText;
    setPhase("editing");
    setErrorMessage("");
    setShowSavedToast(true);
    if (toastTimerRef.current !== null) {
      window.clearTimeout(toastTimerRef.current);
    }
    toastTimerRef.current = window.setTimeout(() => {
      setShowSavedToast(false);
    }, 2000);
  }, [draft, payload]);

  // Ctrl+S 保存、Esc 关闭确认弹窗
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
        event.preventDefault();
        if (phase === "editing") void handleSave(false);
        return;
      }
      if (event.key === "Escape" && confirmDiscardOpen) {
        setConfirmDiscardOpen(false);
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [phase, handleSave, confirmDiscardOpen]);

  const startResize = useCallback((event: ReactPointerEvent<HTMLDivElement>) => {
    event.preventDefault();
    const split = splitRef.current;
    if (!split) return;
    const bounds = split.getBoundingClientRect();
    const move = (moveEvent: PointerEvent) => {
      const width = Math.min(
        640,
        Math.max(300, moveEvent.clientX - bounds.left)
      );
      setChatWidth(width);
    };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  }, []);

  if (phase === "loading") {
    return (
      <div className="plan-editor-loading">
        <span className="plan-editor-loading-spinner" aria-hidden="true" />
        <strong>正在准备编辑器</strong>
        <span>正在从任务窗格同步计划数据…</span>
      </div>
    );
  }

  const busy = phase === "chatting" || phase === "validating";
  const breadcrumb = payload
    ? `${payload.target.kind === "savedTool" ? "我的工具 › 固化工具" : "消息 › 计划脚本"} · ${payload.plan.title} · ${payload.plan.actions.length} 个动作`
    : "";

  return (
    <div className="plan-editor-dialog">
      <header className="plan-editor-dialog-header">
        <div className="plan-editor-dialog-header-left">
          <button
            className="plan-editor-header-close"
            onClick={handleCancel}
            disabled={busy}
            title="关闭 (Esc)"
            aria-label="关闭"
          >
            ×
          </button>
          <div className="plan-editor-dialog-title">
            <div className="plan-editor-dialog-title-row">
              <h1>编辑执行脚本</h1>
              {isDirty && (
                <span className="plan-editor-dirty">已修改</span>
              )}
            </div>
            <p className="plan-editor-breadcrumb">
              <strong>{breadcrumb}</strong>
            </p>
          </div>
        </div>
        <div className="plan-editor-dialog-header-actions">
          <button
            className="secondary-button"
            onClick={handleCancel}
            disabled={busy}
          >
            放弃修改
          </button>
          <button
            className="primary-button"
            onClick={() => void handleSave(true)}
            disabled={busy}
          >
            {phase === "validating" ? "校验中..." : "保存"}
          </button>
        </div>
      </header>

      {errorMessage && phase === "error" && (
        <div className="plan-editor-dialog-error" role="alert">
          {errorMessage}
        </div>
      )}

      <div
        ref={splitRef}
        className={`plan-editor-dialog-split${
          chatPanelCollapsed ? " chat-collapsed" : ""
        }`}
        style={
          chatPanelCollapsed
            ? undefined
            : { gridTemplateColumns: `${chatWidth}px 6px minmax(0, 1fr)` }
        }
      >
        {/* 左侧：AI 对话 */}
        <section className="plan-editor-chat-panel">
          <header className="plan-editor-chat-header">
            <div className="plan-editor-chat-heading">
              <span className="plan-editor-spark" aria-hidden="true">
                ✦
              </span>
              <div>
                <strong>AI 助手</strong>
                <span>用自然语言描述修改</span>
              </div>
            </div>
            <button
              className="plan-editor-chat-collapse"
              onClick={() => setChatPanelCollapsed(true)}
              title="收起 AI 对话"
              aria-label="收起 AI 对话"
            >
              »
            </button>
          </header>

          <div className="plan-editor-chat-messages" ref={messagesRef}>
            {chatMessages.length === 0 && phase !== "chatting" ? (
              <div className="plan-editor-chat-empty">
                <span className="plan-editor-chat-empty-icon" aria-hidden="true">
                  ✦
                </span>
                <p>描述你想要的修改</p>
                <small>
                  AI 会基于当前编辑器内容返回完整计划，确认后可一键应用。
                </small>
              </div>
            ) : (
              <>
                {chatMessages.map((msg, index) => (
                  <div
                    key={index}
                    className={`plan-editor-chat-message ${msg.role}`}
                  >
                    <span className="plan-editor-avatar">
                      {msg.role === "user" ? "我" : "AI"}
                    </span>
                    <div className="plan-editor-chat-meta">
                      <span className="who">
                        {msg.role === "user" ? "你" : "AI 助手"}
                      </span>
                      <div className="plan-editor-chat-bubble">
                        {msg.content}
                      </div>
                      {msg.modifiedPlan && (
                        <button
                          className="apply-plan-button"
                          onClick={() => handleApplyPlan(msg.modifiedPlan!)}
                          disabled={busy}
                        >
                          应用到编辑器 →
                        </button>
                      )}
                    </div>
                  </div>
                ))}
                {phase === "chatting" && (
                  <div className="plan-editor-chat-message assistant">
                    <span className="plan-editor-avatar">AI</span>
                    <div className="plan-editor-chat-meta">
                      <span className="who">AI 助手</span>
                      <div
                        className="plan-editor-chat-bubble plan-editor-typing"
                        aria-label="AI 正在思考"
                      >
                        <span />
                        <span />
                        <span />
                      </div>
                    </div>
                  </div>
                )}
              </>
            )}
          </div>

          <footer className="plan-editor-chat-input">
            <textarea
              value={chatInput}
              onChange={(event) => setChatInput(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === "Enter" && !event.shiftKey) {
                  event.preventDefault();
                  void handleChatSubmit();
                }
              }}
              placeholder="描述你想要的修改...（Enter 发送，Shift+Enter 换行）"
              disabled={busy}
              rows={2}
            />
            <button
              onClick={() => void handleChatSubmit()}
              disabled={busy || !chatInput.trim()}
              className="plan-editor-send"
              aria-label="发送"
              title="发送"
            >
              ↑
            </button>
          </footer>
        </section>

        {/* 折叠后的窄侧栏 */}
        <aside className="plan-editor-chat-rail">
          <button
            onClick={() => setChatPanelCollapsed(false)}
            title="展开 AI 对话"
            aria-label="展开 AI 对话"
          >
            <span aria-hidden="true">✦</span>
            <span>AI 助手</span>
          </button>
        </aside>

        {/* 拖动分隔条 */}
        <div
          className="plan-editor-resizer"
          onPointerDown={startResize}
          title="拖动调整宽度"
        />

        {/* 右侧：JSON 编辑器 */}
        <section className="plan-editor-json-panel">
          <JsonCodeEditor
            value={draft}
            onChange={handleDraftChange}
            disabled={busy}
            ariaLabel="计划 JSON"
            dirty={isDirty}
          />
        </section>
      </div>

      {confirmDiscardOpen && (
        <div className="plan-editor-modal" role="dialog" aria-modal="true">
          <div className="plan-editor-modal-card">
            <h3>放弃未保存的修改？</h3>
            <p>当前计划已被修改，放弃后改动不会保留。确定要放弃吗？</p>
            <div className="plan-editor-modal-actions">
              <button
                className="secondary-button"
                onClick={() => setConfirmDiscardOpen(false)}
              >
                继续编辑
              </button>
              <button
                className="primary-button"
                onClick={() => {
                  setConfirmDiscardOpen(false);
                  closeWindow();
                }}
              >
                放弃修改
              </button>
            </div>
          </div>
        </div>
      )}

      {showSavedToast && (
        <div className="plan-editor-toast" role="status">
          已保存
        </div>
      )}
    </div>
  );
}
