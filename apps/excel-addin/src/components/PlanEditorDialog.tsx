import { useCallback, useEffect, useRef, useState } from "react";
import type { AnalysisPlan, WorkbookSnapshot } from "../contracts";
import { assertEditablePlan } from "../contracts";
import { validatePlan, chatModifyPlan, type PlanChatResponse } from "../api";
import { updateToolPlanTemplate, type SavedTool } from "../storage";
import {
  isToolSchemaFingerprint,
  toolSchemaFingerprintForSnapshot
} from "../excel";
import type { PlanEditorTarget } from "../hooks/usePlanEditor";

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
  };
}

interface PlanEditorCloseMessage {
  type: "plan-editor-close";
}

type DialogPhase = "loading" | "editing" | "validating" | "error" | "chatting";

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

function sendMessageToParent(message: PlanEditorReadyMessage | PlanEditorSaveMessage | PlanEditorCloseMessage) {
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
  const [payload, setPayload] = useState<PlanEditorInitMessage["payload"] | null>(null);
  const [chatMessages, setChatMessages] = useState<ChatMessage[]>([]);
  const [chatInput, setChatInput] = useState("");
  const abortRef = useRef<AbortController | null>(null);

  // Office.js 初始化 + 消息监听
  useEffect(() => {
    Office.onReady(() => {
      // 尝试从 localStorage 读取初始数据（兜底）
      const stored = getStoredPayload();
      if (stored) {
        setPayload(stored);
        setDraft(JSON.stringify(stored.plan, null, 2));
        setPhase("editing");
      }

      // 监听父窗口消息
      if (Office.context?.ui?.addHandlerAsync) {
        Office.context.ui.addHandlerAsync(
          Office.EventType.DialogParentMessageReceived,
          (event: Office.DialogParentMessageReceivedEventArgs) => {
            if (!("message" in event)) return;
            try {
              const message = JSON.parse(event.message) as PlanEditorInitMessage;
              if (message.type === "plan-editor-init") {
                setPayload(message.payload);
                setDraft(JSON.stringify(message.payload.plan, null, 2));
                setPhase("editing");
                setErrorMessage("");
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
  }, []);

  const handleDraftChange = useCallback((text: string) => {
    setDraft(text);
    setErrorMessage("");
  }, []);

  const handleChatSubmit = useCallback(async () => {
    if (!payload || !chatInput.trim()) return;

    const userMessage: ChatMessage = {
      role: "user",
      content: chatInput.trim()
    };
    setChatMessages((prev) => [...prev, userMessage]);
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
      setChatMessages((prev) => [...prev, assistantMessage]);
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

  const handleCancel = useCallback(() => {
    sendMessageToParent({ type: "plan-editor-close" });
    window.close();
  }, []);

  const handleSave = useCallback(async () => {
    if (!payload) return;

    // 第一道门：本地 JSON 解析 + 白名单结构校验
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

    // 第二道门：服务端 Pydantic 规范化
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
        updatedTools
      }
    });

    // 关闭窗口
    window.close();
  }, [draft, payload]);

  if (phase === "loading") {
    return (
      <div className="plan-editor-loading">
        <strong>正在准备编辑器</strong>
        <span>正在从任务窗格同步计划数据…</span>
      </div>
    );
  }

  const busy = phase === "validating" || phase === "chatting";

  return (
    <div className="plan-editor-dialog">
      <header className="plan-editor-dialog-header">
        <div>
          <h1>编辑执行脚本</h1>
          <p>左侧可用自然语言描述修改，右侧可直接编辑 JSON</p>
        </div>
      </header>

      <div className="plan-editor-dialog-split">
        {/* 左侧：AI 对话 */}
        <div className="plan-editor-chat-panel">
          <div className="plan-editor-chat-messages">
            {chatMessages.length === 0 ? (
              <div className="plan-editor-chat-empty">
                <p>💬 描述你想要的修改</p>
                <p className="hint">例如："添加一个筛选步骤，只保留销售额大于 1000 的记录"</p>
              </div>
            ) : (
              chatMessages.map((msg, index) => (
                <div
                  key={index}
                  className={`plan-editor-chat-message ${msg.role}`}
                >
                  <div className="message-role">
                    {msg.role === "user" ? "你" : "AI"}
                  </div>
                  <div className="message-content">{msg.content}</div>
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
              ))
            )}
          </div>
          <div className="plan-editor-chat-input">
            <textarea
              value={chatInput}
              onChange={(e) => setChatInput(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter" && !e.shiftKey) {
                  e.preventDefault();
                  handleChatSubmit();
                }
              }}
              placeholder="描述你想要的修改...（Enter 发送，Shift+Enter 换行）"
              disabled={busy}
              rows={3}
            />
            <button
              onClick={handleChatSubmit}
              disabled={busy || !chatInput.trim()}
              className="primary-button"
            >
              {phase === "chatting" ? "AI 思考中..." : "发送"}
            </button>
          </div>
        </div>

        {/* 右侧：JSON 编辑器 */}
        <div className="plan-editor-json-panel">
          <textarea
            className="plan-editor-dialog-textarea"
            value={draft}
            spellCheck={false}
            disabled={busy}
            aria-label="计划 JSON"
            onChange={(e) => handleDraftChange(e.target.value)}
          />
        </div>
      </div>

      {errorMessage && (
        <div className="plan-editor-dialog-error" role="alert">
          {errorMessage}
        </div>
      )}

      <footer className="plan-editor-dialog-footer">
        <button
          className="secondary-button"
          onClick={handleCancel}
          disabled={busy}
        >
          放弃修改
        </button>
        <button
          className="primary-button"
          onClick={handleSave}
          disabled={busy}
        >
          {busy ? "校验中…" : "保存"}
        </button>
      </footer>
    </div>
  );
}
