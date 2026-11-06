/**
 * WPS 专注窗口的**快照构造**。
 *
 * 只挑展示所需字段：对话文本、计划摘要、结果表头与前 200 行、工具说明。
 * 凭据、连接元数据、SDK 桥引用、完整结果行都不进快照 —— 这条结构就是边界。
 * 内容一律取自任务窗格的当前状态（构造那一刻现取），窗口不会自己读工作表。
 */

import type {
  FocusMessage,
  FocusSnapshot,
  FocusWorkbookRef
} from "@excel-bro/core/focusProtocol";
import { actionLabel } from "../../excel-addin/src/format";
import type { SavedQueryTool, SavedTool } from "../../excel-addin/src/storage";
import type { ResultContext } from "../../excel-addin/src/contracts";

export interface WpsFocusMessageInput {
  id: string;
  role: "assistant" | "user" | "system";
  text: string;
  createdAt: string;
  planPreview?: {
    title: string;
    status: "pending" | "executed" | "cancelled";
    revised?: boolean;
  } | null;
  resultContext?: ResultContext | null;
}

export interface WpsFocusConversationInput {
  id: string;
  title: string;
  updatedAt: string;
  messages: WpsFocusMessageInput[];
}

const PLAN_STATUS_LABEL: Record<"pending" | "executed" | "cancelled", string> = {
  pending: "待处理：尚未执行或确认",
  executed: "已执行",
  cancelled: "已取消"
};

function messageToFocusMessage(message: WpsFocusMessageInput): FocusMessage {
  const result = message.resultContext
    ? {
        title: message.resultContext.title,
        headers: message.resultContext.headers,
        rows: message.resultContext.rows.slice(0, 200)
      }
    : undefined;
  return {
    id: message.id,
    role: message.role,
    text: message.text ?? "",
    createdAt: message.createdAt,
    ...(message.planPreview
      ? {
          plan: {
            title: message.planPreview.title,
            summary: PLAN_STATUS_LABEL[message.planPreview.status],
            steps: []
          }
        }
      : {}),
    ...(result ? { result } : {})
  };
}

export function buildWpsFocusSnapshot(input: {
  initialView: FocusSnapshot["initialView"];
  activeConversationId: string;
  conversations: WpsFocusConversationInput[];
  tools: SavedTool[];
  queryTools: SavedQueryTool[];
  workbook: FocusWorkbookRef;
}): FocusSnapshot {
  return {
    initialView: input.initialView,
    activeConversationId: input.activeConversationId,
    conversations: input.conversations.map((conversation) => ({
      id: conversation.id,
      title: conversation.title,
      updatedAt: conversation.updatedAt,
      messages: conversation.messages.slice(-100).map(messageToFocusMessage)
    })),
    tools: [
      ...input.tools.map((tool) => ({
        id: tool.id,
        kind: "workflow" as const,
        name: tool.name,
        description: tool.description,
        category: "工作流程",
        steps: tool.planTemplate.actions.length,
        stepLabels: tool.planTemplate.actions.map(actionLabel)
      })),
      ...input.queryTools.map((tool) => ({
        id: tool.id,
        kind: "query" as const,
        name: tool.name,
        description: tool.description,
        category: "本地查询",
        steps: 1,
        stepLabels: [] as string[]
      }))
    ],
    workbook: input.workbook
  };
}
