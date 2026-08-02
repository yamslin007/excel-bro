export interface FocusMessage {
  id: string;
  role: "assistant" | "user" | "system";
  text: string;
  createdAt: string;
  plan?: {
    title: string;
    summary: string;
    steps: string[];
  };
  result?: {
    title: string;
    headers: string[];
    rows: Array<Array<string | number | boolean | null>>;
  };
}

export interface FocusConversation {
  id: string;
  title: string;
  updatedAt: string;
  messages: FocusMessage[];
}

export interface FocusTool {
  id: string;
  kind: "workflow" | "query";
  name: string;
  description: string;
  category: string;
  steps: number;
  stepLabels: string[];
  dsl?: string;
}

export interface FocusPayload {
  type: "focus-state";
  workbookName: string;
  initialView: "conversation" | "tools";
  activeConversationId: string;
  conversations: FocusConversation[];
  tools: FocusTool[];
}

export const FOCUS_PAYLOAD_STORAGE_KEY = "excel-bro.focus.v1";
