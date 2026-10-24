import type { Dispatch, SetStateAction } from "react";
import type { ModelConnectionDraft } from "../hooks/useModelManagement";

/** The two host settings editors share protocol and test-purpose semantics. */
export function ModelConnectionProtocolFields({ draft, onChange }: {
  draft: ModelConnectionDraft;
  onChange: Dispatch<SetStateAction<ModelConnectionDraft | null>>;
}) {
  return <>
    <label>
      <span>接口协议</span>
      <select value={draft.protocol ?? "openai_compatible"} onChange={(event) => {
        const protocol = event.target.value as "openai_compatible" | "anthropic";
        onChange((current) => current ? { ...current, protocol, purpose: "agent" } : current);
      }}>
        <option value="openai_compatible">OpenAI 兼容</option>
        <option value="anthropic">Anthropic 兼容（SDK Agent）</option>
      </select>
    </label>
    <label>
      <span>测试用途</span>
      <select value={draft.purpose ?? "agent"} onChange={(event) => {
        const purpose = event.target.value as "agent" | "formula";
        onChange((current) => current ? { ...current, purpose } : current);
      }}>
        <option value="agent">表格 Agent（SDK 与工具调用）</option>
        <option value="formula" disabled={draft.protocol === "anthropic"}>独立公式（OpenAI 接口）</option>
      </select>
    </label>
    <p>表格 Agent 使用 Claude Agent SDK。自定义服务须填写供应商提供的 Anthropic 兼容地址；DeepSeek 官方 OpenAI 地址可自动适配。独立公式测试通过不代表 Agent 可用。</p>
  </>;
}
