import type { Dispatch, SetStateAction } from "react";
import type { ModelConnectionDraft } from "../hooks/useModelManagement";

/** Both hosts configure one connection, independent of how it will be used. */
export function ModelConnectionProtocolFields({ draft, onChange }: {
  draft: ModelConnectionDraft;
  onChange: Dispatch<SetStateAction<ModelConnectionDraft | null>>;
}) {
  return <>
    <label>
      <span>接口协议</span>
      <select value={draft.protocol ?? "openai_compatible"} onChange={(event) => {
        const protocol = event.target.value as "openai_compatible" | "anthropic";
        onChange((current) => current ? { ...current, protocol } : current);
      }}>
        <option value="openai_compatible">OpenAI 兼容</option>
        <option value="anthropic">Anthropic 兼容</option>
      </select>
      <small className="model-protocol-hint">
        按供应商提供的接口选择协议，两种协议均可用于对话和公式。
      </small>
    </label>
    <p>填写服务地址、模型 ID 和 API Key，点击「测试连接」。模型正常回复后即可保存使用。</p>
  </>;
}
