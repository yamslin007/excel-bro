import type { ModelOption } from "../contracts";
import { ModelPickerIcon } from "./ModelPickerIcon";
import { ModelBrandIcon } from "./ModelBrandIcon";

/** Catalog IDs identify connections; only modelId is the configured upstream model. */
export function ModelPickerLabel({ model }: { model: ModelOption | null | undefined }) {
  const modelId = model?.modelId?.trim() ||
    (model && !/^(connection:|claude-agent-sdk)/.test(model.id) ? model.id : "");
  const label = modelId || model?.label || "模型";
  return <>
    <ModelBrandIcon modelId={modelId} fallback={<ModelPickerIcon className="model-picker-symbol" />} />
    <span className="model-picker-name" title={label}>{label}</span>
    <svg className="model-picker-chevron" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="m4 6 4 4 4-4" /></svg>
  </>;
}
