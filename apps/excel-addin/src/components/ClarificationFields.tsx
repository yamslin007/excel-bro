import type { ClarificationField } from "../contracts";

// No hooks: both host workspaces can render this with their own React runtime.
export function clarificationFieldsAnswer(
  fields: ClarificationField[], values: Record<string, string>
): string {
  if (fields.some((field) => !values[field.id]?.trim())) return "";
  return fields.map((field) => `${field.label}：${values[field.id].trim()}`).join("\n");
}

export function ClarificationFields({ fields, values, busy, onChange, onSubmit }: {
  fields: ClarificationField[];
  values: Record<string, string>;
  busy: boolean;
  onChange: (values: Record<string, string>) => void;
  onSubmit: (answer: string) => void;
}) {
  const answer = clarificationFieldsAnswer(fields, values);
  return <div className="clarification-fields">
    {fields.map((field) => <label key={field.id}>
      <span>{field.label}</span>
      <textarea aria-label={field.label} rows={2} maxLength={1000}
        placeholder={field.placeholder || "请填写你的选择或处理方式"}
        value={values[field.id] ?? ""} disabled={busy}
        onChange={(event) => onChange({ ...values, [field.id]: event.target.value })} />
    </label>)}
    <small>补充后将重新生成预览，不会直接修改工作表。</small>
    <button type="button" disabled={busy || !answer}
      onClick={() => { if (answer && !busy) onSubmit(answer); }}>提交补充并生成预览</button>
  </div>;
}
