import {
  clarificationSubmission,
  type ClarificationFieldLike,
  type ClarificationOptionLike
} from "@excel-bro/core/clarificationComposer";
import "./clarificationComposer.css";

/**
 * 澄清回答区（Excel / WPS 共用），嵌入 composer-box 内部。
 *
 * 这一轮把快捷回答从消息里搬到这里，并改成**点选只选中、按发送才提交**：
 *  - 当前问题、选项、多字段和预览在输入框内部，长内容内部滚动；
 *  - 历史问题保留在消息流；不把一次追问伪装成已知总数的多题流程；
 *  - **没有自己的发送按钮**：发送是输入框右下角那颗主操作唯一负责的事。
 *    这里再画一颗，同一个动作就有两个入口，而且其中一颗不受输入框状态影响。
 *
 * 三条边界：
 * 1. 选项只报告"选中了哪个 id"，不在这里拼答案 —— 提交内容由共享层的
 *    `clarificationSubmission` 决定，两端同一份。
 * 2. 必填没填完时，预览行说的是"还差哪几项"，而不是把半张表单当成答案。
 * 3. 已作废的题目**不渲染**：调用方用 `clarificationSubmissionAllowed` 判过绑定，
 *    这里拿不到题就不出现，绝不把用户写给旧问题的答案递给新问题。
 */
export function ClarificationComposer({
  question,
  options,
  selected,
  multiple,
  fields,
  fieldValues,
  freeText,
  busy,
  onToggle,
  onFieldChange
}: {
  question: string;
  options: readonly ClarificationOptionLike[];
  selected: readonly string[];
  multiple: boolean;
  fields?: readonly ClarificationFieldLike[];
  fieldValues?: Record<string, string>;
  freeText: string;
  busy: boolean;
  onToggle: (optionId: string) => void;
  onFieldChange: (values: Record<string, string>) => void;
}) {
  const submission = clarificationSubmission({
    options, selected, freeText, fields, fieldValues
  });
  return (
    <section className="clarification-composer" aria-label="回答待确认的问题">
      <header className="clarification-composer-heading">
        <span className="clarification-composer-mode">需要确认</span>
        <p className="clarification-composer-question">{question}</p>
      </header>
      <div className="clarification-choices-scroll">
      {fields && fields.length > 0 ? (
        <div className="clarification-fields">
          {fields.map((field) => (
            <label key={field.id}>
              <span>{field.label}</span>
              <textarea
                aria-label={field.label}
                rows={2}
                maxLength={1000}
                placeholder="请填写你的选择或处理方式"
                value={fieldValues?.[field.id] ?? ""}
                disabled={busy}
                onChange={(event) => onFieldChange({ ...(fieldValues ?? {}), [field.id]: event.target.value })}
              />
            </label>
          ))}
        </div>
      ) : null}
      {options.length > 0 ? (
        <div className="clarification-options" role="group"
          aria-label={multiple ? "可多选" : "单选"}>
          {options.map((option) => {
            const chosen = selected.includes(option.id);
            return (
              <button
                key={option.id}
                type="button"
                aria-pressed={chosen}
                data-selected={chosen ? "true" : "false"}
                disabled={busy}
                title={chosen ? "再点一次取消选中" : "选中它；按发送才提交"}
                onClick={() => onToggle(option.id)}
              >
                <strong>{option.label}</strong>
                {option.description ? <span>{option.description}</span> : null}
              </button>
            );
          })}
        </div>
      ) : null}
      {(submission.ok || submission.missing.length > 0) && <p className="clarification-composer-preview" role="status" data-ready={submission.ok ? "true" : "false"}>
        {submission.ok
          ? `将发送：${submission.answer.replace(/\n/g, " / ")}`
          : submission.reason}
      </p>}
      </div>
    </section>
  );
}

export default ClarificationComposer;
