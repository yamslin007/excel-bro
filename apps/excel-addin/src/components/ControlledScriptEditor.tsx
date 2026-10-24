import { useState } from "react";
import { compileControlledScript } from "@excel-bro/core/controlledScript";
import { assertEditablePlan, type AnalysisPlan } from "../contracts";
import "./controlledScriptEditor.css";

/** Shared Excel/WPS entry. It only changes an editor draft, never executes a plan. */
export function ControlledScriptEditor({ getAuthorizedSheets, draft, onApply, disabled = false }: {
  getAuthorizedSheets: () => readonly string[];
  draft: string;
  onApply: (draft: string) => void;
  disabled?: boolean;
}) {
  const [source, setSource] = useState("");
  const [feedback, setFeedback] = useState<{ error: boolean; text: string } | null>(null);
  const fromPlan = () => {
    try {
      const plan: unknown = JSON.parse(draft);
      assertEditablePlan(plan);
      setSource(plan.actions.map(action => `emit(${JSON.stringify(action, null, 2)});`).join("\n\n"));
      setFeedback(null);
    } catch (error) { setFeedback({ error: true, text: error instanceof Error ? error.message : "计划无法转换" }); }
  };
  const generate = () => {
    if (disabled) return;
    try {
      // Read scope on the click, not on mounting the editor. The execution path
      // must independently recheck the workbook identity and current permission.
      const actions = compileControlledScript(source, getAuthorizedSheets());
      const sheets = [...new Set(actions.map(action => action.sheet))];
      const plan: unknown = {
        id: `script-${crypto.randomUUID()}`,
        title: "受控 JS 生成的计划",
        summary: `对 ${sheets.join("、")} 执行 ${actions.length} 个操作。`,
        assumptions: [], warnings: [], actions
      };
      assertEditablePlan(plan);
      onApply(JSON.stringify(plan as AnalysisPlan, null, 2));
      setFeedback({ error: false, text: `已生成 ${actions.length} 个操作，涉及 ${sheets.length} 张已勾选表。请核对下方计划；生成计划不会写入工作簿。` });
    } catch (error) { setFeedback({ error: true, text: error instanceof Error ? error.message : "脚本生成失败" }); }
  };
  return (
    <details className="controlled-script-editor">
      <summary>受控 JS · 多表操作</summary>
      <p>用 authorizedSheets 取得已勾选的表名，用 emit(动作) 生成计划。支持 const、for…of、if、数组、对象与基本运算；支持写常量、清空内容、字体、对齐和行高。</p>
      <p>这是受限 JS 语法。当前不运行 VBA、原生宏或任意 JS；不能读取单元格、调用宿主、文件、网络或其他工作簿。</p>
      <div className="controlled-script-actions">
        <button type="button" className="secondary-button" disabled={disabled} onClick={fromPlan}>从当前计划生成代码</button>
        <button type="button" className="secondary-button" disabled={disabled} onClick={() => {
          setSource('for (const sheet of authorizedSheets) {\n  emit({ type: "setFont", sheet, range: "A1:B2", bold: true });\n}');
          setFeedback(null);
        }}>填入多表示例</button>
      </div>
      <textarea aria-label="受控 JS 代码" rows={8} spellCheck={false} value={source} disabled={disabled}
        onChange={event => { setSource(event.target.value); setFeedback(null); }} />
      <button type="button" className="secondary-button" disabled={disabled || !source.trim()} onClick={generate}>生成待核对计划</button>
      {feedback && <p role={feedback.error ? "alert" : "status"}>{feedback.text}</p>}
    </details>
  );
}
