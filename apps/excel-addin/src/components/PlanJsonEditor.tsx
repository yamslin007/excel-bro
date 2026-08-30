import { useEffect, useRef, useState } from "react";
import type { PlanEditorState } from "../hooks/usePlanEditor";
import { JsonCodeEditor } from "./JsonCodeEditor";

interface PlanJsonEditorProps {
  state: PlanEditorState;
  onDraftChange: (text: string) => void;
  onSave: () => void;
  onClose: () => void;
}

/**
 * 计划的 JSON 编辑抽屉。
 *
 * 故意用普通 textarea：Monaco / CodeMirror 会把包体积翻倍，而这里编辑的是几十行
 * JSON，收益撑不起成本。保存前做本地 JSON 语法 + 白名单结构检查。
 */
export function PlanJsonEditor({
  state,
  onDraftChange,
  onSave,
  onClose
}: PlanJsonEditorProps) {
  const [dirty, setDirty] = useState(false);
  const initialDraftRef = useRef<string | null>(null);

  // 每次打开时记录初始内容，用于「已修改」标记。
  useEffect(() => {
    if (state.phase === "editing" && initialDraftRef.current === null) {
      initialDraftRef.current = state.draft;
    }
    if (state.phase === "closed") {
      initialDraftRef.current = null;
      setDirty(false);
    }
  }, [state.phase, state.draft]);

  useEffect(() => {
    if (initialDraftRef.current !== null) {
      setDirty(state.draft !== initialDraftRef.current);
    }
  }, [state.draft]);

  // Ctrl+S 保存
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
        event.preventDefault();
        if (state.phase === "editing") onSave();
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [state.phase, onSave]);

  if (state.phase === "closed") return null;

  return (
    <aside className="tool-drawer plan-editor-drawer" aria-label="编辑执行脚本">
      <div className="tool-drawer-header">
        <div>
          <strong>编辑执行脚本</strong>
          <span>保存前会检查 JSON 格式与白名单结构</span>
        </div>
        <button onClick={onClose} aria-label="关闭">
          ×
        </button>
      </div>

      <div className="plan-editor-body">
        <JsonCodeEditor
          value={state.draft}
          onChange={onDraftChange}
          aria-label="计划 JSON"
          externalError={state.errorMessage}
          dirty={dirty}
        />

        <div className="inline-plan-buttons">
          <button className="secondary-button" onClick={onClose}>
            放弃修改
          </button>
          <button className="primary-button" onClick={onSave}>保存</button>
        </div>
      </div>
    </aside>
  );
}
