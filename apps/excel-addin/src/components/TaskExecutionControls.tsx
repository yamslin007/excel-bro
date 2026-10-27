import type { PauseButtonState } from "@excel-bro/core/pauseControl";
import "./taskExecutionControls.css";

/** Task actions never submit the composer. Each action owns a different DOM node. */
export function TaskExecutionControls({ state, pausing, reviewing, disabled, onPause, onResume }: {
  state: PauseButtonState;
  pausing: boolean;
  reviewing?: boolean;
  disabled?: boolean;
  onPause: () => void;
  onResume: () => void;
}) {
  if (reviewing) return <div className="task-execution-controls" aria-label="当前任务控制"><span>等待审阅结果</span></div>;
  if (state.kind === "send" && !pausing) return null;
  const action = pausing ? "pending" : state.kind;
  const label = pausing ? "正在暂停" : state.label;
  return <div className="task-execution-controls" aria-label="当前任务控制">
    <span>{pausing ? "正在暂停…" : action === "continue" ? "任务已暂停，点击 ▶ 继续" : action === "pause" ? "任务进行中" : state.reason}</span>
    <button key={action} type="button" className="task-execution-button"
      aria-label={label} title={pausing ? "正在停止当前处理，请稍候" : state.title}
      disabled={pausing || disabled || action === "unavailable"}
      onClick={(event) => {
        event.preventDefault();
        event.stopPropagation();
        if (event.detail > 1) return;
        if (action === "pause") onPause();
        else if (action === "continue") onResume();
      }}>
      {pausing ? "…" : state.glyph}
    </button>
  </div>;
}
