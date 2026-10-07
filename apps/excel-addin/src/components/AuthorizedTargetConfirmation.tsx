import { useEffect, useId, useRef, useState } from "react";
import { targetScopeLabel, type AuthorizedTarget, type TargetCheck } from "@excel-bro/core/authorizedTarget";

export function AuthorizedTargetConfirmation({ notice, error, busy, onConfirm, onCancel }: {
  notice: { bound: AuthorizedTarget; current: AuthorizedTarget | null; reason: TargetCheck;
    round: number; retrySelection?: string[] };
  error: string; busy: boolean;
  onConfirm: (choice: "original" | "current", selected: string[]) => void;
  onCancel: () => void;
}) {
  const id = useId();
  const dialog = useRef<HTMLElement>(null);
  const [choice, setChoice] = useState<"original" | "current" | null>(null);
  const [selected, setSelected] = useState<string[]>([]);
  // Reset only when a genuinely new round starts. A rejected retry keeps the user's choice,
  // so re-confirming is one click instead of re-selecting every sheet.
  const appliedRound = useRef<number | null>(null);
  useEffect(() => {
    if (appliedRound.current === notice.round) {
      if (notice.retrySelection) setSelected(notice.retrySelection);
      return;
    }
    appliedRound.current = notice.round;
    setChoice(null);
    setSelected([]);
  }, [notice]);
  useEffect(() => {
    const previous = document.activeElement as HTMLElement | null;
    dialog.current?.focus();
    return () => { if (previous?.isConnected) previous.focus(); };
  }, []);
  return <div className="authorized-target-overlay">
    <section ref={dialog} tabIndex={-1} role="dialog" aria-modal="true" aria-labelledby={id} className="authorized-target-dialog"
      onKeyDown={(event) => {
        if (event.key === "Escape" && !busy) { event.preventDefault(); onCancel(); }
        if (event.key !== "Tab") return;
        const controls = [...(dialog.current?.querySelectorAll<HTMLElement>("button:not(:disabled), input:not(:disabled)") ?? [])];
        const first = controls[0]; const last = controls[controls.length - 1];
        if (!first || !last) { event.preventDefault(); return; }
        if (event.shiftKey && (document.activeElement === first || document.activeElement === dialog.current)) {
          event.preventDefault(); last.focus();
        } else if (!event.shiftKey && (document.activeElement === last || document.activeElement === dialog.current)) {
          event.preventDefault(); first.focus();
        }
      }}>
      <div className="authorized-target-head">
        <span className="authorized-target-icon" aria-hidden="true">
          <svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="1.6"
            strokeLinecap="round" strokeLinejoin="round" focusable="false">
            <circle cx="10" cy="10" r="7.2" />
            <circle cx="10" cy="10" r="2.6" />
            <path d="M10 1.6v2.6M10 15.8v2.6M1.6 10h2.6M15.8 10h2.6" />
          </svg>
        </span>
        <div className="authorized-target-heading">
          <h3 id={id}>{notice.reason === "unavailable" ? "需要重新确认操作目标" : "检测到文件或工作表变化"}</h3>
          <p>操作已暂停，需求与回答仍保留。请选择目标，确认后才会读取和生成预览；不会直接写入。</p>
        </div>
      </div>
      {notice.current?.name === notice.bound.name && notice.reason !== "scope_changed"
        && <p className="authorized-target-note">名称相同不代表同一份文件，请确认你要操作的目标。</p>}
      <div className="authorized-target-options">
        <div className={choice === "original" ? "authorized-target-option is-selected" : "authorized-target-option"}
          onClick={() => { if (!busy) setChoice("original"); }}>
          <label>
            <input type="radio" name={id} checked={choice === "original"} disabled={busy}
              onChange={() => setChoice("original")} />
            <span>继续使用原授权：{targetScopeLabel(notice.bound)}</span>
          </label>
          <span className="authorized-target-badge" aria-hidden="true">原授权</span>
          {choice === "original" && <p className="authorized-target-option-hint">请先切回原文件。系统会重新核对，不会自动操作当前其他文件。</p>}
        </div>
        {notice.current && <div className={choice === "current" ? "authorized-target-option is-selected" : "authorized-target-option"}
          onClick={() => { if (!busy && choice !== "current") { setChoice("current"); setSelected([]); } }}>
          <label>
            <input type="radio" name={id} checked={choice === "current"} disabled={busy}
              onChange={() => { setChoice("current"); setSelected([]); }} />
            <span>改用当前文件：{notice.current.name}</span>
          </label>
          <span className="authorized-target-badge" aria-hidden="true">当前文件</span>
          {choice === "current" && <fieldset><legend>勾选本次允许操作的工作表</legend>
            {notice.current.sheets.map((sheet) => <label key={sheet.id}>
              <input type="checkbox" checked={selected.includes(sheet.id)} disabled={busy}
                onChange={() => setSelected((value) => value.includes(sheet.id)
                  ? value.filter((item) => item !== sheet.id) : [...value, sheet.id])} />{sheet.name}
            </label>)}
            <p>重新读取并生成预览；旧计划和文件专属映射不会套用到新目标。</p>
          </fieldset>}
        </div>}
      </div>
      {error && <p role="alert" className="authorized-target-error">{error}</p>}
      <div className="authorized-target-actions">
        <button type="button" disabled={busy} onClick={onCancel}>暂不操作</button>
        <button type="button" className="primary" disabled={busy || !choice || (choice === "current" && !selected.length)}
          onClick={() => choice && onConfirm(choice, selected)}>确认目标并继续</button>
      </div>
    </section>
  </div>;
}
