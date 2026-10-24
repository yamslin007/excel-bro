import { targetScopeLabel, type AuthorizedTarget, type TargetNoticeReason } from "@excel-bro/core/authorizedTarget";

/** A status beside the existing scope selector, never a second file-choice dialog. */
export function TargetScopeNotice({ notice, error, busy, onOpenScope, onRecheck }: {
  notice: { bound: AuthorizedTarget; current: AuthorizedTarget | null; reason: TargetNoticeReason; round: number };
  error: string; busy: boolean; onOpenScope: () => void;
  onRecheck?: () => void;
}) {
  const unknown = notice.reason === "unavailable";
  const title = unknown ? "暂时无法核实当前目标" : notice.reason === "confirmation_required"
    ? "请确认本次使用的数据范围" : notice.reason === "file_changed" ? "当前文件与已授权文件不同" : "工作表范围发生变化";
  return <section className="target-scope-notice" role="status" aria-label="数据范围提示">
    <strong>{title}</strong>
    {unknown && <span>不代表文件已变化。</span>}
    <span>{unknown
      ? "已保留勾选范围，需求与回答仍保留。重新检查通过后沿用原授权；核实前保持暂停，尚未写入。"
      : "请在数据范围中检查并确认；确认前保持暂停，需求与回答仍保留，尚未写入。"}</span>
    {(notice.reason === "file_changed" || notice.reason === "scope_changed") && <small>
      原范围：{targetScopeLabel(notice.bound)}；当前：{notice.current ? targetScopeLabel(notice.current) : "暂不可读取"}
    </small>}
    {error && <span role="alert">{error}</span>}
    <div>
      {unknown && onRecheck && <button type="button" disabled={busy} onClick={onRecheck}>重新检查</button>}
      <button type="button" disabled={busy} onClick={onOpenScope}>检查数据范围</button>
    </div>
  </section>;
}
