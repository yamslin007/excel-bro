import { targetScopeLabel, type AuthorizedTarget, type TargetNoticeReason } from "@excel-bro/core/authorizedTarget";

/**
 * 输入框上方的范围提示（**只是状态**，不改授权）。
 *
 * `confirmation_required` 这一支是"证据不足、要用户点一下"的情况，与"你改了勾选"
 * 是两件事，所以：
 * - 有原样确认入口时，卡片**直接给**「确认范围并继续」，复用同一个确认处理器
 *   （确认的就是当前这份范围，不需要取消勾选再勾回）；
 * - `restoredScope` 为真（本页面刚重开、勾选是从本地固定范围恢复的）时，明说
 *   "窗格已重新加载"，并点明勾选没有变动 —— 不让用户以为是自己改坏了什么。
 */
export function TargetScopeNotice({ notice, error, busy, onOpenScope, onRecheck, onConfirmOriginal, restoredScope }: {
  notice: { bound: AuthorizedTarget; current: AuthorizedTarget | null; reason: TargetNoticeReason; round: number };
  error: string; busy: boolean; onOpenScope: () => void;
  onRecheck?: () => void;
  /** 原样确认当前范围（复用 `confirm("original")`），不走选择面板。 */
  onConfirmOriginal?: () => void;
  /** 本页面的勾选是从持久化的固定范围恢复来的（任务窗格刚重载）。 */
  restoredScope?: boolean;
}) {
  const unknown = notice.reason === "unavailable";
  const confirmation = notice.reason === "confirmation_required";
  const title = unknown ? "暂时无法核实当前目标" : confirmation
    ? "请确认本次使用的数据范围" : notice.reason === "file_changed" ? "当前文件与已授权文件不同" : "工作表范围发生变化";
  /** 要确认的份数：确认的是**已提交的那份范围**，不是草稿。 */
  const count = (notice.current ?? notice.bound).selected.length;
  return <section className="target-scope-notice" role="status" aria-label="数据范围提示">
    <strong>{title}</strong>
    {unknown && <span>不代表文件已变化。</span>}
    {confirmation && restoredScope && <span>
      窗格已重新加载，请确认继续使用已选的 {count} 张工作表。
    </span>}
    {confirmation && <small>工作表勾选没有变动；确认的就是当前这份范围，不需要取消勾选再勾回。</small>}
    <span>{unknown
      ? "已保留勾选范围，需求与回答仍保留。重新检查通过后沿用原授权；核实前保持暂停，尚未写入。"
      : confirmation
        ? "确认后按原授权继续这一次；同一会话内范围未变时不再重复询问。确认前保持暂停，需求与回答仍保留，尚未写入。"
        : "请在数据范围中检查并确认；确认前保持暂停，需求与回答仍保留，尚未写入。"}</span>
    {(notice.reason === "file_changed" || notice.reason === "scope_changed") && <small>
      原范围：{targetScopeLabel(notice.bound)}；当前：{notice.current ? targetScopeLabel(notice.current) : "暂不可读取"}
    </small>}
    {error && <span role="alert">{error}</span>}
    <div>
      {confirmation && onConfirmOriginal && <button type="button" className="primary" disabled={busy}
        onClick={onConfirmOriginal}>确认范围并继续</button>}
      {unknown && onRecheck && <button type="button" disabled={busy} onClick={onRecheck}>重新检查</button>}
      <button type="button" disabled={busy} onClick={onOpenScope}>检查数据范围</button>
    </div>
  </section>;
}
