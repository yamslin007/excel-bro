import type { ResultReviewState } from '@excel-bro/core/resultReview';

/** Kept outside the scrolling conversation so the decision stays reachable. */
export function ResultReviewBar({ state, title, onCompare, onConfirm, onDiscard }: {
  state: ResultReviewState; title: string;
  onCompare(): void; onConfirm(): void; onDiscard(): void;
}) {
  if (state.phase === 'idle') return null;
  return <section className="result-review-bar" aria-label="本次结果审阅" aria-busy={state.pending}>
    <strong>{state.phase === 'executing' ? '正在执行并准备结果' : state.phase === 'reviewing' ? '请查看主表中的执行结果'
      : state.phase === 'confirmed' ? '已确认变更' : state.phase === 'discarded' ? '已放弃并恢复' : '本次执行需要核对'}</strong>
    <span>{title}</span>
    {state.phase === 'reviewing' && <>
      <p>主表已更新。{state.marked ? '彩色边框是审阅标记；' : ''}双视图中的副本展示执行前状态。</p>
      <div className="result-review-actions">
        <button type="button" disabled={state.pending || state.comparing} onClick={onCompare}>双视图对比</button>
        <button type="button" disabled={state.pending} onClick={onDiscard}>放弃变更</button>
        <button type="button" className="primary-button" disabled={state.pending} onClick={onConfirm}>确认变更</button>
      </div>
    </>}
    {state.phase === 'confirmed' && <p>已接受当前结果并移除审阅标记；保存为工具和撤销入口已开放。确认变更不等于保存磁盘文件。</p>}
    {state.phase === 'discarded' && <p>已按本次执行的恢复点还原主表。</p>}
    {state.notice && <p role="status">{state.notice}</p>}
  </section>;
}
