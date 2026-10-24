import type { ResultReviewState } from '@excel-bro/core/resultReview';

/** Kept outside the scrolling conversation so the decision stays reachable. */
export function ResultReviewBar({ state, title, onCompare, onConfirm, onDiscard, onRetryMarker }: {
  state: ResultReviewState; title: string;
  onCompare(): void; onConfirm(): void; onDiscard(): void;
  /** 只重试"移除本次标记"：不重新执行、不重新撤销。 */
  onRetryMarker?(): void;
}) {
  const settled = state.phase === 'confirmed' || state.phase === 'discarded';
  /*
   * **成功终态不再占着输入框上方。**
   *
   * 用户确认/放弃之后，固定的那一大块只重复"已确认变更"，却把输入区挤掉一截。
   * 业务状态此刻已经由对应消息承载（确认/恢复记录、保存工具与撤销入口），
   * 所以这里收起来；只有"确实没做完"的收尾问题（标记未清、副本未关）才留一条
   * 紧凑提示，并附上它需要的操作。
   */
  const needsAttention = state.pending || Boolean(state.markerPending) || Boolean(state.notice);
  if (state.phase === 'idle' || (settled && !needsAttention)) return null;
  if (settled) {
    return <section className="result-review-bar is-settled" aria-label="本次结果审阅" aria-busy={state.pending}>
      {state.notice && <p role="status">{state.notice}</p>}
      {state.markerPending && onRetryMarker && (
        <div className="result-review-actions">
          <button type="button" className="secondary-button" disabled={state.pending} onClick={onRetryMarker}>重试清理标记</button>
        </div>
      )}
    </section>;
  }
  return <section className="result-review-bar" aria-label="本次结果审阅" aria-busy={state.pending}>
    <strong>{state.phase === 'executing' ? '正在执行并准备结果' : state.phase === 'reviewing' ? '请查看主表中的执行结果'
      : state.phase === 'confirmed' ? '已确认变更' : state.phase === 'discarded' ? '已放弃并恢复' : '本次执行需要核对'}</strong>
    <span>{title}</span>
    {state.phase === 'reviewing' && <>
      <p>主表已更新。{state.marked ? '变更区域用浅黄底色或边框标出，避免遮住实际格式；同时修改填充与边框的区域直接显示实际效果；' : ''}双视图中的副本展示执行前状态，仅比较本次修改区域。公式采用当时的值，不重新计算；图表、合并与条件格式效果不作为本次对比依据。</p>
      <div className="result-review-actions">
        <button type="button" className="secondary-button" disabled={state.pending || state.comparing} onClick={onCompare}>双视图对比</button>
        <button type="button" className="secondary-button" disabled={state.pending} onClick={onDiscard}>放弃变更</button>
        <button type="button" className="primary-button" disabled={state.pending} onClick={onConfirm}>确认变更</button>
      </div>
    </>}
    {state.phase === 'confirmed' && <p>已接受当前结果并移除审阅标记；保存为工具和撤销入口已开放。确认变更不等于保存磁盘文件。</p>}
    {state.phase === 'discarded' && <p>已按本次执行的恢复点还原主表。</p>}
    {state.notice && <p role="status">{state.notice}</p>}
    {/*
      标记清理失败是**收尾问题**，不是业务失败：单独给一个只清理标记的入口。
      少了它，用户要么以为整件事失败了，要么被迫重新执行一遍。
    */}
    {state.markerPending && onRetryMarker && (
      <div className="result-review-actions">
        <button type="button" className="secondary-button" disabled={state.pending} onClick={onRetryMarker}>重试清理标记</button>
      </div>
    )}
  </section>;
}
