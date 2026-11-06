import { ReviewSheetPicker } from './ReviewSheetPicker';
import { useState } from 'react';
import type { ResultReviewState } from '@excel-bro/core/resultReview';
import { isCopyCleanupNotice } from '@excel-bro/core/operationFeedback';
import './resultReviewComposer.css';

/**
 * 展示副本没清理掉的**独立**面板。
 *
 * 「工作簿已恢复 / 已确认变更」是业务结论；这里只说对照用的展示副本还开着。
 * 副本失败按真实身份去重后给数量，逐份原因可展开，既不暗示恢复失败，
 * 也不假称全部清理完成。
 */
function CopyCleanupBlock({ state }: { state: ResultReviewState }) {
  const cleanup = state.copyCleanup;
  if (!cleanup || cleanup.count === 0) return null;
  return (
    <section className="result-review-copy-cleanup" aria-label="展示副本清理结果"
      data-copy-issue-count={cleanup.count}>
      <strong>{cleanup.headline}</strong>
      <p>工作簿的结果已按本次执行结论处理；这里只说展示副本未关闭，不代表恢复失败。</p>
      <details>
        <summary>查看 {cleanup.count} 份副本的原因</summary>
        <ul>
          {cleanup.issues.map((issue) => (
            <li key={issue.identity}>
              <span>{issue.label ?? issue.identity}</span>：{issue.reason}
            </li>
          ))}
        </ul>
      </details>
    </section>
  );
}

/** Keep active answers/IME intact; drafts stay in the parent's state while reviewing. */
export function resultReviewComposerActive(state: ResultReviewState, answering: boolean, composing: boolean): boolean {
  return !answering && !composing && (state.phase === 'reviewing' || Boolean(state.markerPending));
}

/**
 * 审阅栏这一刻会不会画出来 —— **唯一**出处。
 *
 * 固定区域容器（`ComposerActionDock`）要按"有没有这一块"决定展开谁，如果它自己
 * 另判一次（例如只看 `phase === 'reviewing'`），"未执行"/"需要核对"这些形态就会被
 * 判成没有审阅栏，于是被执行摘要或恢复卡顶掉 —— 那正是"算出一个名字、界面却少一块"。
 */
export function resultReviewVisible(state: ResultReviewState): boolean {
  const settled = state.phase === 'confirmed' || state.phase === 'discarded';
  const needsAttention = state.pending || Boolean(state.markerPending) || Boolean(state.decisionPending) || Boolean(state.copyCleanup?.count) || Boolean(state.notice);
  return !(state.phase === 'idle' || (settled && !needsAttention));
}

/**
 * 这块审阅栏是不是**在等你决定**（确认 / 放弃 / 重试清理标记）。
 *
 * 与"会不会画出来"分开：`未执行`、`正在执行`、`需要核对`这些形态也会画出来，
 * 但它们身上没有决定可做 —— 它们是**状态**。固定区域容器据此决定谁是主视图：
 * 一条状态行不该把真正的恢复入口挤成 chip（那会变成"卡片说去恢复，恢复卡片却不在"）。
 */
export function resultReviewNeedsDecision(state: ResultReviewState): boolean {
  const settled = state.phase === 'confirmed' || state.phase === 'discarded';
  const needsAttention = state.pending || Boolean(state.markerPending) || Boolean(state.decisionPending) || Boolean(state.copyCleanup?.count) || Boolean(state.notice);
  return state.phase === 'reviewing' || (settled && needsAttention);
}

/**
 * 只是状态的审阅栏（未执行 / 正在执行 / 需要核对）要不要**因为恢复操作区已经接管
 * 输入框而收起** —— **唯一**出处，Excel / WPS 读同一份。
 *
 * 背景（2026-11-02）：一次执行写到一半失败（回执丢失 / 某步失败），会同时出现
 * 两处都在说同一件事 —— 输入框里那块恢复操作区（已写入 / 尚未写入 / 检查按钮），
 * 和工作簿范围条上方这条「本次执行需要核对」状态卡。这条卡只重复"执行未完成，去
 * 看恢复入口"，而那扇入口此刻就摆在输入框里，用户被要求先读两遍才知道点哪。
 *
 * 收起条件**必须按执行身份匹配**（`executionId === operationId`），不能因为"任意"
 * 一处恢复操作区在场就隐藏另一条任务的状态或告警：旧的审阅状态与后来发现的中断
 * 执行不是同一次执行时，这条卡照旧要留着。两端执行身份都拿不到时（`null`）也
 * **不收起** —— 拿不准就别替用户藏告警。
 */
export function reviewStatusBarSuppressedByRecovery(args: {
  /** 恢复操作区此刻是不是真的接管了输入框（`mode === "recovery"`）。 */
  recoveryComposerActive: boolean;
  /** 审阅归属的那次执行（`resultReviewState.executionId`，失败路径可能为空）。 */
  reviewExecutionId: string | null;
  /** 恢复操作区对应那次执行（恢复入口的 `operationId`）。 */
  recoveryOperationId: string | null;
  /** Failed writes have no receipt; accept their binding only for the current review task. */
  reviewKey?: string | null;
  failedExecution?: { key: string | null; operationId: string } | null;
}): boolean {
  if (!args.recoveryComposerActive) return false;
  const executionId = args.reviewExecutionId ?? (
    args.reviewKey && args.failedExecution?.key === args.reviewKey
      ? args.failedExecution.operationId : null
  );
  return Boolean(executionId && args.recoveryOperationId && executionId === args.recoveryOperationId);
}

/** Kept outside the scrolling conversation so the decision stays reachable. */
export function ResultReviewBar({ state, title, compareUnavailable, onCompare, onExitComparison, onConfirm, onDiscard, onRetryMarker, onRetryCopy, onRetryDecision, onComparisonSheet, paneNeedsResize = false, inComposer = false }: {
  state: ResultReviewState; title: string;
  inComposer?: boolean;
  paneNeedsResize?: boolean;
  onComparisonSheet?(sheet: string): void;
  onExitComparison?(): void;
  /**
   * 本次**没有**执行前副本可对比时的原因（多张写入目标表、跨表公式、区域超限…）。
   *
   * 有它就说明"比对做不了"，于是不画「双视图对比」——画一个点下去只会报错的按钮
   * 等于骗人。执行与确认/放弃都不受影响：比对是审阅的辅助，不是执行许可。
   */
  compareUnavailable?: string;
  onCompare(): void; onConfirm(): void; onDiscard(): void;
  /** 只重试"移除本次标记"：不重新执行、不重新撤销。 */
  onRetryMarker?(): void;
  onRetryCopy?(): void;
  onRetryDecision?(): void;
}) {
  const [hiddenPickerKey, setHiddenPickerKey] = useState<string | null>(null);
  const multipleSheets = (state.comparisonSheets?.length ?? 0) > 1 && Boolean(onComparisonSheet);
  const pickerVisible = state.comparing && multipleSheets && hiddenPickerKey !== state.key;
  const settled = state.phase === 'confirmed' || state.phase === 'discarded';
  /*
   * **成功终态不再占着输入框上方。**
   *
   * 用户确认/放弃之后，固定的那一大块只重复"已确认变更"，却把输入区挤掉一截。
   * 业务状态此刻已经由对应消息承载（确认/恢复记录、保存工具与撤销入口），
   * 所以这里收起来；只有"确实没做完"的收尾问题（标记未清、副本未关）才留一条
   * 紧凑提示，并附上它需要的操作。
   */
  const needsAttention = state.pending || Boolean(state.markerPending) || Boolean(state.decisionPending) || Boolean(state.copyCleanup?.count) || Boolean(state.notice);
  // 副本清理提示已有结构化面板，别再用纯文本把同一句话画第二遍。
  const plainNotice = state.notice && !isCopyCleanupNotice(state.notice, state.copyCleanup?.headline)
    ? state.notice : null;
  const copyCleanup = <>
    <CopyCleanupBlock state={state} />
    {state.copyCleanup && onRetryCopy && <button type="button" className="secondary-button" disabled={state.pending || Boolean(state.decisionPending)} onClick={onRetryCopy}>重试清理副本</button>}
    {state.decisionPending && onRetryDecision && <button type="button" className="secondary-button" disabled={state.pending} onClick={onRetryDecision}>重试记录结果</button>}
  </>;
  if (!resultReviewVisible(state)) return null;
  if (inComposer) {
    return <section className="result-review-bar result-review-composer" aria-label="本次结果审阅" aria-busy={state.pending}>
      <div className="result-review-heading">
        <strong>{state.phase === 'reviewing' ? `待审阅${state.comparisonSheets?.length ? ` · ${state.comparisonSheets.length} 张表` : ''}` : '变更已处理，标记待清理'}</strong>
        {!compareUnavailable && state.phase === 'reviewing' && state.comparing && multipleSheets && <button type="button" className="result-review-picker-toggle" disabled={state.pending} aria-expanded={pickerVisible} onClick={() => setHiddenPickerKey(pickerVisible ? state.key : null)}>{pickerVisible ? '收起选表' : '展开选表'}</button>}
      </div>
      <div className="result-review-composer-scroll">
        {compareUnavailable && <p role="status">{compareUnavailable}</p>}
        {state.phase === 'reviewing' && pickerVisible && onComparisonSheet && <ReviewSheetPicker
          key={`sheets:${state.key}`} sheets={state.comparisonSheets!} selected={state.comparisonSheet} pending={state.pending} onSelect={onComparisonSheet} />}
        <details key={state.key}>
          <summary>查看详情</summary>
          <p>{title}</p>
          <p>确认变更只接受当前结果并移除审阅标记，不会再次执行，也不等于保存文件。放弃变更会按已验证恢复点撤销本次修改。</p>
          {state.marked && <p>本次变更区域用浅黄底色或边框标出；修改过填充或边框的区域显示实际效果。</p>}
          {!compareUnavailable && <details>
            <summary>比对范围</summary>
            {state.comparisonHelp?.scope && <p>{state.comparisonHelp.scope}</p>}
            <p>双视图中的副本展示执行前状态。副本保留到确认或成功放弃，退出双视图后可再次打开。</p>
          </details>}
          {state.comparing && (paneNeedsResize || state.comparisonHelp?.layout) && <details>
            <summary>布局说明</summary>
            {paneNeedsResize && <p>宿主未能自动调窄聊天窗格，请拖动窗格边界，让原表和副本同时可见。确认和放弃变更仍可使用。</p>}
            {state.comparisonHelp?.layout && <p>{state.comparisonHelp.layout}</p>}
          </details>}
        </details>
      </div>
      {plainNotice && <p className="result-review-feedback" role="status">{plainNotice}</p>}
      {copyCleanup}
      {state.pending && <p className="result-review-feedback" role="status">处理中，请稍候…</p>}
      <div className="result-review-actions">
        {state.phase === 'reviewing' && <>
          {!compareUnavailable && !state.comparing && <button type="button" className="secondary-button" disabled={state.pending} onClick={() => { setHiddenPickerKey(null); onCompare(); }}>双视图对比</button>}
          {state.comparing && onExitComparison && <button type="button" className="secondary-button" disabled={state.pending} onClick={onExitComparison}>退出双视图</button>}
          <button type="button" className="secondary-button" disabled={state.pending} onClick={onDiscard}>放弃全部变更</button>
          <button type="button" className="primary-button" disabled={state.pending} onClick={onConfirm}>确认全部变更</button>
        </>}
        {state.markerPending && onRetryMarker && <button type="button" className="secondary-button" disabled={state.pending} onClick={onRetryMarker}>重试清理标记</button>}
      </div>
    </section>;
  }
  if (settled) {
    return <section className="result-review-bar is-settled" aria-label="本次结果审阅" aria-busy={state.pending}>
      {plainNotice && <p role="status">{plainNotice}</p>}
      {copyCleanup}
      {state.markerPending && onRetryMarker && (
        <div className="result-review-actions">
          <button type="button" className="secondary-button" disabled={state.pending} onClick={onRetryMarker}>重试清理标记</button>
        </div>
      )}
    </section>;
  }
  return <section className="result-review-bar" aria-label="本次结果审阅" aria-busy={state.pending}>
    <strong>{state.phase === 'executing' ? '正在执行并准备结果' : state.phase === 'reviewing' ? '请查看主表中的执行结果'
      : state.phase === 'confirmed' ? '已确认变更' : state.phase === 'discarded' ? '已放弃并恢复' : state.phase === 'not_started' ? '未执行' : '本次执行需要核对'}</strong>
    <span>{title}</span>
    {state.phase === 'reviewing' && <>
      <details>
        <summary>查看详情</summary>
        <p>确认变更只接受当前结果；放弃变更按已验证恢复点撤销本次修改。</p>
        {!compareUnavailable && <details><summary>比对范围</summary>
          {state.comparisonHelp?.scope && <p>{state.comparisonHelp.scope}</p>}
          <p>双视图中的副本展示执行前状态。副本保留到确认或成功放弃，退出双视图后可再次打开。</p>
        </details>}
        {state.comparing && (paneNeedsResize || state.comparisonHelp?.layout) && <details><summary>布局说明</summary>
          {paneNeedsResize && <p>请拖动聊天窗格边界，让原表和副本同时可见。确认和放弃变更仍可使用。</p>}
          {state.comparisonHelp?.layout && <p>{state.comparisonHelp.layout}</p>}
        </details>}
      </details>
      {compareUnavailable && <p role="status">{compareUnavailable}</p>}
      <div className="result-review-actions">
        {!compareUnavailable && <button type="button" className="secondary-button" disabled={state.pending || state.comparing} onClick={onCompare}>双视图对比</button>}
        {state.comparing && onExitComparison && <button type="button" className="secondary-button" disabled={state.pending} onClick={onExitComparison}>退出双视图</button>}
        <button type="button" className="secondary-button" disabled={state.pending} onClick={onDiscard}>放弃全部变更</button>
        <button type="button" className="primary-button" disabled={state.pending} onClick={onConfirm}>确认全部变更</button>
      </div>
    </>}
    {state.phase === 'confirmed' && <p>已接受当前结果并移除审阅标记；保存为工具和撤销入口已开放。确认变更不等于保存磁盘文件。</p>}
    {state.phase === 'discarded' && <p>已按本次执行的恢复点还原主表。</p>}
    {plainNotice && <p role="status">{plainNotice}</p>}
    {copyCleanup}
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
