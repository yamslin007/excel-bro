import { useRef, useState } from "react";
import type { AnalysisPlan } from "../contracts";
import { excelNativeReviewMode, type ExcelNativeComparison } from "../excelNativeComparison";
import { excelReviewBlocksNewCopy, type useExcelNativeComparison } from "../hooks/useExcelNativeComparison";
import { excelFormatCapabilityNotice, resolveExcelFormatCapability } from "../excelFormatCapability";
import { reviewCardView, reviewCardHasDiagnostics, reviewChangeSummary, reviewScopeLocations,
  planTargetCells, REVIEW_CARD_COMPAT_DETAIL, REVIEW_CARD_COMPAT_NOTICE,
  REVIEW_CARD_COMPAT_DETAIL_TITLE, REVIEW_CARD_DIAGNOSTICS_TITLE } from "../../../../packages/core/src/reviewCardPresentation";

export function ExcelNativeReviewActions({ messageId, plan, authorization, controller, disabled, commit, paneNeedsResize = false, beforeOperation }: {
  messageId: string; plan: AnalysisPlan; authorization: readonly string[];
  controller: ReturnType<typeof useExcelNativeComparison>; disabled: boolean;
  commit: (session: ExcelNativeComparison | null) => Promise<boolean>;
  paneNeedsResize?: boolean;
  beforeOperation?: () => Promise<boolean>;
}) {
  const [acknowledgedCopy, setAcknowledgedCopy] = useState<string | null>(null);
  const review = controller.review?.messageId === messageId ? controller.review : null;
  const feedback = controller.feedback?.messageId === messageId ? controller.feedback.text : null;
  const current = useRef({ review, plan, authorization, acknowledgedCopy });
  current.current = { review, plan, authorization, acknowledgedCopy };
  const manual = review?.session.mode === 'manual';
  const mode = excelNativeReviewMode();
  const capability = resolveExcelFormatCapability((set, version) =>
    typeof Office !== 'undefined' && Boolean(Office.context?.requirements?.isSetSupported(set, version)));
  const formatNotice = excelFormatCapabilityNotice(plan.actions, capability);
  const reviewing = review?.outcome === 'reviewing';
  const view = review ? reviewCardView({
    outcome: review.outcome,
    scope: { locations: reviewScopeLocations(plan.actions), changes: reviewChangeSummary(plan.actions), totalCells: planTargetCells(plan.actions) },
    compatNotice: manual ? REVIEW_CARD_COMPAT_NOTICE : null,
    compatDetail: manual ? REVIEW_CARD_COMPAT_DETAIL['excel-compat'] : [],
    paneNeedsResize: Boolean(reviewing && paneNeedsResize), paneWidthSettable: true, copy: review.copy,
    syncScrollingEnabled: review.session.synchronousScrolling === 'enabled',
    syncScrollingStatus: '已启用 Excel 原生并排比较与同步滚动，拖动任一表的滚动条可联动查看。',
    diagnostics: {
      restart: review.session.windowLayout === 'service_outdated' ? ['本地窗口服务版本过旧，请重启 Python 服务并刷新加载项。'] : [],
      arrangement: review.session.windowLayout !== 'arranged'
        ? ['可在 Excel「视图 → 并排查看」中选择本次副本与原表，或手动排列窗口。']
        : review.session.synchronousScrolling === 'different_instance'
          ? ['两份工作簿属于不同 Excel 实例，未启用同步滚动。']
          : review.session.synchronousScrolling !== 'enabled' ? ['同步滚动未能确认，请在 Excel 中检查。'] : [],
      identity: ['本次副本工作表名：' + review.session.copySheetName, '原表工作簿标识：' + review.session.workbookIdentity],
      close: review.copy ? [review.copy.detail] : [],
      cleanup: review.session.recoveryCopyPath
        ? ['已保留本次副本的恢复文件；不会覆盖原文件或删除用户另存文件。']
        : ['尚无已确认的恢复文件。']
    }
  }) : null;
  async function confirm() {
    // Capture the exact card, nonce, plan and scope before the external authorization await.
    const issued = { review, plan, authorization, acknowledgedCopy };
    const scope = JSON.stringify([...issued.authorization].sort());
    const planKey = JSON.stringify(issued.plan);
    if (!issued.review || issued.review.outcome !== 'reviewing') return;
    const stillCurrent = () => {
      const now = current.current;
      return now.review === issued.review && now.plan === issued.plan && planKey === JSON.stringify(now.plan) &&
        scope === JSON.stringify([...now.authorization].sort()) && now.acknowledgedCopy === issued.acknowledgedCopy;
    };
    if (!stillCurrent()) { controller.rejectStaleRequest(messageId); return; }
    if (beforeOperation && !(await beforeOperation())) return;
    if (!stillCurrent()) {
      controller.rejectStaleRequest(messageId);
      return;
    }
    await controller.save(messageId, plan, authorization, commit,
      manual ? acknowledgedCopy : issued.review.session.copySheetName, stillCurrent);
  }
  return <section className="excel-native-review" aria-label="Excel 原生并排审阅">
    {formatNotice && <div className="excel-format-capability-notice" role="status">
      <p className="excel-format-capability-summary">{formatNotice.summary}</p>
      <details><summary>技术原因</summary><p>{formatNotice.detail}</p></details>
    </div>}
    {view && review ? <>
      <h4 className="excel-review-title">{view.title}</h4>
      <p className="excel-review-description">{view.description}</p>
      {view.scopeLine && <p className="excel-review-scope">范围：{view.scopeLine}</p>}
      {reviewing && view.compatNotice && <p className="excel-review-compat">{view.compatNotice}</p>}
      {view.paneHint && <details><summary>{view.paneHint.summary}</summary><p>{view.paneHint.detail}</p></details>}
      {reviewing && manual && <label className="excel-manual-review-confirm">
        <input type="checkbox" checked={acknowledgedCopy === review.session.copySheetName} disabled={disabled}
          onChange={event => setAcknowledgedCopy(event.target.checked ? review.session.copySheetName : null)} />
        {view.confirmLabel}
      </label>}
      {reviewing && <div className="excel-native-review-buttons">
        <button className="secondary-button" disabled={disabled} onClick={() => void controller.discard(messageId)}>放弃变更</button>
        <button className="primary-button" disabled={disabled || Boolean(manual && acknowledgedCopy !== review.session.copySheetName)}
          onClick={() => void confirm()}>确认写回原表</button>
      </div>}
      {view.copyStatus && <p className="excel-review-copy-status">{view.copyStatus.headline}</p>}
      {view.copyOpenNotice && <p>{view.copyOpenNotice}</p>}
      {review.session.recoveryCopyPath && <p className="excel-review-recovery">副本恢复文件：<code>{review.session.recoveryCopyPath}</code></p>}
      {view.secondaryAction && <button className="secondary-button" disabled={disabled}
        onClick={() => void controller.discard(messageId)}>{view.secondaryAction.label}</button>}
      {view.syncScrolling && <p>{view.syncScrolling}</p>}
      {view.compatDetail.length > 0 && <details><summary>{REVIEW_CARD_COMPAT_DETAIL_TITLE}</summary>
        {view.compatDetail.map(line => <p key={line}>{line}</p>)}</details>}
      {reviewCardHasDiagnostics(view.diagnostics) && <details><summary>{REVIEW_CARD_DIAGNOSTICS_TITLE}</summary>
        {Object.entries(view.diagnostics).map(([group, entries]) => <div key={group}>{entries.map(line => <p key={line}>{line}</p>)}</div>)}
      </details>}
    </> : <>
      <button className="primary-button" disabled={disabled || excelReviewBlocksNewCopy(controller.review) || mode === 'unavailable'}
        onClick={() => void (async () => {
          const issued = current.current;
          const planKey = JSON.stringify(plan), scopeKey = JSON.stringify([...authorization].sort());
          if (issued.plan !== plan || JSON.stringify([...issued.authorization].sort()) !== scopeKey) return;
          if (beforeOperation && !(await beforeOperation())) return;
          const now = current.current;
          if (now.review !== issued.review || now.plan !== plan || JSON.stringify(now.plan) !== planKey ||
              JSON.stringify([...now.authorization].sort()) !== scopeKey) {
            controller.rejectStaleRequest(messageId); return;
          }
          setAcknowledgedCopy(null);
          await controller.stage(messageId, plan, authorization);
        })()}>生成副本并排审阅</button>
      <p>原表保持不变，副本展示待确认结果；确认后只写回本卡计划。</p>
      {mode === 'unavailable' && <p>当前 Excel 缺少创建副本所需的 API，未生成副本、未写入原表。</p>}
    </>}
    {feedback && <p role="status">{feedback}</p>}
  </section>;
}
