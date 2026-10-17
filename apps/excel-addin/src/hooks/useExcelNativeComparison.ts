import { useRef, useState } from "react";
import type { AnalysisPlan } from "../contracts";
import { assertExcelNativeComparisonUnchanged, closeExcelNativeComparison,
  stageExcelNativeComparison, type ExcelNativeComparison } from "../excelNativeComparison";
import { rearrangeLocalExcelWindows } from "../excelWindowLayout";
import { reviewCopyCleanupState, type ReviewCopyCleanupState, type ReviewCopyCloseOutcome } from "../../../../packages/core/src/reviewCopyLifecycle";
import type { ReviewOutcome } from "../../../../packages/core/src/reviewCardPresentation";

export type ExcelReview = {
  messageId: string;
  session: ExcelNativeComparison;
  outcome: ReviewOutcome;
  copy: ReviewCopyCleanupState['copy'] | null;
  note: string | null;
};
export function excelReviewBlocksNewCopy(review: ExcelReview | null): boolean {
  return Boolean(review && ['reviewing', 'write_failed', 'partial_or_unknown'].includes(review.outcome));
}
type Feedback = { messageId: string; text: string } | null;

export function useExcelNativeComparison() {
  // Host snapshots are deliberately transient: never put them in chat history.
  const live = useRef<ExcelReview | null>(null);
  const running = useRef(false);
  const [review, setReview] = useState<ExcelReview | null>(null);
  const [pending, setPending] = useState(false);
  const [feedback, setFeedback] = useState<Feedback>(null);
  const publish = (value: ExcelReview | null) => { live.current = value; setReview(value); };
  async function operation(messageId: string, job: () => Promise<void>): Promise<boolean> {
    if (running.current) return false;
    running.current = true; setPending(true); setFeedback(null);
    try { await job(); return true; }
    catch (reason) {
      setFeedback({ messageId, text: reason instanceof Error ? reason.message : "审阅操作未完成。" });
      return false;
    } finally { running.current = false; setPending(false); }
  }
  async function close(current: ExcelReview) {
    if (current.outcome === 'reviewing') return;
    let outcome: ReviewCopyCloseOutcome = 'failed';
    try { outcome = await closeExcelNativeComparison(current.session); } catch { /* Keep business outcome. */ }
    const copy = reviewCopyCleanupState({ business: current.outcome, copy: outcome }).copy;
    if (live.current === current) publish({ ...current, copy });
  }
  async function stage(messageId: string, plan: AnalysisPlan, authorization: readonly string[]) {
    return operation(messageId, async () => {
      if (excelReviewBlocksNewCopy(live.current)) throw new Error("请先结束上一份审阅，再生成新的副本。");
      // Freeze input identities before yielding: plan edits cannot inherit the trial.
      const planKey = JSON.stringify(plan), scopeKey = JSON.stringify([...authorization].sort());
      const session = await stageExcelNativeComparison(plan, authorization);
      if (planKey !== JSON.stringify(plan) || scopeKey !== JSON.stringify([...authorization].sort())) {
        publish({ messageId, session, outcome: 'write_failed', copy: null, note: '生成期间计划或授权已变化，请保留副本核对并重新生成。' });
        throw new Error('生成期间计划或授权已变化，未开放写回。');
      }
      publish({ messageId, session, outcome: 'reviewing', copy: null, note: null });
    });
  }
  async function discard(messageId: string) {
    return operation(messageId, async () => {
      const current = live.current;
      if (!current || current.messageId !== messageId) throw new Error("这张卡片没有活跃审阅副本。");
      if (current.outcome === 'write_failed' || current.outcome === 'partial_or_unknown') {
        publish(null);
        setFeedback({ messageId, text: '已保留审阅副本供你核对。请按执行记录检查原表；结束审阅不撤销既有写入。' });
        return;
      }
      // Revoke confirmation synchronously, before any asynchronous cleanup.
      const ended: ExcelReview = { ...current, outcome: current.outcome === 'verified' ? 'verified' : 'discarded' };
      publish(ended);
      await close(ended);
    });
  }
  async function save(messageId: string, plan: AnalysisPlan, authorization: readonly string[],
    commit: (session: ExcelNativeComparison) => Promise<boolean>, acknowledgedCopy: string | null = null,
    isCurrent: () => boolean = () => true) {
    return operation(messageId, async () => {
      const current = live.current;
      if (!current || current.messageId !== messageId || current.outcome !== 'reviewing') throw new Error("没有可保存的已审阅副本。");
      if (!isCurrent()) throw new Error('确认现场已变化，原表未写入。');
      if (acknowledgedCopy !== current.session.copySheetName) {
        throw new Error(current.session.mode === 'manual'
          ? '请先在 Excel 中审阅兼容副本，并确认本次副本；原表未写入。'
          : '确认不属于当前审阅副本，原表未写入。');
      }
      const planKey = JSON.stringify(plan), scopeKey = JSON.stringify([...authorization].sort());
      let commitStarted = false;
      try {
        await Excel.run(context => assertExcelNativeComparisonUnchanged(context, current.session, authorization, plan));
        if (!isCurrent() || live.current !== current || planKey !== JSON.stringify(plan) || scopeKey !== JSON.stringify([...authorization].sort())) {
          throw new Error('计划、授权或审阅现场已变化，请重新核对；原表未写入。');
        }
        commitStarted = true;
        if (!(await commit(current.session))) throw new Error('执行结果未确认，请检查执行记录与恢复点，不能重复提交。');
      } catch (reason) {
        const outcome = commitStarted ? 'partial_or_unknown' : 'write_failed';
        if (live.current === current) publish({ ...current, outcome,
          copy: reviewCopyCleanupState({ business: outcome, copy: 'not_attempted' }).copy,
          note: reason instanceof Error ? reason.message : '执行结果未确认。' });
        throw reason;
      }
      const verified: ExcelReview = { ...current, outcome: 'verified' };
      publish(verified);
      await close(verified);
    });
  }
  // Internal layout operation; there is deliberately no "重新排列" UI button.
  async function rearrange(messageId: string): Promise<boolean> {
    return operation(messageId, async () => {
      const current = live.current;
      if (!current || current.messageId !== messageId || current.outcome !== 'reviewing') throw new Error('没有活跃审阅副本可排列。');
      const layout = await rearrangeLocalExcelWindows(current.session.workbookIdentity, current.session.original.name, current.session.copySheetName);
      if (live.current !== current) throw new Error('审阅现场已变化，未更新排列状态。');
      publish({ ...current, session: { ...current.session, windowLayout: 'arranged', synchronousScrolling: layout.synchronousScrolling } });
    });
  }
  const rejectStaleRequest = (messageId: string) => setFeedback({ messageId, text: '没有可保存的已审阅副本：计划、授权或确认现场已变化，请重新核对。' });
  return { review, pending, feedback, stage, save, discard, rearrange, rejectStaleRequest };
}
