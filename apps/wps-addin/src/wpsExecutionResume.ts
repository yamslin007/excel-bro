import type {HostWriteAudit} from "@excel-bro/core/hostWriteAudit";
import {reportWpsHostWrites} from "../../excel-addin/src/api";
import { sameJsonEvidence } from "@excel-bro/core/jsonEvidence";
import type { WpsExecutionState } from "../../excel-addin/src/contracts";
import type { WpsApplication } from "./wpsHost";
import { assertWpsFormulaSourceCoverage, assertWpsResumeSources, executeWpsPlan, prepareWpsResume, type WpsExecutionProgress } from "./wpsPlanExecutor";
import { claimWpsExecutionResume, completeWpsAction, finishWpsExecution, startWpsAction, suspendWpsExecution } from "./wpsExecutionLedger";

export function wpsResumeUnavailableReason(application: WpsApplication, state: WpsExecutionState | null,
  indexes: readonly number[]): string | null {
  if (!state || state.ownerActive !== false || !state.canResume || state.status === "completed")
    return "执行记录未就绪、已完成或执行权尚未释放。";
  /*
   * 终态闸（2026-10-31）：撤销/放弃之后账本游标与 status 都不变，只有 disposition 两列
   * 记了时刻。`prepareWpsResume` 的 `record.state` 校验挡得住**同会话**（撤销后宿主缓存里
   * 是 `restored`），但跨重启缓存没了、恢复包又被 `retireWpsRecoveryPackage` 删（best-effort，
   * 可能失败），水合会按游标把 state 推回 `partial_write_recoverable` —— 那时只靠材料判据
   * 会放行一次已经回退掉的续跑。所以这里先读 disposition，任一终态都拒绝续跑。
   */
  if (state.disposition?.undoneAt) return "这次执行已经撤销过，不能再续跑剩余动作。";
  if (state.disposition?.abandonedAt) return "这次执行已按「放弃变更」处置过，不能再续跑剩余动作。";
  try { prepareWpsResume(application, state, indexes); return null; }
  catch (error) { return error instanceof Error ? error.message : "本机续跑材料校验失败。"; }
}
export function canResumeWpsExecution(application: WpsApplication, state: WpsExecutionState | null,
  indexes: readonly number[]): boolean {
  return wpsResumeUnavailableReason(application, state, indexes) === null;
}

/** One claim, original action indices, no replanning or registration, no retry after uncertain receipts. */
export async function resumeWpsExecution(application: WpsApplication, initial: WpsExecutionState,
  selectedIndexes: () => readonly number[], onProgress?: (update: WpsExecutionProgress) => void) {
  if (initial.ownerActive !== false || !initial.canResume || initial.status === "completed") throw new Error("执行权尚未释放。");
  /*
   * 终态闸（2026-10-31）：与 `wpsResumeUnavailableReason` 同一判据。旧回调 / 迟到的响应
   * 可能带着一张"撤销前"的旧卡直接走到这里，绕过卡面；这里再挡一道，任何终态都不续跑。
   */
  if (initial.disposition?.undoneAt) throw new Error("这次执行已经撤销过，不能再续跑剩余动作。");
  if (initial.disposition?.abandonedAt) throw new Error("这次执行已按「放弃变更」处置过，不能再续跑剩余动作。");
  prepareWpsResume(application, initial, selectedIndexes());
  let state = await claimWpsExecutionResume(initial.operationId, initial.revision);
  if (!state.ownerToken) throw new Error("未取得执行权，未续写。");
  const hostWriteAudits: HostWriteAudit[] = [];
  try {
    if (state.operationId !== initial.operationId || state.taskId !== initial.taskId ||
        state.nextActionIndex !== initial.nextActionIndex || state.workbookIdentity !== initial.workbookIdentity ||
        !sameJsonEvidence(state.descriptor, initial.descriptor)) throw new Error("领取后登记计划发生变化。");
    const {prepared, record} = prepareWpsResume(application, state, selectedIndexes());
    /*
     * 续跑期间每一拍都复核来源。公式计划多一条**覆盖**判据：指纹只比登记时那块
     * 矩形里的值，矩形之外的一格被填上（有效区域变大）它看不见。
     */
    const checkSources = () => {
      assertWpsResumeSources(application, state, selectedIndexes());
      assertWpsFormulaSourceCoverage(application, state, selectedIndexes(), prepared.plan);
    };
    const resumed = await executeWpsPlan(application, prepared, onProgress, undefined, async (phase, checkpoint) => {
      if (!state.ownerToken) throw new Error("续跑执行权失效。");
      checkSources();
      const common = {expectedRevision: state.revision, ownerToken: state.ownerToken, actionIndex: checkpoint.actionIndex};
      const evidence = {sheetName: checkpoint.sheetName, range: checkpoint.range,
        type: checkpoint.type, beforeFingerprint: checkpoint.beforeFingerprint};
      state = phase === "before"
        ? await startWpsAction(initial.operationId, {...common, beforeSnapshot: evidence})
        : await completeWpsAction(initial.operationId, {...common, checkpoint: {...evidence, afterFingerprint: checkpoint.afterFingerprint}});
    }, initial.operationId, {resume: {record, startActionIndex: state.nextActionIndex,
      check: checkSources}}, entry => { hostWriteAudits.push(entry); });
    await finishWpsExecution(initial.operationId, {expectedRevision: state.revision, ownerToken: state.ownerToken!,
      result: {actionCount: state.totalActions}});
    return resumed;
  } catch (error) {
    try { await suspendWpsExecution(initial.operationId, {expectedRevision: state.revision, ownerToken: state.ownerToken!,
      code: "WPS_RESUME_INTERRUPTED", message: "续跑中断，须重新核对动作现场。"}); } catch { /* No retry; durable in-flight fence stays authoritative. */ }
    throw error;
  } finally {
    // Only the resumed suffix is observed; report after the durable finish/suspend boundary.
    if (hostWriteAudits.length > 0) {
      void reportWpsHostWrites(initial.operationId, {
        ownerToken: state.ownerToken ?? null, entries: [...hostWriteAudits]
      }).catch(() => { /* Observation failure cannot change execution or recovery. */ });
    }
  }
}
