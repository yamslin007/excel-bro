import { sameJsonEvidence } from "@excel-bro/core/jsonEvidence";
import type { WpsExecutionState } from "../../excel-addin/src/contracts";
import type { WpsApplication } from "./wpsHost";
import { assertWpsResumeSources, executeWpsPlan, prepareWpsResume, type WpsExecutionProgress } from "./wpsPlanExecutor";
import { claimWpsExecutionResume, completeWpsAction, finishWpsExecution, startWpsAction, suspendWpsExecution } from "./wpsExecutionLedger";

export function wpsResumeUnavailableReason(application: WpsApplication, state: WpsExecutionState | null,
  indexes: readonly number[]): string | null {
  if (!state || state.ownerActive !== false || !state.canResume || state.status === "completed")
    return "执行记录未就绪、已完成或执行权尚未释放。";
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
  prepareWpsResume(application, initial, selectedIndexes());
  let state = await claimWpsExecutionResume(initial.operationId, initial.revision);
  if (!state.ownerToken) throw new Error("未取得执行权，未续写。");
  try {
    if (state.operationId !== initial.operationId || state.taskId !== initial.taskId ||
        state.nextActionIndex !== initial.nextActionIndex || state.workbookIdentity !== initial.workbookIdentity ||
        !sameJsonEvidence(state.descriptor, initial.descriptor)) throw new Error("领取后登记计划发生变化。");
    const {prepared, record} = prepareWpsResume(application, state, selectedIndexes());
    const resumed = await executeWpsPlan(application, prepared, onProgress, undefined, async (phase, checkpoint) => {
      if (!state.ownerToken) throw new Error("续跑执行权失效。");
      assertWpsResumeSources(application, state, selectedIndexes());
      const common = {expectedRevision: state.revision, ownerToken: state.ownerToken, actionIndex: checkpoint.actionIndex};
      const evidence = {sheetName: checkpoint.sheetName, range: checkpoint.range,
        type: checkpoint.type, beforeFingerprint: checkpoint.beforeFingerprint};
      state = phase === "before"
        ? await startWpsAction(initial.operationId, {...common, beforeSnapshot: evidence})
        : await completeWpsAction(initial.operationId, {...common, checkpoint: {...evidence, afterFingerprint: checkpoint.afterFingerprint}});
    }, initial.operationId, {resume: {record, startActionIndex: state.nextActionIndex,
      check: () => assertWpsResumeSources(application, state, selectedIndexes())}});
    await finishWpsExecution(initial.operationId, {expectedRevision: state.revision, ownerToken: state.ownerToken!,
      result: {actionCount: state.totalActions}});
    return resumed;
  } catch (error) {
    try { await suspendWpsExecution(initial.operationId, {expectedRevision: state.revision, ownerToken: state.ownerToken!,
      code: "WPS_RESUME_INTERRUPTED", message: "续跑中断，须重新核对动作现场。"}); } catch { /* No retry; durable in-flight fence stays authoritative. */ }
    throw error;
  }
}
