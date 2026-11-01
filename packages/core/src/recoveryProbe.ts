/** Explicit, one-shot acceptance fault. Never writes host data or ledger state. */
export const RECOVERY_PROBE_KEY = "excel-bro.recovery-probe.v1";
type Probe = {identity: string; expires: number; phase: "armed" | "bound" | "interrupted" | "not_interrupted"; operationId?: string};
export function readRecoveryProbe(): Probe | null {
  try {
    const p = JSON.parse(localStorage.getItem(RECOVERY_PROBE_KEY) ?? "null") as Probe | null;
    return p && typeof p.identity === "string" && p.expires > Date.now() &&
      ["armed", "bound", "interrupted", "not_interrupted"].includes(p.phase) ? p : null;
  } catch { return null; }
}
function save(p: Probe) {
  localStorage.setItem(RECOVERY_PROBE_KEY, JSON.stringify(p));
  if (typeof window !== "undefined") window.dispatchEvent(new Event("recovery-probe-change"));
}
export function armRecoveryProbe(enabled: boolean, identity: string) {
  if (!enabled || !identity) throw new Error("恢复验收未开放或工作簿尚未保存。");
  save({identity, expires: Date.now() + 10 * 60_000, phase: "armed"});
}
export function cancelRecoveryProbe() {
  localStorage.removeItem(RECOVERY_PROBE_KEY);
  if (typeof window !== "undefined") window.dispatchEvent(new Event("recovery-probe-change"));
}
type Action = {type: string; sheet: string; range?: string; values?: unknown; formulas?: unknown};
export class RecoveryProbeCancelledError extends Error {
  constructor(reason: string) {
    super(`恢复验收已取消：${reason}；未执行写入。`);
    this.name = 'RecoveryProbeCancelledError';
  }
}
/** Called after registration, before the executor. An unsuitable next execution consumes the arm and fails closed. */
export function bindRecoveryProbe(enabled: boolean, identity: string | null | undefined, operationId: string,
  actions: readonly Action[], recoverable: boolean, recoveryReason?: string) {
  const p = enabled ? readRecoveryProbe() : null;
  if (!p || p.phase !== "armed") return;
  cancelRecoveryProbe();
  const keys = actions.map(a => `${a.sheet.toLowerCase()}!${a.range?.replaceAll("$", "").toUpperCase().split(":")[0]}`);
  const singleCell = (range: string) => /^\$?[A-Z]+\$?[1-9]\d*$/i.test(range) ||
    (/^([A-Z]+[1-9]\d*):\1$/i.test(range.replaceAll("$", "")));
  /** 一格一次写：`writeValues` 看 `values`，`writeFormulas` 看 `formulas`，两者都必须是 1×1。 */
  const singleCellWrite = (a: Action) => {
    if (a.type !== "writeValues" && a.type !== "writeFormulas") return false;
    if (!a.range || !singleCell(a.range)) return false;
    const matrix = a.type === "writeValues" ? a.values : a.formulas;
    return Array.isArray(matrix) && matrix.length === 1 && Array.isArray(matrix[0]) && matrix[0].length === 1;
  };
  if (identity !== p.identity) throw new RecoveryProbeCancelledError('工作簿身份与启用验收时不一致');
  if (!recoverable) throw new RecoveryProbeCancelledError(recoveryReason ?? '本次计划未登记可恢复的执行');
  if (actions.length !== 3) throw new RecoveryProbeCancelledError('需要恰好三个独立写入动作');
  if (new Set(keys).size !== 3) throw new RecoveryProbeCancelledError('三个写入目标必须互不相同');
  if (actions.some(a => !singleCellWrite(a))) throw new RecoveryProbeCancelledError('每个动作必须为单格值或公式写入，且矩阵为 1×1');
  save({...p, phase: "bound", operationId});
}
/** Invoke only AFTER native readback, BEFORE completion receipt. Consume BEFORE throwing. */
export function interruptRecoveryProbe(enabled: boolean, operationId: string, actionIndex: number) {
  const p = enabled ? readRecoveryProbe() : null;
  if (!p || p.phase !== "bound" || p.operationId !== operationId || actionIndex !== 1) return;
  save({...p, phase: "interrupted"});
  throw new Error("恢复验收：第二步已写入并回读，已在完成回执前中断；第三步未执行。请重开窗格并核对后续跑。");
}
/** Release only this attempt's unused binding; never rearm or change the execution ledger. */
export function settleRecoveryProbe(enabled: boolean, operationId: string | null | undefined) {
  const p = enabled ? readRecoveryProbe() : null;
  if (!p || p.phase !== "bound" || p.operationId !== operationId) return;
  save({...p, phase: "not_interrupted"});
}
