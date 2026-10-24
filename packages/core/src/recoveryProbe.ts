/** Explicit, one-shot acceptance fault. Never writes host data or ledger state. */
export const RECOVERY_PROBE_KEY = "excel-bro.recovery-probe.v1";
type Probe = {identity: string; expires: number; phase: "armed" | "bound" | "interrupted"; operationId?: string};
export function readRecoveryProbe(): Probe | null {
  try {
    const p = JSON.parse(localStorage.getItem(RECOVERY_PROBE_KEY) ?? "null") as Probe | null;
    return p && typeof p.identity === "string" && p.expires > Date.now() &&
      ["armed", "bound", "interrupted"].includes(p.phase) ? p : null;
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
type Action = {type: string; sheet: string; range?: string; values?: unknown};
/** Called after registration, before the executor. An unsuitable next execution consumes the arm and fails closed. */
export function bindRecoveryProbe(enabled: boolean, identity: string | null | undefined, operationId: string,
  actions: readonly Action[], recoverable: boolean) {
  const p = enabled ? readRecoveryProbe() : null;
  if (!p || p.phase !== "armed") return;
  cancelRecoveryProbe();
  const keys = actions.map(a => `${a.sheet.toLowerCase()}!${a.range?.replaceAll("$", "").toUpperCase().split(":")[0]}`);
  const singleCell = (range: string) => /^\$?[A-Z]+\$?[1-9]\d*$/i.test(range) ||
    (/^([A-Z]+[1-9]\d*):\1$/i.test(range.replaceAll("$", "")));
  if (identity !== p.identity || !recoverable || actions.length !== 3 || new Set(keys).size !== 3 ||
      actions.some(a => a.type !== "writeValues" || !a.range || !singleCell(a.range) ||
        !Array.isArray(a.values) || a.values.length !== 1 || !Array.isArray(a.values[0]) || a.values[0].length !== 1)) {
    throw new Error("恢复验收已取消：需要当前工作簿中已登记可恢复的三个独立单格值写入；未执行写入。");
  }
  save({...p, phase: "bound", operationId});
}
/** Invoke only AFTER native readback, BEFORE completion receipt. Consume BEFORE throwing. */
export function interruptRecoveryProbe(enabled: boolean, operationId: string, actionIndex: number) {
  const p = enabled ? readRecoveryProbe() : null;
  if (!p || p.phase !== "bound" || p.operationId !== operationId || actionIndex !== 1) return;
  save({...p, phase: "interrupted"});
  throw new Error("恢复验收：第二步已写入并回读，已在完成回执前中断；第三步未执行。请重开窗格并核对后续跑。");
}
