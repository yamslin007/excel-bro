/** Only the trusted plan_range_edit response may attest a literal write.
 * The in-memory seal is bound to the exact actions; JSON/history cannot preserve it.
 * Authorization, target checks, undo preparation and readback remain mandatory.
 */
const literalSeal = Symbol("literal-range-edit");
function signature(plan: unknown): string | null {
  if (!plan || typeof plan !== "object") return null;
  const p = plan as { id?: unknown; actions?: unknown; localResultRef?: unknown };
  if (p.localResultRef || !Array.isArray(p.actions) || p.actions.length !== 1) return null;
  const action = p.actions[0];
  if (!action || action.type !== "writeValues" || !Array.isArray(action.values) || !action.values.length) return null;
  const first = action.values[0]?.[0];
  if (first !== null && !["string", "number", "boolean"].includes(typeof first)) return null;
  if (typeof first === "number" && !Number.isFinite(first)) return null;
  if (!action.values.every((row: unknown) => Array.isArray(row) && row.length > 0 && row.every(value => value === first))) return null;
  return JSON.stringify({ id: p.id, actions: p.actions });
}
export function stampLiteralPlanOrigin<T extends object>(plan: T, origin: unknown): T {
  const seal = origin === "literal_range_edit" ? signature(plan) : null;
  return { ...plan, [literalSeal]: seal };
}
export function isTrustedLiteralPlan(plan: unknown): boolean {
  if (!plan || typeof plan !== "object") return false;
  const seal = (plan as { [literalSeal]?: unknown })[literalSeal];
  return typeof seal === "string" && seal === signature(plan);
}
