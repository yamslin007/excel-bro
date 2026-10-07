/** Local, in-memory metadata only. Never serialize this into a model request/history. */
export type AuthorizedTarget = {
  fileKey: string | null;
  fileRef?: object;
  name: string;
  sheets: { id: string; name: string }[];
  selected: string[];
};

export type TargetCheck = "same" | "file_changed" | "scope_changed" | "unavailable";

/** An early change detector, NOT a replacement for execution/clarification identity proofs. */
export function checkAuthorizedTarget(bound: AuthorizedTarget, current: AuthorizedTarget | null): TargetCheck {
  if (!current) return "unavailable";
  if (bound.fileKey && current.fileKey) {
if (bound.fileKey !== current.fileKey) return "file_changed";
  } else if (bound.fileKey !== current.fileKey || !bound.fileRef || bound.fileRef !== current.fileRef) {
    // New host wrapper / missing identity is absence of evidence, not proof of a changed file.
    return "unavailable";
  }
  const selected = [...bound.selected].sort();
  if (JSON.stringify(selected) !== JSON.stringify([...current.selected].sort())) return "scope_changed";
  for (const id of selected) {
    const before = bound.sheets.find((sheet) => sheet.id === id);
    const after = current.sheets.find((sheet) => sheet.id === id);
    if (!before || !after || before.name !== after.name) return "scope_changed";
  }
  return "same";
}

export function targetScopeLabel(target: AuthorizedTarget): string {
  return `${target.name} · ${target.sheets.filter((sheet) => target.selected.includes(sheet.id)).map((sheet) => sheet.name).join("、") || "尚未选择工作表"}`;
}
/** A changed target pauses work for UI confirmation; it is not a failed task. */
export class TargetConfirmationRequiredError extends Error {
  constructor() { super("需要确认操作目标"); this.name = "TargetConfirmationRequiredError"; }
}
