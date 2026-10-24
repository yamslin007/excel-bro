/** Local, in-memory metadata only. Never serialize this into a model request/history. */
export type AuthorizedTarget = {
  fileKey: string | null;
  fileRef?: object;
  name: string;
  sheets: { id: string; name: string }[];
  selected: string[];
};

export type TargetCheck = "same" | "file_changed" | "scope_changed" | "unavailable";

/**
 * 授权集合的**规范形**：只消除顺序。
 *
 * 用户勾表的先后不是授权的一部分：「Sheet11 → Sheet1 → Sheet10」与
 * 「Sheet1 → Sheet10 → Sheet11」是同一份授权。凡是"这是不是同一组表"的判断，
 * 都必须先过它 —— 少了它，同一份授权会产出两个不同的证据键，
 * 分析前置检查把两个键读成"证据冲突"，请求被卡在「暂时无法核实当前目标」。
 *
 * 它**只做排序**：
 * - 返回**新数组**，绝不原地 `sort` 调用方的数组（调用方可能正按点选顺序映射名字）；
 * - **不去重**：重复 ID 与缺失表是真实差异，"排序无关"不许把它们吞掉；
 * - 不合并、不改写成员：增删一张表照旧是 `scope_changed`。
 */
export function canonicalTargetSelection(selected: readonly string[]): string[] {
  return [...selected].sort();
}

/**
 * 证据键：`身份` + 授权集合的规范形。
 *
 * 两端（Office / WPS）都从这里取键。各自拼一份"看起来一样"的 `JSON.stringify`
 * 是最容易失手的写法：一端忘了规范化，同一份授权就有两个键。
 */
export function authorizedTargetEvidenceKey(identity: string, members: readonly string[]): string {
  return JSON.stringify([identity, canonicalTargetSelection(members)]);
}

/** 已授权表的**名字**，按 `selected` 的顺序（用户点选顺序），不是工作表目录顺序。 */
export function selectedTargetSheetNames(target: AuthorizedTarget): string[] {
  return target.selected
    .map((id) => target.sheets.find((sheet) => sheet.id === id)?.name)
    .filter((name): name is string => Boolean(name));
}
/** Local UI state only: needing consent is not a failed identity check. */
export type TargetNoticeReason = Exclude<TargetCheck, "same"> | "confirmation_required";

/** An early change detector, NOT a replacement for execution/clarification identity proofs. */
export function checkAuthorizedTarget(bound: AuthorizedTarget, current: AuthorizedTarget | null): TargetCheck {
  if (!current) return "unavailable";
  if (bound.fileKey && current.fileKey) {
    if (bound.fileKey !== current.fileKey) return "file_changed";
  } else if (bound.fileKey !== current.fileKey || !bound.fileRef || bound.fileRef !== current.fileRef) {
    // New host wrapper / missing identity is absence of evidence, not proof of a changed file.
    return "unavailable";
  }
  const selected = canonicalTargetSelection(bound.selected);
  if (JSON.stringify(selected) !== JSON.stringify(canonicalTargetSelection(current.selected))) return "scope_changed";
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
