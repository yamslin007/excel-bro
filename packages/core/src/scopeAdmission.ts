import { checkAuthorizedTarget, type AuthorizedTarget } from "./authorizedTarget";

/** Selection preflight is local, ephemeral, and is NOT an authorization or write token. */
export type ScopeAdmissionReason = "none" | "read_only" | "protected" | "write_unverified" |
  "host_write_blocked" | "identity_unverified" | "target_changed" | "missing" |
  "unreadable" | "unsupported" | "locked" | "directory_not_writable";
export type ScopeAdmissionItem = {
  id: string;
  status: "ready" | "read_only" | "unavailable";
  reason: ScopeAdmissionReason;
};
export type ScopeAdmissionReport = {
  items: ScopeAdmissionItem[];
  /** Local file-mode response carries a content/selection binding, never raw data. */
  fingerprint?: string | null;
  /** Host references must never be persisted or sent to the server/model. */
  target?: AuthorizedTarget;
  evidenceKey?: string;
  conclusive?: boolean;
};
export const SCOPE_ADMISSION_MESSAGES: Record<ScopeAdmissionReason, string> = {
  none: "可分析 · 修改仍需预览确认",
  read_only: "可分析 · 文件只读，不能写回",
  protected: "可分析 · 工作表受保护，写入受限",
  write_unverified: "可分析 · 写入状态待核实",
  host_write_blocked: "可分析 · 当前宿主未开放写入",
  identity_unverified: "暂时无法核实目标，请刷新后重试",
  target_changed: "检查期间目标发生变化，请刷新后重新选择",
  missing: "文件或工作表已不存在，请取消该项或刷新",
  unreadable: "无法读取文件，请检查访问权限、加密或文件损坏",
  unsupported: "当前通道不支持此文件或工作表",
  locked: "可分析 · 文件被占用或无写回权限，请关闭文件或检查权限",
  directory_not_writable: "可分析 · 保存目录不可写，不能写回"
};
const readOnly = new Set<ScopeAdmissionReason>(["read_only", "protected", "write_unverified",
  "host_write_blocked", "locked", "directory_not_writable"]);
export function admissionItem(id: string, reason: ScopeAdmissionReason = "none"): ScopeAdmissionItem {
  return { id, reason, status: reason === "none" ? "ready" : readOnly.has(reason) ? "read_only" : "unavailable" };
}
export function admissionComplete(ids: readonly string[], report: ScopeAdmissionReport | null): boolean {
  return ids.length > 0 && new Set(ids).size === ids.length && Boolean(report &&
    report.items.length === ids.length && new Set(report.items.map((item) => item.id)).size === ids.length &&
    ids.every((id) => report.items.some((item) => item.id === id &&
      Object.prototype.hasOwnProperty.call(SCOPE_ADMISSION_MESSAGES, item.reason) &&
      item.status === admissionItem(id, item.reason).status && item.status !== "unavailable")));
}
export function sameAdmissionTarget(before: ScopeAdmissionReport, after: ScopeAdmissionReport): boolean {
  if (before.fingerprint && before.fingerprint !== after.fingerprint) return false;
  if (before.target) {
    if (!after.target || checkAuthorizedTarget(before.target, after.target) !== "same") return false;
    if (before.conclusive && before.target.fileRef !== after.target.fileRef) return false;
    if (before.evidenceKey !== after.evidenceKey) return false;
  }
  return true;
}
