import { SCOPE_ADMISSION_MESSAGES, type ScopeAdmissionItem } from "@excel-bro/core/scopeAdmission";

/*
 * 这里**故意没有**逐行徽标（原 `ScopeAdmissionBadge`）。
 *
 * 列表行只承载「勾选框 + 工作表名」两件事；逐行的授权文案（「可分析 · 修改仍需
 * 预览确认」等）会把这行撑成两行、行高跟着跳，视觉上还压过表名。授权状态是
 * **整次选择**的属性，所以只在 `ScopeAdmissionSummary` 里汇总一次，
 * 由它按 reason 展开明细。
 */
export function ScopeAdmissionSummary({ items, count, checking }: {
  items: readonly ScopeAdmissionItem[]; count: number; checking: boolean;
}) {
  const unavailable = items.filter((item) => item.status === "unavailable");
  const readOnly = items.filter((item) => item.status === "read_only" && item.reason !== "write_unverified").length;
  const pending = checking || items.length !== count;
  return <div className="scope-admission-summary" role="status">
    <span>{!count ? "请先选择工作表" : pending ? `已选 ${count} 张表 · 正在检查…`
      : unavailable.length ? `已选 ${count} 张表 · ${unavailable.length} 张不可用`
      : `已选 ${count} 张表 · 可分析${readOnly ? ` · ${readOnly} 张写入受限` : ""}`}</span>
    {!pending && [...new Set(items.filter((item) => item.reason !== "none").map((item) => item.reason))].map((reason) =>
      <span key={reason} className={unavailable.some((item) => item.reason === reason) ? "scope-admission-error" : "scope-admission-detail"}>
        {SCOPE_ADMISSION_MESSAGES[reason]}
      </span>)}
  </div>;
}
