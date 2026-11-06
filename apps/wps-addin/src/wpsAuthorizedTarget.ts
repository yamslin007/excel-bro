import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { getWpsApplication } from "./wpsHost";
import { isWpsReviewCopyWorkbook, wpsReadWorkbook } from "./wpsResultReviewTarget";
import type { WpsWorkbook } from "./wpsHost";
import { listWpsWorksheets, workbookFullName } from "./wpsWorkbookReader";
import { captureWpsObjectIdentity, verifyWpsObjectIdentity } from "./wpsObjectIdentity";
import { wpsConfirmationTargetKey } from "./wpsClarificationTargetConfirmation";
import { checkAuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { admissionItem, type ScopeAdmissionReport, type ScopeAdmissionReason } from "@excel-bro/core/scopeAdmission";
import { wpsWriteVersionStatus } from "./wpsWriteVersionGate";
import type { TargetStableIdGrade } from "@excel-bro/core/targetDiagnostics";
import capabilities from "../../../config/capabilities.json";

/**
 * 这是不是**本次审阅**的副本（只用于识别角色，绝不构成授权）。
 *
 * 判据收敛到 `wpsResultReviewTarget` 的绑定：创建时登记的完整身份 + 未保存 +
 * 单表 + 与原表并存。旧实现挂在 `stagedReview` 上，新流程里它是 null，
 * 于是副本一获得焦点就被判成"换了工作簿"。
 */
export function isWpsReviewCopyTarget(target: AuthorizedTarget): boolean {
  return isWpsReviewCopyWorkbook(target.fileRef as WpsWorkbook | undefined);
}

/** Change detection only; existing stronger clarification/write guards remain authoritative. */
export async function inspectWpsAuthorizedTarget(selected: readonly number[] | null): Promise<AuthorizedTarget> {
  const application = getWpsApplication();
  /*
   * 当前活动簿是**本次审阅副本**时按原表核对：副本不是新授权目标。
   * 无关工作簿仍按活动簿核对，因此"真的换了文件"依旧被拦住。
   */
  const workbook = wpsReadWorkbook(application);
  if (!workbook) throw new Error("无法读取当前文件身份。");
  const fullName = workbookFullName(workbook);
  // Some hosts report an unsaved workbook name as FullName; that is not a path identity.
  const fileKey = fullName && /^(?:[a-z]:[\\/]|\\\\)/i.test(fullName) ? fullName : null;
  const metadata = listWpsWorksheets(application);
  const current = wpsReadWorkbook(application);
  if (!current || fullName !== workbookFullName(current) || workbook.Name !== current.Name) {
    throw new Error("文件身份在检查期间发生变化。");
  }
  return { fileKey, fileRef: workbook, name: metadata.name,
    sheets: metadata.sheets.map((sheet) => ({ id: String(sheet.index), name: sheet.name })),
    selected: (selected ?? metadata.sheets.filter((sheet) => sheet.isActive).map((sheet) => sheet.index)).map(String) };
}

/**
 * Detect unknown wrappers BEFORE starting a task; consent only permits fresh reads/planning.
 *
 * 证据分两级：
 *
 * - `conclusive`（对象级）：同一瞬间两次读拿到**同一个宿主对象** ⇒ 会话内能证明
 *   "还是那个对象"。它只在宿主包装对象稳定时成立，跨刷新必然失效。
 * - `documentBound`（文档级）：**路径身份 + 每张授权表都有宿主声明的稳定 ID**
 *   （`CodeName` / `SId`）。宿主换包装对象不影响它，所以它**在宿主真的报出稳定 ID 时**
 *   跨刷新仍成立。**注意这不是"与 Excel 同强度"**：Excel 的文档身份无条件成立，
 *   这里是有条件的 —— 宿主不给稳定 ID、或工作簿未保存，`documentBound` 就是 false，
 *   照旧要用户确认一次（2026-11-01 真机复测即落在这一支）。
 *
 * 它**也不做刷新前的基线比较**：`documentBound` 只回答"本次读数是否具备文档级要素"，
 * 不回答"稳定 ID 是否与刷新前那份相同"。因此它**不**能识别"删掉已授权表再新建同名表"——
 * 那种情况只有 `conclusive`（对象引用）看得见。不要宣称它实现了"同名表替换降级确认"。
 *
 * 它**不**冒充对象级：`conclusive` 的含义一字未改，写入/恢复/数据读取仍只看它；
 * 只有"要不要再做一次范围确认"这一处允许文档级证据直接放行（见
 * `useAuthorizedTarget.checkForAnalysis`）。
 */
export async function wpsAnalysisTargetEvidence(target: AuthorizedTarget): Promise<
  {
    key: string; conclusive: boolean; documentBound?: boolean;
    /** 只用于诊断：稳定 ID 读到几成 / 本次是否拿到路径身份。不参与判定。 */
    diagnostic?: { stableId: TargetStableIdGrade; pathIdentity: boolean };
  } | null
> {
  const identity = captureWpsObjectIdentity(getWpsApplication(), target.selected.map(Number));
  if (identity.workbookName !== target.name ||
      (target.fileKey ? identity.workbookFullName !== target.fileKey : identity.workbookRef !== target.fileRef) ||
      identity.sheets.some((sheet) => !target.sheets.some((allowed) =>
        allowed.id === String(sheet.index) && allowed.name === sheet.name))) return null;
  const conclusive = verifyWpsObjectIdentity(getWpsApplication(), identity) === "same";
  const key = wpsConfirmationTargetKey(identity);
  const readable = identity.sheets.filter((sheet) => typeof sheet.id === "string" && sheet.id.length > 0).length;
  const stableId: TargetStableIdGrade = readable === 0 || identity.sheets.length === 0 ? "none"
    : readable === identity.sheets.length ? "all" : "partial";
  // 文档级必须有路径身份（key 非空）且每张授权表都能报出稳定 ID；缺一即不成立。
  const documentBound = Boolean(key) && identity.sheets.length > 0 && stableId === "all";
  const diagnostic = { stableId, pathIdentity: Boolean(key) };
  if (key) return { key, conclusive, documentBound, diagnostic };
  return conclusive ? { key: "session-object-reference", conclusive, documentBound: false, diagnostic } : null;
}

/** No cells, data rows, model calls or consent mutations. Only the selected sheets are inspected. */
export async function inspectWpsScopeAdmission(selected: readonly number[], expected?: AuthorizedTarget): Promise<ScopeAdmissionReport> {
  const target = await inspectWpsAuthorizedTarget(selected);
  const ids = selected.map(String);
  if (expected && checkAuthorizedTarget(expected, target) !== "same") {
    return { items: ids.map((id) => admissionItem(id, target.sheets.some((s) => s.id === id) ? "target_changed" : "missing")) };
  }
  if (ids.some((id) => !target.sheets.some((sheet) => sheet.id === id))) {
    return { items: ids.map((id) => admissionItem(id, target.sheets.some((s) => s.id === id) ? "write_unverified" : "missing")) };
  }
  const evidence = await wpsAnalysisTargetEvidence(target);
  if (!evidence) return { items: ids.map((id) => admissionItem(id, "identity_unverified")) };
  const app = getWpsApplication();
  const book = wpsReadWorkbook(app);
  const items = ids.map((id) => {
    let reason: ScopeAdmissionReason = "write_unverified";
    try {
      const sheet = book!.Worksheets.Item(Number(id));
      if (book!.ReadOnly === true) reason = "read_only";
      else if (sheet.ProtectContents === true) reason = "protected";
      else if (book!.ReadOnly === false && sheet.ProtectContents === false) {
        reason = capabilities.wpsWrite.enabled && wpsWriteVersionStatus(app.Version) !== "blocked" ? "none" : "host_write_blocked";
      }
    } catch { /* Unknown write capability is not a false writable promise. */ }
    return admissionItem(id, reason);
  });
  const after = await inspectWpsAuthorizedTarget(selected);
  const repeated = await wpsAnalysisTargetEvidence(after);
  if (checkAuthorizedTarget(target, after) !== "same" || !repeated || evidence.key !== repeated.key ||
      (evidence.conclusive && target.fileRef !== after.fileRef)) {
    return { items: ids.map((id) => admissionItem(id, "target_changed")) };
  }
  return { items, target, evidenceKey: evidence.key, conclusive: evidence.conclusive };
}
