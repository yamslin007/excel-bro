import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { currentOfficeWorkbookIdentity, isEBSystemSheet } from "./excel";
import { workbookNameFromDocumentUrl } from "./workbookIdentity";
import {
  authorizedTargetEvidenceKey,
  checkAuthorizedTarget
} from "@excel-bro/core/authorizedTarget";
import { admissionItem, type ScopeAdmissionReport, type ScopeAdmissionReason } from "@excel-bro/core/scopeAdmission";

/** Document-scoped Office metadata only; no cells, used ranges or arbitrary other files. */
export async function inspectOfficeAuthorizedTarget(selected: readonly string[] | null): Promise<AuthorizedTarget> {
  const before = currentOfficeWorkbookIdentity();
  const documentRef = Office.context.document;
  const result = await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    const active = worksheets.getActiveWorksheet();
    worksheets.load("items/id,items/name");
    active.load("id,name");
    await context.sync();
    const sheets = worksheets.items.filter((sheet) => !isEBSystemSheet(sheet.name))
      .map((sheet) => ({ id: sheet.id, name: sheet.name }));
    return { sheets, selected: selected === null ? [active.id]
      : selected.map((name) => sheets.find((sheet) => sheet.name === name)?.id ?? `missing:${name}`) };
  });
  if (before !== currentOfficeWorkbookIdentity() || documentRef !== Office.context.document) {
    throw new Error("文件身份在检查期间发生变化。");
  }
  return { ...result, fileKey: before ?? null, fileRef: documentRef,
    name: workbookNameFromDocumentUrl(Office.context.document.url) };
}

/** Strong document-scoped evidence for analysis only. No cells and no write permission. */
export async function officeAnalysisTargetEvidence(target: AuthorizedTarget) {
  const identity = currentOfficeWorkbookIdentity() ?? null;
  if (!target.fileRef || target.fileRef !== Office.context.document || identity !== target.fileKey) return null;
  // 键必须只由"哪个文件 + 授权了哪些表"决定 —— 用户点选的先后顺序不是身份。
  return { key: authorizedTargetEvidenceKey(identity ?? "document-reference", target.selected), conclusive: true };
}

/** The range picker's explicit confirmation shares the same live authorization guard.
 * Metadata only: no headers, data rows, model requests or workbook writes. */
export async function confirmOfficeAuthorizedScope(
  selected: readonly string[] | null,
  confirm: (target: AuthorizedTarget) => Promise<boolean>,
  expected?: AuthorizedTarget
): Promise<boolean> {
  const target = await inspectOfficeAuthorizedTarget(selected);
  if (expected && (checkAuthorizedTarget(expected, target) !== "same" || expected.fileRef !== target.fileRef)) return false;
  return confirm(target);
}

/** Selection eligibility: only identities and supported protection/read-only metadata. */
export async function inspectOfficeScopeAdmission(
  selected: readonly string[],
  expectedFile?: Pick<AuthorizedTarget, "fileKey" | "fileRef">
): Promise<ScopeAdmissionReport> {
  const target = await inspectOfficeAuthorizedTarget(selected);
  if (expectedFile && (expectedFile.fileKey !== target.fileKey || expectedFile.fileRef !== target.fileRef)) {
    return { items: selected.map((id) => admissionItem(id, "target_changed")) };
  }
  const evidence = await officeAnalysisTargetEvidence(target);
  if (!evidence) return { items: selected.map((id) => admissionItem(id, "identity_unverified")) };
  if (selected.some((name) => !target.sheets.some((sheet) => sheet.name === name))) {
    return { items: selected.map((id) => admissionItem(id,
      target.sheets.some((sheet) => sheet.name === id) ? "write_unverified" : "missing")) };
  }
  const statuses = new Map<string, ScopeAdmissionReason>();
  const supported = (version: string) => Office.context.requirements?.isSetSupported("ExcelApi", version) === true;
  // Unavailable host properties do not prevent analysis, but never imply writable.
  await Excel.run(async (context) => {
    const book = context.workbook;
    const canReadOnly = supported("1.8");
    const canProtection = supported("1.2");
    if (canReadOnly) book.load("readOnly");
    const sheets = selected.map((name) => target.sheets.find((sheet) => sheet.name === name)).filter((s) => s !== undefined);
    const protections = new Map<string, Excel.WorksheetProtection>();
    if (canProtection) for (const sheet of sheets) {
      const protection = book.worksheets.getItem(sheet.id).protection;
      protection.load("protected"); protections.set(sheet.name, protection);
    }
    await context.sync();
    for (const name of selected) {
      statuses.set(name, !target.sheets.some((sheet) => sheet.name === name) ? "missing"
        : (canReadOnly && book.readOnly === true) || String(Office.context.document.mode) === "readOnly" ? "read_only"
        : protections.get(name)?.protected === true ? "protected"
        : canReadOnly && book.readOnly === false && protections.get(name)?.protected === false ? "none" : "write_unverified");
    }
  });
  const after = await inspectOfficeAuthorizedTarget(selected);
  if (target.fileKey === after.fileKey && target.fileRef === after.fileRef &&
      selected.some((name) => !after.sheets.some((sheet) => sheet.name === name))) {
    return { items: selected.map((id) => admissionItem(id,
      after.sheets.some((sheet) => sheet.name === id) ? "write_unverified" : "missing")) };
  }
  if (checkAuthorizedTarget(target, after) !== "same" || !(await officeAnalysisTargetEvidence(target))) {
    return { items: selected.map((id) => admissionItem(id, "target_changed")) };
  }
  return { items: selected.map((id) => admissionItem(id, statuses.get(id) ?? "identity_unverified")),
    target, evidenceKey: evidence.key, conclusive: evidence.conclusive };
}
