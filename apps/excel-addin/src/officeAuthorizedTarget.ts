import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { currentOfficeWorkbookIdentity, isEBSystemSheet } from "./excel";
import { workbookNameFromDocumentUrl } from "./workbookIdentity";

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
  return { key: JSON.stringify([identity ?? "document-reference", target.selected]), conclusive: true };
}
