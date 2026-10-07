import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { checkAuthorizedTarget, TargetConfirmationRequiredError } from "@excel-bro/core/authorizedTarget";
import type { PythonDataHost } from "@excel-bro/core/pythonData";
import { inspectOfficeAuthorizedTarget, officeAnalysisTargetEvidence } from "./officeAuthorizedTarget";
import { createOfficePrimitiveOps } from "./officeHostAdapter";

/** No UI activation: caller supplies the currently committed authorization. */
export function createOfficePythonDataHost(bound: AuthorizedTarget,
  getAuthorization: () => AuthorizedTarget | null): PythonDataHost {
  bound = { ...bound, sheets: bound.sheets.map((sheet) => ({ ...sheet })), selected: [...bound.selected] };
  const selected = bound.sheets.filter((sheet) => bound.selected.includes(sheet.id));
  const host: PythonDataHost = {
    hostId: "excel-office", identity: bound.fileKey ?? `session:${crypto.randomUUID()}`,
    authorizedSources: selected,
    async assertCurrent() {
      if (checkAuthorizedTarget(bound, getAuthorization()) !== "same" ||
          checkAuthorizedTarget(bound, await inspectOfficeAuthorizedTarget(selected.map((sheet) => sheet.name))) !== "same" ||
          !(await officeAnalysisTargetEvidence(bound))?.conclusive) throw new TargetConfirmationRequiredError();
    },
    async read(source) {
      await host.assertCurrent();
      if (!selected.some((sheet) => sheet.id === source.sheetId && sheet.name === source.sheetName)) throw new TargetConfirmationRequiredError();
      // WorksheetCollection.getItem accepts an ID. Pin the already-authorized
      // object, not a name that could resolve to a newly recreated worksheet.
      const values = await Excel.run((context) => createOfficePrimitiveOps(context).readValues({ name: source.sheetId }, source.range));
      await host.assertCurrent();
      return values;
    }
  };
  return host;
}
