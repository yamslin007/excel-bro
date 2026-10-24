import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { checkAuthorizedTarget, TargetConfirmationRequiredError } from "@excel-bro/core/authorizedTarget";
import type { PythonDataHost } from "@excel-bro/core/pythonData";
import { inspectWpsAuthorizedTarget, wpsAnalysisTargetEvidence } from "./wpsAuthorizedTarget";
import { createWpsPrimitiveOps } from "./wpsHostAdapter";
import { getWpsApplication } from "./wpsHost";
import { wpsReadWorkbook } from "./wpsResultReviewTarget";
import { authorizeWpsSheets } from "./wpsWorkbookReader";
import { captureWpsObjectIdentity, verifyWpsObjectIdentity } from "./wpsObjectIdentity";

/** Mirrors Office; a new/unknown wrapper is not a fresh implicit consent. */
export function createWpsPythonDataHost(bound: AuthorizedTarget,
  getAuthorization: () => AuthorizedTarget | null): PythonDataHost {
  bound = { ...bound, sheets: bound.sheets.map((sheet) => ({ ...sheet })), selected: [...bound.selected] };
  const selected = bound.sheets.filter((sheet) => bound.selected.includes(sheet.id));
  let identity: ReturnType<typeof captureWpsObjectIdentity> | null = null;
  const host: PythonDataHost = {
    hostId: "wps-et", identity: bound.fileKey ?? `session:${crypto.randomUUID()}`,
    authorizedSources: selected,
    async assertCurrent() {
      const application = getWpsApplication();
      // AuthorizedTarget has only sheet indexes, not the consent-time WPS
      // stable object identity. Do not adopt a replacement workbook as baseline.
      // 本次副本被激活时仍按原表：那不是换了工作簿，只是多了个窗口。
      if (!identity && wpsReadWorkbook(application) !== bound.fileRef) throw new TargetConfirmationRequiredError();
      if (checkAuthorizedTarget(bound, getAuthorization()) !== "same" ||
          checkAuthorizedTarget(bound, await inspectWpsAuthorizedTarget(bound.selected.map(Number))) !== "same" ||
          !(await wpsAnalysisTargetEvidence(bound))?.conclusive) throw new TargetConfirmationRequiredError();
      identity ??= captureWpsObjectIdentity(application, bound.selected.map(Number));
      if (verifyWpsObjectIdentity(application, identity) !== "same") throw new TargetConfirmationRequiredError();
    },
    async read(source) {
      await host.assertCurrent();
      if (!selected.some((sheet) => sheet.id === source.sheetId && sheet.name === source.sheetName)) throw new TargetConfirmationRequiredError();
      const application = getWpsApplication();
      const authorization = authorizeWpsSheets(application, bound.selected.map(Number));
      const values = await createWpsPrimitiveOps(application, authorization).readValues({ name: source.sheetName }, source.range);
      await host.assertCurrent();
      return values;
    }
  };
  return host;
}
