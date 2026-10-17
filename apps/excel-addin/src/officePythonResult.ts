import type { CellValue, PythonDataSource, PythonResultReceipt } from "@excel-bro/core/contracts";
import type { PythonDataHost } from "@excel-bro/core/pythonData";
import { PythonResultUnavailable, stagePythonResultPreview, type PythonResultSession } from "@excel-bro/core/pythonResult";
import { preflightPlanActions } from "./excel";

/** In-memory preview preparation only: no native copy, ledger or write route. */
export async function prepareOfficePythonResult(session: PythonResultSession, host: PythonDataHost,
  receipt: PythonResultReceipt, target: PythonDataSource) {
  await session.assertHost(host);
  if (host.hostId !== "excel-office" || session.scope.hostId !== host.hostId) throw new PythonResultUnavailable();
  const destination = {...target};
  const readTarget = async () => {
    await session.assertHost(host);
    if (!host.authorizedSources.some((sheet) => sheet.id === destination.sheetId && sheet.name === destination.sheetName)) throw new PythonResultUnavailable();
    const snapshot = await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getItem(destination.sheetId).getRange(destination.range);
      range.load("values,formulas");
      await context.sync();
      const values = range.values as CellValue[][], formulas = range.formulas as CellValue[][];
      if (![values, formulas].every((matrix) => Array.isArray(matrix) && matrix.length === receipt.rows &&
          matrix.every((row) => Array.isArray(row) && row.length === receipt.columns))) throw new PythonResultUnavailable();
      return structuredClone({values,formulas});
    });
    await session.assertHost(host);
    return snapshot;
  };
  const preview = await stagePythonResultPreview(session, receipt, destination, async (plan, sources) => {
    const names = host.authorizedSources.map((sheet) => sheet.name);
    const issues = preflightPlanActions(plan, names,
      (version) => Office.context.requirements.isSetSupported("ExcelApi",version), {}, names);
    if (issues.length) throw new PythonResultUnavailable();
    return { sources, before: await readTarget() };
  });
  const original = JSON.stringify(preview.prepared);
  const assertCurrent = async () => {
    await preview.assertCurrent();
    if (JSON.stringify(preview.prepared) !== original ||
        JSON.stringify(await readTarget()) !== JSON.stringify(preview.prepared.before)) throw new PythonResultUnavailable();
    /*
     * 来源核对必须在**写前那一刻**重跑，与 WPS 端同语义。
     *
     * `stagePythonResultPreview` 只在生成预览时核对一次来源快照；预览到用户确认
     * 之间，来源表可能已经被改。只复读目标是不够的 —— 那会让"来源变了、结果却是
     * 基于旧源算的"这种计划照样通过写前复核。这里按 `preview.prepared.sources`
     * 的**声明快照**重读一遍：快照与现场不一致即拒绝（fail-closed）。
     */
    const current = await session.captureSources();
    const expected = preview.prepared.sources;
    if (current.length !== expected.length ||
        current.some((snapshot, index) =>
          snapshot.source.sheetId !== expected[index].source.sheetId ||
          snapshot.source.sheetName !== expected[index].source.sheetName ||
          snapshot.source.range !== expected[index].source.range ||
          JSON.stringify(snapshot.values) !== JSON.stringify(expected[index].values)
        )) throw new PythonResultUnavailable();
  };
  await assertCurrent();
  return {...preview,assertCurrent};
}
