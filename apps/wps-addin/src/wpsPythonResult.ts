import type { PythonDataSource, PythonResultReceipt } from "@excel-bro/core/contracts";
import type { PythonDataHost } from "@excel-bro/core/pythonData";
import { PythonResultUnavailable, stagePythonResultPreview, type PythonResultSession } from "@excel-bro/core/pythonResult";
import { getWpsApplication } from "./wpsHost";
import { prepareWpsPlan, assertWpsPreparedPlanUnchanged } from "./wpsPlanExecutor";

/** Same compiler/liveness semantics as Office; prepareWpsPlan never writes. */
export async function prepareWpsPythonResult(session: PythonResultSession, host: PythonDataHost,
  receipt: PythonResultReceipt, target: PythonDataSource) {
  await session.assertHost(host);
  if (host.hostId !== "wps-et" || session.scope.hostId !== host.hostId) throw new PythonResultUnavailable();
  const selected = [...new Set(session.scope.sources.map((source) => source.sheetId))];
  if (selected.some((id) => !/^[1-9][0-9]*$/.test(id) || !Number.isSafeInteger(Number(id)))) throw new PythonResultUnavailable();
  const preview = await stagePythonResultPreview(session,receipt,target,async (plan,sources) => {
    await session.assertHost(host);
    const prepared = prepareWpsPlan(getWpsApplication(),selected.map(Number),plan,
      sources.map(({source,values}) => ({sheetName:source.sheetName,range:source.range,values})));
    await session.assertHost(host);
    return prepared;
  });
  const assertCurrent = async () => {
    await session.assertHost(host);
    await preview.assertCurrent();
    assertWpsPreparedPlanUnchanged(preview.prepared);
    const current = prepareWpsPlan(getWpsApplication(),selected.map(Number),preview.plan,preview.prepared.sourceSnapshots);
    if (current.actions.some((action,index) => action.previewFingerprint !== preview.prepared.actions[index].previewFingerprint) ||
        current.sources.some((source,index) => source.previewFingerprint !== preview.prepared.sources[index].previewFingerprint)) throw new PythonResultUnavailable();
    await session.assertHost(host);
  };
  await assertCurrent();
  return {...preview,assertCurrent};
}
