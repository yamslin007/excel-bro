import {afterEach,beforeEach,describe,expect,it,vi} from "vitest";
import capabilities from "../../../config/capabilities.json";
import type {PythonResultSession} from "@excel-bro/core/pythonResult";
import type {PythonDataHost} from "@excel-bro/core/pythonData";
import {prepareOfficePythonResult} from "./officePythonResult";

const source = {sheetId:"id1",sheetName:"Sheet1",range:"A1:B1"},target = {...source,range:"E1:F1"};
const receipt = {resultHandle:"result://"+"r".repeat(43),contentFingerprint:"c".repeat(64),rows:1,columns:2};
beforeEach(() => {capabilities.pythonExecution.enabled = true;});
afterEach(() => {capabilities.pythonExecution.enabled = false;vi.unstubAllGlobals();});
function fixture() {
  let before = [[0,0]],formulas:(string|number)[][] = [[0,0]];const write = vi.fn();
  const range = {load:vi.fn(),get values(){return before;},set values(_v){write();},get formulas(){return formulas;},set formulas(_v){write();}};
  const getItem = vi.fn(() => ({getRange:vi.fn(() => range)}));
  vi.stubGlobal("Excel",{run:vi.fn(async (work) => work({workbook:{worksheets:{getItem}},sync:async () => {}}))});
  const supported = vi.fn(() => true);vi.stubGlobal("Office",{context:{requirements:{isSetSupported:supported}}});
  const host:PythonDataHost = {hostId:"excel-office",identity:"file",authorizedSources:[{id:"id1",name:"Sheet1"}],assertCurrent:vi.fn(async () => {}),read:vi.fn(async () => [[1,2]])};
  const session:PythonResultSession = {scope:{version:1,taskId:"task",hostId:host.hostId,workbookIdentityHash:"a".repeat(64),authorizationVersion:"b".repeat(64),sources:[source]},
    assertHost:async (candidate) => {if(candidate !== host) throw Error();await host.assertCurrent();},
    getResult:vi.fn(async () => ({kind:"matrix" as const,values:[[3,4]]})),captureSources:async () => [{source,values:[[1,2]]}],revalidate:async () => {}};
  return {host,session,write,getItem,supported,change:() => {before=[[9,0]];},changeFormula:() => {formulas=[["=1",0]];}};
}
describe("Office local result preview preparation",() => {
  it("real preflight and target snapshot, no write or remote validator",async () => {
    const f=fixture(),preview=await prepareOfficePythonResult(f.session,f.host,receipt,target);
    expect(preview.plan.actions[0].values).toEqual([[3,4]]);
    expect(preview.prepared.before.values).toEqual([[0,0]]);
    expect(f.getItem).toHaveBeenCalledWith("id1");expect(f.write).not.toHaveBeenCalled();
  });
  it.each(["target","formula","file","plan","snapshot","result"])("invalidates %s before any write",async (change) => {
    const f=fixture(),preview=await prepareOfficePythonResult(f.session,f.host,receipt,target);
    if(change === "target") f.change();
    if(change === "formula") f.changeFormula();
    if(change === "file") vi.mocked(f.host.assertCurrent).mockRejectedValue(Error());
    if(change === "plan") preview.plan.actions[0].values[0][0]=99;
    if(change === "snapshot") preview.prepared.before.values[0][0]=99;
    if(change === "result") vi.mocked(f.session.getResult).mockRejectedValue(Error());
    await expect(preview.assertCurrent()).rejects.toThrow();expect(f.write).not.toHaveBeenCalled();
  });
  it("does not prepare with wrong host, unauthorized target or absent API",async () => {
    const f=fixture();
    await expect(prepareOfficePythonResult(f.session,{...f.host},receipt,target)).rejects.toThrow();
    await expect(prepareOfficePythonResult(f.session,f.host,receipt,{...target,sheetId:"other"})).rejects.toThrow();
    f.supported.mockReturnValue(false);
    await expect(prepareOfficePythonResult(f.session,f.host,receipt,target)).rejects.toThrow();
    expect(f.write).not.toHaveBeenCalled();
  });
});
