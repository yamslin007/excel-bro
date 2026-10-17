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

/**
 * **两来源、单目标**的审阅（与 WPS 端同一套业务语义）。
 *
 * 两端必须给出同一结论：审阅副本只针对确切的目标表；两张来源表的身
 * 份/授权/源版本核对一条都不能省；任一来源变化都让预览失效。
 */
describe("Office local result preview: two sources, one target",() => {
  function twoSourceFixture() {
    const s1={sheetId:"id1",sheetName:"Sheet1",range:"A1:B1"};
    const s2={sheetId:"id2",sheetName:"Sheet2",range:"A1:B1"};
    const dest={sheetId:"id1",sheetName:"Sheet1",range:"E1:F1"};
    let before:unknown=[[0,0]],formulas:unknown=[[0,0]];
    let s1Values:unknown=[[1,2]];
    let s2Values:unknown=[[3,4]];
    const write=vi.fn();
    function rangeFor(sheetId:string) {
      if(sheetId === "id2") return {load:vi.fn(),get values(){return s2Values;},set values(_v){write();},get formulas(){return s2Values;},set formulas(_v){write();}};
      // id1 既是来源也是目标：用地址区分（源 A1:B1 / 目标 E1:F1）。
      return {load:vi.fn(),get values(){return before;},set values(_v){write();},get formulas(){return formulas;},set formulas(_v){write();}};
    }
    const getItem=vi.fn((sheetId:string) => ({
      getRange:vi.fn(() => rangeFor(sheetId))
    }));
    vi.stubGlobal("Excel",{run:vi.fn(async (work) => work({workbook:{worksheets:{getItem}},sync:async () => {}}))});
    vi.stubGlobal("Office",{context:{requirements:{isSetSupported:vi.fn(() => true)}}});
    const host:PythonDataHost={hostId:"excel-office",identity:"file",
      authorizedSources:[{id:"id1",name:"Sheet1"},{id:"id2",name:"Sheet2"}],
      assertCurrent:vi.fn(async () => {}),read:vi.fn(async () => [[1,2]])};
    const session:PythonResultSession={scope:{version:1,taskId:"task",hostId:host.hostId,
      workbookIdentityHash:"a".repeat(64),authorizationVersion:"b".repeat(64),sources:[s1,s2]},
      assertHost:async (candidate) => {if(candidate !== host) throw Error();await host.assertCurrent();},
      getResult:vi.fn(async () => ({kind:"matrix" as const,values:[[3,4]]})),
      captureSources:async () => [{source:s1,values:s1Values as number[][]},{source:s2,values:s2Values as number[][]}],
      revalidate:async () => {}};
    return {host,session,dest,write,
      changeSecondSource:() => {s2Values=[[9,4]];},
      changeFirstSource:() => {s1Values=[[9,2]];}};
  }

  it("只对确切目标表预检，两张来源快照都在",async () => {
    const f=twoSourceFixture();
    const preview=await prepareOfficePythonResult(f.session,f.host,receipt,f.dest);
    expect(preview.plan.actions).toHaveLength(1);
    expect(preview.plan.actions[0].sheet).toBe("Sheet1");
    expect(preview.prepared.sources).toHaveLength(2);
    expect(f.write).not.toHaveBeenCalled();
  });

  it("任一来源变化都让预览失效",async () => {
    const first=twoSourceFixture();
    const previewA=await prepareOfficePythonResult(first.session,first.host,receipt,first.dest);
    first.changeFirstSource();
    await expect(previewA.assertCurrent()).rejects.toThrow();

    const second=twoSourceFixture();
    const previewB=await prepareOfficePythonResult(second.session,second.host,receipt,second.dest);
    second.changeSecondSource();
    await expect(previewB.assertCurrent()).rejects.toThrow();
    expect(first.write).not.toHaveBeenCalled();
    expect(second.write).not.toHaveBeenCalled();
  });
});
