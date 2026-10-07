import {afterEach,beforeEach,describe,expect,it,vi} from "vitest";
import capabilities from "../../../config/capabilities.json";
import type {PythonResultSession} from "@excel-bro/core/pythonResult";
import type {PythonDataHost} from "@excel-bro/core/pythonData";
import type {WpsApplication} from "./wpsHost";
import {getWpsApplication} from "./wpsHost";
import {prepareWpsPythonResult} from "./wpsPythonResult";
vi.mock("./wpsHost",() => ({getWpsApplication:vi.fn()}));
const source={sheetId:"1",sheetName:"Sheet1",range:"A1:B1"},target={...source,range:"E1:F1"};
const receipt={resultHandle:"result://"+"r".repeat(43),contentFingerprint:"c".repeat(64),rows:1,columns:2};
beforeEach(() => {capabilities.pythonExecution.enabled=true;});
afterEach(() => {capabilities.pythonExecution.enabled=false;vi.clearAllMocks();});
function fixture() {
  let before=[[0,0]],formulas:(string|number)[][]=[[0,0]];const write=vi.fn();
  const targetRange={Rows:{Count:1},Columns:{Count:2},MergeCells:false,get Value2(){return before;},set Value2(_v){write();},get Formula(){return formulas;},set Formula(_v){write();}};
  const sourceRange={Rows:{Count:1},Columns:{Count:2},Value2:[[1,2]],Formula:[[1,2]]};
  const sheet={Name:"Sheet1",ProtectContents:false,Range:(address:string) => address===source.range?sourceRange:targetRange};
  const workbook={Name:"file.xlsx",FullName:"C:\\Data\\file.xlsx",Path:"C:\\Data",ReadOnly:false,Worksheets:{Count:1,Item:() => sheet},ActiveSheet:sheet};
  const app={Version:"12.0",ActiveWorkbook:workbook,ActiveSheet:sheet,PluginStorage:{getItem:vi.fn(),setItem:write}} as unknown as WpsApplication;
  vi.mocked(getWpsApplication).mockReturnValue(app);
  const host:PythonDataHost={hostId:"wps-et",identity:"file",authorizedSources:[{id:"1",name:"Sheet1"}],assertCurrent:vi.fn(async () => {}),read:vi.fn(async () => [[1,2]])};
  const session:PythonResultSession={scope:{version:1,taskId:"task",hostId:host.hostId,workbookIdentityHash:"a".repeat(64),authorizationVersion:"b".repeat(64),sources:[source]},
    assertHost:async (candidate) => {if(candidate!==host) throw Error();await host.assertCurrent();},
    getResult:vi.fn(async () => ({kind:"matrix" as const,values:[[3,4]]})),captureSources:async () => [{source,values:[[1,2]]}],revalidate:async () => {}};
  return {host,session,write,sheet,targetRange,sourceRange,change:() => {before=[[9,0]];},changeFormula:() => {formulas=[["=1",0]];}};
}
describe("WPS local result preview preparation",() => {
  it("uses actual prepareWpsPlan with source and target snapshots, zero writes/storage",async () => {
    const f=fixture(),preview=await prepareWpsPythonResult(f.session,f.host,receipt,target);
    expect(preview.prepared.actions[0].beforeValues).toEqual([[0,0]]);
    expect(preview.plan.actions[0].values).toEqual([[3,4]]);expect(f.write).not.toHaveBeenCalled();
  });
  it.each(["target","formula","file","plan","snapshot","result","source","protected","merged"])("invalidates %s before any write",async (change) => {
    const f=fixture(),preview=await prepareWpsPythonResult(f.session,f.host,receipt,target);
    if(change==="target") f.change();if(change==="formula") f.changeFormula();
    if(change==="file") vi.mocked(f.host.assertCurrent).mockRejectedValue(Error());
    if(change==="plan") preview.plan.actions[0].values[0][0]=99;
    if(change==="snapshot") preview.prepared.actions[0].beforeValues[0][0]=99;
    if(change==="result") vi.mocked(f.session.getResult).mockRejectedValue(Error());
    if(change==="source") f.sourceRange.Value2=[[9,2]];
    if(change==="protected") f.sheet.ProtectContents=true;
    if(change==="merged") f.targetRange.MergeCells=true;
    await expect(preview.assertCurrent()).rejects.toThrow();expect(f.write).not.toHaveBeenCalled();
  });
  it("rejects wrong host and unauthorized sheet",async () => {
    const f=fixture();await expect(prepareWpsPythonResult(f.session,{...f.host},receipt,target)).rejects.toThrow();
    await expect(prepareWpsPythonResult(f.session,f.host,receipt,{...target,sheetId:"other"})).rejects.toThrow();
    expect(f.write).not.toHaveBeenCalled();
  });
});
