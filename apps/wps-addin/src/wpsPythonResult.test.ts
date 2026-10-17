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

/**
 * **两来源、单目标**的审阅（与 Excel 端同一套业务语义）。
 *
 * 这一组专门证明"WPS 不是把单表限制整条删掉"：
 * - 审阅副本只针对**确切的目标表**生成（`actions[].sheet` 就是目标）；
 * - 但**两张来源表**的身份/授权/源版本核对一条都不能少（`prepared.sources` 两张、
 *   任一来源变化都必须让预览失效）；
 * - 未授权表既不进授权集合，也不会被读。
 */
describe("WPS local result preview: two sources, one target",() => {
  /** 两张来源（Sheet1、Sheet2），目标落在 Sheet1 上。 */
  function twoSourceFixture() {
    const s1={sheetId:"1",sheetName:"Sheet1",range:"A1:B1"};
    const s2={sheetId:"2",sheetName:"Sheet2",range:"A1:B1"};
    const dest={sheetId:"1",sheetName:"Sheet1",range:"D1:E1"};
    let s1Values:(string|number)[][]=[[1,2]];
    let s2Values:(string|number)[][]=[[3,4]];
    let before:(string|number)[][]=[[0,0]];
    let s2Before:(string|number)[][]=[[7,8]];
    const write=vi.fn();
    const targetRange={Rows:{Count:1},Columns:{Count:2},MergeCells:false,get Value2(){return before;},set Value2(_v){write();},get Formula(){return before;},set Formula(_v){write();}};
    const s1Range={Rows:{Count:1},Columns:{Count:2},get Value2(){return s1Values;},get Formula(){return s1Values;}};
    const s2Range={Rows:{Count:1},Columns:{Count:2},get Value2(){return s2Values;},get Formula(){return s2Values;}};
    const sheet1={Name:"Sheet1",ProtectContents:false,Range:(address:string) => address==="A1:B1"?s1Range:targetRange};
    const sheet2={Name:"Sheet2",ProtectContents:false,Range:(_address:string) => s2Range};
    const worksheets=[sheet1,sheet2];
    const workbook={Name:"file.xlsx",FullName:"C:\\Data\\file.xlsx",Path:"C:\\Data",ReadOnly:false,
      Worksheets:{Count:2,Item:(index:number) => worksheets[index-1]},ActiveSheet:sheet1};
    const app={Version:"12.0",ActiveWorkbook:workbook,ActiveSheet:sheet1,PluginStorage:{getItem:vi.fn(),setItem:write}} as unknown as WpsApplication;
    vi.mocked(getWpsApplication).mockReturnValue(app);
    const host:PythonDataHost={hostId:"wps-et",identity:"file",
      authorizedSources:[{id:"1",name:"Sheet1"},{id:"2",name:"Sheet2"}],
      assertCurrent:vi.fn(async () => {}),read:vi.fn(async () => [[1,2]])};
    const session:PythonResultSession={scope:{version:1,taskId:"task",hostId:host.hostId,
      workbookIdentityHash:"a".repeat(64),authorizationVersion:"b".repeat(64),sources:[s1,s2]},
      assertHost:async (candidate) => {if(candidate!==host) throw Error();await host.assertCurrent();},
      getResult:vi.fn(async () => ({kind:"matrix" as const,values:[[3,4]]})),
      captureSources:async () => [{source:s1,values:s1Values as number[][]},{source:s2,values:s2Values as number[][]}],
      revalidate:async () => {}};
    return {host,session,dest,receipt,write,
      changeFirstSource:() => {s1Values=[[9,2]];},
      changeSecondSource:() => {s2Values=[[9,4]];},
      restoreSecondSource:() => {s2Values=[[7,8]];}};
  }

  it("只对确切目标表生成一个写入动作，来源仍是两张",async () => {
    const f=twoSourceFixture();
    const preview=await prepareWpsPythonResult(f.session,f.host,f.receipt,f.dest);
    // 审阅副本只有一个动作，且落在确切的目标表上。
    expect(preview.plan.actions).toHaveLength(1);
    expect(preview.plan.actions[0].sheet).toBe("Sheet1");
    // 两张来源快照都在：来源核对没有因为"只写一张表"而被省掉。
    expect(preview.prepared.sources).toHaveLength(2);
    expect(f.write).not.toHaveBeenCalled();
  });

  it("第二张来源变化同样让预览失效（来源核对覆盖两来源）",async () => {
    const f=twoSourceFixture();
    const preview=await prepareWpsPythonResult(f.session,f.host,f.receipt,f.dest);
    f.changeSecondSource();
    await expect(preview.assertCurrent()).rejects.toThrow();
    expect(f.write).not.toHaveBeenCalled();
  });

  it("第一张来源变化同样让预览失效",async () => {
    const f=twoSourceFixture();
    const preview=await prepareWpsPythonResult(f.session,f.host,f.receipt,f.dest);
    f.changeFirstSource();
    await expect(preview.assertCurrent()).rejects.toThrow();
    expect(f.write).not.toHaveBeenCalled();
  });

  it("计划里出现未授权表时预览失败关闭（不复制未授权表）",async () => {
    const f=twoSourceFixture();
    const preview=await prepareWpsPythonResult(f.session,f.host,f.receipt,f.dest);
    // 篡改目标表名到未授权表：源版本核对与目标快照都不得通过。
    (preview.plan.actions[0] as {sheet:string}).sheet="Sheet9";
    await expect(preview.assertCurrent()).rejects.toThrow();
    expect(f.write).not.toHaveBeenCalled();
  });
});

/**
 * **来源与写入目标同表、区域不重叠**的真实支持程度（与 Excel 端同一判据）。
 *
 * 默认 `fixture()` 就是"同表不重叠"：来源 `Sheet1!A1:B1`、目标 `Sheet1!E1:F1`。
 * 这里要证的是：**同表不放松任何一道闸** —— 来源区域照样逐格核对，来源一变
 * 预览必须失效；而不是"因为同表了所以来源就不用看了"。
 */
describe("WPS local result preview: 同表不重叠也不放松来源核对",() => {
  it("来源与目标同表不重叠：仍逐格核对来源，来源一变预览即失效",async () => {
    const f=fixture();
    // 前提：来源与目标确实落在同一张表、但区域不同。
    expect(f.session.scope.sources[0].sheetId).toBe(target.sheetId);
    expect(f.session.scope.sources[0].range).not.toBe(target.range);
    const preview=await prepareWpsPythonResult(f.session,f.host,receipt,target);
    expect(preview.prepared.sources).toHaveLength(1);
    // 同表不重叠的写入不改来源区：预览立即可用、且一个单元格都没写。
    await expect(preview.assertCurrent()).resolves.toBeUndefined();
    expect(f.write).not.toHaveBeenCalled();

    // 改**同一张表上的来源区域**（不是目标区）：预览照样必须失效。
    f.sourceRange.Value2=[[9,2]] as never;
    await expect(preview.assertCurrent()).rejects.toThrow();
    expect(f.write).not.toHaveBeenCalled();
  });
});
