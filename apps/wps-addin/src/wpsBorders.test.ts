import { describe, it, expect } from "vitest";
import type { AnalysisPlan, ExcelAction } from "../../excel-addin/src/contracts";
import { assertWpsRecoveryDescriptor } from "../../excel-addin/src/contracts";
import { parseExplicitA1Rect } from "@excel-bro/core/planCore";
import { capabilityHost } from "./wpsCapabilityTestHost";
import { prepareWpsPlan, executeWpsPlan, undoWpsExecution, latestWpsUndo } from "./wpsPlanExecutor";
import { wpsRecoveryDescriptor } from "./wpsExecutionLedger";
import { borderTarget, readWpsBorder, type BorderAction } from "./wpsBorders";
import type { WpsRange } from "./wpsHost";
const plan=(actions:ExcelAction[]):AnalysisPlan=>({id:"borders",title:"边框",summary:"边框",actions,assumptions:[],warnings:[],acceptanceCriteria:[]});
const action=(range="A1:B2",sides:BorderAction["sides"]=["top","bottom","left","right","insideHorizontal","insideVertical"]):BorderAction=>({type:"setBorders",sheet:"明细",range,sides,color:"#123456",style:"dash",weight:"medium"});
// Opposite sides share physical edges. Setters also model automatic/theme colour provenance.
function borderHost() {
 const h=capabilityHost();
 const edges=new Map<string,Record<string,number>>();let writes=0;
 const controls:{missing?:string;fail?:string}={};
 const state=(key:string)=>{
  if(!edges.has(key)) edges.set(key,{LineStyle:-4142,Weight:2,Color:0,ColorIndex:-4105,ThemeColor:0,TintAndShade:0});
  return edges.get(key)!;
 };
 const edge=(r:number,c:number,index:number)=>{
  const key=index===8?`h:${r}:${c}`:index===9?`h:${r+1}:${c}`:index===7?`v:${r}:${c}`:`v:${r}:${c+1}`;
  const s=state(key);
  return new Proxy(s,{get:(target,k)=>controls.missing===k?undefined:target[String(k)],set:(target,k,v)=>{
   if(controls.fail===`${key}:${String(k)}`) throw new Error("border setter refused");
   target[String(k)]=v;writes++;
   if(k==="Color") {target.ThemeColor=0;target.ColorIndex=3;}
   if(k==="ThemeColor") {target.Color=1000+v;target.ColorIndex=3;}
   if(k==="ColorIndex"&&v===-4105) {target.Color=0;target.ThemeColor=0;}
   if(k==="Weight"||k==="Color") target.LineStyle=1;
   return true;
  }});
 };
 const range=h.sheets[0].Range!;
 h.sheets[0].Range=(address:string)=>{
  const result=range(address);const rect=parseExplicitA1Rect(address)!;const item=result.Item!;
  const attach=(cell:WpsRange,r:number,c:number)=>Object.assign(cell,{Borders:{Item:(i:number)=>edge(r,c,i)}});
  result.Item=(r:number,c:number)=>attach(item(r,c),rect.startRow+r-1,rect.startColumn+c-1);
  return attach(result,rect.startRow,rect.startColumn);
 };
 return {...h,edges,state,controls,borderWrites:()=>writes,cell:(a:string)=>h.sheets[0].Range!(a)};
}
describe("WPS recoverable borders",()=>{
 it("sets outer and inner edges, verifies descriptors and restores mixed automatic/theme/RGB borders",async()=>{
  const h=borderHost();
  Object.assign(h.state("h:1:1"),{LineStyle:4,Weight:4,Color:1005,ColorIndex:3,ThemeColor:5,TintAndShade:0.25});
  Object.assign(h.state("v:2:3"),{LineStyle:-4118,Weight:1,Color:1234,ColorIndex:3});
  const beforeValues=h.grids[0].map(row=>row.map(cell=>[cell.value,cell.formula]));
  const prepared=prepareWpsPlan(h.application,[1],plan([action()]));
  const before=[...h.edges].map(([key,value])=>[key,{...value}] as const);
  const descriptor=wpsRecoveryDescriptor(prepared);assertWpsRecoveryDescriptor(descriptor);
  expect(descriptor.actions.every(a=>a.formatFields?.every(f=>f.startsWith("border")))).toBe(true);
  expect(JSON.stringify(descriptor)).not.toContain("1005");
  const record=await executeWpsPlan(h.application,prepared);
  for(const key of ["h:1:1","h:1:2","h:2:1","h:2:2","h:3:1","h:3:2","v:1:1","v:1:2","v:1:3","v:2:1","v:2:2","v:2:3"]) {
   expect(h.state(key).LineStyle,key).toBe(-4115);expect(h.state(key).Weight,key).toBe(-4138);
  }
  expect(readWpsBorder(h.cell("A1"),"borderBottom")).toBe(borderTarget(action()));
  h.grids[0][0][0].bold=true;
  await expect(undoWpsExecution(h.application,record.executionId)).resolves.toBeDefined();
  for(const [key,value] of before) {
   const actual=h.state(key);expect({...actual,ColorIndex:value.ColorIndex},key).toEqual(value);
  }
  expect(h.grids[0][0][0].bold).toBe(true);
  expect(h.grids[0].map(row=>row.map(cell=>[cell.value,cell.formula]))).toEqual(beforeValues);
 });
 it("preserves undeclared borders and rejects external changes to the owned physical edge before any undo write",async()=>{
  const h=borderHost();const record=await executeWpsPlan(h.application,prepareWpsPlan(h.application,[1],plan([action("A1",["bottom"])])));
  h.state("v:1:1").LineStyle=4;
  h.cell("A2").Borders!.Item!(8)!.LineStyle=1; // The same physical edge, other cell.
  const writes=h.borderWrites();await expect(undoWpsExecution(h.application,record.executionId)).rejects.toThrow();
  expect(h.borderWrites()).toBe(writes);expect(h.state("v:1:1").LineStyle).toBe(4);
 });
 it.each(["ThemeColor","TintAndShade","ColorIndex"])("refuses missing %s before changing any edge",missing=>{
  const h=borderHost();h.controls.missing=missing;
  expect(()=>prepareWpsPlan(h.application,[1],plan([action()]))).toThrow();expect(h.borderWrites()).toBe(0);expect(h.writes()).toBe(0);
 });
 it("retains a partial recovery record when a later border setter throws",async()=>{
  const h=borderHost();const prepared=prepareWpsPlan(h.application,[1],plan([action("A1:B1",["top"])]));
  h.controls.fail="h:1:2:LineStyle";
  await expect(executeWpsPlan(h.application,prepared)).rejects.toThrow("border setter refused");
  expect(h.state("h:1:1").LineStyle).toBe(-4115);
  const record=latestWpsUndo(h.application);expect(record).not.toBeNull();h.controls.fail=undefined;
  await undoWpsExecution(h.application,record!.executionId);
  expect(h.state("h:1:1").LineStyle).toBe(-4142);expect(h.state("h:1:2").LineStyle).toBe(-4142);
 });
 it("rejects conflicting adjacent ownership, while repeated writes through the same owner remain recoverable",async()=>{
  const h=borderHost();
  expect(()=>prepareWpsPlan(h.application,[1],plan([action("A1",["bottom"]),action("A2",["top"])]))).toThrow("相邻单元格");expect(h.borderWrites()).toBe(0);
  const record=await executeWpsPlan(h.application,prepareWpsPlan(h.application,[1],plan([action("A1",["bottom"]),{...action("A1",["bottom"]),style:"double"}])));
  expect(h.state("h:2:1").LineStyle).toBe(-4119);
  await expect(undoWpsExecution(h.application,record.executionId)).resolves.toBeDefined();expect(h.state("h:2:1").LineStyle).toBe(-4142);
 });
 it("checks authorized sheets and changes in flight before any write",async()=>{
  const h=borderHost();expect(()=>prepareWpsPlan(h.application,[],plan([action()]))).toThrow();
  const prepared=prepareWpsPlan(h.application,[1],plan([action("A1",["top"])]));
  await expect(executeWpsPlan(h.application,prepared,undefined,undefined,async phase=>{if(phase==="before")h.state("h:1:1").LineStyle=4;})).rejects.toThrow();
  expect(h.borderWrites()).toBe(0);
 });
});
