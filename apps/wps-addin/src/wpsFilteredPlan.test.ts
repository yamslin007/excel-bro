import { describe, it, expect } from "vitest";
import type { AnalysisPlan, ExcelAction } from "../../excel-addin/src/contracts";
import { capabilityHost } from "./wpsCapabilityTestHost";
import { prepareWpsPlan, executeWpsPlan, undoWpsExecution } from "./wpsPlanExecutor";
const plan = (actions:ExcelAction[]):AnalysisPlan=>({id:"filters",title:"条件操作",summary:"条件操作",assumptions:[],warnings:[],acceptanceCriteria:[],actions});
const filters = [{field:"类别",operator:"equals" as const,value:"甲"}];
function fixture() {
 const h=capabilityHost(["来源","目标"]);
 [["类别","金额"],["甲",3],["乙",7],["甲",9]].forEach((row,r)=>row.forEach((v,c)=>{h.grids[0][r][c].value=v;h.grids[0][r][c].formula=v;}));
 return h;
}
describe("WPS filtered operation execution",()=>{
 it("copies selected rows and header, preserves link addresses through transposition, and undoes",async()=>{
  const h=fixture(),before=structuredClone(h.grids);
  const action:ExcelAction={type:"copyRange",sheet:"目标",sourceSheet:"来源",sourceRange:"A1:B4",targetRange:"C2",copyType:"link",transpose:true,skipBlanks:false,hasHeaders:true,filters};
  const prepared=prepareWpsPlan(h.application,[1,2],plan([action]));
  expect(prepared.originalPlan?.actions[0]).toEqual(action);
  const record=await executeWpsPlan(h.application,prepared);
  expect(h.grids[1][1].slice(2,5).map(cell=>cell.formula)).toEqual(["='来源'!A1","='来源'!A2","='来源'!A4"]);
  expect(h.grids[1][2].slice(2,5).map(cell=>cell.formula)).toEqual(["='来源'!B1","='来源'!B2","='来源'!B4"]);
  await undoWpsExecution(h.application,record.executionId);expect(h.grids).toEqual(before);
 });
 it("only clears matching data rows, preserves headers, nonmatching rows and formatting, and restores",async()=>{
  const h=fixture(),before=structuredClone(h.grids);
  const prepared=prepareWpsPlan(h.application,[1],plan([{type:"clearRange",sheet:"来源",range:"A1:B4",applyTo:"contents",filters}]));
  const record=await executeWpsPlan(h.application,prepared);
  expect(h.grids[0][0]).toEqual(before[0][0]);expect(h.grids[0][2]).toEqual(before[0][2]);
  for(const r of [1,3]) {expect(h.grids[0][r][0].value).toBe(null);expect(h.grids[0][r][0].fillColor).toBe(before[0][r][0].fillColor);}
  await undoWpsExecution(h.application,record.executionId);expect(h.grids).toEqual(before);
 });
 it("does not clear rows if the condition source changes after preparation",async()=>{
  const h=fixture(),prepared=prepareWpsPlan(h.application,[1],plan([{type:"clearRange",sheet:"来源",range:"A1:B4",applyTo:"contents",filters}]));
  h.grids[0][2][0].value="甲";h.grids[0][2][0].formula="甲";
  await expect(executeWpsPlan(h.application,prepared)).rejects.toThrow("源区域");expect(h.writes()).toBe(0);
 });
 it.each(["missing","ambiguous","headerless"])("rejects %s fields instead of guessing an empty match",kind=>{
  const h=fixture();if(kind==="ambiguous"){h.grids[0][0][1].value="类别";h.grids[0][0][1].formula="类别";}
  expect(()=>prepareWpsPlan(h.application,[1],plan([{type:"clearRange",sheet:"来源",range:"A1:B4",applyTo:"contents",hasHeaders:kind!=="headerless",filters:kind==="missing"?[{...filters[0],field:"没有"}]:filters}]))).toThrow(/字段|表头/);
  expect(h.writes()).toBe(0);
 });
});
