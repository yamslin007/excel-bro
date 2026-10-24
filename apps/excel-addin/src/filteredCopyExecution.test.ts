import { describe, expect, it } from "vitest";
import type { AnalysisPlan, ExcelAction } from "./contracts";
import { executePlan, undoExecution } from "./excel";
import { createOfficeTestHost, installOfficeTestHost } from "./officeTestHost";
const plan=(action:ExcelAction):AnalysisPlan=>({id:"filtered",title:"条件复制",summary:"条件复制",assumptions:[],warnings:[],actions:[action]});
const copy:Extract<ExcelAction,{type:"copyRange"}>={type:"copyRange",sheet:"目标",sourceSheet:"来源",sourceRange:"A1:B4",targetRange:"C2",copyType:"values",transpose:false,skipBlanks:false,hasHeaders:true,filters:[{field:"类别",operator:"equals",value:"甲"}]};
function host(){const h=createOfficeTestHost({sheets:[{name:"来源",values:[["类别","金额"],["甲",3],["乙",7],["甲",9]]},{name:"目标",values:[["before"]]}]});installOfficeTestHost(h);return h;}
describe("Excel filtered copy",()=>{
 it("copies header and matching rows then restores the target",async()=>{
  const h=host(),before=h.matrix("目标","C2:D4");
  const result=await executePlan(plan(copy),{authorizedSheetNames:["来源","目标"]});
  expect(h.matrix("目标","C2:D4")).toEqual([["类别","金额"],["甲",3],["甲",9]]);
  expect(result.verification.checks.length).toBeGreaterThan(0);
  expect(result.verification.checks.every(check=>check.passed)).toBe(true);
  await undoExecution(result.undoSnapshot!);expect(h.matrix("目标","C2:D4")).toEqual(before);
 });
 it("transposes links while retaining actual source row addresses",async()=>{
  const h=host(),result=await executePlan(plan({...copy,copyType:"link",transpose:true}),{authorizedSheetNames:["来源","目标"]});
  const formulas=await Excel.run(async context=>{const range=context.workbook.worksheets.getItem("目标").getRange("C2:E3");range.load("formulas");await context.sync();return range.formulas;});
  expect(formulas).toEqual([["='来源'!A1","='来源'!A2","='来源'!A4"],["='来源'!B1","='来源'!B2","='来源'!B4"]]);
  expect(result.verification.checks.every(check=>check.passed)).toBe(true);
  await undoExecution(result.undoSnapshot!);expect(h.cell("目标","C2") ?? "").toBe("");
 });
 it("keeps skipped cells and their later edits outside undo",async()=>{
  const h=host();h.setCell("来源","B2",null);h.setCell("目标","D3","keep");
  const result=await executePlan(plan({...copy,skipBlanks:true}),{authorizedSheetNames:["来源","目标"]});
  expect(h.cell("目标","D3")).toBe("keep");h.setCell("目标","D3","later");
  await undoExecution(result.undoSnapshot!);expect(h.cell("目标","D3")).toBe("later");
 });
 it("rejects a missing field and unsupported mode before any writes",async()=>{
  for(const action of [{...copy,filters:[{...copy.filters![0],field:"missing"}]},{...copy,copyType:"all" as const}]){
   const h=host();await expect(executePlan(plan(action),{authorizedSheetNames:["来源","目标"]})).rejects.toThrow();expect(h.writes()).toHaveLength(0);
  }
 });
});
