import {describe,it,expect} from "vitest";
import type {AnalysisPlan,ExcelAction} from "../../excel-addin/src/contracts";
import {prepareWpsPlan,executeWpsPlan,undoWpsExecution} from "./wpsPlanExecutor";
import {wpsRecoveryDescriptor} from "./wpsExecutionLedger";
import {capabilityHost} from "./wpsCapabilityTestHost";
const plan=(actions:ExcelAction[]):AnalysisPlan=>({id:"copy",title:"复制",summary:"复制",assumptions:[],warnings:[],acceptanceCriteria:[],actions});
const action:Extract<ExcelAction,{type:"copyRange"}>={type:"copyRange",sheet:"目标",sourceSheet:"来源",sourceRange:"A1:B2",targetRange:"C3:D4",copyType:"values",skipBlanks:false,transpose:false};
describe("WPS authorized copy",()=>{
 it("copies values across selected sheets, transposes and restores original formulas/formats",async()=>{
   const host=capabilityHost(["来源","目标","私有"]);const before=structuredClone(host.grids);
   const prepared=prepareWpsPlan(host.application,[1,2],plan([{...action,transpose:true}]));
   expect(host.writes()).toBe(0);
   const record=await executeWpsPlan(host.application,prepared);
   expect(host.grids[1][2][2].value).toBe(before[0][0][0].value);
   expect(host.grids[1][2][3].value).toBe(before[0][1][0].value);
   expect(host.grids[1][3][2].value).toBe(before[0][0][1].value);
   expect(host.grids[0]).toEqual(before[0]);expect(host.grids[2]).toEqual(before[2]);
   await undoWpsExecution(host.application,record.executionId);expect(host.grids).toEqual(before);
 });
 it("skips empty source cells without rewriting target formulas",async()=>{
   const host=capabilityHost(["来源","目标"]);
   host.grids[0][0][0].value=null;host.grids[0][1][1].value="";
   host.grids[1][2][2].formula="=SUM(A1:A2)";
   const before=structuredClone(host.grids);
   const record=await executeWpsPlan(host.application,prepareWpsPlan(host.application,[1,2],plan([{...action,skipBlanks:true}])));
   expect(host.grids[1][2][2].formula).toBe("=SUM(A1:A2)");
   expect(host.grids[1][3][3]).toEqual(before[1][3][3]);
   await undoWpsExecution(host.application,record.executionId);expect(host.grids).toEqual(before);
 });
 it("creates exact authorized link formulas with quoted names",async()=>{
   const host=capabilityHost(["来源'表","目标"]);
   const record=await executeWpsPlan(host.application,prepareWpsPlan(host.application,[1,2],plan([{...action,sourceSheet:"来源'表",copyType:"link"}])));
   expect(host.grids[1][2][2].formula).toBe("='来源''表'!A1");
   expect(host.grids[1][3][3].formula).toBe("='来源''表'!B2");
   await undoWpsExecution(host.application,record.executionId);
 });
 it("source changes after preparation block all writes and outbound evidence contains no values",async()=>{
   const host=capabilityHost(["来源","目标"]),prepared=prepareWpsPlan(host.application,[1,2],plan([action]));
   const evidence=JSON.stringify(wpsRecoveryDescriptor(prepared));expect(evidence).not.toContain("old-");
   host.grids[0][0][0].value="changed";
   await expect(executeWpsPlan(host.application,prepared)).rejects.toThrow("源区域");expect(host.writes()).toBe(0);
 });
 it("checks all selected source/target identities before reading a private source",()=>{
   const host=capabilityHost(["来源","目标"]);let reads=0;
   const range=host.sheets[0].Range!;host.sheets[0].Range=address=>{reads++;return range(address);};
   expect(()=>prepareWpsPlan(host.application,[2],plan([action]))).toThrow("未授权");expect(reads).toBe(0);expect(host.writes()).toBe(0);
 });
 it("rejects unsupported variants and wrong shapes before writes",()=>{
   const host=capabilityHost(["来源","目标"]);
   for(const patch of [{copyType:"all"},{copyType:"formats"},{copyType:"formulas"},{targetRange:"A1:B3"},{targetRange:"XFD1048576"}] as Partial<typeof action>[]){
     expect(()=>prepareWpsPlan(host.application,[1,2],plan([{...action,...patch}]))).toThrow();expect(host.writes()).toBe(0);
   }
 });
 it("retains a dynamic copy template and protects it against mutation",async()=>{
   const host=capabilityHost(["来源","目标"]),prepared=prepareWpsPlan(host.application,[1,2],plan([action]));
   expect(prepared.originalPlan?.actions).toEqual([action]);expect(prepared.originalPlan?.sourceDependencySheets).toEqual(["来源"]);
   const saved=structuredClone(prepared.originalPlan!);
   host.grids[0][0][0].value="new source";host.grids[0][0][0].formula="new source";
   const fresh=prepareWpsPlan(host.application,[1,2],saved);
   await executeWpsPlan(host.application,fresh);expect(host.grids[1][2][2].value).toBe("new source");
   const edited=prepareWpsPlan(host.application,[1,2],saved);edited.originalPlan!.actions[0].sheet="私有";
   const writes=host.writes();await expect(executeWpsPlan(host.application,edited)).rejects.toThrow("计划或预览已变化");expect(host.writes()).toBe(writes);
 });
 it("does not pad an incomplete source matrix with invented blanks",()=>{
   const host=capabilityHost(["来源","目标"]),range=host.sheets[0].Range!;
   host.sheets[0].Range=address=>{const result=range(address);return {...result,Value2:[["partial"]]};};
   expect(()=>prepareWpsPlan(host.application,[1,2],plan([action]))).toThrow("矩阵不完整");expect(host.writes()).toBe(0);
 });
 it("does not use stale source data after an earlier write in the same plan",()=>{
   const host=capabilityHost(["来源","目标"]);
   expect(()=>prepareWpsPlan(host.application,[1,2],plan([{type:"writeValues",sheet:"来源",range:"A1",values:[[99]]},action]))).toThrow("前序数据修改");
   expect(host.writes()).toBe(0);
 });
});
