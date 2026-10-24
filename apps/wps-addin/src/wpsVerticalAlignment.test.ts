import { describe, it, expect } from "vitest";
import type { AnalysisPlan, ExcelAction } from "../../excel-addin/src/contracts";
import { assertWpsRecoveryDescriptor } from "../../excel-addin/src/contracts";
import { capabilityHost } from "./wpsCapabilityTestHost";
import { prepareWpsPlan, executeWpsPlan, undoWpsExecution, inspectWpsUndoRecord, latestWpsUndo } from "./wpsPlanExecutor";
import { wpsRecoveryDescriptor } from "./wpsExecutionLedger";
const plan=(actions:ExcelAction[]):AnalysisPlan=>({id:"vertical",title:"垂直对齐",summary:"对齐",actions,assumptions:[],warnings:[],acceptanceCriteria:[]});
describe("WPS vertical alignment",()=>{
 it("restores mixed vertical positions after overlapping actions and preserves unrelated later edits",async()=>{
  const h=capabilityHost();h.grids[0][0][0].verticalAlignment=-4160;h.grids[0][1][0].verticalAlignment=-4107;
  const prepared=prepareWpsPlan(h.application,[1],plan([{type:"setAlignment",sheet:"明细",range:"A1:A2",vertical:"center"},{type:"setAlignment",sheet:"明细",range:"A2",vertical:"top"}]));
  const descriptor=wpsRecoveryDescriptor(prepared);assertWpsRecoveryDescriptor(descriptor);
  expect(descriptor.actions[0].formatFields).toEqual(["verticalAlignment"]);
  const record=await executeWpsPlan(h.application,prepared);
  expect(h.grids[0][0][0].verticalAlignment).toBe(-4108);expect(h.grids[0][1][0].verticalAlignment).toBe(-4160);
  h.grids[0][0][0].fillColor=123;h.grids[0][0][0].bold=true;
  await expect(undoWpsExecution(h.application,record.executionId)).resolves.toBeDefined();
  expect(h.grids[0][0][0].verticalAlignment).toBe(-4160);expect(h.grids[0][1][0].verticalAlignment).toBe(-4107);
  expect(h.grids[0][0][0].fillColor).toBe(123);expect(h.grids[0][0][0].bold).toBe(true);
 });
 it("detects external changes to the declared attribute before any restoration",async()=>{
  const h=capabilityHost(),prepared=prepareWpsPlan(h.application,[1],plan([{type:"setAlignment",sheet:"明细",range:"A1:A2",vertical:"center"}]));
  const record=await executeWpsPlan(h.application,prepared);h.grids[0][1][0].verticalAlignment=-4160;
  const writes=h.writes();expect(inspectWpsUndoRecord(h.application,record.executionId).kind).not.toBe("completed");
  await expect(undoWpsExecution(h.application,record.executionId)).rejects.toThrow();expect(h.writes()).toBe(writes);
 });
 it("refuses an intervening format edit while awaiting the checkpoint",async()=>{
  const h=capabilityHost(),prepared=prepareWpsPlan(h.application,[1],plan([{type:"setAlignment",sheet:"明细",range:"A1:A2",vertical:"center"}]));
  await expect(executeWpsPlan(h.application,prepared,undefined,undefined,async phase=>{
   if(phase==="before") h.grids[0][0][0].verticalAlignment=-4160;
  })).rejects.toThrow("检查期间目标");
  expect(h.writes()).toBe(0);expect(h.grids[0][0][0].verticalAlignment).toBe(-4160);
 });
 it("retains recovery when a later vertical setter throws after an earlier cell changed",async()=>{
  const h=capabilityHost();const prepared=prepareWpsPlan(h.application,[1],plan([{type:"setAlignment",sheet:"明细",range:"A1:A2",vertical:"center"}]));
  let current=-4107;let fail=true;
  Object.defineProperty(h.grids[0][1][0],"verticalAlignment",{get:()=>current,set:value=>{if(fail)throw new Error("host refused second cell");current=value;},configurable:true});
  await expect(executeWpsPlan(h.application,prepared)).rejects.toThrow("host refused");
  expect(h.grids[0][0][0].verticalAlignment).toBe(-4108);
  const recovery=latestWpsUndo(h.application);expect(recovery, "partial format writes must retain their recovery record").not.toBeNull();
  fail=false;await expect(undoWpsExecution(h.application,recovery!.executionId)).resolves.toBeDefined();
  expect(h.grids[0][0][0].verticalAlignment).toBe(-4107);expect(current).toBe(-4107);
 });
 it("restores overlapping formats even when the second action stops halfway",async()=>{
  const h=capabilityHost();const prepared=prepareWpsPlan(h.application,[1],plan([{type:"setAlignment",sheet:"明细",range:"A1:A2",vertical:"center"},{type:"setAlignment",sheet:"明细",range:"A2:A3",vertical:"top"}]));
  let current=-4107,fail=true;
  Object.defineProperty(h.grids[0][2][0],"verticalAlignment",{get:()=>current,set:value=>{if(fail)throw new Error("second action interrupted");current=value;},configurable:true});
  await expect(executeWpsPlan(h.application,prepared)).rejects.toThrow("interrupted");
  expect(h.grids[0][1][0].verticalAlignment).toBe(-4160);
  const recovery=latestWpsUndo(h.application)!;expect(recovery).not.toBeNull();fail=false;
  await expect(undoWpsExecution(h.application,recovery.executionId)).resolves.toBeDefined();
  expect(h.grids[0].slice(0,3).map(row=>row[0].verticalAlignment)).toEqual([-4107,-4107,-4107]);
 });
 it("does not treat missing post-write format evidence as permission to restore",async()=>{
  const h=capabilityHost(),prepared=prepareWpsPlan(h.application,[1],plan([{type:"setAlignment",sheet:"明细",range:"A1",vertical:"center"}]));
  let current=-4107,readable=true;
  Object.defineProperty(h.grids[0][0][0],"verticalAlignment",{get:()=>readable?current:undefined,set:value=>{current=value;readable=false;},configurable:true});
  await expect(executeWpsPlan(h.application,prepared)).rejects.toThrow();
  const recovery=latestWpsUndo(h.application)!;expect(recovery).not.toBeNull();
  const writes=h.writes();await expect(undoWpsExecution(h.application,recovery.executionId)).rejects.toThrow();
  expect(h.writes(),"unreadable attributes must block restoration before any setter").toBe(writes);expect(current).toBe(-4108);
 });
 it("does not write when one cell cannot provide its previous vertical alignment",()=>{
  const h=capabilityHost();Object.defineProperty(h.grids[0][1][0],"verticalAlignment",{get:()=>undefined});
  expect(()=>prepareWpsPlan(h.application,[1],plan([{type:"setAlignment",sheet:"明细",range:"A1:A2",vertical:"center"}]))).toThrow("verticalAlignment");expect(h.writes()).toBe(0);
 });
});
