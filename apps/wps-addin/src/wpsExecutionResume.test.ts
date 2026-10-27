import {reportWpsHostWrites} from "../../excel-addin/src/api";
import {armRecoveryProbe,bindRecoveryProbe,interruptRecoveryProbe} from "@excel-bro/core/recoveryProbe";
import {inspectWpsRecoveryState} from "./wpsRecoveryInspection";
import {fixture, operationId} from "./wpsResume.testHost";
import {beforeEach, describe, expect, it, vi} from "vitest";
import type {WpsExecutionState} from "../../excel-addin/src/contracts";
import {assertWpsRecoveryDescriptor, assertWpsExecutionResult} from "../../excel-addin/src/contracts";
import {executeWpsPlan, prepareWpsPlan, prepareWpsResume, undoWpsExecution, wpsUndoRecordFor} from "./wpsPlanExecutor";
import {canResumeWpsExecution, resumeWpsExecution} from "./wpsExecutionResume";
import {wpsReconcileVerdict, wpsReconcileActions} from "./wpsExecutionReconcile";
import * as ledger from "./wpsExecutionLedger";

vi.mock("../../excel-addin/src/api", async original => ({...await original<typeof import("../../excel-addin/src/api")>(), reportWpsHostWrites: vi.fn(async () => ({}))}));
vi.mock("./wpsExecutionLedger", async original => ({...await original<typeof import("./wpsExecutionLedger")>(),
  claimWpsExecutionResume: vi.fn(), startWpsAction: vi.fn(), completeWpsAction: vi.fn(),
  finishWpsExecution: vi.fn(), suspendWpsExecution: vi.fn()}));
async function interrupted(next=2) {
  const f=fixture();
  await expect(executeWpsPlan(f.application,f.prepared,undefined,undefined,async (phase,c) => {
    if (phase === "after" && c.actionIndex === next-1) throw new Error("lost receipt");
  },operationId,{taskId:"task-1"})).rejects.toThrow("lost receipt");
  f.state.nextActionIndex=next; // Production reconciliation has proved the last applied action.
  f.writes.length=0;
  return f;
}
function wire(state: WpsExecutionState, onClaim?:()=>void, onStart?:()=>void) {
  let owned={...state,ownerActive:true,ownerToken:"owner-new",revision:state.revision+1};
  vi.mocked(ledger.claimWpsExecutionResume).mockImplementation(async()=>{onClaim?.();return owned;});
  vi.mocked(ledger.startWpsAction).mockImplementation(async(_id,req)=>{onStart?.();
    expect(req.ownerToken).toBe("owner-new"); expect(req.actionIndex).toBe(owned.nextActionIndex);
    owned={...owned,inFlightActionIndex:req.actionIndex,revision:owned.revision+1};return owned;});
  vi.mocked(ledger.completeWpsAction).mockImplementation(async(_id,req)=>{
    expect(req.actionIndex).toBe(owned.inFlightActionIndex);
    owned={...owned,inFlightActionIndex:null,nextActionIndex:req.actionIndex+1,revision:owned.revision+1};return owned;});
  vi.mocked(ledger.finishWpsExecution).mockImplementation(async(_id, request)=>{assertWpsExecutionResult(request.result);return {...owned,status:"completed"};});
  vi.mocked(ledger.suspendWpsExecution).mockImplementation(async()=>({...owned,ownerActive:false}));
}
beforeEach(()=>vi.clearAllMocks());
describe("WPS exact remaining action recovery",()=>{
  it("real acceptance fault leaves second in flight, fresh host checks it and resumes only third",async()=>{
    const f=fixture(), m=new Map<string,string>();
    vi.stubGlobal("localStorage",{getItem:(k:string)=>m.get(k)??null,setItem:(k:string,v:string)=>m.set(k,v),removeItem:(k:string)=>m.delete(k)});
    try {
      armRecoveryProbe(true,f.state.workbookIdentity);
      bindRecoveryProbe(true,f.state.workbookIdentity,operationId,f.prepared.plan.actions,true);
      const receipts:number[]=[]; f.state.nextActionIndex=0;
      await expect(executeWpsPlan(f.application,f.prepared,undefined,undefined,async(phase,c)=>{
        if(phase==="before") {f.state.inFlightActionIndex=c.actionIndex;return;}
        interruptRecoveryProbe(true,operationId,c.actionIndex);
        receipts.push(c.actionIndex); f.state.nextActionIndex=c.actionIndex+1;f.state.inFlightActionIndex=null;
      },operationId,{taskId:"task-1"})).rejects.toThrow("第二步已写入并回读");
      expect(receipts).toEqual([0]);expect(f.state.nextActionIndex).toBe(1);expect(f.state.inFlightActionIndex).toBe(1);
      expect([1,2,3].map(i=>f.cells.get(`Target!A${i}`))).toEqual([10,20,"old3"]);
      const reopened={...f.application};
      expect(inspectWpsRecoveryState(reopened,f.state,[1,2])).toBe("completed");
      // API transition only after the production native inspection, same as the panel's reconciliation endpoint.
      f.state.nextActionIndex=2;f.state.inFlightActionIndex=null;wire(f.state);
      const before=[...f.writes];await resumeWpsExecution(reopened,f.state,()=>[1,2]);
      expect(f.writes.slice(before.length)).toEqual(["Target!A3"]);
      expect(vi.mocked(ledger.startWpsAction).mock.calls.map(c=>c[1].actionIndex)).toEqual([2]);
      expect(vi.mocked(ledger.completeWpsAction).mock.calls.map(c=>c[1].actionIndex)).toEqual([2]);
    } finally {vi.unstubAllGlobals();}
  });

  it("three value copies use the production lowering, execute once and undo all three",async()=>{
    const f=fixture();
    [111,222,333].forEach((v,i)=>f.cells.set(`Source!A${i+1}`,v));
    const prepared=prepareWpsPlan(f.application,[1,2],{...f.prepared.plan,sourceDependencySheets:undefined,
      actions:[1,2,3].map(r=>({type:"copyRange",sheet:"Target",sourceSheet:"Source",sourceRange:`A${r}`,
        targetRange:`A${r}`,copyType:"values",skipBlanks:false,transpose:false}))});
    await executeWpsPlan(f.application,prepared,undefined,undefined,undefined,operationId,{taskId:"task-1"});
    expect(f.writes).toEqual(["Target!A1","Target!A2","Target!A3"]);
    expect([1,2,3].map(r=>f.cells.get(`Target!A${r}`))).toEqual([111,222,333]);
    await undoWpsExecution(f.application,operationId);
    expect([1,2,3].map(r=>f.cells.get(`Target!A${r}`))).toEqual(["old1","old2","old3"]);
    expect([1,2,3].map(r=>f.cells.get(`Source!A${r}`))).toEqual([111,222,333]);
  });

  it("dependent sequential copies remain blocked before any writes",()=>{
    const f=fixture();
    expect(()=>prepareWpsPlan(f.application,[1,2],{...f.prepared.plan,actions:[
      {type:"copyRange",sheet:"Target",sourceSheet:"Source",sourceRange:"A1",targetRange:"A1",copyType:"values",skipBlanks:false,transpose:false},
      {type:"copyRange",sheet:"Target",sourceSheet:"Target",sourceRange:"A1",targetRange:"A2",copyType:"values",skipBlanks:false,transpose:false}
    ]})).toThrow("复制来源必须");
    expect(f.writes).toEqual([]);
  });
  it("independent copy batch with formula sources remains blocked before writes",()=>{
    const f=fixture(); f.cells.set("Source!A1","=1+1");
    expect(()=>prepareWpsPlan(f.application,[1,2],{...f.prepared.plan,actions:[1,2].map(r=>({
      type:"copyRange",sheet:"Target",sourceSheet:"Source",sourceRange:"A1",targetRange:`A${r}`,copyType:"values",skipBlanks:false,transpose:false
    }))})).toThrow("公式依赖");
    expect(f.writes).toEqual([]);
  });
  it.each([1,2])("after %i completed actions, writes only the suffix and undo restores the entire original plan",async next=>{
    const f=await interrupted(next); wire(f.state);
    // Pane recreation: use a fresh application object with the same host and PluginStorage.
    const reopened={...f.application};
    expect(canResumeWpsExecution(reopened,f.state,[2,1])).toBe(true);
    const verdict=wpsReconcileVerdict({state:f.state,identityMatch:true,localInspection:"not_started",resumeAvailable:true});
    expect(wpsReconcileActions(verdict)).toContain("resume");
    const record=await resumeWpsExecution(reopened,f.state,()=>[1,2]);
    expect(f.writes).toEqual(Array.from({length:3-next},(_,i)=>`Target!A${next+i+1}`));
    expect(record.executionId).toBe(operationId); expect(record.state).toBe("verified");expect(record.resume).toBeUndefined();
    expect(ledger.startWpsAction).toHaveBeenCalledTimes(3-next);expect(ledger.finishWpsExecution).toHaveBeenCalledTimes(1);
    expect(reportWpsHostWrites).toHaveBeenCalledTimes(1);
    const report = vi.mocked(reportWpsHostWrites).mock.calls[0];
    expect(report[0]).toBe(operationId);
    expect(report[1].entries.map(entry => entry.actionIndex)).toEqual(Array.from({length: 3-next}, (_, i) => next+i));
    for (const entry of report[1].entries) {
      expect(entry).toMatchObject({outcome: "written", readback: "verified"});
      expect(entry.hostWriteCalls).toBeGreaterThan(0);
    }

    await undoWpsExecution(reopened,operationId);
    expect([1,2,3].map(i=>f.cells.get(`Target!A${i}`))).toEqual(["old1","old2","old3"]);
  });
  it("SQLite sorted object keys preserve original evidence and only resume the third action", async () => {
    const f = await interrupted();
    const sorted = (value: unknown): unknown => Array.isArray(value) ? value.map(sorted) :
      value && typeof value === "object" ? Object.fromEntries(Object.entries(value).sort(([a],[b]) => a.localeCompare(b)).map(([k,v]) => [k, sorted(v)])) : value;
    f.state.descriptor = sorted(f.state.descriptor) as typeof f.state.descriptor;
    wire(f.state);
    expect(canResumeWpsExecution(f.application, f.state, [1,2])).toBe(true);
    await resumeWpsExecution(f.application, f.state, () => [1,2]);
    expect(f.writes).toEqual(["Target!A3"]);
    expect(ledger.startWpsAction).toHaveBeenCalledTimes(1);
  });
  it.each(["source","completed target","remaining target","authorization","expired","task","matrix","old record","inflight"])("refuses %s before claiming or writing",async issue=>{
    const f=await interrupted(); let indexes=[1,2];
    const records=JSON.parse(String(f.storage.get("excel-bro.wps.undo.v1")));
    if(issue==="source")f.cells.set("Source!A1","changed");
    if(issue==="completed target")f.cells.set("Target!A1","changed");
    if(issue==="remaining target")f.cells.set("Target!A3","changed");
    if(issue==="authorization")indexes=[2];
    if(issue==="expired")records[0].createdAt="2000-01-01T00:00:00Z";
    if(issue==="task")f.state.taskId="other";
    if(issue==="matrix")records[0].actions[2].values=[[999]];
    if(issue==="old record")delete records[0].resume;
    if(issue==="inflight")f.state.inFlightActionIndex=2;
    f.storage.set("excel-bro.wps.undo.v1",JSON.stringify(records));
    expect(canResumeWpsExecution(f.application,f.state,indexes)).toBe(false);
    await expect(resumeWpsExecution(f.application,f.state,()=>indexes)).rejects.toThrow();
    expect(ledger.claimWpsExecutionResume).not.toHaveBeenCalled();expect(f.writes).toEqual([]);
  });
  it.each([
    ["claim","source"], ["checkpoint","source"],
    ["claim","authorization"], ["checkpoint","authorization"]
  ])("rechecks after awaited %s when %s changes",async (point, issue)=>{
    const f=await interrupted(); let indexes=[1,2];
    const change=()=>{
      if(issue==="source") f.cells.set("Source!A1","changed during await");
      else indexes=[2];
    };
    wire(f.state,point==="claim"?change:undefined,point==="checkpoint"?change:undefined);
    await expect(resumeWpsExecution(f.application,f.state,()=>indexes)).rejects.toThrow();
    expect(f.writes).toEqual([]);expect(ledger.suspendWpsExecution).toHaveBeenCalledTimes(1);
  });
  it("lost claim / concurrent owner never writes",async()=>{
    const f=await interrupted();vi.mocked(ledger.claimWpsExecutionResume).mockRejectedValueOnce(new Error("409 owner active"));
    await expect(resumeWpsExecution(f.application,f.state,()=>[1,2])).rejects.toThrow("409");expect(f.writes).toEqual([]);
  });
  it("production copyRange expansion supplies complete evidence without invented dependencies",()=>{
    const f=fixture();
    const prepared=prepareWpsPlan(f.application,[1,2],{...f.prepared.plan,sourceDependencySheets:undefined,
      actions:[{type:"copyRange",sourceSheet:"Source",sourceRange:"A1:A1",sheet:"Target",targetRange:"A3",copyType:"values",skipBlanks:false,transpose:false}]});
    const descriptor=ledger.wpsRecoveryDescriptor(prepared);
    expect(descriptor.resumeContext?.sources.map(source=>source.sheetName)).toEqual(["Source"]);
    expect(descriptor.actions.map(action=>action.type)).toEqual(["writeValues"]);
  });
  it("cannot smuggle raw values in resume metadata; old descriptors still accepted",()=>{
    const f=fixture(),d=f.state.descriptor;
    expect(JSON.stringify(d)).not.toContain("private-source");expect(d.resumeContext).toBeTruthy();
    expect(()=>assertWpsRecoveryDescriptor(d)).not.toThrow();
    expect(()=>assertWpsRecoveryDescriptor({actions:d.actions})).not.toThrow();
    expect(()=>assertWpsRecoveryDescriptor({...d,resumeContext:{...d.resumeContext,values:[["private"]]}})).toThrow();
  });
  it("preserves original age and rollback contents on a second interruption",async()=>{
    const f=await interrupted(1),before=wpsUndoRecordFor(f.application,operationId)!;wire(f.state);
    vi.mocked(ledger.completeWpsAction).mockRejectedValueOnce(new Error("lost again"));
    await expect(resumeWpsExecution(f.application,f.state,()=>[1,2])).rejects.toThrow("lost again");
    const after=wpsUndoRecordFor(f.application,operationId)!;
    expect(after.createdAt).toBe(before.createdAt);expect(after.actions.map(a=>a.beforeValues)).toEqual(before.actions.map(a=>a.beforeValues));
    expect(f.writes).toEqual(["Target!A2"]);expect(after.resume).toBeTruthy();
    expect(()=>prepareWpsResume(f.application,{...f.state,nextActionIndex:2},[1,2])).not.toThrow();
  });
});
