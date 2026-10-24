import {describe,it,expect} from "vitest";
import capabilities from "../../../config/capabilities.json";
import type {AnalysisPlan} from "../../excel-addin/src/contracts";
import {prepareWpsPlan,executeWpsPlan,undoWpsExecution,expandWpsPlanActions} from "./wpsPlanExecutor";
import {wpsRecoveryDescriptor} from "./wpsExecutionLedger";
import {capabilityHost} from "./wpsCapabilityTestHost";
const plan=(actions:AnalysisPlan["actions"]):AnalysisPlan=>({id:"write-table",title:"写表",summary:"写表",assumptions:[],warnings:[],acceptanceCriteria:[],actions});
const action={type:"writeTable",sheet:"明细",startCell:"B2",headers:["名称","数量"],rows:[["甲",2],["乙",3]]} as const;
const table=()=>({...action,headers:[...action.headers],rows:action.rows.map(row=>[...row])});
describe("WPS writeTable actual effects",()=>{
  it("writes headers/data once, styles only headers, checkpoints each effect and restores mixed originals",async()=>{
    const host=capabilityHost();host.grids[0][1][1].bold=true;host.grids[0][1][2].fillPattern=3;
    host.grids[0][2][1].formula="=1+2";host.grids[0][2][1].value="=1+2";
    const before=structuredClone(host.grids);
    const prepared=prepareWpsPlan(host.application,[1],plan([table()]));
    expect(host.writes()).toBe(0);
    expect(prepared.plan.actions.map(a=>a.type)).toEqual(capabilities.wpsWrite.actionExpansion.writeTable);
    const descriptor=wpsRecoveryDescriptor(prepared), events:Array<{phase:string;fingerprint:string|null}>=[];
    const record=await executeWpsPlan(host.application,prepared,undefined,undefined,async(phase,checkpoint)=>{
      events.push({phase,fingerprint:phase==="before"?checkpoint.beforeFingerprint:checkpoint.afterFingerprint});
    });
    expect(record.state).toBe("verified");
    expect(events.map(x=>x.fingerprint)).toEqual(descriptor.actions.flatMap(a=>[a.previewFingerprint,a.expectedFingerprint]));
    expect(host.grids[0][1][1].value).toBe("名称");
    expect(host.grids[0][3][2].value).toBe(3);
    expect(host.grids[0][1][2].bold).toBe(true);
    expect(host.grids[0][1][2].fillColor).toBe(0xe4f3df);
    expect(host.grids[0][2][2].fillColor).toBe(before[0][2][2].fillColor);
    await undoWpsExecution(host.application,record.executionId);
    expect(host.grids).toEqual(before);
  });
  it("rejects unavailable header format before writing any data",()=>{
    const host=capabilityHost();const range=host.sheets[0].Range!;
    host.sheets[0].Range=(address)=>{
      const r=range(address);if(address==="B2:C2") r.Item=(row,column)=>{const c=range(`${column===1?"B":"C"}2`);delete c.Interior;return c;};return r;
    };
    expect(()=>prepareWpsPlan(host.application,[1],plan([table()]))).toThrow();
    expect(host.writes()).toBe(0);
  });
  it("rejects unauthorized tables, invalid rectangles and formula injection before writing",()=>{
    const host=capabilityHost(["明细","未选"]);
    for(const candidate of [{...table(),sheet:"未选"},{...table(),rows:[[1]]},{...table(),startCell:"XFD1048576"},{...table(),headers:["=1+1","x"]}]){
      expect(()=>prepareWpsPlan(host.application,[1],plan([candidate]))).toThrow();
      expect(host.writes()).toBe(0);
    }
  });
  it("continues to reject overlapping data writes, including an expanded table",()=>{
    const host=capabilityHost();
    expect(()=>prepareWpsPlan(host.application,[1],plan([table(),{type:"writeValues",sheet:"明细",range:"B2",values:[["overwrite"]]}]))).toThrow("重叠");
    expect(host.writes()).toBe(0);
  });
  it("retains value and format undo for either action order",async()=>{
    for(const reverse of [false,true]){
      const host=capabilityHost(),before=structuredClone(host.grids);
      const actions:AnalysisPlan["actions"]=[{type:"writeValues",sheet:"明细",range:"A1",values:[["after"]]},{type:"setNumberFormat",sheet:"明细",range:"A1",formatCode:"0.00"}];
      if(reverse) actions.reverse();
      const record=await executeWpsPlan(host.application,prepareWpsPlan(host.application,[1],plan(actions)));
      await undoWpsExecution(host.application,record.executionId);expect(host.grids).toEqual(before);
    }
  });
  it("expansion does not modify the submitted plan",()=>{
    const input=plan([table()]),before=structuredClone(input);
    expect(expandWpsPlanActions(input).actions).toHaveLength(3);expect(input).toEqual(before);
  });
});
