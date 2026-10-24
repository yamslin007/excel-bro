import { describe, it, expect } from "vitest";
import type { AnalysisPlan, ExcelAction } from "../../excel-addin/src/contracts";
import { capabilityHost } from "./wpsCapabilityTestHost";
import { prepareWpsPlan, executeWpsPlan, undoWpsExecution, latestWpsUndo } from "./wpsPlanExecutor";
import { wpsRecoveryDescriptor } from "./wpsExecutionLedger";
const plan = (actions: ExcelAction[]): AnalysisPlan => ({ id: "compute", title: "汇总", summary: "本地汇总", assumptions: [], warnings: [], acceptanceCriteria: [], actions });
const action: Extract<ExcelAction, {type: "computeValues"}> = { type: "computeValues", sheet: "结果", sourceSheet: "来源", sourceRange: "A1:B4", hasHeaders: true,
  query: { mode: "aggregate", groupBy: ["类别"], metrics: [{ field: "金额", operation: "sum", outputName: "合计" }] }, destinationCell: "C2", includeHeaders: true };
function fixture() {
  const h = capabilityHost(["来源", "结果", "私有"]);
  [["类别", "金额"], ["甲", 2], ["乙", 5], ["甲", 3]].forEach((row, r) => row.forEach((v, c) => { h.grids[0][r][c].value = v; h.grids[0][r][c].formula = v; }));
  return h;
}
describe("WPS computed result execution", () => {
  it("computes locally, restores mixed destination formulas and keeps the dynamic tool", async () => {
    const h = fixture(); h.grids[1][1][2].formula = "=1+1";
    const before = structuredClone(h.grids);
    const prepared = prepareWpsPlan(h.application, [1, 2], plan([action]));
    expect(prepared.originalPlan?.actions).toEqual([action]); expect(h.writes()).toBe(0);
    expect(prepared.plan.actions).toEqual([{type: "writeValues", sheet: "结果", range: "C2:D4", values: [["类别", "合计"], ["甲",5], ["乙",5]]},
      {type:"setFont",sheet:"结果",range:"C2:D2",bold:true},{type:"setFill",sheet:"结果",range:"C2:D2",color:"#DFF3E4"}]);
    expect(JSON.stringify(wpsRecoveryDescriptor(prepared))).not.toContain("合计");
    const record = await executeWpsPlan(h.application, prepared);
    expect(h.grids[1][1][2].bold).toBe(true); expect(h.grids[1][1][2].fillColor).toBe(0xE4F3DF);
    expect(h.grids[1][2][3].value).toBe(5); expect(h.grids[0]).toEqual(before[0]); expect(h.grids[2]).toEqual(before[2]);
    await undoWpsExecution(h.application, record.executionId);
    expect(h.grids[1][1][2].formula).toBe("=1+1");
    before[1][1][2].value = "=1+1"; // This host records formulas without a calculation engine.
    expect(h.grids).toEqual(before);
    h.grids[0][1][1].value = 20; h.grids[0][1][1].formula = 20;
    const fresh = prepareWpsPlan(h.application, [1, 2], prepared.originalPlan!);
    expect(fresh.actions[0].values[1][1]).toBe(23);
  });
  it("checks source and destination authorization before any source read", () => {
    const h = fixture(); let reads = 0; const original = h.sheets[0].Range!;
    h.sheets[0].Range = address => { reads++; return original(address); };
    expect(() => prepareWpsPlan(h.application, [2], plan([action]))).toThrow("未授权");
    expect(() => prepareWpsPlan(h.application, [1], plan([action]))).toThrow("未授权");
    expect(reads).toBe(0); expect(h.writes()).toBe(0);
  });
  it.each(["value", "formula", "numberFormat"] as const)("changed source %s prevents every write", async key => {
    const h = fixture(), prepared = prepareWpsPlan(h.application, [1,2], plan([action]));
    h.grids[0][1][1][key] = key === "numberFormat" ? "yyyy-mm-dd" : key === "formula" ? "=99" : "changed";
    await expect(executeWpsPlan(h.application, prepared)).rejects.toThrow(/来源|源区域/); expect(h.writes()).toBe(0);
  });
  it("supports headerless sources without dropping their first data row", () => {
    const h = fixture(); const prepared = prepareWpsPlan(h.application, [1,2], plan([{...action, sourceRange: "B2:B4", hasHeaders:false, includeHeaders:false,
      query:{ mode:"aggregate", metrics:[{field:"列1",operation:"sum",outputName:"合计"}] }}]));
    expect(prepared.actions[0].values).toEqual([[10]]);
  });
  it("rejects unavailable fields, unknown ranges, oversized outputs and in-plan stale inputs without writes", () => {
    const h = fixture();
    for (const patch of [{sourceRange:null},{destinationCell:"A1:B2"},{destinationCell:"XFD1048576"},{query:{mode:"rows",fields:["不存在"]}}]) {
      expect(() => prepareWpsPlan(h.application,[1,2],plan([{...action,...patch} as typeof action]))).toThrow();
    }
    expect(() => prepareWpsPlan(h.application,[1,2],plan([{type:"writeValues",sheet:"来源",range:"B2",values:[[99]]},action]))).toThrow("其他修改前");
    expect(h.writes()).toBe(0);
  });
  it("source changes while the durable checkpoint is in flight still cause zero writes", async () => {
    const h = fixture(), prepared = prepareWpsPlan(h.application,[1,2],plan([action]));
    await expect(executeWpsPlan(h.application, prepared, undefined, undefined, async phase => {
      if (phase === "before") { h.grids[0][1][1].value = 99; h.grids[0][1][1].formula = 99; }
    })).rejects.toThrow("源区域");
    expect(h.writes()).toBe(0);
  });
  it.each(["contents", "readOnly", "protection"])("rechecks target %s after awaiting the checkpoint", async kind => {
    const h = fixture(), prepared = prepareWpsPlan(h.application,[1,2],plan([action]));
    await expect(executeWpsPlan(h.application, prepared, undefined, undefined, async phase => {
      if (phase !== "before") return;
      if (kind === "contents") { h.grids[1][1][2].value = "manual"; h.grids[1][1][2].formula = "manual"; }
      if (kind === "readOnly") h.application.ActiveWorkbook!.ReadOnly = true;
      if (kind === "protection") h.sheets[1].ProtectContents = true;
    })).rejects.toThrow(/目标|只读|编辑|保护/);
    expect(h.writes()).toBe(0);
    if (kind === "contents") expect(h.grids[1][1][2].value).toBe("manual");
    expect(latestWpsUndo(h.application)).toBeNull();
  });
  it("rejects a truncated re-read even when the absent cell was blank", async () => {
    const h = fixture(); h.grids[0][3][1].value = null; h.grids[0][3][1].formula = null;
    const prepared = prepareWpsPlan(h.application,[1,2],plan([action]));
    const original = h.sheets[0].Range!;
    h.sheets[0].Range = address => address === "A1:B4" ? {...original(address), Value2:[["类别","金额"],["甲",2],["乙",5],["甲"]]} : original(address);
    await expect(executeWpsPlan(h.application, prepared)).rejects.toThrow("完整重读");
    expect(h.writes()).toBe(0);
  });
  it("does not adopt external edits of an earlier completed action into its recovery", async()=>{
    const h=fixture();
    const prepared=prepareWpsPlan(h.application,[2],plan([{type:"writeValues",sheet:"结果",range:"A1",values:[[1]]},{type:"writeValues",sheet:"结果",range:"A2",values:[[2]]}]));
    await expect(executeWpsPlan(h.application,prepared,undefined,undefined,async(phase,checkpoint)=>{
      if(phase==="before"&&checkpoint.actionIndex===1){h.grids[1][0][0].value="manual";h.grids[1][0][0].formula="manual";}
    })).rejects.toThrow("外部变化");
    const recovery=latestWpsUndo(h.application)!;expect(recovery).not.toBeNull();
    const writes=h.writes();await expect(undoWpsExecution(h.application,recovery.executionId)).rejects.toThrow();
    expect(h.writes()).toBe(writes);expect(h.grids[1][0][0].value).toBe("manual");expect(h.grids[1][1][0].value).toBe("old-1-0");
  });
  it("never invents blanks for a truncated source matrix", () => {
    const h = fixture(), range = h.sheets[0].Range!;
    h.sheets[0].Range = address => ({...range(address),Value2:[["partial"]]});
    expect(() => prepareWpsPlan(h.application,[1,2],plan([action]))).toThrow("矩阵不完整"); expect(h.writes()).toBe(0);
  });

});
