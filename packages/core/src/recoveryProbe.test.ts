import {beforeEach, afterEach, expect, it, vi} from "vitest";
import {armRecoveryProbe, bindRecoveryProbe, interruptRecoveryProbe, settleRecoveryProbe, readRecoveryProbe, RECOVERY_PROBE_KEY} from "./recoveryProbe";
const actions = [1,2,3].map(i => ({type:"writeValues",sheet:"Target",range:`A${i}`,values:[[i]]}));
beforeEach(() => {const m = new Map<string,string>(); vi.stubGlobal("localStorage", {
  getItem:(k:string)=>m.get(k)??null, setItem:(k:string,v:string)=>m.set(k,v), removeItem:(k:string)=>m.delete(k)});});
afterEach(()=>vi.unstubAllGlobals());
it("disabled builds cannot arm, bind or interrupt",()=>{
  expect(()=>armRecoveryProbe(false,"book")).toThrow();
  armRecoveryProbe(true,"book");bindRecoveryProbe(false,"book","op",actions,true);
  expect(readRecoveryProbe()?.phase).toBe("armed");
  bindRecoveryProbe(true,"book","op",actions,true);
  expect(()=>interruptRecoveryProbe(false,"op",1)).not.toThrow();
});
it("one operation, exactly after second readback, consumed before error, reload does not rearm",()=>{
  armRecoveryProbe(true,"book");bindRecoveryProbe(true,"book","op",actions,true);
  interruptRecoveryProbe(true,"other",1);interruptRecoveryProbe(true,"op",0);
  expect(()=>interruptRecoveryProbe(true,"op",1)).toThrow("第二步已写入并回读");
  expect(readRecoveryProbe()?.phase).toBe("interrupted");
  interruptRecoveryProbe(true,"op",1);interruptRecoveryProbe(true,"op",2);
  bindRecoveryProbe(true,"book","new",actions,true);
  expect(readRecoveryProbe()?.operationId).toBe("op");
  expect(localStorage.getItem(RECOVERY_PROBE_KEY)).not.toContain("values");
});
it.each(["identity","recovery","length","overlap","matrix","type"])("rejects %s before executor and consumes arm",kind=>{
  armRecoveryProbe(true,"book");const a=structuredClone(actions);
  if(kind==="length")a.pop();if(kind==="overlap")a[1].range="$A$1:$A$1";
  if(kind==="matrix")a[1].values=[[1,2]];if(kind==="type")a[1].type="setFormula";
  expect(()=>bindRecoveryProbe(true,kind==="identity"?"other":"book","op",a,kind!=="recovery")).toThrow();
  expect(readRecoveryProbe()).toBeNull();
});
/*
 * 公式计划也要能走这条确定性中断（2026-10-30）：原来只认 `writeValues`，于是
 * "第二步是公式写入"的计划在绑定时就被拒 —— 公式续跑的实机验收根本没有确定性入口。
 * 放宽的只是**验收工具**认定"一格一次写"的形状（值或公式，都是 1×1），
 * 条数、互异单格、可恢复资格、工作簿身份一条没动。
 */
it("formula plans bind: three independent single-cell formula writes on three cells",()=>{
  const formulas = [1,2,3].map(i=>({type:"writeFormulas",sheet:`T${i}`,range:"B2",formulas:[[`=S!B${i}*1`]]}));
  armRecoveryProbe(true,"book");bindRecoveryProbe(true,"book","op",formulas,true);
  expect(readRecoveryProbe()?.phase).toBe("bound");
  expect(()=>interruptRecoveryProbe(true,"op",1)).toThrow("第二步已写入并回读");
  expect(readRecoveryProbe()?.phase).toBe("interrupted");
});
it("mixed value/formula steps bind too; the interrupt still lands on step 2 only",()=>{
  const mixed = [
    {type:"writeValues",sheet:"T1",range:"B2",values:[[1]]},
    {type:"writeFormulas",sheet:"T2",range:"B2",formulas:[["=S!B2*2"]]},
    {type:"writeValues",sheet:"T3",range:"B2",values:[[3]]}];
  armRecoveryProbe(true,"book");bindRecoveryProbe(true,"book","op",mixed,true);
  interruptRecoveryProbe(true,"op",0);interruptRecoveryProbe(true,"op",2);
  expect(readRecoveryProbe()?.phase).toBe("bound");
  expect(()=>interruptRecoveryProbe(true,"op",1)).toThrow("第二步已写入并回读");
  expect(readRecoveryProbe()?.phase).toBe("interrupted");
});
it.each(["wide","no-matrix","r1c1-only"])("formula actions that cannot prove one-cell-one-write are rejected: %s",kind=>{
  type ProbeAction = {type: string; sheet: string; range?: string; formulas?: string[][]; formulaR1C1?: string};
  const ok: ProbeAction[] = [1,2,3].map(i=>({type:"writeFormulas",sheet:`T${i}`,range:"B2",formulas:[[`=S!B${i}`]]}));
  const broken = structuredClone(ok);
  if(kind==="wide") broken[1].formulas=[["=S!B1","=S!B2"]];
  if(kind==="no-matrix") delete broken[1].formulas;
  if(kind==="r1c1-only") {delete broken[1].formulas;broken[1].formulaR1C1="=RC[-1]*2";}
  armRecoveryProbe(true,"book");
  expect(()=>bindRecoveryProbe(true,"book","op",broken,true)).toThrow("恢复验收已取消");
  expect(readRecoveryProbe()).toBeNull();
});
it("expired arms do not affect later execution",()=>{
  armRecoveryProbe(true,"book");vi.spyOn(Date,"now").mockReturnValue(Date.now()+600001);
  try {bindRecoveryProbe(true,"book","op",actions,true);expect(readRecoveryProbe()).toBeNull();}
  finally {vi.restoreAllMocks();}
});
it('preserves the recovery refusal reason and consumes the probe', () => {
  armRecoveryProbe(true, 'book');
  expect(() => bindRecoveryProbe(true, 'book', 'op', actions, false, '来源读取失败')).toThrow('来源读取失败；未执行写入');
  expect(readRecoveryProbe()).toBeNull();
});
it("a finished attempt cannot keep the control bound to an old execution",()=>{
  armRecoveryProbe(true,"book");bindRecoveryProbe(true,"book","old-op",actions,true);
  settleRecoveryProbe(false,"old-op");settleRecoveryProbe(true,"another-op");
  expect(readRecoveryProbe()?.phase).toBe("bound");
  settleRecoveryProbe(true,"old-op");
  expect(readRecoveryProbe()?.phase).toBe("not_interrupted");
  // Finishing an unsuccessful probe does not arm a new write implicitly.
  bindRecoveryProbe(true,"book","next-op",actions,true);
  expect(readRecoveryProbe()?.operationId).toBe("old-op");
  armRecoveryProbe(true,"book");bindRecoveryProbe(true,"book","new-op",actions,true);
  expect(()=>interruptRecoveryProbe(true,"new-op",1)).toThrow("第二步已写入并回读");
  settleRecoveryProbe(true,"new-op");
  expect(readRecoveryProbe()?.phase).toBe("interrupted");
});
