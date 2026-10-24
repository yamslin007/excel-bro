import {beforeEach, afterEach, expect, it, vi} from "vitest";
import {armRecoveryProbe, bindRecoveryProbe, interruptRecoveryProbe, readRecoveryProbe, RECOVERY_PROBE_KEY} from "./recoveryProbe";
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
it("expired arms do not affect later execution",()=>{
  armRecoveryProbe(true,"book");vi.spyOn(Date,"now").mockReturnValue(Date.now()+600001);
  try {bindRecoveryProbe(true,"book","op",actions,true);expect(readRecoveryProbe()).toBeNull();}
  finally {vi.restoreAllMocks();}
});
