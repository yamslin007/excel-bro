import {act} from "react";
import {createRoot} from "react-dom/client";
import {afterEach, expect, it, vi} from "vitest";
import {RecoveryProbeControl} from "@excel-bro/core/RecoveryProbeControl";
import {armRecoveryProbe, bindRecoveryProbe} from "@excel-bro/core/recoveryProbe";

(globalThis as {IS_REACT_ACT_ENVIRONMENT?: boolean}).IS_REACT_ACT_ENVIRONMENT = true;
afterEach(()=>{vi.useRealTimers();localStorage.clear();});

it("an expired old binding stops claiming it is waiting for readback without another click", async()=>{
  vi.useFakeTimers();
  armRecoveryProbe(true,"book");
  bindRecoveryProbe(true,"book","old-operation",[1,2,3].map(i=>({type:"writeValues",sheet:"Target",range:`B${i}`,values:[[i]]})),true);
  const container=document.createElement("div");
  const root=createRoot(container);
  try {
    await act(async()=>root.render(<RecoveryProbeControl identity="book" disabled={false}/>));
    expect(container.textContent).toContain("已绑定本次执行，等待第二步回读");
    expect(container.querySelector("button")?.disabled).toBe(true);
    await act(async()=>{vi.advanceTimersByTime(600_001);});
    expect(container.textContent).not.toContain("等待第二步回读");
    expect(container.textContent).toContain("当前工作簿未启用中断");
    expect(container.querySelector("button")?.disabled).toBe(false);
  } finally {await act(async()=>root.unmount());}
});
