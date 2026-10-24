import {createElement as h, useEffect, useState} from "react";
import {armRecoveryProbe, cancelRecoveryProbe, readRecoveryProbe} from "./recoveryProbe";
export function RecoveryProbeControl({identity, disabled}: {identity: string | null; disabled: boolean}) {
  const [probe, setProbe] = useState(readRecoveryProbe);
  const [error, setError] = useState("");
  useEffect(() => {
    const refresh = () => setProbe(readRecoveryProbe());
    window.addEventListener("recovery-probe-change", refresh);
    return () => window.removeEventListener("recovery-probe-change", refresh);
  }, []);
  const act = (fn: () => void) => { try {fn(); setError("");} catch(e) {setError(String(e));} };
  return h("details", {className: "recovery-probe-control"},
    h("summary", null, "恢复验收（开发专用）"),
    h("p", null, "仅在独立测试工作簿中使用。下一次三动作执行将在第二步写入回读后、完成回执前中断。10 分钟内有效。"),
    h("button", {type: "button", disabled: disabled || !identity || probe?.phase === "bound",
      onClick: () => act(() => armRecoveryProbe(true, identity!))}, "启用下一次第二步中断"),
    h("button", {type: "button", disabled: disabled || !probe,
      onClick: () => act(cancelRecoveryProbe)}, "取消验收中断"),
    h("p", {role: error || probe ? "status" : undefined}, error || (probe?.identity !== identity ? "当前工作簿未启用中断。" :
      probe?.phase === "armed" ? "已启用：仅作用于当前工作簿的下一次执行。" :
      probe?.phase === "bound" ? "已绑定本次执行，等待第二步回读。" :
      "验收中断已触发。请重开窗格，核对后续跑第三步；不会再次触发中断。")));
}
