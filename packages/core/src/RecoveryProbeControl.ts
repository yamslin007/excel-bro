import {createElement as h, useEffect, useState} from "react";
import {armRecoveryProbe, cancelRecoveryProbe, readRecoveryProbe} from "./recoveryProbe";

/**
 * 恢复验收的**开发专用**控制区（只在 DEV 或 `VITE_ENABLE_RECOVERY_PROBE=true` 时渲染）。
 *
 * 里面两件事都必须**确定性**，不能靠人碰时机：
 *
 * 1. 「启用下一次第二步中断」：三动作执行在第 2 步写入回读之后、完成回执之前中断，
 *    第 3 步不执行（判据在 `recoveryProbe.ts`，不适合的计划 fail-closed）。
 * 2. 「模拟本机索引丢失」：把本机挂接记录（`excel-bro.pending-office-execution.v1`）
 *    清掉 —— 换机器、清缓存、换端口之后就是这个样子。**只清本机挂接，不动账本、
 *    不写单元格**；不要求验收者自己去 DevTools 里删 localStorage。
 */
export function RecoveryProbeControl({identity, disabled, onClearIndex}: {
  identity: string | null;
  disabled: boolean;
  /** 清掉本机挂接记录；不传则整个按钮不出现（非验收构建里没有这条出口）。 */
  onClearIndex?: () => void;
}) {
  const [probe, setProbe] = useState(readRecoveryProbe);
  const [error, setError] = useState("");
  const [cleared, setCleared] = useState(false);
  useEffect(() => {
    const refresh = () => setProbe(readRecoveryProbe());
    window.addEventListener("recovery-probe-change", refresh);
    return () => window.removeEventListener("recovery-probe-change", refresh);
  }, []);
  useEffect(() => {
    if (!probe) return;
    const timer = window.setTimeout(() => setProbe(readRecoveryProbe()), Math.max(0, probe.expires - Date.now()));
    return () => window.clearTimeout(timer);
  }, [probe]);
  const act = (fn: () => void) => { try {fn(); setError("");} catch(e) {setError(String(e));} };
  return h("details", {className: "recovery-probe-control"},
    h("summary", null, "恢复验收（开发专用）"),
    h("p", null, "仅在独立测试工作簿中使用。下一次三动作执行（三个独立单格写入：值或公式）将在第二步写入回读后、完成回执前中断。10 分钟内有效。"),
    h("button", {type: "button", disabled: disabled || !identity || probe?.phase === "bound",
      onClick: () => act(() => armRecoveryProbe(true, identity!))}, "启用下一次第二步中断"),
    h("button", {type: "button", disabled: disabled || !probe,
      onClick: () => act(cancelRecoveryProbe)}, "取消验收中断"),
    onClearIndex && h("button", {type: "button", disabled: disabled,
      onClick: () => act(() => {onClearIndex(); setCleared(true);})}, "模拟本机索引丢失"),
    onClearIndex && h("p", null, "只清本机的挂接记录（换机器/清缓存后的样子）；账本与单元格都不动。"),
    h("p", {role: error || probe || cleared ? "status" : undefined},
      error || (cleared ? "已清空本机挂接记录（账本未改）。请重开窗格，让它重新发现。" :
        probe?.identity !== identity ? "当前工作簿未启用中断。" :
        probe?.phase === "armed" ? "已启用：仅作用于当前工作簿的下一次执行。" :
        probe?.phase === "bound" ? "已绑定本次执行，等待第二步回读。" :
        probe?.phase === "not_interrupted" ? "本次执行已结束，验收中断未触发；请重新启用下一次中断。" :
        "验收中断已触发。请重开窗格，核对后续跑第三步；不会再次触发中断。")));
}
