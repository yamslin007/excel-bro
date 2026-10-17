import React from "react";
import WpsProbePanel from "./WpsProbePanel";
import WideViewPocPanel from "./WideViewPocPanel";
import WpsFormatProbePanel from "./WpsFormatProbePanel";
import { useWpsHashRoute } from "./useWpsHashRoute";
import { isDev } from "./isDev";

/**
 * 按当前 hash 选择页面。DEV 门控：诊断路由仅在 DEV 下可达，
 * 非 DEV 即使手改 hash 也一律落到主面板，打不开诊断页。
 * 页面导航本身不触发任何格式探测或工作簿写入（各面板内部有自己的运行前校验与用户确认）。
 *
 * 任务窗格顶部原先还有一条 DEV 诊断导航条（「主面板」/「宽幅审阅 · 能力探针」/
 * 「WPS 格式能力 · 受控探测」三颗按钮）。它在真实宿主里独占一整行、又只是开发诊断入口，
 * 已按产品要求**整条移除**。诊断页本身保留：DEV 下直接改地址栏 hash
 * （`#/wide-view-poc` / `#/dev-format-probe`）仍可进入，非 DEV 依旧打不开。
 */
export function AppRoutes() {
  const { hash } = useWpsHashRoute();

  if (isDev() && hash === "/wide-view-poc") return <WideViewPocPanel />;
  if (isDev() && hash === "/dev-format-probe") return <WpsFormatProbePanel />;
  // 主面板（含未知 hash / 非 DEV 的诊断地址）都兜底到主面板。
  return <WpsProbePanel />;
}
