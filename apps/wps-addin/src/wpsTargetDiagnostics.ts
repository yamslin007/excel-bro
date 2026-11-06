import type { TargetDiagnostic } from "@excel-bro/core/targetDiagnostics";
import { reportDevProbeDetail } from "./devProbe";

/**
 * 「这次为什么要求确认范围」的记录端（WPS）。
 *
 * 存在理由（2026-11-01 真机）：提示标题只有几种文案，窗格又开不了 devtools，
 * 于是"证据不足""键冲突""调用方强制确认""旧提示没清掉"在界面上长得一样，
 * 只能靠猜。这里把共享守卫算出来的**布尔值/原因码**留一份。
 *
 * 纪律（不可放宽）：
 * - **只收** `targetDiagnostics.ts` 定义的字段：布尔值、原因码、计数。
 *   文件名、路径、工作表名、身份键、证据键、单元格内容一律不进这里；
 * - 诊断不是判据：写失败、读失败都不许影响授权与执行；
 * - 不进对话历史、不进请求体（可选出口只在 DEV 且显式开关时打开）。
 */

const MAX_EVENTS = 8;
/**
 * 落一份到 localStorage 的**唯一**理由：真机窗格没有 devtools，而这份 profile 的
 * `Local Storage` 在磁盘上可读（`%APPDATA%/kingsoft/wps/addons/data/win-x64/cef/2/jsapi/cache`），
 * 于是"现场布尔值"能在不打断用户的前提下取回来。DEV 与测试构建都不写正式键空间。
 */
const STORAGE_KEY = "excel-bro.wps.target-diag.v1";
const events: TargetDiagnostic[] = [];

function persist(): void {
  if (!import.meta.env.DEV || import.meta.env.MODE === "test") return;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(events));
  } catch {
    // 存不下就只在内存里留着；诊断永远不是判据。
  }
}

/** 记一条。字段由 `TargetDiagnostic` 的类型钉住，多一个也放不进来。 */
export function recordWpsTargetDiagnostic(diagnostic: TargetDiagnostic): void {
  /*
   * 定时巡检每轮都会产出一条"元数据判定通过"。它没有信息量，却会把真正要看的
   * 那条挤出环形缓冲、还会每 30 秒写一次 localStorage。所以只留"被拦下"和
   * "走到分析闸"的判定 —— 序列号仍由守卫单调分配，跳号本身就说明中间有巡检。
   */
  if (diagnostic.outcome === "allowed" && diagnostic.gate === "metadata") return;
  events.push(diagnostic);
  if (events.length > MAX_EVENTS) events.shift();
  if (import.meta.env.MODE !== "test") {
    try {
      console.debug(`[wps-target] ${JSON.stringify(diagnostic)}`);
    } catch {
      // 控制台不可用不影响判定。
    }
  }
  if (import.meta.env.VITE_ENABLE_TARGET_DIAG === "true") {
    try {
      reportDevProbeDetail("taskpane", { target: JSON.stringify(diagnostic) });
    } catch {
      // 发不出去就只在本地留着。
    }
  }
  persist();
}

/** 只读快照（测试与"开发诊断"共用）。 */
export function wpsTargetDiagnosticEvents(): readonly TargetDiagnostic[] {
  return events;
}

/** 一行一条的可读报告；只有原因码与布尔值。 */
export function wpsTargetDiagnosticReport(): string {
  if (!events.length) return "（本轮还没有范围判定）";
  return events.map((event) =>
    `#${event.seq} r${event.round} ${event.outcome}/${event.gate}` +
    `${event.branch ? ` ${event.branch}` : ""}${event.via ? ` via=${event.via}` : ""}` +
    ` conclusive=${event.conclusive} documentBound=${event.documentBound}` +
    ` consent=${event.hasConsent} explicit=${event.explicitRequired}` +
    ` stableId=${event.stableId} path=${event.pathIdentity}`
  ).join("\n");
}

/** 测试与"重开窗格"用：清掉内存与本地副本。 */
export function resetWpsTargetDiagnostics(): void {
  events.length = 0;
  if (!import.meta.env.DEV || import.meta.env.MODE === "test") return;
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch {
    // 同上。
  }
}
