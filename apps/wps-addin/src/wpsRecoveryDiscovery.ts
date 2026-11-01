/**
 * WPS 侧的"重新发现未完成执行"。
 *
 * 2026-10-27 真实现场：第二步写入并回读后精确中断，用户保存、完整退出 WPS、再打开同一个
 * 工作簿。账本里 `needs_reconciliation / nextActionIndex=1 / inFlightActionIndex=1`，原表
 * 仍是 111、222、-303 —— 进度与现场都在，**恢复入口没了**：浏览器 `localStorage` 的
 * 单条待恢复索引是唯一入口，完整退出后它不在了。
 *
 * 这里把那个索引降级为缓存：宿主身份在本机读，发现走后端（按身份精确查找）。只负责
 * "找到入口"，不负责"恢复写入资格" —— 认领、材料水合、身份/授权/来源/现场复核仍由
 * 面板既有的那一条路径做，本模块不复制任何恢复判据。
 *
 * 三条不许猜的边界：
 *
 * - **身份不可读就不发请求**：未保存簿、非 `.xlsx`、宿主读失败一律 `no_identity`，
 *   不许退回文件名或"最近一条"；
 * - **服务不可达不是"没有"**：返回 `unavailable`，界面必须说"暂时读不到"，
 *   不许显示成"没有未完成任务"；
 * - **不自动选**：选择规则在共享层（`pendingExecutionSelection`），多于一条或结果被
 *   截断时交给用户。
 */

import {
  PENDING_DISCOVERY_UNAVAILABLE_DETAIL,
  assertPendingExecutionDiscovery,
  pendingExecutionSelection,
  type PendingExecutionDiscoveryResponse,
  type PendingExecutionSelection
} from "@excel-bro/core/pendingExecutionDiscovery";
import { listPendingWpsExecutions } from "../../excel-addin/src/api";
import type { WpsApplication } from "./wpsHost";
import { wpsReadWorkbook } from "./wpsResultReviewTarget";
import { wpsWorkbookIdentity } from "./wpsExecutionLedger";

/**
 * 当前活动工作簿的稳定身份（`wps-path:` + 小写磁盘完整路径）。
 *
 * 与登记、状态核对用的是**同一个**函数：发现如果自己拼一份"差不多的身份"，就会出现
 * "登记时认得出、发现时认不出"这种只在真机上才发现的错位。
 */
export function currentWpsWorkbookIdentity(application: WpsApplication): string | null {
  try {
    return wpsWorkbookIdentity(wpsReadWorkbook(application)?.FullName);
  } catch {
    return null;
  }
}

export type WpsRecoveryDiscoveryOutcome =
  /** 没有可用身份：未保存、非 `.xlsx` 或宿主读失败 —— 一个请求都不发。 */
  | { kind: "no_identity" }
  /** 查不到（服务不可达、身份被拒、响应不合规）。**不是**"没有未完成任务"。 */
  | { kind: "unavailable"; detail: string; reason: string }
  | { kind: "resolved"; selection: PendingExecutionSelection };

function reasonOf(error: unknown, maximum = 200): string {
  const text = error instanceof Error ? error.message : String(error ?? "");
  const trimmed = text.trim();
  return trimmed.length > maximum ? `${trimmed.slice(0, maximum)}…` : trimmed;
}

export async function discoverWpsPendingExecutions(
  application: WpsApplication,
  options: { load?: (workbookIdentity: string) => Promise<PendingExecutionDiscoveryResponse> } = {}
): Promise<WpsRecoveryDiscoveryOutcome> {
  const identity = currentWpsWorkbookIdentity(application);
  if (!identity) return { kind: "no_identity" };
  const load = options.load ?? listPendingWpsExecutions;
  try {
    const response = await load(identity);
    /*
     * 再验一次形状：`api.ts` 那条路已经验过，但注入的 `load`（测试、将来的其它传输）
     * 未必。发现结果决定了界面会不会给出"可以核对"的入口，这里宽松收下的代价是把
     * 泄漏或畸形数据当成正常现场。
     */
    assertPendingExecutionDiscovery(response);
    return { kind: "resolved", selection: pendingExecutionSelection(response, identity) };
  } catch (error) {
    return { kind: "unavailable", detail: PENDING_DISCOVERY_UNAVAILABLE_DETAIL, reason: reasonOf(error) };
  }
}
