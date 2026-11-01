/**
 * Excel 侧的「重新发现未完成执行」。
 *
 * 背景（2026-10-29）：WPS 端在 2026-10-27 就有了后端发现（`wpsRecoveryDiscovery.ts`），
 * Excel 端一直只有**一条** `localStorage` 索引（`officeOperationRecovery.ts`）把用户带回断点。
 * 索引不在（换机器 / 清缓存 / 窗格重开 / 换端口）时，账本里的进度与写了一半的动作都还在，
 * 面板却再也找不到这次执行 —— 与 WPS 那个现场同源。
 *
 * 这里只负责"找到入口"：身份在本机读，发现走后端（`GET /api/office/executions/pending`，
 * 按 `url:` 完整身份精确查找）。**不**负责"恢复写入资格"——身份/授权/现场核对由
 * `pendingResumeReadiness`（共享判据）与面板既有续跑路径做，本模块不复制任何恢复判据。
 *
 * 三条不许猜的边界（与 WPS 端逐条同义，判据来自同一个共享模块）：
 *
 * - **身份不可读就不发请求**：未保存簿（Office.js 的 `document.url` 为空）一律 `no_identity`，
 *   不许退回文件名或"最近一条"；未保存簿的会话内对象身份（`unsaved:`）也不用于发现 ——
 *   账本登记的永远是 `url:`。
 * - **服务不可达不是"没有"**：返回 `unavailable`，界面必须说"暂时读不到"。
 * - **不自动选**：选择规则在共享层（`pendingExecutionSelection`），多于一条或结果被截断时
 *   交给用户。
 */

import {
  PENDING_DISCOVERY_UNAVAILABLE_DETAIL,
  assertPendingExecutionDiscovery,
  pendingExecutionSelection,
  type PendingExecutionDiscoveryResponse,
  type PendingExecutionSelection
} from "@excel-bro/core/pendingExecutionDiscovery";
import { listPendingOfficeExecutions } from "./api";
import { currentOfficeWorkbookIdentity } from "./excel";

/**
 * 当前活动工作簿的**已保存**身份（`url:` + 规范化磁盘/站点路径）。
 *
 * 与登记、状态核对用的是同一个函数：发现如果自己拼一份"差不多的身份"，就会出现
 * "登记时认得出、发现时认不出"这种只在真机上才发现的错位。
 *
 * 未保存 → `null`（`currentOfficeWorkbookIdentity` 在 `document.url` 为空时返回
 * `undefined`）。与 `url:` 命名空间不同，`unsaved:` 令牌只在本次加载内有效，
 * 拿它去查账本永远是"查不到"，会被读成"没有未完成的任务"。
 */
export function currentOfficeRecoveryIdentity(): string | null {
  try {
    const identity = currentOfficeWorkbookIdentity();
    if (!identity) return null;
    // 账本只登记 `url:`；别的命名空间不属于发现范围，宁可当作读不到身份。
    return identity.startsWith("url:") ? identity : null;
  } catch {
    return null;
  }
}

export type OfficeRecoveryDiscoveryOutcome =
  /** 没有可用身份：未保存或宿主读失败 —— 一个请求都不发。 */
  | { kind: "no_identity" }
  /** 查不到（服务不可达、身份被拒、响应不合规）。**不是**"没有未完成任务"。 */
  | { kind: "unavailable"; detail: string; reason: string }
  | { kind: "resolved"; selection: PendingExecutionSelection };

function reasonOf(error: unknown, maximum = 200): string {
  const text = error instanceof Error ? error.message : String(error ?? "");
  const trimmed = text.trim();
  return trimmed.length > maximum ? `${trimmed.slice(0, maximum)}…` : trimmed;
}

export async function discoverOfficePendingExecutions(
  options: {
    identity?: () => string | null;
    load?: (workbookIdentity: string) => Promise<PendingExecutionDiscoveryResponse>;
  } = {}
): Promise<OfficeRecoveryDiscoveryOutcome> {
  const readIdentity = options.identity ?? currentOfficeRecoveryIdentity;
  const identity = readIdentity();
  if (!identity) return { kind: "no_identity" };
  const load = options.load ?? listPendingOfficeExecutions;
  try {
    const response = await load(identity);
    /*
     * 再验一次形状：`api.ts` 那条路已经验过，但注入的 `load`（测试、将来的其它传输）
     * 未必。发现结果决定界面会不会给出"可以核对"的入口，宽松收下的代价是把泄漏或
     * 畸形数据当成正常现场。
     */
    assertPendingExecutionDiscovery(response);
    if (response.namespace !== "office") {
      // WPS 的候选不得跨命名空间出现在 Excel 面板里（同一个身份串前缀不同）。
      return {
        kind: "unavailable",
        detail: "服务端返回的不是 Office 写入的记录，不据此恢复；请在原宿主中核对。",
        reason: `namespace=${response.namespace}`
      };
    }
    return { kind: "resolved", selection: pendingExecutionSelection(response, identity) };
  } catch (error) {
    return {
      kind: "unavailable",
      detail: PENDING_DISCOVERY_UNAVAILABLE_DETAIL,
      reason: reasonOf(error)
    };
  }
}
