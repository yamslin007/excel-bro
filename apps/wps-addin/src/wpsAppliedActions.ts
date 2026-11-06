import { parseExplicitA1Rect } from "../../excel-addin/src/planCore";

/**
 * 「这次执行到底写了哪些动作」的**唯一**判据（WPS 端，纯函数）。
 *
 * 账本游标 `nextActionIndex` 只覆盖"已经拿到完成回执"的动作。真机上还有第二种事实：
 * 第 N 步的宿主写入**已经发生**、但完成回执没落（进程被杀、写后读回失败、调用超时…），
 * 此时 `nextActionIndex` 仍停在 N、`inFlightActionIndex === N`，而账本 `checkpoint`
 * 里留着这一步的**写后指纹**。
 *
 * 撤销与"这次到底写了几格"的核对都必须把这一步算作已写入 —— 只按游标算，撤销就会
 * 漏掉它，然后**把"漏了一格"报成"全部恢复成功"**（2026-11-02 真机：结果2!B2 的 8
 * 留在原处，卡面却说已撤销并回读确认）。
 *
 * 反方向的边界同样重要，而且是**不许猜**的那一侧：只有 `checkpoint` 真的指向同一块
 * 区域（表名 + 矩形相等）、且带着写后指纹时才算证据。拿不到证据时返回 `appliedRows = 0`，
 * 让撤销收尾的"全目标回读"去发现残留 —— 宁可报"这次没能全部恢复"，也不替用户断言
 * 一个没被证明的"没写过"。
 */

export interface WpsAppliedActionEvidence {
  /** 这次动作**确实改动了**的行数；0 = 没有证据表明它写过。 */
  appliedRows: number;
  /** 与"当前现场"比对的写后指纹；`null` = 没有可用的写后证据。 */
  afterFingerprint: string | null;
}

function sameRect(first: string, second: string): boolean {
  const a = parseExplicitA1Rect(first);
  const b = parseExplicitA1Rect(second);
  if (!a || !b) return false;
  return (
    a.startRow === b.startRow &&
    a.startColumn === b.startColumn &&
    a.endRow === b.endRow &&
    a.endColumn === b.endColumn
  );
}

/**
 * 账本 `checkpoint` 里那条写后证据，**只**当它指向 `target` 那块区域时交出来。
 *
 * `checkpoint` 是单条（最后进入写入的那一步），不带 `actionIndex` —— 所以归属只能按
 * `sheetName` + `range` 认，不按"离谁近"猜。表名两边都按真机写法比（大小写敏感：
 * WPS 的表名本来就区分）。
 */
export function wpsCheckpointAfterFingerprint(
  checkpoint: Record<string, unknown> | null | undefined,
  target: { sheetName: string; range: string }
): string | null {
  if (!checkpoint || typeof checkpoint !== "object") return null;
  const sheetName = checkpoint.sheetName;
  const range = checkpoint.range;
  const after = checkpoint.afterFingerprint;
  if (typeof sheetName !== "string" || typeof range !== "string") return null;
  if (typeof after !== "string" || after.trim() === "") return null;
  if (sheetName !== target.sheetName) return null;
  if (!sameRect(range, target.range)) return null;
  return after;
}

/**
 * 单个动作的已写入证据。
 *
 * 两条来源，**互不冒充**：
 * - `actionIndex < nextActionIndex`：完成回执已落，写后指纹按登记的**目标**指纹
 *   （`expectedFingerprint`）比对 —— 这是原来就有的口径，不改。
 * - `actionIndex === inFlightActionIndex` 且 checkpoint 指向它：写过了但回执没落，
 *   写后指纹用 **checkpoint 记下的真实指纹**，不是"应该写到什么"。
 *
 * 两种都不成立 ⇒ `appliedRows = 0`、`afterFingerprint = null`（没有证据，不猜）。
 */
export function wpsAppliedActionEvidence(
  facts: {
    nextActionIndex: number;
    inFlightActionIndex?: number | null;
    /** 见 `wpsCheckpointAfterFingerprint`；只有指向本动作时才非空。 */
    checkpointAfterFingerprint: string | null;
  },
  action: {
    actionIndex: number;
    rowCount: number;
    /**
     * 登记的**目标**指纹（`matrixFingerprintV2(目标矩阵, 目标矩阵)`）。`undefined` = 这个
     * 调用点手上没有登记值（例如给已存在的记录补进度）：那时游标分支不回填指纹，
     * 由调用方保留记录里执行期回读到的真实值。
     */
    expectedFingerprint?: string | null;
  }
): WpsAppliedActionEvidence {
  if (action.actionIndex < facts.nextActionIndex) {
    return { appliedRows: action.rowCount, afterFingerprint: action.expectedFingerprint ?? null };
  }
  const inFlight = facts.inFlightActionIndex ?? null;
  if (inFlight !== null && action.actionIndex === inFlight && facts.checkpointAfterFingerprint) {
    return { appliedRows: action.rowCount, afterFingerprint: facts.checkpointAfterFingerprint };
  }
  return { appliedRows: 0, afterFingerprint: null };
}

/**
 * 撤销收尾要对齐的**全部**目标（用于"回读每个应恢复的区域"）。
 *
 * 不止 `appliedRows > 0` 的那些：一个被判成"没写过"的动作，如果现场其实已经不是写前值，
 * 那正是本模块要防的那种残留。撤销必须能发现它，而不是只看自己恢复过的那几格。
 *
 * 返回**原对象**（不折成精简结构）：调用方还要用 `sheetIndex`/`rowCount`/`columnCount`
 * 去定位与回读，另造一份投影就等于给定位信息开第二个来源。
 */
export function wpsUndoReadbackTargets<
  A extends { type: string; beforeFingerprint: string | null }
>(actions: readonly A[]): A[] {
  return actions.filter(
    (action) =>
      /*
       * 值动作 = 会改单元格内容的那些。`clearRange` 也算：它写的是"空"，
       * 漏掉它等于把"没清干净的表格区"判成已恢复。格式动作不算 —— 它们的恢复
       * 由逐格格式快照自己核对，不在这里按值指纹判。
       */
      (action.type === "writeValues" || action.type === "writeFormulas" || action.type === "clearRange") &&
      typeof action.beforeFingerprint === "string" &&
      action.beforeFingerprint.trim() !== ""
  );
}
