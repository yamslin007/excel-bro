/**
 * 「备份与恢复」抽屉**空态**该说什么（纯呈现判据，Excel / WPS 共用）。
 *
 * 为什么要有这一层：抽屉里原本只有一句
 * 「还没有恢复点 · 只有确认并完成写入后才会创建，不会生成空备份。」
 * 这句话把生命周期说**反**了 —— 写前快照是在**第一次写入之前**保存的（那正是它能
 * 撤销的前提），不是"确认变更之后才创建"。用户读到它，会以为在确认之前没有保护。
 *
 * 而且空态并不只有一种意思。至少两种真实情形必须分开：
 *
 * 1. `none` —— 当前工作簿**没有待撤销的写入**（还没写过，或更早的恢复点已过期/被用掉）。
 * 2. `undone` —— **本次执行刚刚被撤销**，它的恢复点已被用掉并回收。此时说"还没有恢复点"
 *    会让人以为"这次根本没被保护过"，恰恰相反：正是因为它有快照才撤得回来。
 *
 * 判据只认调用方能证明的事实：有没有可选中的恢复点、最近一次执行的处置是不是已撤销。
 * 拿不到就退回 `none`，**不编第三种状态**，也不把"读不到"说成"没有"（那是调用方的事，
 * 本模块不替它下结论）。
 */

/** 抽屉空态的两种真实含义。 */
export type BackupEmptyKind = "none" | "undone";

export interface BackupEmptyFacts {
  /** 抽屉里是否已经有一条可选中的恢复点（有就根本不显示空态）。 */
  hasUndoRecord: boolean;
  /**
   * 最近一次**与写入有关**的执行，其处置是不是"已撤销"。
   *
   * 只认消息/结果卡上那条权威留痕（`verification.status === "undone"`）——
   * 不要在调用方按"点了撤销按钮"猜，那会在失败时也显示成已撤销。
   */
  latestExecutionUndone?: boolean;
}

export interface BackupEmptyView {
  kind: BackupEmptyKind;
  title: string;
  hint: string;
}

/**
 * 空态文案。
 *
 * 返回 `null` = **不该显示空态**（已经有恢复点了）。
 */
export function backupEmptyState(facts: BackupEmptyFacts): BackupEmptyView | null {
  if (facts.hasUndoRecord) return null;
  if (facts.latestExecutionUndone) {
    return {
      kind: "undone",
      title: "本次执行已撤销",
      hint: "写前快照在第一次写入之前就已保存；撤销完成后它被用掉并回收，所以这里不再显示。"
    };
  }
  return {
    kind: "none",
    title: "暂无恢复点",
    hint: "每次写入前都会先保存完整快照；这里为空表示当前没有待撤销的写入。"
  };
}
