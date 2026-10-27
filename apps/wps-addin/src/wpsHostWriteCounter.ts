/**
 * WPS 端「宿主写 API 调用次数」的计数钩子（A1-E3 第三批）。
 *
 * 为什么需要它：`actions/start` 的行数只是**申请**写动作的次数；认领成功也可能在
 * 碰到宿主之前就失败，一个动作还可能分块调用多次。要回答"到底向宿主发起了几次写
 * 调用"，只能在**真正赋值给宿主对象的那一行**上数 —— 也就是本模块唯一的用途。
 *
 * 作用域（scope）是这里的核心：计数只在 `executeWpsPlan` 的**动作窗口**内开启。
 * 于是下面这些**辅助写**天然不会被算成"计划写调用"：
 *
 * - 审阅副本的搭建与比对（`wpsNativeComparison`，不同工作簿）；
 * - 撤销 / 恢复（`undoWpsExecution`、格式回滚）；
 * - 审查标记的清写；
 * - 恢复点快照写入（`PluginStorage`，根本不是宿主写）。
 *
 * 两个计数不能混：
 * - `calls`：进入宿主写 API 的次数（`hostWriteCalls`）。**含**同一动作内为了让
 *   文本保持原样而临时改 `NumberFormat` 这类辅助调用 —— 它们确实也写了宿主；
 * - `mutating`：真正承载**计划内容**的调用（值/公式写入、格式动作声明的属性）。
 *
 * 没有活动作用域时 `countHostWrite` 是空操作 —— 这就是"辅助写不混计"的实现方式，
 * 不需要每个调用点再判断"我现在是不是在撤销"。
 */

export interface WpsHostWriteTotals {
  /** 进入宿主写 API 的次数（含同动作内的辅助调用）。 */
  calls: number;
  /** 承载计划内容的那部分调用次数。 */
  mutating: number;
}

interface ActiveScope {
  calls: number;
  mutating: number;
  previous: ActiveScope | null;
}

let active: ActiveScope | null = null;

/** 记一次真实写调用。只在作用域内生效。 */
export function countHostWrite(mutating = true): void {
  if (!active) return;
  active.calls += 1;
  if (mutating) active.mutating += 1;
}

/**
 * 开启一个计数窗口。嵌套时内层独立计数、退出后恢复外层（当前没有嵌套调用点，
 * 但恢复比"假设不会嵌套"安全）。
 */
export function beginHostWriteScope(): ActiveScope {
  const scope: ActiveScope = { calls: 0, mutating: 0, previous: active };
  active = scope;
  return scope;
}

export function endHostWriteScope(scope: ActiveScope): WpsHostWriteTotals {
  if (active !== scope) {
    // 作用域被交叉关闭：不把数字记到别人头上，如实返回 0。
    return { calls: 0, mutating: 0 };
  }
  active = scope.previous;
  return { calls: scope.calls, mutating: scope.mutating };
}

/** 当前有没有在计数（诊断/测试用；业务判断不要读它）。 */
export function hostWriteScopeActive(): boolean {
  return active !== null;
}
