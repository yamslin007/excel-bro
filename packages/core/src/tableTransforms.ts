/**
 * 表格区域变换（排序 / 去重）的**确定性判据**，宿主无关，Excel / WPS 共用。
 *
 * 为什么必须放在共享层：排序顺序、空白单元格怎么排、重复行保留哪一条，
 * 这些是**用户可见的结果语义**。两端各写一份 `sort` 比较器，迟早会出现
 * "同一个计划在 Excel 和 WPS 上排出来的顺序不一样"，而用户无法从界面分辨
 * 是谁错了。这里只有纯函数：输入矩阵 + 声明，输出行顺序 / 行集合，不碰宿主。
 *
 * 硬性约束（调用方必须自己守住，本模块不替它兜底）：
 *  - 区域必须是矩形、每一行长度一致；
 *  - 关键列索引以**区域左边界**为 0 起算；
 *  - `hasHeaders` 为真时第 0 行不参与排序/去重。
 */

import type { CellValue } from "./contracts";

/** 排序的一个键。`column` 以区域左边界为 0。 */
export type SortKeySpec = { column: number; ascending: boolean };

/** 空白单元格（null / undefined / 空串）在排序里的固定档位：永远排在最后。 */
const BLANK_RANK = 4;

/**
 * 单元格值在排序中的类型档。
 *
 * 与常见表格软件一致：数字 < 文本 < 逻辑值 < 其它 < 空白。
 * **空白永远最后**（不随升/降序翻转）—— 这是用户最容易察觉的行为：
 * 降序时把空行顶到第一屏是错误的。
 */
function sortRank(value: CellValue): number {
  if (value === null || value === undefined || value === "") return BLANK_RANK;
  if (typeof value === "number") return 0;
  if (typeof value === "string") return 1;
  if (typeof value === "boolean") return 2;
  return 3;
}

function isBlank(value: CellValue): boolean {
  return value === null || value === undefined || value === "";
}

/**
 * 比较两个单元格值的先后。返回值 <0 表示 a 在前。
 *
 * 只在**同档**内比较；跨档由 `sortRank` 决定。文本用运行环境的语言规则
 * （`localeCompare`）逐字比较，因此中文按本地排序习惯，而不是码点顺序。
 */
export function compareSortValues(a: CellValue, b: CellValue): number {
  const rankA = sortRank(a);
  const rankB = sortRank(b);
  if (rankA !== rankB) return rankA - rankB;
  if (rankA === BLANK_RANK) return 0;
  if (rankA === 0) return (a as number) - (b as number);
  if (rankA === 2) return Number(a as boolean) - Number(b as boolean);
  if (rankA === 1) return String(a).localeCompare(String(b));
  return 0;
}

/**
 * 计算**稳定排序**后的数据行顺序。
 *
 * 返回数组的第 i 项 = 输出第 i 行取自原来的第几行（相对矩阵首行，已含表头偏移）。
 * 稳定性：所有键都相等时保持原始先后 —— 用户的次级顺序不会被无谓打乱。
 */
export function sortRowOrder(
  values: readonly CellValue[][],
  keys: readonly SortKeySpec[],
  hasHeaders: boolean
): number[] {
  if (!keys.length) throw new Error("排序至少需要一个键。");
  const dataStart = hasHeaders ? 1 : 0;
  const order: number[] = [];
  for (let row = dataStart; row < values.length; row += 1) order.push(row);
  order.sort((left, right) => {
    for (const key of keys) {
      const a = values[left]?.[key.column] ?? null;
      const b = values[right]?.[key.column] ?? null;
      const blankA = isBlank(a);
      const blankB = isBlank(b);
      // 两者都空：交给下一个键，不让"空"本身决定先后。
      if (blankA && blankB) continue;
      // **空白永远最后，与升降序无关** —— 这一条必须在方向翻转之前处理，
      // 否则降序会把空行顶到第一屏。
      if (blankA) return 1;
      if (blankB) return -1;
      const compared = compareSortValues(a, b);
      if (compared !== 0) return key.ascending ? compared : -compared;
    }
    // 完全相等：保持原始先后（`Array.prototype.sort` 在 ES2019 起稳定，
    // 这里仍显式按行号兜底，避免运行环境差异）。
    return left - right;
  });
  return order;
}

/** 按行顺序重排矩阵；表头行原样留在第一行。 */
export function reorderRows(
  values: readonly CellValue[][],
  order: readonly number[],
  hasHeaders: boolean
): CellValue[][] {
  const rows: CellValue[][] = [];
  if (hasHeaders && values.length) rows.push([...values[0]]);
  for (const source of order) rows.push([...(values[source] ?? [])]);
  return rows;
}

/**
 * 排序后应当出现的完整矩阵（表头 + 按序数据行）。
 * 调用方用它做回读比对：**回读到的顺序必须与这份矩阵逐格相等**，
 * 否则不能声称"排序已完成"。
 */
export function sortedMatrix(
  values: readonly CellValue[][],
  keys: readonly SortKeySpec[],
  hasHeaders: boolean
): CellValue[][] {
  return reorderRows(values, sortRowOrder(values, keys, hasHeaders), hasHeaders);
}

/**
 * 去重：按 `columns` 的组合判断重复，**保留首次出现**的那一行。
 *
 * 返回的索引都相对矩阵首行。`kept` 保持原有先后；`duplicates` 也要保持先后，
 * 因为删除行必须从后往前，才能让前面的行号不因删除而失效。
 */
export function duplicateRowPartition(
  values: readonly CellValue[][],
  columns: readonly number[],
  hasHeaders: boolean
): { kept: number[]; duplicates: number[] } {
  if (!columns.length) throw new Error("去重至少需要一列依据。");
  const dataStart = hasHeaders ? 1 : 0;
  const kept: number[] = [];
  const duplicates: number[] = [];
  const seen = new Set<string>();
  for (let row = dataStart; row < values.length; row += 1) {
    const key = JSON.stringify(columns.map((column) => {
      const value = values[row]?.[column] ?? null;
      // `null` / `undefined` / `""` 在去重里是**同一种**：用户看到的就是"空的"。
      // 把空串折成 null，否则 `=IF(A1>0,"","x")` 这类公式产出的空串会和真空格
      // 被当成两行不同数据，去重看起来"没去掉"。
      return value === undefined || value === "" ? null : value;
    }));
    if (seen.has(key)) duplicates.push(row);
    else { seen.add(key); kept.push(row); }
  }
  return { kept, duplicates };
}

/**
 * 去重后应当保留下来的矩阵（表头 + 保留下来的数据行）。
 *
 * 注意：这只是"保留哪些行"的判定，**不代表**已经完成真正的行删除 ——
 * 行删除会移动下方数据，属于结构变更，必须由宿主适配层另行执行与恢复。
 */
export function deduplicatedMatrix(
  values: readonly CellValue[][],
  columns: readonly number[],
  hasHeaders: boolean
): CellValue[][] {
  return reorderRows(values, duplicateRowPartition(values, columns, hasHeaders).kept, hasHeaders);
}
