import type { CellValue } from "../../excel-addin/src/contracts";
import type { WpsRange } from "./wpsHost";
import { countHostWrite } from "./wpsHostWriteCounter";

/** Preserve literal strings during WPS coercion, then restore the original formats.
 * All callers must still verify the exact values/formulas after this operation.
 * Older hosts without cell-format access retain the existing strict readback path.
 *
 * 计数口径见 `wpsHostWriteCounter`：真正赋值给 `range[property]` 的那一次记为
 * **计划写**（`mutating`），前后为了保持文本原样而改 `NumberFormat` 的记为
 * **同动作内的辅助调用**（只进 `hostWriteCalls`）。撤销/审阅副本调用本函数时
 * 没有活动作用域 ⇒ 一次都不计，不会被当成"重放"。
 */
export function writeWpsMatrix(
  range: WpsRange,
  values: CellValue[][],
  property: "Value2" | "Formula" = "Value2"
): void {
  const formats: { cell: WpsRange; format: string }[] = [];
  const single = values.length === 1 && values[0].length === 1;
  // Read every format before changing any cell, so an unavailable getter cannot
  // leave a half-prepared range. Real formulas must retain their original format.
  values.forEach((row, r) => row.forEach((value, c) => {
    if (typeof value !== "string" || (property === "Formula" && value.startsWith("="))) return;
    const cell = single ? range : range.Item?.(r + 1, c + 1);
    if (!cell) return;
    const format = cell.NumberFormat;
    if (typeof format === "string" && format !== "@") formats.push({ cell, format });
  }));
  const touched: typeof formats = [];
  try {
    for (const item of formats) {
      touched.push(item);
      countHostWrite(false);
      item.cell.NumberFormat = "@";
    }
    countHostWrite(true);
    range[property] = single ? values[0][0] : values;
  } finally {
    const failures: unknown[] = [];
    for (const item of touched.reverse()) {
      try {
        countHostWrite(false);
        item.cell.NumberFormat = item.format;
        if (item.cell.NumberFormat !== item.format) throw new Error("WPS 数字格式恢复回读不一致");
      } catch (error) { failures.push(error); }
    }
    if (failures.length) {
      throw new AggregateError(failures, "WPS 未能恢复写入区域的原数字格式，不能确认写入成功。");
    }
  }
}
