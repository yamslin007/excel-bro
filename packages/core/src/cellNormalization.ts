import type { CellValue } from "./contracts";

/**
 * 单元格值归一（**两端同源**）。
 *
 * 隔离 Python 拿到的是"值"，不是"屏幕上的样子"，所以日期序列值必须还原成日期、
 * 前导零编号必须保留字符串形态 —— 否则 `001` 会变成 `1`、`2027/7/4` 会变成
 * `46572`，模型算出来的东西和用户看到的不是一回事。
 *
 * 这份实现**只此一份**：Excel 与 WPS 两侧都从这里取，不许各写一份"看起来一样"
 * 的副本。
 */
const DATE_TOKENS = /(^|[^\\])[ymdhis]/i;
const TEXT_CODE_FORMAT = /^0+$/;

function excelSerialDate(serial: number): string | null {
  if (!Number.isFinite(serial) || serial < 0 || serial > 2_958_465) return null;
  const milliseconds = Math.round((serial - 25_569) * 86_400_000);
  const date = new Date(milliseconds);
  if (Number.isNaN(date.getTime())) return null;
  return date.toISOString().slice(0, 10);
}

export function normalizeCellValue(
  value: unknown,
  displayText = "",
  numberFormat = "General"
): CellValue {
  if (value === null || value === undefined || value === "") return value === "" ? "" : null;
  if (typeof value === "boolean" || typeof value === "string") return value;
  if (typeof value !== "number") return String(value);

  const format = numberFormat.trim();
  if (DATE_TOKENS.test(format)) {
    return excelSerialDate(value) ?? value;
  }
  if (
    (format === "@" || TEXT_CODE_FORMAT.test(format)) &&
    /^0\d+$/.test(displayText.trim())
  ) {
    return displayText.trim();
  }
  return Number.isFinite(value) ? value : String(value);
}

export function displayCellValue(value: CellValue, displayText = ""): CellValue {
  return displayText === "" ? value : displayText;
}
