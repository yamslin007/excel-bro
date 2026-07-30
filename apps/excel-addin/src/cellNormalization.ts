import type { CellValue } from "./contracts";

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

export function displayCellValue(
  value: CellValue,
  displayText = ""
): CellValue {
  return displayText === "" ? value : displayText;
}
