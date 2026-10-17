import type { WpsRange } from "./wpsHost";
import { normalizeWpsMatrix } from "./wpsWorkbookReader";

function cell(value: unknown): boolean {
  return value === null || typeof value === "string" || typeof value === "boolean" ||
    (typeof value === "number" && Number.isFinite(value));
}

/** Explicit blank (null) is distinct from a missing or partial Value2 response. */
export function readWpsRangeValues(range: WpsRange, rows: number, columns: number) {
  let raw = range.Value2;
  let captured = false;
  if (Array.isArray(raw)) {
    const flat = raw as unknown[];
    captured = flat.some(Array.isArray)
      ? flat.length === rows && Array.from({ length: rows }, (_, r) => {
          const row = flat[r];
          return Array.isArray(row) && row.length === columns &&
            Array.from({ length: columns }, (_, c) => cell(row[c])).every(Boolean);
        }).every(Boolean)
      : flat.length === rows * columns &&
        Array.from({ length: rows * columns }, (_, i) => cell(flat[i])).every(Boolean);
  } else {
    captured = rows === 1 && columns === 1 && cell(raw);
  }
  // Multi-cell scalar null does not prove every cell is empty. Read individual
  // cells when supported, without padding an unavailable response into blanks.
  if (!captured && rows * columns > 1 && typeof range.Item === "function") {
    raw = Array.from({ length: rows }, (_, r) =>
      Array.from({ length: columns }, (_, c) => range.Item!(r + 1, c + 1).Value2));
    captured = (raw as unknown[][]).every((row) => row.every(cell));
  }
  return { values: normalizeWpsMatrix(raw, rows, columns), valuesCaptured: captured };
}
