import type { CellValue } from "./contracts";
import capabilities from "../../../config/capabilities.json";

const HEADER_SCAN_ROWS = capabilities.queryTable.headerScanRows;

export function normalizeField(value: unknown): string {
  return String(value ?? "")
    .replace(/[\s_\-:：()（）]+/g, "")
    .toLocaleLowerCase();
}

export function detectHeaderIndex(
  values: CellValue[][],
  wantedFields: string[] = []
): number {
  const wanted = new Set(wantedFields.map(normalizeField).filter(Boolean));
  let bestIndex = 0;
  let bestScore = -1;
  const scanCount = Math.min(values.length, HEADER_SCAN_ROWS);
  for (let index = 0; index < scanCount; index += 1) {
    const row = values[index] ?? [];
    const normalized = row.map(normalizeField).filter(Boolean);
    const matches = normalized.filter((field) => wanted.has(field)).length;
    const unique = new Set(normalized).size;
    const score =
      (wanted.size > 0 ? matches * 100 : 0) +
      unique -
      Math.max(0, normalized.length - unique) * 2;
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
    if (wanted.size > 0 && matches === wanted.size) return index;
  }
  return bestIndex;
}

export function detectSheetFields(values: CellValue[][]): CellValue[] {
  if (values.length === 0) return [];
  const header = values[detectHeaderIndex(values)] ?? [];
  const fields: CellValue[] = [];
  const seen = new Set<string>();
  for (const value of header) {
    const text = String(value ?? "").trim();
    const normalized = normalizeField(text);
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    fields.push(text);
  }
  return fields;
}
