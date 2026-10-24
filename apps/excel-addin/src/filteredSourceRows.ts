import type { CellValue, DataFilter } from "./contracts";
import { resolveFilteredRowIndices } from "./dataTools";
import { normalizeField } from "./tableSchema";

/** Explicit field resolution: unknown/ambiguous fields are not equivalent to no matching rows. */
export function filteredSourceRows(values: CellValue[][], filters: DataFilter[], hasHeaders: boolean): number[] {
  if (!hasHeaders) throw new Error("条件操作需要明确表头以定位字段，请先确认表头。");
  const header = values[0] ?? [];
  for (const filter of filters) {
    const matches = header.filter(value => normalizeField(String(value ?? "").trim()) === normalizeField(filter.field));
    if (matches.length !== 1) throw new Error(`筛选字段「${filter.field}」不存在或不唯一，请明确字段后重新生成计划。`);
  }
  return resolveFilteredRowIndices(values, 0, filters);
}
