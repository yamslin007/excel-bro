import type { CellValue } from "./contracts";
import capabilities from "../../../config/capabilities.json";

const HEADER_SCAN_ROWS = capabilities.queryTable.headerScanRows;

/** 我们替无表头区域合成的列名：`列1`、`列2`…… */
const SYNTHETIC_FIELD = /^列[1-9][0-9]*$/;

export function normalizeField(value: unknown): string {
  return String(value ?? "")
    .replace(/[\s_\-:：()（）]+/g, "")
    .toLocaleLowerCase();
}

/**
 * 这一行**能不能**当表头。
 *
 * 判据只有一条：至少有一个非空单元格，且至少有一个是**文本**。
 * 整行都是数值（`1/2/3/4/5` 这种无表头数字区域）按数据处理。
 *
 * 为什么要有这条判据：`detectHeaderIndex` 原来的打分是「非空唯一值越多越像
 * 表头」，而 `[[1],[2],[3],[4],[5]]` 每行得分都相同、平局取最早一行，
 * 于是**首行必然被当成表头吞掉**。后果是静默的：用户问「五个数的合计」，
 * 本地算出来的是 14（少了首行），而界面上看不出少了哪一行。
 * 数值行与文本行的区别是这里唯一能拿到的证据，所以把它当作闸门：
 * 一行全是数字时不敢断言「它是表头」，宁可把首行当数据。
 */
export function looksLikeHeaderRow(row: readonly CellValue[] | undefined): boolean {
  let nonEmpty = 0;
  let hasText = false;
  for (const value of row ?? []) {
    if (value === null || value === undefined || value === "") continue;
    nonEmpty += 1;
    if (typeof value === "string") hasText = true;
  }
  return nonEmpty > 0 && hasText;
}

/** 字段名是不是我们合成的「列N」。合成名只可能对应无表头区域。 */
export function isSyntheticField(value: unknown): boolean {
  return SYNTHETIC_FIELD.test(String(value ?? "").trim());
}

/** 无表头区域的合成列名，与 `recordsFromSheet` 合成的那一套完全一致。 */
export function syntheticFields(columnCount: number): string[] {
  return Array.from(
    { length: Math.max(0, columnCount) },
    (_unused, index) => `列${index + 1}`
  );
}

/**
 * 找出表头行下标；**没有表头时返回 -1**。
 *
 * - 请求里点名了合成列名（`列2`）⇒ 默认这块区域**没有表头**：那是本端自己合成的
 *   名字，只可能来自无表头区域。**但真表头恰好就叫「列1/列2」时以事实为准** ——
 *   有行真的写着这个名字，就用那一行（否则真实业务表会被误判成无表头）。
 *   调用方显式声明 `hasHeaders: true` 时根本不走这里，那条更优先。
 * - 候选行必须过 `looksLikeHeaderRow`：纯数值行不当表头。
 * - 指定了字段名时仍优先选「命中字段最多的行」（既有行为）；命中不到任何行
 *   且没有一行像表头时，返回 -1（无表头），而不是假定第 0 行是表头。
 */
export function detectHeaderIndex(
  values: CellValue[][],
  wantedFields: string[] = []
): number {
  const wanted = new Set(wantedFields.map(normalizeField).filter(Boolean));
  const wantsSynthetic = [...wanted].some(isSyntheticField);
  // 有指定字段（且不是合成名）时保留「找不到就用第 0 行」的既有行为；
  // 点名合成名时默认无表头 —— 只有真写着该名字的行才能推翻这个默认。
  let bestIndex = wanted.size > 0 && !wantsSynthetic ? 0 : -1;
  let bestScore = -1;
  const scanCount = Math.min(values.length, HEADER_SCAN_ROWS);
  for (let index = 0; index < scanCount; index += 1) {
    const row = values[index] ?? [];
    if (!looksLikeHeaderRow(row)) continue;
    const normalized = row.map(normalizeField).filter(Boolean);
    const matches = normalized.filter((field) => wanted.has(field)).length;
    // 合成名场景：没写着这个名字的行不能成为表头，否则等于替调用方猜。
    if (wantsSynthetic && matches === 0) continue;
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
  const headerIndex = detectHeaderIndex(values);
  if (headerIndex < 0) {
    // 无表头：如实报出合成列名，并让调用方知道「所有行都是数据」
    // （`列1/列2` 就是 query_table 里引用它们的方式）。
    const width = values.reduce(
      (max, row) => Math.max(max, Array.isArray(row) ? row.length : 0),
      0
    );
    return syntheticFields(width);
  }
  const header = values[headerIndex] ?? [];
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
