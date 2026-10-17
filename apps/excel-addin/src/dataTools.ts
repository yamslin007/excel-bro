import type {
  CellValue,
  DataFilter,
  DataMetric,
  DataToolRequest,
  DataToolResult,
  QueryTableArguments
} from "./contracts";
import capabilities from "../../../config/capabilities.json";
import { detectHeaderIndex, normalizeField } from "./tableSchema";
import {
  assertPythonPrepareRectWithin,
  parsePythonPrepareRect,
  PythonPrepareRegionRefused
} from "@excel-bro/core/pythonPrepare";
import { normalizeCellValue } from "./cellNormalization";

const MAX_TOOL_ROWS = capabilities.queryTable.maxRows;
const MAX_TOOL_COLUMNS = capabilities.queryTable.maxColumns;
const MAX_TOOL_CELLS = capabilities.queryTable.maxCells;
const TOOL_CHUNK_ROWS = capabilities.queryTable.chunkRows;
const HEADER_SCAN_ROWS = capabilities.queryTable.headerScanRows;

export interface SheetValues {
  name: string;
  values: CellValue[][];
}

export type DataToolErrorCode =
  | "UNSUPPORTED_TOOL"
  | "SCOPE_VIOLATION"
  | "FIELD_NOT_FOUND"
  | "TOOL_LIMIT_EXCEEDED"
  | "CANCELLED";

export interface QueryToolProgress {
  scannedRows: number;
  totalRows: number;
  sheet: string;
}

export interface QueryToolOptions {
  signal?: AbortSignal;
  onProgress?: (progress: QueryToolProgress) => void;
}

export class DataToolExecutionError extends Error {
  constructor(
    public readonly code: DataToolErrorCode,
    message: string,
    public readonly retryable: boolean,
    public readonly availableFields: string[] = []
  ) {
    super(message);
    this.name = "DataToolExecutionError";
  }
}

interface RecordSchema {
  headers: string[];
  fieldIndexes: ReadonlyMap<string, number>;
  requestedFieldIndexes: Map<string, number | null>;
}

interface RecordRow {
  values: CellValue[];
  schema: RecordSchema;
}

function comparableNumber(value: CellValue): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return null;
  const trimmed = value.trim().replace(/,/g, "");
  if (!trimmed) return null;
  const percent = trimmed.endsWith("%");
  const parsed = Number(percent ? trimmed.slice(0, -1) : trimmed);
  if (!Number.isFinite(parsed)) return null;
  return percent ? parsed / 100 : parsed;
}

function compareValues(left: CellValue, right: CellValue): number {
  const leftNumber = comparableNumber(left);
  const rightNumber = comparableNumber(right);
  if (leftNumber !== null && rightNumber !== null) {
    return leftNumber - rightNumber;
  }
  return String(left ?? "").localeCompare(String(right ?? ""), "zh-CN", {
    numeric: true,
    sensitivity: "base"
  });
}

function requiredFields(arguments_: QueryTableArguments): string[] {
  const metricOutputs = new Set(
    (arguments_.metrics ?? []).flatMap((metric) => [
      normalizeField(metric.outputName),
      ...(metric.ratioOutputName
        ? [normalizeField(metric.ratioOutputName)]
        : [])
    ])
  );
  return [
    ...(arguments_.fields ?? []),
    ...(arguments_.groupBy ?? []),
    ...(arguments_.filters ?? []).map((filter) => filter.field),
    ...(arguments_.metrics ?? [])
      .map((metric) => metric.field)
      .filter((field): field is string => Boolean(field)),
    ...(arguments_.profileField ? [arguments_.profileField] : []),
    ...(arguments_.sortBy &&
    !metricOutputs.has(normalizeField(arguments_.sortBy))
      ? [arguments_.sortBy]
      : [])
  ];
}

function* recordsFromSheet(
  sheet: SheetValues,
  required: string[],
  warnings: string[]
): Generator<RecordRow> {
  if (sheet.values.length === 0) return;
  const headerIndex = detectHeaderIndex(sheet.values, required);
  const rawHeaders = sheet.values[headerIndex] ?? [];
  const headers = rawHeaders.map((value, index) => {
    const text = String(value ?? "").trim();
    return text || `未命名列${index + 1}`;
  });
  const recordHeaders = ["工作表", ...headers];
  const fieldIndexes = new Map<string, number>();
  recordHeaders.forEach((header, index) => {
    const normalized = normalizeField(header);
    if (!fieldIndexes.has(normalized)) {
      fieldIndexes.set(normalized, index);
    }
  });
  const wanted = required.map(normalizeField);
  const missing = wanted.filter(
    (field) => field && !fieldIndexes.has(field)
  );
  if (missing.length > 0) {
    warnings.push(
      `「${sheet.name}」未找到字段：${[...new Set(missing)].join("、")}`
    );
    return;
  }
  const schema: RecordSchema = {
    headers: recordHeaders,
    fieldIndexes,
    requestedFieldIndexes: new Map()
  };
  for (const row of sheet.values.slice(headerIndex + 1)) {
    if (!row.some((value) => value !== null && value !== "")) continue;
    yield {
      values: [
        sheet.name,
        ...headers.map((_, index) => row[index] ?? null)
      ],
      schema
    };
  }
}

function fieldValue(record: RecordRow, field: string): CellValue {
  let index = record.schema.requestedFieldIndexes.get(field);
  if (index === undefined) {
    index =
      record.schema.fieldIndexes.get(normalizeField(field)) ?? null;
    record.schema.requestedFieldIndexes.set(field, index);
  }
  return index === null ? null : record.values[index] ?? null;
}

function matchesFilterValue(actual: CellValue, filter: DataFilter): boolean {
  const expected = filter.value ?? null;
  switch (filter.operator) {
    case "isBlank":
      return actual === null || actual === "";
    case "isNotBlank":
      return actual !== null && actual !== "";
    case "contains":
      return String(actual ?? "")
        .toLocaleLowerCase()
        .includes(String(expected ?? "").toLocaleLowerCase());
    case "equals":
      return compareValues(actual, expected) === 0;
    case "notEquals":
      return compareValues(actual, expected) !== 0;
    case "greaterThan":
      return compareValues(actual, expected) > 0;
    case "greaterThanOrEqual":
      return compareValues(actual, expected) >= 0;
    case "lessThan":
      return compareValues(actual, expected) < 0;
    case "lessThanOrEqual":
      return compareValues(actual, expected) <= 0;
  }
}

function matchesFilter(record: RecordRow, filter: DataFilter): boolean {
  return matchesFilterValue(fieldValue(record, filter.field), filter);
}

// 在 values（二维数组）中解析 filters：headerRowIndex 指向表头行（无表头时传 -1，
// 此时字段名无法解析，返回空）。返回所有命中行（AND 逻辑）的下标数组。
export function resolveFilteredRowIndices(
  values: CellValue[][],
  headerRowIndex: number,
  filters: DataFilter[]
): number[] {
  if (filters.length === 0 || values.length === 0) return [];
  if (headerRowIndex < 0 || headerRowIndex >= values.length) return [];
  const headers = values[headerRowIndex];
  const fieldIndexes = new Map<string, number>();
  headers.forEach((header, index) => {
    const text = String(header ?? "").trim();
    const normalized = normalizeField(text || `未命名列${index + 1}`);
    if (!fieldIndexes.has(normalized)) fieldIndexes.set(normalized, index);
  });
  const resolved: Array<{ column: number; filter: DataFilter }> = [];
  for (const filter of filters) {
    const index = fieldIndexes.get(normalizeField(filter.field));
    if (index === undefined) return [];
    resolved.push({ column: index, filter });
  }
  const result: number[] = [];
  for (let row = headerRowIndex + 1; row < values.length; row++) {
    const cells = values[row];
    const matches = resolved.every(({ column, filter }) =>
      matchesFilterValue(cells[column] ?? null, filter)
    );
    if (matches) result.push(row);
  }
  return result;
}

function metricValue(rows: RecordRow[], metric: DataMetric): CellValue {
  if (metric.operation === "countRows") return rows.length;
  if (metric.operation === "countDistinct") {
    const distinct = new Set<string>();
    for (const row of rows) {
      const value = fieldValue(row, metric.field ?? "");
      if (value !== null && value !== "") distinct.add(String(value));
    }
    return distinct.size;
  }
  let count = 0;
  let total = 0;
  let minimum = Infinity;
  let maximum = -Infinity;
  for (const row of rows) {
    const number = comparableNumber(
      fieldValue(row, metric.field ?? "")
    );
    if (number === null) continue;
    count += 1;
    total += number;
    if (number < minimum) minimum = number;
    if (number > maximum) maximum = number;
  }
  if (count === 0) return null;
  if (metric.operation === "sum") return total;
  if (metric.operation === "average") return total / count;
  if (metric.operation === "min") return minimum;
  return maximum;
}

function sortAndLimit(
  headers: string[],
  rows: CellValue[][],
  arguments_: QueryTableArguments
): CellValue[][] {
  const sortIndex = arguments_.sortBy
    ? headers.findIndex(
        (header) =>
          normalizeField(header) === normalizeField(arguments_.sortBy)
      )
    : -1;
  const sorted =
    sortIndex < 0
      ? rows
      : [...rows].sort((left, right) => {
          const compared = compareValues(left[sortIndex], right[sortIndex]);
          return arguments_.sortDirection === "asc" ? compared : -compared;
        });
  return sorted.slice(
    0,
    arguments_.limit ?? capabilities.queryTable.defaultLimit
  );
}

export interface QueryTableAccumulator {
  addSheet(sheet: SheetValues): void;
  finish(): DataToolResult;
}

/** 一块待读取的表格区域，用索引描述，便于分块取子区域。 */
export interface SheetReadRegion {
  name: string;
  worksheet: Excel.Worksheet;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
}

export interface SheetReadOptions {
  /** 用于识别表头行：优先挑包含这些字段名的行。 */
  queryArguments: QueryTableArguments;
  /**
   * detect（默认）= 按 queryArguments 的字段名在前几行里识别表头；
   * synthesize = 源数据没有表头，用「列1、列2……」补一行，所有行都算数据。
   */
  headerMode?: "detect" | "synthesize";
  signal?: AbortSignal;
  /** 每读完一块回调一次，参数是这一块的行数。 */
  onChunk?: (rowsRead: number) => void;
}

/**
 * 把一张表分块读进累加器。
 *
 * 分块是为了避免一次 load 整表撑爆 Office.js 的请求负载。累加器的 `addSheet`
 * 把每次调用都当成一张独立的表来解析，所以第二块起必须手动把首块识别出的表头
 * 行拼回去 —— 否则后续块的第一行数据会被当成表头吞掉，分组统计会静默算错，
 * 而且只在数据超过一个分块（chunkRows）时才暴露。
 */
export async function readSheetIntoAccumulator(
  context: Excel.RequestContext,
  region: SheetReadRegion,
  accumulator: QueryTableAccumulator,
  options: SheetReadOptions
): Promise<number> {
  if (region.rowCount <= 0 || region.columnCount <= 0) {
    accumulator.addSheet({ name: region.name, values: [] });
    return 0;
  }

  let header: CellValue[] | null = null;
  let rowsRead = 0;
  for (
    let rowOffset = 0;
    rowOffset < region.rowCount;
    rowOffset += TOOL_CHUNK_ROWS
  ) {
    if (options.signal?.aborted) {
      throw new DataToolExecutionError(
        "CANCELLED",
        "本地查询已取消，未生成或写入任何结果",
        false
      );
    }
    const rowCount = Math.min(TOOL_CHUNK_ROWS, region.rowCount - rowOffset);
    const range = region.worksheet.getRangeByIndexes(
      region.rowIndex + rowOffset,
      region.columnIndex,
      rowCount,
      region.columnCount
    );
    range.load("values,text,numberFormat");
    await context.sync();
    const values = range.values.map((row, rowIndex) =>
      row.map((value, columnIndex) =>
        normalizeCellValue(
          value,
          range.text[rowIndex]?.[columnIndex] ?? "",
          range.numberFormat[rowIndex]?.[columnIndex] ?? "General"
        )
      )
    );
    if (header === null) {
      if (options.headerMode === "synthesize") {
        // 源数据没有表头：造一行「列1、列2……」，这样所有真实行都算数据，
        // 且查询里的字段名可以用这套合成名引用。
        header = Array.from(
          { length: region.columnCount },
          (_, index) => `列${index + 1}`
        );
        accumulator.addSheet({ name: region.name, values: [header, ...values] });
      } else {
        const headerIndex = detectHeaderIndex(
          values,
          requiredFields(options.queryArguments)
        );
        header = values[headerIndex] ?? [];
        accumulator.addSheet({ name: region.name, values });
      }
    } else {
      accumulator.addSheet({ name: region.name, values: [header, ...values] });
    }
    rowsRead += rowCount;
    options.onChunk?.(rowCount);
  }
  return rowsRead;
}

export function createQueryTableAccumulator(
  request: DataToolRequest
): QueryTableAccumulator {
  const arguments_ = request.arguments;
  if (arguments_.combine && arguments_.combine.mode !== "union") {
    throw new DataToolExecutionError(
      "UNSUPPORTED_TOOL",
      "当前工作簿暂不支持多表去重或关联，请切换到文件夹模式",
      false
    );
  }
  type MetricAccumulator = {
    count: number;
    total: number;
    minimum: number;
    maximum: number;
    distinct: Set<string>;
  };
  const required = requiredFields(arguments_);
  const warnings: string[] = [];
  const availableFields = new Set<string>();
  const sourceSheets: string[] = [];
  const fields =
    arguments_.fields && arguments_.fields.length > 0
      ? arguments_.fields
      : null;
  let effectiveFields: string[] = fields ?? [];
  const rowCandidates: CellValue[][] = [];
  const profileCounts = new Map<string, { value: CellValue; count: number }>();
  const groupBy = arguments_.groupBy ?? [];
  const metrics = arguments_.metrics ?? [];
  const groups = new Map<
    string,
    { values: CellValue[]; metrics: MetricAccumulator[] }
  >();
  let scannedRows = 0;
  let filteredRows = 0;
  const limit = arguments_.limit ?? capabilities.queryTable.defaultLimit;
  const newMetricAccumulator = (): MetricAccumulator => ({
    count: 0,
    total: 0,
    minimum: Infinity,
    maximum: -Infinity,
    distinct: new Set()
  });
  const updateMetric = (
    state: MetricAccumulator,
    record: RecordRow,
    metric: DataMetric
  ) => {
    if (metric.operation === "countRows") {
      state.count += 1;
      return;
    }
    const value = fieldValue(record, metric.field ?? "");
    if (metric.operation === "countDistinct") {
      if (value !== null && value !== "") {
        state.distinct.add(JSON.stringify(value));
      }
      return;
    }
    const numeric = comparableNumber(value);
    if (numeric === null) return;
    state.count += 1;
    state.total += numeric;
    if (numeric < state.minimum) state.minimum = numeric;
    if (numeric > state.maximum) state.maximum = numeric;
  };
  const metricResult = (
    state: MetricAccumulator,
    metric: DataMetric
  ): CellValue => {
    if (metric.operation === "countRows") return state.count;
    if (metric.operation === "countDistinct") return state.distinct.size;
    if (state.count === 0) return null;
    if (metric.operation === "sum") return state.total;
    if (metric.operation === "average") return state.total / state.count;
    if (metric.operation === "min") return state.minimum;
    return state.maximum;
  };

  return {
    addSheet(sheet) {
      if (!sourceSheets.includes(sheet.name)) sourceSheets.push(sheet.name);
      sheet.values
        .slice(0, HEADER_SCAN_ROWS)
        .flat()
        .map((value) => String(value ?? "").trim())
        .filter(Boolean)
        .forEach((field) => availableFields.add(field));
      const sheetWarnings: string[] = [];
      for (const record of recordsFromSheet(sheet, required, sheetWarnings)) {
        scannedRows += 1;
        if (
          !(arguments_.filters ?? []).every((filter) =>
            matchesFilter(record, filter)
          )
        ) {
          continue;
        }
        filteredRows += 1;
        if (arguments_.mode === "profile") {
          const value = fieldValue(record, arguments_.profileField!);
          const key = JSON.stringify(value);
          const current = profileCounts.get(key);
          profileCounts.set(key, {
            value,
            count: (current?.count ?? 0) + 1
          });
          continue;
        }
        if (arguments_.mode === "aggregate") {
          const values = groupBy.map((field) => fieldValue(record, field));
          const key = JSON.stringify(values);
          const group = groups.get(key) ?? {
            values,
            metrics: metrics.map(newMetricAccumulator)
          };
          metrics.forEach((metric, index) =>
            updateMetric(group.metrics[index], record, metric)
          );
          groups.set(key, group);
          continue;
        }
        if (effectiveFields.length === 0) {
          effectiveFields = record.schema.headers.slice(0, 30);
        }
        rowCandidates.push(
          effectiveFields.map((field) => fieldValue(record, field))
        );
        if (rowCandidates.length > limit * 2) {
          rowCandidates.splice(
            0,
            rowCandidates.length,
            ...sortAndLimit(effectiveFields, rowCandidates, arguments_)
          );
        }
      }
      for (const warning of sheetWarnings) {
        if (!warnings.includes(warning)) warnings.push(warning);
      }
    },
    finish() {
      if (scannedRows === 0 && warnings.length > 0 && required.length > 0) {
        throw new DataToolExecutionError(
          "FIELD_NOT_FOUND",
          warnings.join("；"),
          true,
          [...availableFields].slice(0, 100)
        );
      }
      if (arguments_.mode === "profile") {
        const field = arguments_.profileField!;
        const headers = [field, "数量", "占比"];
        const rows = [...profileCounts.values()].map(({ value, count }) => [
          value,
          count,
          filteredRows > 0 ? count / filteredRows : 0
        ]);
        return {
          requestId: request.id,
          tool: "query_table",
          title: `${field}分布`,
          headers,
          rows: sortAndLimit(headers, rows, {
            ...arguments_,
            sortBy: arguments_.sortBy ?? "数量"
          }),
          sourceSheets,
          scannedRows,
          complete: warnings.length === 0,
          calculation: `按「${field}」统计数量，占比 = 该值数量 ÷ 筛选后总记录数。`,
          warnings
        };
      }
      if (arguments_.mode === "aggregate") {
        const metricHeaders = metrics.flatMap((metric) => [
          metric.outputName,
          ...(metric.ratioOutputName ? [metric.ratioOutputName] : [])
        ]);
        const headers = [...groupBy, ...metricHeaders];
        const baseRows = [...groups.values()].map((group) => ({
          values: group.values,
          metrics: metrics.map((metric, index) =>
            metricResult(group.metrics[index], metric)
          )
        }));
        const totals = metrics.map((_, index) =>
          baseRows.reduce(
            (total, row) =>
              total + (comparableNumber(row.metrics[index]) ?? 0),
            0
          )
        );
        const rows = baseRows.map((row) => {
          const output: CellValue[] = [...row.values];
          metrics.forEach((metric, index) => {
            const value = row.metrics[index];
            output.push(value);
            if (metric.ratioOutputName) {
              const numeric = comparableNumber(value) ?? 0;
              output.push(totals[index] > 0 ? numeric / totals[index] : 0);
            }
          });
          return output;
        });
        return {
          requestId: request.id,
          tool: "query_table",
          title: "分组统计结果",
          headers,
          rows: sortAndLimit(headers, rows, arguments_),
          sourceSheets,
          scannedRows,
          complete: warnings.length === 0,
          calculation: `筛选后按 ${groupBy.join("、") || "全部记录"} 分组，执行 ${metrics
            .map((metric) => `${metric.outputName}=${metric.operation}`)
            .join("；")}。${warnings.length === 0 ? "筛选后总记录数" : "已处理记录数（结果不完整）"}：${filteredRows}。`,
          warnings
        };
      }
      return {
        requestId: request.id,
        tool: "query_table",
        title: "查询结果",
        headers: effectiveFields,
        rows: sortAndLimit(effectiveFields, rowCandidates, arguments_),
        sourceSheets,
        scannedRows,
        complete: warnings.length === 0,
        calculation: `按 ${arguments_.filters?.length ?? 0} 个条件筛选，并返回指定字段。`,
        warnings
      };
    }
  };
}

export function executeQueryTableData(
  request: DataToolRequest,
  sheets: SheetValues[]
): DataToolResult {
  const accumulator = createQueryTableAccumulator(request);
  for (const sheet of sheets) accumulator.addSheet(sheet);
  return accumulator.finish();
}

function executeQueryTableDataLegacy(
  request: DataToolRequest,
  sheets: SheetValues[]
): DataToolResult {
  const arguments_ = request.arguments;
  if (arguments_.combine && arguments_.combine.mode !== "union") {
    throw new DataToolExecutionError(
      "UNSUPPORTED_TOOL",
      "当前工作簿暂不支持多表去重或关联，请切换到文件夹模式",
      false
    );
  }
  const warnings: string[] = [];
  const required = requiredFields(arguments_);
  type MetricAccumulator = {
    count: number;
    total: number;
    minimum: number;
    maximum: number;
    distinct: Set<string>;
  };
  const newMetricAccumulator = (): MetricAccumulator => ({
    count: 0,
    total: 0,
    minimum: Infinity,
    maximum: -Infinity,
    distinct: new Set()
  });
  const updateMetric = (
    state: MetricAccumulator,
    record: RecordRow,
    metric: DataMetric
  ) => {
    if (metric.operation === "countRows") {
      state.count += 1;
      return;
    }
    const value = fieldValue(record, metric.field ?? "");
    if (metric.operation === "countDistinct") {
      if (value !== null && value !== "") state.distinct.add(JSON.stringify(value));
      return;
    }
    const numeric = comparableNumber(value);
    if (numeric === null) return;
    state.count += 1;
    state.total += numeric;
    if (numeric < state.minimum) state.minimum = numeric;
    if (numeric > state.maximum) state.maximum = numeric;
  };
  const metricResult = (
    state: MetricAccumulator,
    metric: DataMetric
  ): CellValue => {
    if (metric.operation === "countRows") return state.count;
    if (metric.operation === "countDistinct") return state.distinct.size;
    if (state.count === 0) return null;
    if (metric.operation === "sum") return state.total;
    if (metric.operation === "average") return state.total / state.count;
    if (metric.operation === "min") return state.minimum;
    return state.maximum;
  };

  const fields = arguments_.fields && arguments_.fields.length > 0
    ? arguments_.fields
    : null;
  let effectiveFields: string[] = fields ?? [];
  const rowCandidates: CellValue[][] = [];
  const profileCounts = new Map<string, { value: CellValue; count: number }>();
  const groupBy = arguments_.groupBy ?? [];
  const metrics = arguments_.metrics ?? [];
  const groups = new Map<
    string,
    { values: CellValue[]; metrics: MetricAccumulator[] }
  >();
  let scannedRows = 0;
  let filteredRows = 0;
  const limit = arguments_.limit ?? capabilities.queryTable.defaultLimit;

  for (const sheet of sheets) {
    for (const record of recordsFromSheet(sheet, required, warnings)) {
      scannedRows += 1;
      if (
        !(arguments_.filters ?? []).every((filter) =>
          matchesFilter(record, filter)
        )
      ) {
        continue;
      }
      filteredRows += 1;
      if (arguments_.mode === "profile") {
        const value = fieldValue(record, arguments_.profileField!);
        const key = JSON.stringify(value);
        const current = profileCounts.get(key);
        profileCounts.set(key, {
          value,
          count: (current?.count ?? 0) + 1
        });
        continue;
      }
      if (arguments_.mode === "aggregate") {
        const values = groupBy.map((field) => fieldValue(record, field));
        const key = JSON.stringify(values);
        const group = groups.get(key) ?? {
          values,
          metrics: metrics.map(newMetricAccumulator)
        };
        metrics.forEach((metric, index) =>
          updateMetric(group.metrics[index], record, metric)
        );
        groups.set(key, group);
        continue;
      }
      if (effectiveFields.length === 0) {
        effectiveFields = record.schema.headers.slice(0, 30);
      }
      rowCandidates.push(
        effectiveFields.map((field) => fieldValue(record, field))
      );
      if (rowCandidates.length > limit * 2) {
        rowCandidates.splice(
          0,
          rowCandidates.length,
          ...sortAndLimit(effectiveFields, rowCandidates, arguments_)
        );
      }
    }
  }

  if (scannedRows === 0 && warnings.length > 0 && required.length > 0) {
    const availableFields = [
      ...new Set(
        sheets.flatMap((sheet) =>
          sheet.values
            .slice(0, HEADER_SCAN_ROWS)
            .flat()
            .map((value) => String(value ?? "").trim())
            .filter(Boolean)
        )
      )
    ].slice(0, 100);
    throw new DataToolExecutionError(
      "FIELD_NOT_FOUND",
      warnings.join("；"),
      true,
      availableFields
    );
  }

  if (arguments_.mode === "profile") {
    const field = arguments_.profileField!;
    const headers = [field, "数量", "占比"];
    const rows = [...profileCounts.values()].map(({ value, count }) => [
      value,
      count,
      filteredRows > 0 ? count / filteredRows : 0
    ]);
    return {
      requestId: request.id,
      tool: "query_table",
      title: `${field}分布`,
      headers,
      rows: sortAndLimit(headers, rows, {
        ...arguments_,
        sortBy: arguments_.sortBy ?? "数量"
      }),
      sourceSheets: sheets.map((sheet) => sheet.name),
      scannedRows,
      complete: warnings.length === 0,
      calculation: `按「${field}」统计数量，占比 = 该值数量 ÷ 筛选后总记录数。`,
      warnings
    };
  }

  if (arguments_.mode === "aggregate") {
    const metricHeaders = metrics.flatMap((metric) => [
      metric.outputName,
      ...(metric.ratioOutputName ? [metric.ratioOutputName] : [])
    ]);
    const headers = [...groupBy, ...metricHeaders];
    const baseRows = [...groups.values()].map((group) => ({
      values: group.values,
      metrics: metrics.map((metric, index) =>
        metricResult(group.metrics[index], metric)
      )
    }));
    const totals = metrics.map((_, index) =>
      baseRows.reduce(
        (total, row) => total + (comparableNumber(row.metrics[index]) ?? 0),
        0
      )
    );
    const rows = baseRows.map((row) => {
      const output: CellValue[] = [...row.values];
      metrics.forEach((metric, index) => {
        const value = row.metrics[index];
        output.push(value);
        if (metric.ratioOutputName) {
          const numeric = comparableNumber(value) ?? 0;
          output.push(totals[index] > 0 ? numeric / totals[index] : 0);
        }
      });
      return output;
    });
    return {
      requestId: request.id,
      tool: "query_table",
      title: "分组统计结果",
      headers,
      rows: sortAndLimit(headers, rows, arguments_),
      sourceSheets: sheets.map((sheet) => sheet.name),
      scannedRows,
      complete: warnings.length === 0,
      calculation: `筛选后按 ${groupBy.join("、") || "全部记录"} 分组，执行 ${metrics
        .map((metric) => `${metric.outputName}=${metric.operation}`)
        .join("；")}。${warnings.length === 0 ? "筛选后总记录数" : "已处理记录数（结果不完整）"}：${filteredRows}。`,
      warnings
    };
  }

  return {
    requestId: request.id,
    tool: "query_table",
    title: "查询结果",
    headers: effectiveFields,
    rows: sortAndLimit(effectiveFields, rowCandidates, arguments_),
    sourceSheets: sheets.map((sheet) => sheet.name),
    scannedRows,
    complete: warnings.length === 0,
    calculation: `按 ${arguments_.filters?.length ?? 0} 个条件筛选，并返回指定字段。`,
    warnings
  };
}

export async function executeQueryTableTool(
  request: DataToolRequest,
  allowedSheetNames: string[],
  activeWorksheet: string,
  options: QueryToolOptions = {}
): Promise<DataToolResult> {
  if (request.tool !== "query_table") {
    throw new DataToolExecutionError(
      "UNSUPPORTED_TOOL",
      "不支持的数据工具",
      false
    );
  }
  if (
    request.arguments.combine &&
    request.arguments.combine.mode !== "union"
  ) {
    throw new DataToolExecutionError(
      "UNSUPPORTED_TOOL",
      "当前工作簿暂不支持多表去重或关联，请切换到文件夹模式",
      false
    );
  }
  const allowed = new Set(allowedSheetNames);
  const requestedNames =
    request.arguments.scope === "active"
      ? [activeWorksheet]
      : allowedSheetNames;
  if (
    requestedNames.length === 0 ||
    requestedNames.some((name) => !allowed.has(name))
  ) {
    throw new DataToolExecutionError(
      "SCOPE_VIOLATION",
      "数据工具请求超出了用户选择的工作表范围",
      false
    );
  }

  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    const pending = requestedNames.map((name) => {
      const worksheet = worksheets.getItem(name);
      const usedRange = worksheet.getUsedRangeOrNullObject(true);
      usedRange.load(
        "address,rowCount,columnCount,rowIndex,columnIndex,isNullObject"
      );
      return { name, worksheet, usedRange };
    });
    await context.sync();

    let totalRows = 0;
    let totalCells = 0;
    for (const { name, usedRange } of pending) {
      if (usedRange.isNullObject) continue;
      if (usedRange.columnCount > MAX_TOOL_COLUMNS) {
        throw new DataToolExecutionError(
          "TOOL_LIMIT_EXCEEDED",
          `「${name}」有 ${usedRange.columnCount} 列，超过本地工具的 ${MAX_TOOL_COLUMNS} 列安全上限。请缩小数据范围。`,
          false
        );
      }
      totalRows += usedRange.rowCount;
      totalCells += usedRange.rowCount * usedRange.columnCount;
    }
    if (totalRows > MAX_TOOL_ROWS || totalCells > MAX_TOOL_CELLS) {
      throw new DataToolExecutionError(
        "TOOL_LIMIT_EXCEEDED",
        `所选数据包含约 ${totalRows} 行、${totalCells} 个单元格，超过单次本地查询上限。请缩小工作表范围。`,
        false
      );
    }
    const accumulator = createQueryTableAccumulator(request);
    let scannedRows = 0;
    for (const { name, worksheet, usedRange } of pending) {
      if (usedRange.isNullObject) {
        accumulator.addSheet({ name, values: [] });
        continue;
      }
      await readSheetIntoAccumulator(
        context,
        {
          name,
          worksheet,
          rowIndex: usedRange.rowIndex,
          columnIndex: usedRange.columnIndex,
          rowCount: usedRange.rowCount,
          columnCount: usedRange.columnCount
        },
        accumulator,
        {
          queryArguments: request.arguments,
          signal: options.signal,
          onChunk: (rowsRead) => {
            scannedRows += rowsRead;
            options.onProgress?.({ scannedRows, totalRows, sheet: name });
          }
        }
      );
    }
    return accumulator.finish();
  });
}

/** ``prepare_python_data`` 的单个来源：授权工作表名 + A1 区域。 */
export interface PreparePythonDataSource {
  sheet: string;
  range: string;
}

/**
 * 读授权表指定区域的**原始值矩阵**（供隔离 Python 数据登记）。
 *
 * 与 ``query_table`` 不同：这里不做表头识别、不做分组/画像，直接按来源声明
 * 的 A1 区域读原始值。每个来源的 sheet 必须落在已授权工作表内（复核由桥的
 * ``verify`` 端口兜底，这里再做一次防御性拒绝）。返回与服务端
 * ``BridgeDatasetReceipt`` 对齐的 ``sourceIndex → values`` 块。
 */
export async function executePreparePythonData(
  sources: PreparePythonDataSource[],
  allowedSheetNames: string[],
  options: QueryToolOptions = {}
): Promise<{ sourceIndex: number; values: CellValue[][] }[]> {
  if (!Array.isArray(sources) || sources.length === 0) {
    throw new DataToolExecutionError(
      "UNSUPPORTED_TOOL",
      "数据登记缺少来源",
      false
    );
  }
  const allowed = new Set(allowedSheetNames);
  for (const source of sources) {
    if (
      source === null ||
      typeof source !== "object" ||
      typeof source.sheet !== "string" ||
      typeof source.range !== "string"
    ) {
      throw new DataToolExecutionError(
        "SCOPE_VIOLATION",
        "数据登记来源无效",
        false
      );
    }
    if (!allowed.has(source.sheet)) {
      throw new DataToolExecutionError(
        "SCOPE_VIOLATION",
        "数据登记请求超出了用户选择的工作表范围",
        false
      );
    }
  }

  const rects = sources.map((source) => {
    const rect = parsePythonPrepareRect(source.range);
    if (!rect) {
      throw new DataToolExecutionError(
        "SCOPE_VIOLATION",
        "数据登记区域无效",
        false
      );
    }
    return rect;
  });

  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    const blocks: { sourceIndex: number; values: CellValue[][] }[] = [];
    for (let index = 0; index < sources.length; index++) {
      if (options.signal?.aborted) {
        throw new DataToolExecutionError(
          "CANCELLED",
          "本地数据登记已取消",
          false
        );
      }
      const source = sources[index];
      const worksheet = worksheets.getItem(source.sheet);
      /*
       * 先用**已用区域**把越界请求挡掉，再读精确区域。
       *
       * `getRange` 拿绝对地址，落在已用区域之外时会**静默返回空白**而不是报错 ——
       * WPS 侧是明确拒绝。不先挡，同一张表在两端就会一个给空值、一个拒绝，
       * 语义分叉且不可见。挡掉之后，"表外"在两端都是拒绝，"表内的真空白"
       * 才是空值 —— 两者不再互相冒充。
       */
      const used = worksheet.getUsedRange();
      used.load("rowIndex,columnIndex,rowCount,columnCount");
      await context.sync();
      try {
        assertPythonPrepareRectWithin(
          rects[index],
          {
            startRow: Number(used.rowIndex ?? 0) + 1,
            startColumn: Number(used.columnIndex ?? 0) + 1
          },
          Number(used.rowCount ?? 0),
          Number(used.columnCount ?? 0)
        );
      } catch (error) {
        if (error instanceof PythonPrepareRegionRefused) {
          throw new DataToolExecutionError(
            "SCOPE_VIOLATION",
            error.message,
            false
          );
        }
        throw error;
      }
      const range = worksheet.getRange(source.range);
      range.load("values,text,numberFormat");
      await context.sync();
      const values: CellValue[][] = range.values.map((row, rowIndex) =>
        row.map((value, columnIndex) =>
          normalizeCellValue(
            value,
            range.text[rowIndex]?.[columnIndex] ?? "",
            range.numberFormat[rowIndex]?.[columnIndex] ?? "General"
          )
        )
      );
      blocks.push({ sourceIndex: index, values });
    }
    return blocks;
  });
}

