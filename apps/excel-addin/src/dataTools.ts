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

function matchesFilter(record: RecordRow, filter: DataFilter): boolean {
  const actual = fieldValue(record, filter.field);
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
            .join("；")}。`,
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
        .join("；")}。`,
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
      let header: CellValue[] | null = null;
      if (!usedRange.isNullObject) {
        for (
          let rowOffset = 0;
          rowOffset < usedRange.rowCount;
          rowOffset += TOOL_CHUNK_ROWS
        ) {
          if (options.signal?.aborted) {
            throw new DataToolExecutionError(
              "CANCELLED",
              "本地查询已取消，未生成或写入任何结果",
              false
            );
          }
          const rowCount = Math.min(
            TOOL_CHUNK_ROWS,
            usedRange.rowCount - rowOffset
          );
          const range = worksheet.getRangeByIndexes(
            usedRange.rowIndex + rowOffset,
            usedRange.columnIndex,
            rowCount,
            usedRange.columnCount
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
            const headerIndex = detectHeaderIndex(
              values,
              requiredFields(request.arguments)
            );
            header = values[headerIndex] ?? [];
            accumulator.addSheet({ name, values });
          } else {
            accumulator.addSheet({ name, values: [header, ...values] });
          }
          scannedRows += rowCount;
          options.onProgress?.({ scannedRows, totalRows, sheet: name });
        }
      } else {
        accumulator.addSheet({ name, values: [] });
      }
    }
    return accumulator.finish();
  });
}
