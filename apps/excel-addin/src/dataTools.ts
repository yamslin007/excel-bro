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

const MAX_TOOL_ROWS = capabilities.queryTable.maxRows;
const MAX_TOOL_COLUMNS = capabilities.queryTable.maxColumns;
const MAX_TOOL_CELLS = capabilities.queryTable.maxCells;
const HEADER_SCAN_ROWS = capabilities.queryTable.headerScanRows;

export interface SheetValues {
  name: string;
  values: CellValue[][];
}

export type DataToolErrorCode =
  | "UNSUPPORTED_TOOL"
  | "SCOPE_VIOLATION"
  | "FIELD_NOT_FOUND"
  | "TOOL_LIMIT_EXCEEDED";

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

type RecordRow = Record<string, CellValue>;

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

function recordsFromSheet(
  sheet: SheetValues,
  arguments_: QueryTableArguments,
  warnings: string[]
): RecordRow[] {
  if (sheet.values.length === 0) return [];
  const headerIndex = detectHeaderIndex(
    sheet.values,
    requiredFields(arguments_)
  );
  const rawHeaders = sheet.values[headerIndex] ?? [];
  const headers = rawHeaders.map((value, index) => {
    const text = String(value ?? "").trim();
    return text || `未命名列${index + 1}`;
  });
  const normalizedHeaders = [
    ...headers.map(normalizeField),
    normalizeField("工作表")
  ];
  const wanted = requiredFields(arguments_).map(normalizeField);
  const missing = wanted.filter(
    (field) => field && !normalizedHeaders.includes(field)
  );
  if (missing.length > 0) {
    warnings.push(
      `「${sheet.name}」未找到字段：${[...new Set(missing)].join("、")}`
    );
    return [];
  }
  return sheet.values
    .slice(headerIndex + 1)
    .filter((row) => row.some((value) => value !== null && value !== ""))
    .map((row) => {
      const record: RecordRow = {
        工作表: sheet.name
      };
      headers.forEach((header, index) => {
        record[header] = row[index] ?? null;
      });
      return record;
    });
}

function fieldValue(record: RecordRow, field: string): CellValue {
  const normalized = normalizeField(field);
  const key = Object.keys(record).find(
    (candidate) => normalizeField(candidate) === normalized
  );
  return key ? record[key] : null;
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
  const values = rows
    .map((row) => fieldValue(row, metric.field ?? ""))
    .filter((value) => value !== null && value !== "");
  if (metric.operation === "countDistinct") {
    return new Set(values.map((value) => String(value))).size;
  }
  const numbers = values
    .map(comparableNumber)
    .filter((value): value is number => value !== null);
  if (numbers.length === 0) return null;
  if (metric.operation === "sum") {
    return numbers.reduce((total, value) => total + value, 0);
  }
  if (metric.operation === "average") {
    return numbers.reduce((total, value) => total + value, 0) / numbers.length;
  }
  if (metric.operation === "min") return Math.min(...numbers);
  return Math.max(...numbers);
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

export function executeQueryTableData(
  request: DataToolRequest,
  sheets: SheetValues[]
): DataToolResult {
  const arguments_ = request.arguments;
  const warnings: string[] = [];
  const allRecords = sheets.flatMap((sheet) =>
    recordsFromSheet(sheet, arguments_, warnings)
  );
  if (
    allRecords.length === 0 &&
    warnings.length > 0 &&
    requiredFields(arguments_).length > 0
  ) {
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
  const records = allRecords.filter((record) =>
    (arguments_.filters ?? []).every((filter) =>
      matchesFilter(record, filter)
    )
  );

  if (arguments_.mode === "profile") {
    const field = arguments_.profileField!;
    const counts = new Map<string, { value: CellValue; count: number }>();
    for (const record of records) {
      const value = fieldValue(record, field);
      const key = JSON.stringify(value);
      const current = counts.get(key);
      counts.set(key, {
        value,
        count: (current?.count ?? 0) + 1
      });
    }
    const headers = [field, "数量", "占比"];
    const rows = [...counts.values()].map(({ value, count }) => [
      value,
      count,
      records.length > 0 ? count / records.length : 0
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
      scannedRows: allRecords.length,
      complete: warnings.length === 0,
      calculation: `按「${field}」统计数量，占比 = 该值数量 ÷ 筛选后总记录数。`,
      warnings
    };
  }

  if (arguments_.mode === "aggregate") {
    const groupBy = arguments_.groupBy ?? [];
    const groups = new Map<string, { values: CellValue[]; rows: RecordRow[] }>();
    for (const record of records) {
      const values = groupBy.map((field) => fieldValue(record, field));
      const key = JSON.stringify(values);
      const group = groups.get(key) ?? { values, rows: [] };
      group.rows.push(record);
      groups.set(key, group);
    }
    const metrics = arguments_.metrics ?? [];
    const metricHeaders = metrics.flatMap((metric) => [
      metric.outputName,
      ...(metric.ratioOutputName ? [metric.ratioOutputName] : [])
    ]);
    const headers = [...groupBy, ...metricHeaders];
    const baseRows = [...groups.values()].map((group) => ({
      values: group.values,
      metrics: metrics.map((metric) => metricValue(group.rows, metric))
    }));
    const totals = metrics.map((_, index) =>
      baseRows.reduce((total, row) => {
        const value = comparableNumber(row.metrics[index]);
        return total + (value ?? 0);
      }, 0)
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
      scannedRows: allRecords.length,
      complete: warnings.length === 0,
      calculation: `筛选后按 ${groupBy.join("、") || "全部记录"} 分组，执行 ${metrics
        .map((metric) => `${metric.outputName}=${metric.operation}`)
        .join("；")}。`,
      warnings
    };
  }

  const fields =
    arguments_.fields && arguments_.fields.length > 0
      ? arguments_.fields
      : Object.keys(records[0] ?? {}).slice(0, 30);
  const rows = records.map((record) =>
    fields.map((field) => fieldValue(record, field))
  );
  return {
    requestId: request.id,
    tool: "query_table",
    title: "查询结果",
    headers: fields,
    rows: sortAndLimit(fields, rows, arguments_),
    sourceSheets: sheets.map((sheet) => sheet.name),
    scannedRows: allRecords.length,
    complete: warnings.length === 0,
    calculation: `按 ${arguments_.filters?.length ?? 0} 个条件筛选，并返回指定字段。`,
    warnings
  };
}

export async function executeQueryTableTool(
  request: DataToolRequest,
  allowedSheetNames: string[],
  activeWorksheet: string
): Promise<DataToolResult> {
  if (request.tool !== "query_table") {
    throw new DataToolExecutionError(
      "UNSUPPORTED_TOOL",
      "不支持的数据工具",
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
    const ranges = pending.map(({ name, worksheet, usedRange }) => {
      if (usedRange.isNullObject) return { name, range: null };
      if (usedRange.columnCount > MAX_TOOL_COLUMNS) {
        throw new DataToolExecutionError(
          "TOOL_LIMIT_EXCEEDED",
          `「${name}」有 ${usedRange.columnCount} 列，超过本地工具的 ${MAX_TOOL_COLUMNS} 列安全上限。请缩小数据范围。`,
          false
        );
      }
      totalRows += usedRange.rowCount;
      totalCells += usedRange.rowCount * usedRange.columnCount;
      const range = worksheet.getRangeByIndexes(
        usedRange.rowIndex,
        usedRange.columnIndex,
        usedRange.rowCount,
        usedRange.columnCount
      );
      range.load("values");
      return { name, range };
    });
    if (totalRows > MAX_TOOL_ROWS || totalCells > MAX_TOOL_CELLS) {
      throw new DataToolExecutionError(
        "TOOL_LIMIT_EXCEEDED",
        `所选数据包含约 ${totalRows} 行、${totalCells} 个单元格，超过单次本地查询上限。请缩小工作表范围。`,
        false
      );
    }
    await context.sync();

    const sheets: SheetValues[] = ranges.map(({ name, range }) => ({
      name,
      values: (range?.values ?? []).map((row) =>
        row.map((value) =>
          value === null ||
          typeof value === "string" ||
          typeof value === "number" ||
          typeof value === "boolean"
            ? value
            : String(value)
        )
      )
    }));
    return executeQueryTableData(request, sheets);
  });
}
