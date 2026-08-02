import type { CellValue, ExcelAction } from "./contracts";

export type SplitAggregateAction = Extract<
  ExcelAction,
  { type: "splitGroupAggregate" }
>;

export interface SplitAggregateOutput {
  splitValue: CellValue;
  headers: CellValue[];
  rows: CellValue[][];
  ratioColumnIndexes: number[];
}

export function batchSplitAggregateOutputs(
  outputs: SplitAggregateOutput[],
  batchSize: number
): SplitAggregateOutput[][] {
  if (!Number.isInteger(batchSize) || batchSize < 1) {
    throw new Error("拆分聚合同步批次大小必须是正整数");
  }
  const batches: SplitAggregateOutput[][] = [];
  for (let index = 0; index < outputs.length; index += batchSize) {
    batches.push(outputs.slice(index, index + batchSize));
  }
  return batches;
}

export function safeWorksheetBaseName(value: CellValue): string {
  let cleaned = String(value ?? "")
    .replace(/[\u0000-\u001f:：\\＼/／?？*＊\[\]［］]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^'+|'+$/g, "")
    .trim();
  if (!cleaned) cleaned = "未命名";
  if (cleaned.toLocaleLowerCase() === "history") cleaned = "History 结果";
  return cleaned.slice(0, 31);
}

function text(value: CellValue): string {
  return String(value ?? "").trim();
}

function normalizedHeader(value: CellValue): string {
  return text(value).toLocaleLowerCase();
}

function blank(value: CellValue): boolean {
  return value === null || (typeof value === "string" && value.trim() === "");
}

function key(values: CellValue[]): string {
  return JSON.stringify(
    values.map((value) => [
      value === null ? "null" : typeof value,
      typeof value === "string" ? value.trim() : value
    ])
  );
}

function numeric(value: CellValue): number {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const parsed = Number(String(value ?? "").replace(/,/g, "").trim());
  return Number.isFinite(parsed) ? parsed : 0;
}

function metricValue(
  operation: SplitAggregateAction["metrics"][number]["operation"],
  value: CellValue
): number {
  if (operation === "countRows") return 1;
  if (operation === "countNonBlank") return blank(value) ? 0 : 1;
  return numeric(value);
}

export function buildSplitAggregateOutputs(
  values: CellValue[][],
  action: SplitAggregateAction
): SplitAggregateOutput[] {
  if (values.length === 0) {
    throw new Error("源数据范围为空");
  }

  const requiredFields = [
    action.splitBy,
    ...action.groupBy,
    ...action.metrics
      .map((metric) => metric.field)
      .filter((field): field is string => Boolean(field))
  ];
  const requiredNormalized = [...new Set(requiredFields.map(normalizedHeader))];
  const headerRowIndex = values.findIndex((row) => {
    const rowHeaders = new Set(row.map(normalizedHeader).filter(Boolean));
    return requiredNormalized.every((field) => rowHeaders.has(field));
  });
  if (headerRowIndex < 0) {
    throw new Error(`未找到包含这些字段的表头：${requiredFields.join("、")}`);
  }

  const headerIndexes = new Map<string, number>();
  values[headerRowIndex].forEach((value, index) => {
    const normalized = normalizedHeader(value);
    if (normalized && !headerIndexes.has(normalized)) {
      headerIndexes.set(normalized, index);
    }
  });
  const indexOf = (field: string): number => {
    const index = headerIndexes.get(normalizedHeader(field));
    if (index === undefined) throw new Error(`未找到字段「${field}」`);
    return index;
  };
  const splitIndex = indexOf(action.splitBy);
  const groupIndexes = action.groupBy.map(indexOf);
  const metricIndexes = action.metrics.map((metric) =>
    metric.field ? indexOf(metric.field) : -1
  );

  type Group = {
    values: CellValue[];
    totals: number[];
  };
  type Pair = {
    groupKey: string;
    values: number[];
  };
  type Split = {
    value: CellValue;
    pairs: Map<string, Pair>;
  };
  const groups = new Map<string, Group>();
  const splits = new Map<string, Split>();

  for (const row of values.slice(headerRowIndex + 1)) {
    const splitValue = row[splitIndex] ?? null;
    const groupValues = groupIndexes.map((index) => row[index] ?? null);
    if (
      (!action.includeBlankSplitValues && blank(splitValue)) ||
      groupValues.some(blank)
    ) {
      continue;
    }
    const groupKey = key(groupValues);
    const splitKey = key([splitValue]);
    let group = groups.get(groupKey);
    if (!group) {
      group = {
        values: groupValues,
        totals: action.metrics.map(() => 0)
      };
      groups.set(groupKey, group);
    }
    let split = splits.get(splitKey);
    if (!split) {
      split = { value: splitValue, pairs: new Map() };
      splits.set(splitKey, split);
    }
    let pair = split.pairs.get(groupKey);
    if (!pair) {
      pair = {
        groupKey,
        values: action.metrics.map(() => 0)
      };
      split.pairs.set(groupKey, pair);
    }
    action.metrics.forEach((metric, metricIndex) => {
      const cell =
        metricIndexes[metricIndex] < 0
          ? null
          : (row[metricIndexes[metricIndex]] ?? null);
      const increment = metricValue(metric.operation, cell);
      group.totals[metricIndex] += increment;
      pair.values[metricIndex] += increment;
    });
  }

  if (splits.size === 0) {
    throw new Error("表头下没有可拆分的有效数据行");
  }
  if (splits.size > action.maxOutputSheets) {
    throw new Error(
      `将生成 ${splits.size} 个工作表，超过安全上限 ${action.maxOutputSheets} 个`
    );
  }

  const headers: CellValue[] = [
    ...action.groupBy,
    action.splitBy,
    ...action.metrics.flatMap((metric) => [
      metric.outputName,
      ...(metric.ratioOutputName ? [metric.ratioOutputName] : [])
    ])
  ];
  const ratioColumnIndexes: number[] = [];
  let outputIndex = action.groupBy.length + 1;
  for (const metric of action.metrics) {
    outputIndex += 1;
    if (metric.ratioOutputName) {
      ratioColumnIndexes.push(outputIndex);
      outputIndex += 1;
    }
  }

  return [...splits.values()].map((split) => ({
    splitValue: split.value,
    headers,
    ratioColumnIndexes,
    rows: [...split.pairs.values()].map((pair) => {
      const group = groups.get(pair.groupKey);
      if (!group) throw new Error("拆分聚合内部状态异常");
      const metrics: CellValue[] = [];
      action.metrics.forEach((metric, metricIndex) => {
        const aggregate = pair.values[metricIndex];
        metrics.push(aggregate);
        if (metric.ratioOutputName) {
          const total = group.totals[metricIndex];
          metrics.push(total === 0 ? null : aggregate / total);
        }
      });
      return [...group.values, split.value, ...metrics];
    })
  }));
}
