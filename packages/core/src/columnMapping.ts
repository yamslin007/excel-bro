/**
 * 参数化列映射的确定性计算（L2，宿主无关）。
 *
 * **与 `server/app/column_mapping.py` 是镜像关系**，行为必须逐格一致。
 * 一致性不靠人工对读源码,靠共享夹具
 * `config/column_mapping_cases.json` —— 两侧各自加载同一批用例并断言同一批
 * 期望值,任何单边修改都会让一侧变红。改这里必须同时改 Python 侧。
 *
 * 为什么两份实现而不是一份:服务端需要它（文件夹通道由服务端直接读写 xlsx),
 * 两端客户端也需要它（WPS 的源值不上传服务端,`wpsReadonlyAgent.ts` 送的
 * `dataRows` 恒为空 —— 这是刻意的隐私/体积取舍,不是缺陷）。
 * 让服务端算 = WPS 永远算不出来;让客户端算 = 文件夹通道拿不到。
 *
 * 三处刻意的"拒绝"而非"兜底",理由同 Python 侧:
 * 1. 来源值不在映射表内 → 抛错,不猜。错误的分类在单元格层面看不出来。
 * 2. 源/目标在同一表上重叠 → 抛错。逐格落地时读写先后决定结果,
 *    而这个顺序在三条通道上不一致。
 * 3. 形状不符 → 抛错。截断或补齐都会写出用户没要求的格子。
 */

import type { CellValue } from "./contracts";

export class ColumnMappingError extends Error {}

/** 来源值未命中映射表。携带**全部**未命中值,便于一次补全。 */
export class UnmappedCategoryError extends ColumnMappingError {
  readonly values: readonly CellValue[];

  constructor(values: readonly CellValue[]) {
    const rendered = values.map((value) => JSON.stringify(value)).join("、");
    super(`来源值 ${rendered} 不在映射表内，已拒绝猜测。请补全映射表后重试。`);
    this.values = values;
  }
}

export class ShapeMismatchError extends ColumnMappingError {}

export class OverlappingRangeError extends ColumnMappingError {}

export interface ColumnMappingInput {
  sourceSheet: string;
  targetSheet: string;
  sourceRange: string;
  targetRange: string;
  sourceValues: readonly (readonly CellValue[])[];
  /** 来源值 → 目标标签。键顺序不影响结果。 */
  mapping: Readonly<Record<string, string>>;
  blankLabel: string;
}

export interface ColumnMappingAction {
  type: "writeValues";
  sheet: string;
  range: string;
  values: string[][];
}

const CELL = String.raw`\$?([A-Za-z]+)\$?(\d+)`;
const RANGE_PATTERN = new RegExp(`^${CELL}(?::${CELL})?$`);

function columnNumber(name: string): number {
  let result = 0;
  for (const character of name.toUpperCase()) {
    result = result * 26 + (character.charCodeAt(0) - 64);
  }
  return result;
}

interface RangeBox {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}

/** 1 基闭区间。与 Python 侧 `_range_box` 同语义。 */
function rangeBox(address: string): RangeBox {
  const match = RANGE_PATTERN.exec(address.trim());
  if (!match) throw new ColumnMappingError(`无法解析区域「${address}」。`);
  const startColumn = columnNumber(match[1]);
  const startRow = Number(match[2]);
  const endColumn = columnNumber(match[3] ?? match[1]);
  const endRow = Number(match[4] ?? match[2]);
  if (endRow < startRow || endColumn < startColumn) {
    throw new ColumnMappingError(`区域「${address}」的起止顺序颠倒。`);
  }
  return { startRow, startColumn, endRow, endColumn };
}

/** 空白 = 没有值,或纯空白字符串。数字 0 与 false **不是**空白。 */
function isBlank(value: CellValue): boolean {
  if (value === null || value === undefined) return true;
  return typeof value === "string" && value.trim() === "";
}

/**
 * 逐格映射。未命中值**全部**收集后一次抛出。
 *
 * 一次只报一个会让调用方反复试错;映射表通常人工维护,一轮拿到完整缺口
 * 才能一次补完。
 */
export function mapColumnValues(input: {
  sourceValues: readonly (readonly CellValue[])[];
  mapping: Readonly<Record<string, string>>;
  blankLabel: string;
}): string[][] {
  const mapped: string[][] = [];
  const unmapped: CellValue[] = [];
  const seen = new Set<string>();

  for (const row of input.sourceValues) {
    const mappedRow: string[] = [];
    for (const cell of row) {
      // 显式映射项优先于空白策略：调用方若把 "" 写进 mapping,
      // 说明它要区分"空串"与"真空格",这个意图应当被尊重。
      if (typeof cell === "string" && Object.hasOwn(input.mapping, cell)) {
        mappedRow.push(input.mapping[cell]);
        continue;
      }
      if (isBlank(cell)) {
        mappedRow.push(input.blankLabel);
        continue;
      }
      const key = typeof cell === "string" ? cell : stringifyKey(cell);
      if (Object.hasOwn(input.mapping, key)) {
        mappedRow.push(input.mapping[key]);
        continue;
      }
      if (!seen.has(key)) {
        seen.add(key);
        unmapped.push(cell);
      }
      mappedRow.push(input.blankLabel); // 占位，本次调用不会被返回
    }
    mapped.push(mappedRow);
  }

  if (unmapped.length > 0) throw new UnmappedCategoryError(unmapped);
  return mapped;
}

/**
 * 非字符串来源值的键化。**必须与 Python 的 `str()` 对齐**,
 * 否则同一张表在两端会得出不同的未命中集合。
 *
 * `true` 在 JS 里 `String(true) === "true"`,Python 是 `"True"`。
 * 布尔按 Python 形态键化;数字两端一致（整数不带 `.0`,JS 本来就不带）。
 */
function stringifyKey(value: CellValue): string {
  if (value === true) return "True";
  if (value === false) return "False";
  return String(value);
}

export function buildColumnMappingActions(
  input: ColumnMappingInput
): ColumnMappingAction[] {
  if (!input.sourceSheet.trim() || !input.targetSheet.trim()) {
    throw new ColumnMappingError("源或目标工作表名为空。");
  }
  if (!input.blankLabel) {
    throw new ColumnMappingError("空白标签为空，无法确定空白单元格应写入什么。");
  }

  const source = rangeBox(input.sourceRange);
  const target = rangeBox(input.targetRange);
  const sourceRows = source.endRow - source.startRow + 1;
  const sourceColumns = source.endColumn - source.startColumn + 1;
  const targetRows = target.endRow - target.startRow + 1;
  const targetColumns = target.endColumn - target.startColumn + 1;

  if (
    input.sourceValues.length !== sourceRows ||
    input.sourceValues.some((row) => row.length !== sourceColumns)
  ) {
    throw new ShapeMismatchError(
      `源值矩阵形状与区域「${input.sourceRange}」（${sourceRows}×${sourceColumns}）不一致。`
    );
  }
  if (targetRows !== sourceRows || targetColumns !== sourceColumns) {
    throw new ShapeMismatchError(
      `目标区域「${input.targetRange}」（${targetRows}×${targetColumns}）` +
        `与源区域「${input.sourceRange}」（${sourceRows}×${sourceColumns}）形状不一致。`
    );
  }

  if (input.sourceSheet === input.targetSheet) {
    const rowsOverlap =
      source.startRow <= target.endRow && target.startRow <= source.endRow;
    const columnsOverlap =
      source.startColumn <= target.endColumn &&
      target.startColumn <= source.endColumn;
    if (rowsOverlap && columnsOverlap) {
      throw new OverlappingRangeError(
        `源区域「${input.sourceRange}」与目标区域「${input.targetRange}」` +
          "在同一工作表上重叠，读写先后顺序在各通道不一致，已拒绝。"
      );
    }
  }

  return [
    {
      type: "writeValues",
      sheet: input.targetSheet,
      range: input.targetRange,
      values: mapColumnValues({
        sourceValues: input.sourceValues,
        mapping: input.mapping,
        blankLabel: input.blankLabel
      })
    }
  ];
}
