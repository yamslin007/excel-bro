/**
 * 把服务端下发的列映射**规则**在本端算成可执行计划（方案 §3，阶段 4）。
 *
 * 为什么这一步必须在客户端：WPS 的规划快照 `dataRows` 恒为空
 * （`wpsReadonlyAgent.ts` 里唯一的赋值处），配 `planningReadPolicy: "none"`
 * —— 这是刻意的隐私取舍。服务端手上没有源值，算不出目标矩阵，所以它只下发
 * 规则，由这里用 `packages/core` 的 `buildColumnMappingActions` 就地算。
 *
 * 两侧算出的值必然相同，这一点由 `config/column_mapping_cases.json` 的共享
 * 夹具锁住，不靠人工对读源码。
 *
 * **源值不上传**：本模块只接收已在内存里的授权工作表值，产出的计划直接交给
 * `prepareWpsPlan` 执行。算好的矩阵不回传服务端 —— 单射映射可逆，标签列加上
 * 映射表足以还原源列，回传就等于把源列上传了。
 */

import { buildColumnMappingActions, columnMappingFormulaCandidate, parseExplicitA1Rect } from "@excel-bro/core";
import capabilities from "../../../config/capabilities.json";
import { assertSafeFormula } from "../../excel-addin/src/formulaSafety";
import type { AnalysisPlan, CellValue, ColumnMappingRule } from "../../excel-addin/src/contracts";

const FORMULA_WARNING = "右侧副本将先试算公式并核对每格结果；公式只识别映射表中的整格类别和该行本次出现的组合，其他新组合显示 #N/A，须重新预览。";
const FIXED_WARNING = "公式无法可靠表达当前规则或超出长度上限，改用固定值；源列之后变化，目标列不会自动更新。";

/** 一张已读入内存的授权工作表。与 `WpsSheetReadResult` 的相关字段同形。 */
export type MappingSourceSheet = {
  name: string;
  /** 值矩阵左上角在工作表里的 1 基行/列号。 */
  startRow: number;
  startColumn: number;
  /**
   * 宿主报告的**声明宽度**。列边界按它判定，不按各行实际长度 ——
   * 读取侧会裁掉尾部空值，按行长判定会把真实空单元格误判成越界。
   */
  columnCount: number;
  values: readonly (readonly CellValue[])[];
};

/**
 * 从已读入的工作表里切出规则指定的源区域。
 *
 * 越界**拒绝，不补 `null`**：补出来的空值会被映射成空白标签，写进实际上没有
 * 数据的行，用户看到凭空多出来的一列标签。
 *
 * 四个方向（行/列 × 早于原点/超出范围）**都要拒**，这一点由
 * `config/mapping_slice_cases.json` 与服务端 `_snapshot_region` 共用同一批
 * 用例锁住。此前两侧各漏一个方向：本端不拦右侧列溢出、服务端不拦左侧列早于
 * 原点，各自静默折成空值 —— 正是夹具跑起来才暴露的。
 */
export function sliceMappingSource(
  sheet: MappingSourceSheet,
  sourceRange: string
): CellValue[][] {
  const rect = parseExplicitA1Rect(sourceRange);
  if (!rect) {
    throw new Error(`无法解析源区域「${sourceRange}」，或其结束位置早于开始位置。`);
  }
  const lastRow = sheet.startRow + sheet.values.length - 1;
  if (rect.startRow < sheet.startRow) {
    throw new Error(
      `区域 ${sourceRange} 起始于第 ${rect.startRow} 行，` +
        `早于「${sheet.name}」已读入区域的起点第 ${sheet.startRow} 行。`
    );
  }
  if (rect.endRow > lastRow) {
    throw new Error(
      `「${sheet.name}」已读入到第 ${lastRow} 行，` +
        `区域 ${sourceRange} 要到第 ${rect.endRow} 行，超出已读取范围。` +
        "映射结果会写入工作表，不能对没读到的行补空值。"
    );
  }
  if (rect.startColumn < sheet.startColumn) {
    throw new Error(
      `区域 ${sourceRange} 起始于第 ${rect.startColumn} 列，` +
        `早于「${sheet.name}」已读入区域的起点第 ${sheet.startColumn} 列。`
    );
  }
  const lastColumn = sheet.startColumn + Math.max(sheet.columnCount, 1) - 1;
  if (rect.endColumn > lastColumn) {
    throw new Error(
      `「${sheet.name}」已读入到第 ${lastColumn} 列，` +
        `区域 ${sourceRange} 要到第 ${rect.endColumn} 列，超出已读取范围。`
    );
  }
  const rows: CellValue[][] = [];
  for (let row = rect.startRow; row <= rect.endRow; row += 1) {
    const sourceRow = sheet.values[row - sheet.startRow] ?? [];
    const cells: CellValue[] = [];
    for (let column = rect.startColumn; column <= rect.endColumn; column += 1) {
      const offset = column - sheet.startColumn;
      // 行内缺项折成 null：读取侧会因尾部空值把行裁短，区域内缺的那几格
      // 是真实空单元格。这与"列越界"是两件事 —— 后者已按声明宽度拒掉。
      cells.push(offset < sourceRow.length ? sourceRow[offset] : null);
    }
    rows.push(cells);
  }
  return rows;
}

/**
 * 规则 + 已读入的源表 → 可执行计划。
 *
 * 计划的 `warnings` / `assumptions` 直接沿用服务端下发的那几条：换了计算位置
 * 不改变"这是一张快照"的事实，告知不能因此丢掉。
 */
export function buildWpsColumnMappingPlan(
  rule: ColumnMappingRule,
  sheet: MappingSourceSheet,
  context: { summary: string; warnings: readonly string[]; assumptions: readonly string[] }
): AnalysisPlan {
  if (sheet.name !== rule.sourceSheet) {
    throw new Error(
      `规则要求读取「${rule.sourceSheet}」，但传入的是「${sheet.name}」。`
    );
  }
  const sourceValues = sliceMappingSource(sheet, rule.sourceRange);
  const mapping: Record<string, string> = {};
  for (const pair of rule.mapping) {
    mapping[pair.source] = pair.target;
  }
  const actions = buildColumnMappingActions({
    sourceSheet: rule.sourceSheet,
    targetSheet: rule.targetSheet,
    sourceRange: rule.sourceRange,
    targetRange: rule.targetRange,
    sourceValues,
    mapping,
    // 空白策略优先取结构化的那个；没给 = 未定。真读到空白时
    // `buildColumnMappingActions` 抛 `BlankPolicyNeededError`，由调用方
    // 翻成澄清卡 —— 这里不替用户挑默认标签。
    blankPolicy: rule.blankPolicy,
    blankLabel: rule.blankLabel,
    maxSplitActions: capabilities.columnMapping.maxSplitActions,
    // 多值语义参数原样透传：它们是**规则数据**，本端不认识任何具体分隔符
    // 或等级名。与 Excel 侧逐字一致，两侧才不会算出不同的矩阵。
    delimiters: rule.delimiters,
    valueRank: rule.valueRank,
    undeclaredSeparators: rule.undeclaredSeparators,
    clarification: rule.clarification
  });
  /*
   * 公式通道**只在目标区域是一整块连续写入**时才考虑。
   *
   * 需要跳过（保留原值）或清空时，动作已被拆成多段，而公式会覆盖区域内的
   * 每一个格子 —— 它表达不了"这一格别动"。硬上公式等于把"保留原值"实现成
   * "写了个空值/公式"，那正是这里要避免的。
   */
  const firstAction = actions[0];
  const fixedAction =
    actions.length === 1 && firstAction.type === "writeValues" ? firstAction : null;
  const candidate = fixedAction && rule.sourceSheet === rule.targetSheet
    ? columnMappingFormulaCandidate({
        sourceRange: rule.sourceRange,
        sourceValues,
        expected: fixedAction.values,
        mapping: rule.mapping,
        maxCases: capabilities.columnMappingFormula.maxCases,
        maxCharacters: capabilities.columnMappingFormula.maxCharactersPerCell
      })
    : null;
  let formulaIsSafe = candidate !== null;
  if (candidate) {
    try {
      candidate.formulas.forEach((row, rowIndex) => row.forEach((formula, columnIndex) =>
        assertSafeFormula(formula, `${rule.targetSheet}!${rule.targetRange} 第 ${rowIndex + 1} 行第 ${columnIndex + 1} 列`)
      ));
    } catch {
      formulaIsSafe = false;
    }
  }
  const useFormula = candidate !== null && formulaIsSafe;
  return {
    // 计划 id 在本端生成：服务端没产计划，也就没有可沿用的 id。
    id: `wps-map-column-${Date.now().toString(36)}`,
    title: `按映射表填写 ${rule.targetSheet}!${rule.targetRange}`,
    summary: context.summary,
    assumptions: [...context.assumptions],
    warnings: [
      ...context.warnings,
      useFormula
        ? FORMULA_WARNING
        : FIXED_WARNING
    ],
    actions: useFormula && candidate
      ? [{ type: "writeFormulas", sheet: rule.targetSheet, range: rule.targetRange, formulas: candidate.formulas }]
      : actions,
    acceptanceCriteria: useFormula
      ? [{ type: "rangeEquals", sheet: rule.targetSheet, range: rule.targetRange, expected: candidate.expected }]
      : undefined
  };
}

/** Only a locally generated mapping formula can be downgraded, and only to its local expected matrix. */
export function fixedValueFallbackFromFormula(plan: AnalysisPlan): AnalysisPlan | null {
  if (!plan.id.startsWith("wps-map-column-") || plan.actions.length !== 1) return null;
  const action = plan.actions[0];
  if (action.type !== "writeFormulas") return null;
  const criterion = plan.acceptanceCriteria?.find((item) =>
    item.type === "rangeEquals" && item.sheet === action.sheet && item.range === action.range
  );
  if (!criterion || criterion.type !== "rangeEquals" ||
      criterion.expected.length !== action.formulas.length ||
      criterion.expected.some((row, index) =>
        row.length !== action.formulas[index].length || row.some((value) => typeof value !== "string")
      )) return null;
  return {
    ...plan,
    id: `${plan.id}-fixed`,
    warnings: [
      ...plan.warnings.filter((warning) => warning !== FORMULA_WARNING),
      "右侧副本的公式试算或结果回读不符合计划；已重新生成固定值副本。源列之后变化，目标列不会自动更新。"
    ],
    actions: [{ type: "writeValues", sheet: action.sheet, range: action.range,
      values: criterion.expected.map((row) => row.map((value) => value as string)) }],
    acceptanceCriteria: [{ ...criterion }]
  };
}
