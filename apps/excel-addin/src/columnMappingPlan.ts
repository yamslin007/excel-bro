/**
 * 把服务端下发的列映射**规则**在本端算成可执行计划（Office 通道）。
 *
 * 与 WPS 的 `wpsColumnMappingPlan.ts` 是**同一套分工**，只是宿主读取方式不同，
 * 所以计算部分刻意放在共享 core（`buildColumnMappingActions`），这里只做：
 *
 * 1. 从已读入内存的授权工作表里切出规则指定的源区域（`sliceMappingSource`）；
 * 2. 组装 `AnalysisPlan`（白名单 `writeValues` + 服务端给的 summary/warnings）。
 *
 * 为什么必须在客户端算：当 `planningReadPolicy: "none"` 时，服务端手上没有源值
 * （这是刻意的隐私取舍，不是缺陷），它只能下发规则。源值只在本端读取、只在本端
 * 计算，**算好的矩阵不回传服务端** —— 单射映射可逆，标签列加上映射表足以还原
 * 源列，回传就等于把源列上传了。
 *
 * 两侧算出的值必然相同：由 `config/column_mapping_cases.json` 的共享夹具锁住，
 * 不靠人工对读源码。
 */

import { buildColumnMappingActions, parseExplicitA1Rect } from "@excel-bro/core";
import capabilities from "../../../config/capabilities.json";
import type {
  AnalysisPlan,
  CellValue,
  ColumnMappingRule,
  WorkbookSnapshot
} from "./contracts";
/** 一张已读入内存的授权工作表。与 WPS 侧 `MappingSourceSheet` 同形。 */
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
 * 四个方向（行/列 × 早于原点/超出范围）都要拒。与 WPS 的 `sliceMappingSource`
 * 及服务端 `_snapshot_region` 共用 `config/mapping_slice_cases.json` 的同一批
 * 用例锁定 —— 此前两侧各漏一个方向，正是夹具跑起来才暴露的。
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
 *
 * 计划 id 在本端生成（前缀 `office-map-column-`，与 WPS 的 `wps-map-column-`
 * 区分），服务端没产计划也就没有可沿用的 id。
 */
export function buildColumnMappingPlan(
  rule: ColumnMappingRule,
  sheet: MappingSourceSheet,
  context: {
    summary: string;
    warnings: readonly string[];
    assumptions: readonly string[];
  }
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
    // 空白策略**优先**取结构化的那个；只有旧规则才靠 `blankLabel` 折。
    // 没给任何策略时是"未定" —— 源区域真有空白会抛 `BlankPolicyNeededError`，
    // 由调用方翻成一张澄清卡，而不是在这里替用户挑一个默认标签。
    blankPolicy: rule.blankPolicy,
    blankLabel: rule.blankLabel,
    maxSplitActions: capabilities.columnMapping.maxSplitActions,
    // 多值语义参数原样透传：它们是**规则数据**，本端不认识任何具体分隔符
    // 或等级名。少传一个字段就会静默退化成"整格当作一个值"，
    // 而那种结果在单元格层面看不出来 —— 所以这里刻意逐个点名。
    delimiters: rule.delimiters,
    valueRank: rule.valueRank,
    undeclaredSeparators: rule.undeclaredSeparators,
    clarification: rule.clarification
  });
  return {
    id: `office-map-column-${Date.now().toString(36)}`,
    title: `按映射表填写 ${rule.targetSheet}!${rule.targetRange}`,
    summary: context.summary,
    assumptions: [...context.assumptions],
    warnings: [
      ...context.warnings,
      "当前 Excel 通道尚无写入前的隔离公式试算；本次使用已在本地算准的固定值，源列之后变化不会自动更新。"
    ],
    actions
  };
}

/**
 * 本地映射通道专用的规划快照：**剥掉所有数据行**。
 *
 * Excel 默认走 `planningReadPolicy: "ifNeeded"`（快照带值，服务端直接算）——
 * 那是所有其他写入动作依赖的通道，不能因为加了本地映射就改成全体不发数据。
 * 但"这一条需求要不要本地算"，在拿到服务端响应之前并不知道（模型可能给出
 * 映射规则，也可能给出别的计划）。若此时把整份带数据行的快照发出去，
 * 就违反了"意图阶段不向模型发送原始数据行"。
 *
 * 因此本地映射通道**独立构造**一份 planning 请求：工作表结构（表头、区域、
 * 行列数）保留 —— 服务端据此判断来源表是否存在、区域是否越界并产出规则；
 * `dataRows` 一律清空 —— 源值留在本端，服务端拿不到就算不出目标矩阵，
 * 只能下发规则，正好走本地通道。
 *
 * 这不改变"白名单 `writeValues` + 先预览后确认"的闸门：算出的计划仍走既有
 * 预览卡与用户确认，只是计算位置从服务端挪到了本端。
 */
export function stripDataRows(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  return {
    ...snapshot,
    worksheets: snapshot.worksheets.map((sheet) => ({
      ...sheet,
      dataRows: [],
      truncated: sheet.rowCount > 0
    }))
  };
}
