/**
 * 本地列映射的结构化补全。规则和来源值只留在当前任务窗格内存；
 * 用户确认后由宿主重新读取源列，不能把这里的结果直接当作写入计划。
 */
import type { BlankPolicy, CellValue } from "@excel-bro/core";
import type { ColumnMappingRule } from "./contracts";

export function localMappingSourceKey(value: CellValue): string {
  if (value === true) return "True";
  if (value === false) return "False";
  return String(value);
}

export function mappingTargets(rule: ColumnMappingRule): string[] {
  return [...new Set(rule.mapping.map((pair) => pair.target))];
}

/** Reuse only an explicit, complete and unambiguous order. Never infer severity from labels. */
export function declaredMappingPriority(rule: ColumnMappingRule): string[] | null {
  const targets = mappingTargets(rule);
  if (targets.length === 1) return targets;
  const ranks = new Map(rule.valueRank?.map((entry) => [entry.value, entry.rank]));
  const targetRanks = new Map<string, number>();
  for (const pair of rule.mapping) {
    const rank = ranks.get(pair.source);
    if (rank === undefined || !Number.isFinite(rank)) return null;
    if (targetRanks.has(pair.target) && targetRanks.get(pair.target) !== rank) return null;
    targetRanks.set(pair.target, rank);
  }
  if (!targetRanks.size || new Set(targetRanks.values()).size !== targets.length) return null;
  return targets.sort((left, right) => targetRanks.get(right)! - targetRanks.get(left)!);
}

export function mappingInputError(source: string, target: string, rule: ColumnMappingRule): string | null {
  if (!target.trim()) return `请填写「${source}」的目标值`;
  if (target.trim().startsWith("=")) return `「${source}」不能填写公式，请填写目标标签`;
  if (rule.mapping.some((pair) => pair.source === source)) return `「${source}」已有映射，不能在补全中覆盖`;
  return null;
}

/** 只给候选，不自动采用；无法从已知类别明确推断时留空供用户填写。 */
export function suggestMappingDelimiter(
  rule: ColumnMappingRule,
  values: readonly CellValue[]
): string {
  const strings = values.map(localMappingSourceKey);
  const declaredCandidate = rule.undeclaredSeparators?.find(
    (candidate) => candidate && strings.some((value) => value.includes(candidate))
  );
  if (declaredCandidate) return declaredCandidate;
  const sources = rule.mapping.map((pair) => pair.source).filter(Boolean);
  for (const value of strings) {
    for (const left of sources) {
      if (!value.startsWith(left)) continue;
      for (const right of sources) {
        if (!value.endsWith(right)) continue;
        const middle = value.slice(left.length, value.length - right.length);
        if (middle && !sources.includes(middle) && middle.length <= 8) return middle;
      }
    }
  }
  return rule.delimiters?.[0] ?? "";
}

export function uncoveredSplitItems(
  rule: ColumnMappingRule,
  values: readonly CellValue[],
  delimiter: string
): string[] {
  if (!delimiter) return [];
  const known = new Set(rule.mapping.map((pair) => pair.source));
  const missing = new Set<string>();
  const delimiters = [...new Set([...(rule.delimiters ?? []), delimiter])];
  for (const value of values) {
    const text = localMappingSourceKey(value);
    if (known.has(text)) continue; // The executor always checks whole-cell mappings first.
    // Mirror the executor's earliest declared delimiter, not just the newest candidate.
    let selected = "";
    let earliest = Infinity;
    for (const candidate of delimiters) {
      if (!candidate) continue;
      const position = text.indexOf(candidate);
      if (position >= 0 && position < earliest) { earliest = position; selected = candidate; }
    }
    for (const part of selected ? text.split(selected) : [text]) {
      const item = part.trim();
      if (item && !known.has(item)) missing.add(item);
    }
  }
  return [...missing];
}

export function addExactMapping(
  rule: ColumnMappingRule,
  values: readonly CellValue[],
  targets: readonly string[]
): ColumnMappingRule {
  if (values.length !== targets.length || values.length === 0) {
    throw new Error("请逐项确认来源值与目标值。");
  }
  const existing = new Set(rule.mapping.map((pair) => pair.source));
  const additions = values.map((value, index) => {
    const source = localMappingSourceKey(value);
    const target = targets[index]?.trim() ?? "";
    const error = mappingInputError(source, target, rule);
    if (!source || error || existing.has(source)) throw new Error(error ?? `「${source}」的映射重复`);
    existing.add(source);
    return { source, target };
  });
  return { ...rule, mapping: [...rule.mapping, ...additions], clarification: undefined };
}

/**
 * 把用户选定的**空白策略**写进规则。
 *
 * 这是"替换"，不是"追加"：用户改主意时新策略整体取代旧的，不会留下两条
 * 互相冲突的空白约定。旧字段 `blankLabel` 同步清掉 —— 它只能表达"写个标签"，
 * 留着它会在 `blankPolicy` 缺席时被当成兜底，语义上正好是用户刚否掉的那个。
 */
export function applyBlankPolicy(
  rule: ColumnMappingRule,
  policy: BlankPolicy
): ColumnMappingRule {
  if (policy.kind === "value" && policy.value.trim() === "" && policy.value !== "") {
    throw new Error("空白写入值不能只有空白字符；要写真正的空字符串请选「写入空字符串」。");
  }
  return { ...rule, blankPolicy: policy, blankLabel: undefined, clarification: undefined };
}

export function addSplitMapping(
  rule: ColumnMappingRule,
  input: {
    delimiter: string;
    missingItems: readonly string[];
    targets: readonly string[];
    /** 目标标签从最高优先级到最低优先级；用户必须显式确认这个顺序。 */
    priorityTargets: readonly string[];
  }
): ColumnMappingRule {
  const delimiter = input.delimiter;
  if (!delimiter || !delimiter.trim()) throw new Error("请明确填写拆分符。");
  if (input.missingItems.length !== input.targets.length) {
    throw new Error("请补齐拆分项的映射。");
  }
  const existing = new Set(rule.mapping.map((pair) => pair.source));
  const additions = input.missingItems.map((source, index) => {
    const target = input.targets[index]?.trim() ?? "";
    const error = mappingInputError(source, target, rule);
    if (!source || error || existing.has(source)) throw new Error(error ?? `「${source}」的映射重复`);
    existing.add(source);
    return { source, target };
  });
  const mapping = [...rule.mapping, ...additions];
  const requiredTargets = [...new Set(mapping.map((pair) => pair.target))];
  const priority = [...input.priorityTargets];
  if (
    priority.length !== requiredTargets.length ||
    new Set(priority).size !== priority.length ||
    requiredTargets.some((target) => !priority.includes(target))
  ) {
    throw new Error("请为所有目标标签明确排序，不能遗漏或重复。");
  }
  const ranks = new Map(priority.map((target, index) => [target, priority.length - index]));
  return {
    ...rule,
    mapping,
    delimiters: [...new Set([...(rule.delimiters ?? []), delimiter])],
    undeclaredSeparators: (rule.undeclaredSeparators ?? []).filter((item) => item !== delimiter),
    valueRank: mapping.map((pair) => ({ value: pair.source, rank: ranks.get(pair.target)! })),
    clarification: undefined
  };
}
