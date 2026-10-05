/**
 * 来源格为**空白**时的结构化处理策略（L2，宿主无关）。
 *
 * 为什么不再用一个 `blankLabel: string` 表达全部情况：它只能表达"写一个
 * 什么标签"，而这实际上是**四种互不相同的意图**，且旧协议里的空串被禁止
 * （`min_length=1`），于是"清空"和"写空串"都表达不出来，"保留目标原值"
 * 更是完全无法表达 —— 只能靠模型去读回原值再写回去，那等于把该位置过一遍
 * 写入通道：公式、数字格式、条件格式都可能被替换掉。
 *
 * 四种意图必须分开表达，且**不许互相冒充**：
 *
 * - `keep`：保留目标原值 —— 这个位置**根本不进入写入范围**，不是"读回原值
 *   再写回"。后者是一次真实写入，会改公式与格式。
 * - `clear`：清空目标内容 —— 走 `clearRange:contents`，不是"写空串"。
 *   空串在部分宿主上会留下常量空字符串而不是真空单元格。
 * - `value`：写入用户指定的值。`value` **允许是空串**，那是"显式写一个空
 *   字符串"，与 `clear` 不是一回事，与 `keep` 更不是一回事。
 * - `undecided`：还没决定。**这不是错误**：只有在真的读到空白来源格时才需要
 *   停下来问；源区域一格空白都没有时，它不该挡住预览。
 *
 * 与"拒绝而非兜底"是同一个立场：`undecided` 绝不静默降级成某个默认标签
 * （曾经的实际后果就是把空白一律写成「/」或「未分类」，用户在单元格层面
 * 看不出那是实现替他选的）。
 */

import type { CellValue } from "./contracts";

export type BlankPolicy =
  | { kind: "keep" }
  | { kind: "clear" }
  | { kind: "value"; value: string }
  | { kind: "undecided" };

/** 缺省策略：未定。只有实际读到空白时才升级成一次澄清。 */
export const UNDECIDED_BLANK_POLICY: BlankPolicy = { kind: "undecided" };

/**
 * 旧协议 `blankLabel` → 新策略。
 *
 * 映射刻意是**唯一且无损**的：非空串 = `{kind:"value"}`。旧计划里写了
 * 「未分类」就还是写「未分类」，语义一点没动 —— 兼容不是"换个名字重解释"。
 *
 * 旧协议里空串本来就是**被拒绝**的（服务端 `min_length=1`、core 抛错），
 * 所以这里把它折成 `undecided` **不会改变任何一条曾经合法的旧计划的语义**。
 */
export function blankPolicyFromLegacy(
  blankLabel: string | null | undefined
): BlankPolicy {
  if (typeof blankLabel === "string" && blankLabel.length > 0) {
    return { kind: "value", value: blankLabel };
  }
  return { kind: "undecided" };
}

/**
 * 严格解析。**认不出来的 `kind` 一律抛错**，不折成 `undecided` ——
 * 一条写错但被静默接受的策略会让"保留原值"变成"未定"，下一轮又去问用户
 * 一个他已经回答过的问题。
 */
export function parseBlankPolicy(value: unknown): BlankPolicy {
  if (!value || typeof value !== "object") return { kind: "undecided" };
  const candidate = value as Record<string, unknown>;
  switch (candidate.kind) {
    case "keep":
      return { kind: "keep" };
    case "clear":
      return { kind: "clear" };
    case "undecided":
      return { kind: "undecided" };
    case "value": {
      if (typeof candidate.value !== "string") {
        throw new BlankPolicyError("空白策略 kind=value 缺少字符串类型的 value。");
      }
      return { kind: "value", value: candidate.value };
    }
    default:
      throw new BlankPolicyError(`无法识别的空白策略 kind：${String(candidate.kind)}`);
  }
}

export class BlankPolicyError extends Error {}

/** 空白 = 没有值，或纯空白字符串。数字 0 与 false **不是**空白。 */
export function isBlankValue(value: CellValue): boolean {
  if (value === null || value === undefined) return true;
  return typeof value === "string" && value.trim() === "";
}

/**
 * 单个空白格在给定策略下的**落地方式**。
 *
 * `undecided` 走到这里是**实现缺陷**（调用方应当先按"是否真有空白"拦掉），
 * 所以抛错而不是挑一个默认值 —— 静默兜底正是本模块要消灭的行为。
 */
export type BlankCellOutcome =
  | { kind: "skip" }
  | { kind: "clear" }
  | { kind: "write"; value: string };

export function blankCellOutcome(policy: BlankPolicy): BlankCellOutcome {
  switch (policy.kind) {
    case "keep":
      return { kind: "skip" };
    case "clear":
      return { kind: "clear" };
    case "value":
      return { kind: "write", value: policy.value };
    case "undecided":
      throw new BlankPolicyError(
        "空白策略尚未确定，不能落地任何写入动作。"
      );
  }
}

/** 源矩阵里所有空白格的位置（0 基行列）。 */
export function blankCellPositions(
  sourceValues: readonly (readonly CellValue[])[]
): { row: number; column: number }[] {
  const positions: { row: number; column: number }[] = [];
  sourceValues.forEach((row, rowIndex) => {
    row.forEach((cell, columnIndex) => {
      if (isBlankValue(cell)) positions.push({ row: rowIndex, column: columnIndex });
    });
  });
  return positions;
}

/** 面向用户的短标签。**不进模型 prompt**，只用于卡片与日志。 */
export const BLANK_POLICY_LABEL: Record<BlankPolicy["kind"], string> = {
  keep: "保留目标原值（该位置不写入）",
  clear: "清空目标内容",
  value: "写入指定值",
  undecided: "尚未决定"
};

/**
 * 澄清卡上的四个动作，顺序即推荐顺序。
 *
 * 与 `ColumnMappingResolution` 分开，是两件不同的事：那些是**类别/分隔符**
 * 的补救动作，这些是**空白**的补救动作。合成一组会让"映射表缺类别"的卡片
 * 冒出"保留目标原值"这种无关按钮 —— 与"归因决定可选项"同一条纪律。
 */
export type BlankPolicyResolution =
  | "keepTarget"
  | "clearTarget"
  | "writeBlankValue"
  | "writeEmptyString";

export const BLANK_POLICY_CHOICE_TEXT: Record<
  BlankPolicyResolution,
  { label: string; description: string }
> = {
  keepTarget: {
    label: "保留目标原值",
    description: "这些位置不写入任何内容，目标格现在的公式、值与格式都不动。"
  },
  clearTarget: {
    label: "清空目标内容",
    description: "把这些目标格清成真空单元格（清除内容，不是写一个空字符串）。"
  },
  writeBlankValue: {
    label: "写入指定值",
    description: "给空白来源格统一写一个你指定的标签，需要填写具体写什么。"
  },
  writeEmptyString: {
    label: "写入空字符串",
    description: "明确写一个空字符串常量，与「清空」不同 —— 单元格不是真空的。"
  }
};
