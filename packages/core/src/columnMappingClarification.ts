/**
 * 把「列映射待澄清的归因」翻成**一张澄清卡的内容**（L2，宿主无关）。
 *
 * 为什么不放在各端的 UI 里各自拼：Excel/WPS 两端要展示的是同一件事，
 * 归因与选项的对应关系又是**语义**（不是排版）—— 分开写必然漂移，
 * 而漂移的后果是同一份规则在两端给用户看到的补救动作不一样。
 *
 * 核心立场：**归因决定能做什么，不能做什么**。
 *
 * - `matchedDelimiterMissingCategories`：分隔符已经声明且对上了，缺的是类别。
 *   这时给"指定其他分隔符"只会让用户去改一个本来没问题的设置，
 *   所以选项里**只有**「补充类别映射」。
 * - `unresolvedSeparatorOrMapping`：整格处理不了，而且分不清是"用了别的
 *   分隔符"还是"缺这个类别的映射"。归因不确定就必须把选择权交给用户，
 *   三个选项都给全。
 *
 * **原文不出现在卡片里**：卡片内容可能被写进聊天历史、导出、上报，
 * 而单元格原文属于本端数据。要展示原文的界面有自己的源值，就地读即可。
 * 这里只带"有几格处理不了"这个计数。
 */

import type {
  ColumnMappingClarification,
  ColumnMappingResolution
} from "./columnMapping";

/**
 * 澄清卡的可选动作，带**面向用户的文案**。
 *
 * `resolution` 是回传给服务端的枚举值（与协议里的
 * `ColumnMappingResolution` 逐字对齐）；`label` / `description` 只用于展示。
 * 两者分开，界面文案改动不会改到协议值。
 */
export interface ColumnMappingClarificationChoice {
  resolution: ColumnMappingResolution;
  label: string;
  description: string;
}

export interface ColumnMappingClarificationCard {
  /** 卡片标题，一句话说明"停在哪了"。 */
  title: string;
  /** 缺失的类别（已去重）。只在归因确定时非空。 */
  missingCategories: string[];
  /** 归因不确定时，无法处理的整格数量；确定时为 0。 */
  unresolvedCellCount: number;
  /** 界面应提供的动作，顺序即推荐顺序。 */
  choices: ColumnMappingClarificationChoice[];
  /** 展示用的一句话说明，直接取自规则（不含原文）。 */
  message: string;
}

const CHOICE_TEXT: Record<
  ColumnMappingResolution,
  Omit<ColumnMappingClarificationChoice, "resolution">
> = {
  addCategories: {
    label: "补充类别映射",
    description: "把上面列出的类别补进映射表，然后重新生成预览。"
  },
  specifyOtherDelimiter: {
    label: "指定其他分隔符",
    description: "这一格可能用了别的分隔写法，确认后按它拆分再逐项匹配。"
  },
  treatWholeCellAsCategory: {
    label: "将整格视为一个类别",
    description: "不拆分，把这一格的完整内容当成一个类别去映射。"
  },
  keepTarget: {
    label: "保留目标原值",
    description:
      "空白来源格对应的目标位置不写入任何内容，那里现在的公式、值与格式都不动。"
  },
  clearTarget: {
    label: "清空目标内容",
    description: "把这些目标格清成真空单元格（清除内容，不是写一个空字符串）。"
  },
  writeBlankValue: {
    label: "写入指定值",
    description: "给空白来源格统一写一个你指定的标签，选这项后还要填具体写什么。"
  },
  writeEmptyString: {
    label: "写入空字符串",
    description: "明确写一个空字符串常量，与「清空」不同 —— 单元格不是真空的。"
  }
};

/** 归因 → 允许的动作。**归因不同，可选项就不同**，不能统一给全量。 */
const ALLOWED_RESOLUTIONS: Record<
  ColumnMappingClarification["reason"],
  readonly ColumnMappingResolution[]
> = {
  matchedDelimiterMissingCategories: ["addCategories"],
  unresolvedSeparatorOrMapping: [
    "specifyOtherDelimiter",
    "addCategories",
    "treatWholeCellAsCategory"
  ],
  /*
   * 空白的四个动作**自成一组**，与类别/分隔符那组不混：
   * 卡片上冒出无关按钮，用户会以为"补类别映射"也能解决空白。
   */
  blankPolicyNeeded: [
    "keepTarget",
    "clearTarget",
    "writeBlankValue",
    "writeEmptyString"
  ]
};

const TITLE: Record<ColumnMappingClarification["reason"], string> = {
  matchedDelimiterMissingCategories: "映射表缺这几个类别",
  unresolvedSeparatorOrMapping: "有单元格判断不了怎么拆",
  blankPolicyNeeded: "源区域有空白，要先定空白怎么处理"
};

/**
 * 规则里的澄清信息 → 卡片内容。
 *
 * 可选项**只按 `reason` 推导**，不读服务端传下来的 `resolvedBy`：
 * 那个字段是给协议留的口子，若界面直接信它，一个错写的 `resolvedBy`
 * 就能让"分隔符已对上"的卡片冒出"指定其他分隔符"，而这条恰恰是
 * 本模块要防的误导。`resolvedBy` 由本函数**反向**产出，供调用方回传。
 */
export function columnMappingClarificationCard(
  clarification: ColumnMappingClarification
): ColumnMappingClarificationCard {
  const allowed = ALLOWED_RESOLUTIONS[clarification.reason];
  return {
    title: TITLE[clarification.reason],
    missingCategories: [...clarification.missingCategories],
    // 归因确定时这个计数没有意义（缺的是类别，不是"处理不了的格子"），
    // 一律归零，免得界面显示一个会让人误判规模的数字。
    unresolvedCellCount:
      clarification.reason === "unresolvedSeparatorOrMapping"
        ? clarification.unresolvedCellCount
        : 0,
    choices: allowed.map((resolution) => ({
      resolution,
      ...CHOICE_TEXT[resolution]
    })),
    message: clarification.message
  };
}

/**
 * 卡片上被选中的动作 → 回传给服务端的 `resolvedBy` 列表。
 *
 * 与 `columnMappingClarificationCard` 是同一张表的两个方向：
 * 卡片只列允许的，回传也只回传允许的。放在一起是为了让"界面能点什么"
 * 与"回传什么"共用一个来源 —— 分开写就会出现界面给了某个按钮、
 * 但回传时被丢掉（或反之）的静默不一致。
 */
export function columnMappingResolutionsFor(
  reason: ColumnMappingClarification["reason"]
): ColumnMappingResolution[] {
  return [...ALLOWED_RESOLUTIONS[reason]];
}
