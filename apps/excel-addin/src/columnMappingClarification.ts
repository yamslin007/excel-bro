/**
 * 把服务端下发的**列映射澄清**接进既有的需求澄清卡（Office 通道）。
 *
 * 为什么不另起一张卡片：澄清卡已经有一套完整的状态机（pending → resolving →
 * resolved / cancelled / invalidated）、续聊语义（带着补充说明重跑同一轮）、
 * 以及"范围变了就作废"的判定。列映射的澄清要的正是这些 —— 另起一套必然
 * 在状态与失效判定上与它漂移，用户会看到两种行为不一致的卡片。
 *
 * 本模块只做**翻译**：`ColumnMappingClarification` → `IntentClarification`。
 * 归因 → 可选项的映射放在 `@excel-bro/core`（Excel/WPS 共用），
 * 这里只管接上本端的消息契约 `MessageClarification`。
 *
 * **原文不进卡片**：卡片的 `summary` / `question` / `reason` 会进聊天历史与
 * 本地存储，而单元格原文属于本端数据。缺哪些类别、几格判断不了，够了。
 */

import { columnMappingClarificationCard } from "@excel-bro/core";
import type { ColumnMappingClarification, IntentOption } from "./contracts";
import type { MessageClarification } from "./types/chat";

/**
 * 澄清 → 卡片选项。
 *
 * `resolutionMode: "append"`：用户选了动作之后，是把这句补充说明**追加**到
 * 原始需求后面重跑，而不是替换掉原需求 —— 原需求里的映射表、区域、
 * 空白标签都还要用，替换掉它们会让第二轮丢失全部上下文。
 *
 * `id` 用 `resolution` 本身（而非序号）：它是协议枚举，天然稳定，
 * 也让回传路径可以直接把 `id` 当 `resolvedBy` 的值用，不需要再建一张映射表。
 */
function optionsFrom(clarification: ColumnMappingClarification): IntentOption[] {
  const card = columnMappingClarificationCard(clarification);
  return card.choices.map((choice) => ({
    id: choice.resolution,
    label: choice.label,
    description: choice.description,
    resolution: `${card.title}：${choice.label}。${
      choice.resolution === "addCategories" && card.missingCategories.length > 0
        ? `需要补充的类别：${card.missingCategories.join("、")}。`
        : ""
    }`,
    resolutionMode: "append"
  }));
}
/**
 * 组装 `MessageClarification`。
 *
 * `scopeFingerprint` / `originalPrompt` / `round` 由调用方给：它们描述的是
 * **这一轮请求的上下文**，不是澄清本身。分开传避免本函数去猜当前范围。
 */
export function columnMappingClarificationMessage(input: {
  clarification: ColumnMappingClarification;
  /** 规则指向的源区域，如 `订单!B2:B200`，用作卡片上的范围标签。 */
  scopeLabel: string;
  originalPrompt: string;
  scopeFingerprint: string;
  round: number;
}): MessageClarification {
  const { clarification } = input;
  const card = columnMappingClarificationCard(clarification);
  const detail =
    clarification.reason === "matchedDelimiterMissingCategories"
      ? "分隔符已按规则声明切分，但拆出的类别里有映射表没覆盖的。"
      : "有整格既没命中映射表、又看不出该按哪个已声明的分隔符拆。";
  return {
    id: `column-mapping-${clarification.reason}`,
    summary: card.title,
    question: card.message,
    reason: detail,
    scopeLabel: input.scopeLabel,
    options: optionsFrom(clarification),
    originalPrompt: input.originalPrompt,
    scopeFingerprint: input.scopeFingerprint,
    round: input.round,
    status: "pending"
  };
}
