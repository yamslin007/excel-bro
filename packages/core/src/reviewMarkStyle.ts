/**
 * 审阅标记的**呈现方式**判据（宿主无关，Excel / WPS 两端共用）。
 *
 * 只回答一个语义问题：本次计划里，某块变更区域是否**会修改单元格填充**。
 *
 * 为什么这一条要抽成共享纯函数：两端各自的标记实现（`resultReviewMarkers.ts` /
 * `wpsResultReviewMarkers.ts`）都要据此在「浅黄临时底色 + 醒目边框」与
 * 「只画边框」之间二选一。若各写各的 `type === "setFill"` 判断，两端口径就会
 * 漂移 —— 同一个计划在 Excel 是底色标记、在 WPS 是边框标记，用户看到两种
 * "本次改动"的视觉，会以为一边改了填充另一边没改。
 *
 * 规则（用户口径，不可放宽）：
 *  - 本次动作**不修改填充** ⇒ 用浅黄临时底色 + 醒目边框，让变更区域一眼可见；
 *  - 本次动作**涉及填充变化**（`setFill`）⇒ 只用边框标记，保证实际填充结果
 *    可见、不被审阅底色盖掉。
 *
 * 判据只看动作**类型**，不看颜色值：`setFill` 会把区域底色改成本次业务结果，
 * 无论改成的颜色是什么，都不得再叠一层审阅底色去干扰它。
 */

import { parseExplicitA1Rect } from "./planCore";

/** 一个动作会改单元格填充（底色）时为 true。 */
export function actionTouchesFill(action: { type: string }): boolean {
  return action.type === "setFill";
}

/**
 * 计划里**任何**动作涉及填充变化时为 true。
 *
 * 一个计划可能同时含写值动作与 `setFill` 动作、且落在不同区域；但只要有一处
 * 会改填充，标记实现就必须知道"不能对那块区域叠底色"。调用方应按**动作/区域**
 * 逐块判定（见下方 `reviewMarkStyleForPlan`），而不是只拿整份计划的布尔值。
 */
export function planTouchesFill(actions: readonly { type: string }[]): boolean {
  return actions.some(actionTouchesFill);
}

export type ReviewMarkStyle = "fill+border" | "border-only" | "fill-only" | "none";

/**
 * 给定一个动作，返回它那一块变更区域该用哪种标记。
 *
 * `setFill` ⇒ `border-only`（别叠底色）；其余（写值/公式/清空/字体/对齐/行高等）
 * ⇒ `fill+border`（浅黄底色让变更区一眼可见）。
 *
 * 注意：这只看**单个动作**。同一区域若同时/重叠落了一个 `setFill`，最终底色
 * 会被填充动作覆盖，此时**那块区域**必须 `border-only` —— 用 `reviewMarkStyleForPlan`
 * 按整份计划的重叠关系算，不要用这个函数去决定整份计划的每块区域。
 */
export function reviewMarkStyleForAction(action: { type: string }): ReviewMarkStyle {
  return actionTouchesFill(action) ? "border-only" : action.type === "setBorders" ? "fill-only" : "fill+border";
}

/** 动作区域的稳定键（同 sheet + 同区域 = 同一块）。 */
export type ReviewMarkRegion = { sheet: string; range: string };

function regionKey(region: { sheet: string; range: string }): string {
  return `${region.sheet}\u0000${region.range}`;
}

/** 两个矩形是否相交（含边界相接）。解析失败（null）一律按"不相交"处理。 */
function rectsOverlap(
  a: ReturnType<typeof parseExplicitA1Rect>,
  b: ReturnType<typeof parseExplicitA1Rect>
): boolean {
  if (!a || !b) return false;
  return !(a.endColumn < b.startColumn || b.endColumn < a.startColumn ||
           a.endRow < b.startRow || b.endRow < a.startRow);
}

/**
 * 按**整份计划**的区域重叠关系，算出每个变更区域该用哪种标记。
 *
 * 规则（用户口径，不可放宽）：
 *  - 一块区域只要与**任何** `setFill` 动作的区域（同一工作表内）相交 —— 即使只是
 *    部分重叠 —— 最终都会有单元格的填充被本次动作改写，此时这块区域**只画边框**，
 *    不得叠浅黄底色，否则会遮住实际填充结果。
 *  - 其余区域（只写值/公式/字体/对齐/行高，且不被任何 setFill 碰到）→ 浅黄底色 + 边框。
 *
 * 为什么不能只按"单个动作是不是 setFill"判：`setFont` 和 `setFill` 落在同一区域、
 * 或写值区域与 setFill 区域部分重叠时，`setFont`/写值那块区域其实也包含会被改填充的
 * 单元格，按单动作判就会给它叠上底色 → 黄色盖住真实填充。两端必须读这一个函数。
 *
 * 区域解析失败（不是明确 A1）时：无法证明"不被 setFill 碰到"，保守按 `border-only`
 * 处理（宁可少叠一层底色，也不遮住真实填充）。
 */
export function reviewMarkStyleForPlan(
  actions: readonly { type: string; sheet: string; range: string }[]
): Map<string, ReviewMarkStyle> {
  // Choose a temporary property only when no overlapping business action owns it.
  const result = new Map<string, ReviewMarkStyle>();
  for (const action of actions) {
    const key = regionKey(action);
    if (result.has(key)) continue;
    const rect = parseExplicitA1Rect(action.range);
    const overlapsType = (type: string) => actions.some(other => {
      if (other.sheet !== action.sheet || other.type !== type) return false;
      const otherRect = parseExplicitA1Rect(other.range);
      if (!rect || !otherRect) return true;
      // An adjacent region's marker can hide the other side of a business border.
      if (type === "setBorders") return !(
        rect.endColumn + 1 < otherRect.startColumn || otherRect.endColumn + 1 < rect.startColumn ||
        rect.endRow + 1 < otherRect.startRow || otherRect.endRow + 1 < rect.startRow
      );
      return rectsOverlap(rect, otherRect);
    });
    const fillOwned = !rect || overlapsType("setFill");
    const borderOwned = overlapsType("setBorders");
    result.set(key, fillOwned ? (borderOwned ? "none" : "border-only") : (borderOwned ? "fill-only" : "fill+border"));
  }
  return result;
}

/**
 * 标记的固定配色。放共享层，两端读同一份，避免一边黄一边别的色。
 *
 * `fill` 是"临时审阅底色"，不是业务填充：它是挂在**独立条件格式规则**上的
 * `Interior.Color` / `FormatFill`，删除该规则即随标记一起消失，绝不写进
 * 单元格的 `Interior` 本身。`border` 是醒目边框色，与底色一起让变更区可辨。
 */
export const REVIEW_MARK_COLORS = {
  /** 浅黄临时底色（fill+border / fill-only 用）。 */
  fill: "#FFF3B0",
  /** 醒目边框（fill+border / border-only 用）。 */
  border: "#B8860B"
} as const;

/** 图例文本：与 `ResultReviewBar`/结果卡里"彩色边框/黄色底色"的口径一致。 */
export const REVIEW_MARK_LEGEND = {
  fill: "浅黄色底色 + 边框：本次变更区域",
  border: "边框：本次变更区域（该区本次改了填充，故不叠底色）"
} as const;
