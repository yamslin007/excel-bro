import { describe, expect, it } from "vitest";
import { reviewWriteBackTargetChanged, type ReviewWriteBackTarget } from "./reviewWriteBackGate";

/** 身份是**对象引用**：所以"同一份现场"要用同一批对象，每次新建对象就是另一个版本。 */
const REVIEW = { copy: "Book1" };
const PLAN = { revision: 1 };
const WORKBOOK = "url:C:/data/格式验收.xlsx";

function target(overrides: Partial<ReviewWriteBackTarget> = {}): ReviewWriteBackTarget {
  return { review: REVIEW, plan: PLAN, workbookIdentity: WORKBOOK, ...overrides };
}

/**
 * 判据要能**分开**判三件事：换了副本 / 换了计划 / 换了原工作簿。
 *
 * 每条只动一个字段，另外两个逐字不动 —— 这样"变红"只可能归因于那一个字段，
 * 不会靠别的字段顺带通过。
 */
describe("写回请求的现场身份（同一次审阅 / 同一计划版本 / 同一原工作簿）", () => {
  it("三项都还是同一份 ⇒ 放行", () => {
    expect(reviewWriteBackTargetChanged(target(), target())).toBe(false);
    const same = target();
    expect(reviewWriteBackTargetChanged(same, same)).toBe(false);
  });

  it("只有副本换了（放弃后又生成新副本）⇒ 已变化", () => {
    expect(reviewWriteBackTargetChanged(target(), target({ review: { copy: "Book2" } }))).toBe(true);
  });

  it("只有计划换了（重新预览 / 编辑后重新校验）⇒ 已变化", () => {
    expect(reviewWriteBackTargetChanged(target(), target({ plan: { revision: 2 } }))).toBe(true);
  });

  it("只有原工作簿换了（切文件 / 另存为另一条路径）⇒ 已变化", () => {
    expect(reviewWriteBackTargetChanged(target(), target({ workbookIdentity: "url:C:/data/另一本.xlsx" }))).toBe(true);
  });

  it("任一侧缺席（放弃之后计划与副本都不在场）⇒ 已变化（fail-closed）", () => {
    expect(reviewWriteBackTargetChanged(target(), null)).toBe(true);
    expect(reviewWriteBackTargetChanged(null, target())).toBe(true);
    expect(reviewWriteBackTargetChanged(null, null)).toBe(true);
  });

  it("内容一模一样但**不是同一个对象** ⇒ 仍然算换了（按对象身份判，不按内容判）", () => {
    // 两份内容完全相同的计划对象属于两个版本：前一次核对过的证据不能顺延给后一次。
    expect(reviewWriteBackTargetChanged(
      target({ review: { copy: "Book1" }, plan: { revision: 1 } }),
      target({ review: { copy: "Book1" }, plan: { revision: 1 } })
    )).toBe(true);
  });
});
