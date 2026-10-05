import { describe, expect, it } from "vitest";
import type { ChangeReviewStatus, RangeChangeReview } from "./changeReview";
import {
  buildRangeChangeReview,
  canClaimWriteSuccess,
  cellDisplayText,
  changeReviewSummary
} from "./changeReview";

/**
 * 这批用例对应 docs/DUAL_VIEW_CHANGE_REVIEW_GUIDE_2026-09-29.md §8。
 * 它们是**先于实现**写下的：在共用模型落地之前这里应该全红，
 * 红灯的作用是把"空单元格 vs 空字符串""公式覆盖""回读不符"这些缺口钉住，
 * 不允许事后用"大概是等价的"糊过去。
 */

type Values = (string | number | boolean | null)[][];
type Formulas = (string | null)[][];

function snapshot(
  values: Values,
  formulas: Formulas,
  valuesCaptured = true
) {
  return { values, formulas, valuesCaptured };
}

function base(input: Partial<Parameters<typeof buildRangeChangeReview>[0]> = {}) {
  return {
    sheet: "Sheet1",
    range: "A1:A1",
    phase: "actual" as const,
    status: "applied" as ChangeReviewStatus,
    before: snapshot([["x"]], [[null]]),
    after: snapshot([["y"]], [[null]]),
    ...input
  };
}

describe("双栏变更对照 · 空单元格与空字符串", () => {
  it("空字符串 → 空单元格 必须算一次改动", () => {
    const review = buildRangeChangeReview(
      base({
        before: snapshot([[""]], [[null]]),
        after: snapshot([[null]], [[null]])
      })
    );
    expect(review.rows[0].changed).toBe(true);
    expect(review.rows[0].before.value).toBe("");
    expect(review.rows[0].after.value).toBe(null);
  });

  it("空单元格 → 空字符串 也必须算一次改动", () => {
    const review = buildRangeChangeReview(
      base({
        before: snapshot([[null]], [[null]]),
        after: snapshot([[""]], [[null]])
      })
    );
    expect(review.rows[0].changed).toBe(true);
  });

  it("两个空状态各自相同 → 不算改动（且不能被折叠成同一个状态）", () => {
    const sameEmpty = buildRangeChangeReview(
      base({
        before: snapshot([[null]], [[null]]),
        after: snapshot([[null]], [[null]])
      })
    );
    const sameBlankString = buildRangeChangeReview(
      base({
        before: snapshot([[""]], [[null]]),
        after: snapshot([[""]], [[null]])
      })
    );
    expect(sameEmpty.rows[0].changed).toBe(false);
    expect(sameBlankString.rows[0].changed).toBe(false);
    // 折叠成同一状态会让前两条用例失效，这里反过来锁住区分度。
    expect(sameEmpty.rows[0].before.value).not.toBe(sameBlankString.rows[0].before.value);
  });
});

describe("双栏变更对照 · 公式", () => {
  it("覆盖已有公式时必须单独标记 overwritesFormula", () => {
    const review = buildRangeChangeReview(
      base({
        before: snapshot([[7]], [["=A2+1"]]),
        after: snapshot([[9]], [[null]])
      })
    );
    expect(review.rows[0].overwritesFormula).toBe(true);
    expect(review.rows[0].kind).toBe("formula");
    expect(review.rows[0].changed).toBe(true);
  });

  it("公式没变、只有计算值漂移 → 判为值变化，不得误报成公式变化", () => {
    const review = buildRangeChangeReview(
      base({
        before: snapshot([[7]], [["=A2+1"]]),
        after: snapshot([[8]], [["=A2+1"]])
      })
    );
    expect(review.rows[0].kind).toBe("value");
    expect(review.rows[0].overwritesFormula).toBe(false);
  });

  it("写入成功与否按公式本身核对，不能只看计算值", () => {
    const sameValueDifferentFormula = buildRangeChangeReview(
      base({
        before: snapshot([[7]], [["=A2+1"]]),
        after: snapshot([[7]], [["=A2+2"]])
      })
    );
    expect(sameValueDifferentFormula.rows[0].changed).toBe(true);
    expect(sameValueDifferentFormula.rows[0].kind).toBe("formula");
  });
});

describe("双栏变更对照 · 清空", () => {
  it("原非空 → 空 必须渲染成一行，不许被当成未改动过滤掉", () => {
    const review = buildRangeChangeReview(
      base({
        before: snapshot([["旧值"]], [[null]]),
        after: snapshot([[null]], [[null]])
      })
    );
    expect(review.rows).toHaveLength(1);
    expect(review.rows[0].kind).toBe("empty");
    expect(review.rows[0].changed).toBe(true);
    expect(review.changedCount).toBe(1);
  });

  it("本来就是空的清空 → 不算改动", () => {
    const review = buildRangeChangeReview(
      base({
        before: snapshot([[null]], [[null]]),
        after: snapshot([[null]], [[null]])
      })
    );
    expect(review.rows[0].changed).toBe(false);
    expect(review.changedCount).toBe(0);
  });
});

describe("双栏变更对照 · 部分写入", () => {
  it("未写入的行标记为 not_executed，且不进改动计数", () => {
    const review = buildRangeChangeReview(
      base({
        range: "A1:A3",
        before: snapshot([["a"], ["b"], ["c"]], [[null], [null], [null]]),
        after: snapshot([["A"], ["b"], ["c"]], [[null], [null], [null]]),
        appliedRows: 1
      })
    );
    expect(review.rows).toHaveLength(3);
    expect(review.rows[0].changed).toBe(true);
    expect(review.rows[1].kind).toBe("not_executed");
    expect(review.rows[2].kind).toBe("not_executed");
    expect(review.changedCount).toBe(1);
  });
});

describe("双栏变更对照 · 回读不符", () => {
  it("readback_failed / readback_mismatch 一律不得宣称成功", () => {
    expect(canClaimWriteSuccess("readback_failed")).toBe(false);
    expect(canClaimWriteSuccess("readback_mismatch")).toBe(false);
    expect(canClaimWriteSuccess("not_executed")).toBe(false);
    expect(canClaimWriteSuccess("ready")).toBe(false);
    expect(canClaimWriteSuccess("applied")).toBe(true);
  });

  it("写完了但独立验收没过（executed_unverified）也不得宣称成功", () => {
    // Excel 的 VerificationReport 有 executed_unverified 这一态：
    // 前后矩阵是真的，但验收条件没能全部核对，双栏不能替它背书。
    expect(canClaimWriteSuccess("executed_unverified")).toBe(false);
  });

  it("回读失败时 rows 仍然产出，但状态要带上失败", () => {
    const review = buildRangeChangeReview(
      base({ status: "readback_mismatch" })
    );
    expect(review.status).toBe("readback_mismatch");
    expect(review.rows).toHaveLength(1);
    expect(canClaimWriteSuccess(review.status)).toBe(false);
  });
});

describe("双栏变更对照 · 显示值未采集", () => {
  it("没有采集到计算值时不得把公式冒充值，只能标注未采集", () => {
    const review = buildRangeChangeReview(
      base({
        before: snapshot([[null]], [["=SUM(A:A)"]], false),
        after: snapshot([[null]], [["=SUM(A:A)"]], false)
      })
    );
    const row = review.rows[0];
    expect(row.before.valueCaptured).toBe(false);
    const text = cellDisplayText(row.before);
    expect(text.unresolved).toBe(true);
    expect(text.text).not.toContain("=SUM");
    expect(text.text).toBe("显示值未采集");
  });

  it("未采集状态下只按公式比对", () => {
    const review = buildRangeChangeReview(
      base({
        before: snapshot([[null]], [["=A1"]], false),
        after: snapshot([[null]], [["=A1"]], false)
      })
    );
    expect(review.rows[0].changed).toBe(false);
  });
});

describe("双栏变更对照 · 限额", () => {
  it("超出 maxRows 时 truncated=true 并带上真实计数", () => {
    const values: Values = Array.from({ length: 10 }, (_unused, index) => [index]);
    const formulas: Formulas = Array.from({ length: 10 }, () => [null]);
    const review = buildRangeChangeReview(
      base({
        range: "A1:A10",
        before: snapshot(values, formulas),
        after: snapshot(values.map((row) => [`c${row[0]}`]), formulas),
        limit: { maxRows: 3, maxTotalCells: 100 }
      })
    );
    expect(review.truncated).toBe(true);
    expect(review.rows).toHaveLength(3);
    // 计数必须反映**全部**改动，而不是被截断后的行数。
    expect(review.changedCount).toBe(10);
    expect(review.totalCount).toBe(10);
  });
});

describe("双栏变更对照 · 隐私出口", () => {
  it("摘要里不得携带任何单元格内容", () => {
    const review = buildRangeChangeReview(
      base({
        range: "A1:A2",
        before: snapshot([["机密甲"], ["机密乙"]], [[null], [null]]),
        after: snapshot([["x"], ["y"]], [[null], [null]])
      })
    );
    const serialized = JSON.stringify(changeReviewSummary(review));
    expect(serialized).not.toContain("机密甲");
    expect(serialized).not.toContain("机密乙");
    expect(serialized).toContain("\"changedCount\":2");
    // 摘要允许保留"范围"用于定位，但不允许保留逐格前后值。
    expect(changeReviewSummary(review as RangeChangeReview)).toEqual({
      sheet: "Sheet1",
      range: "A1:A2",
      status: "applied",
      phase: "actual",
      changedCount: 2,
      totalCount: 2,
      truncated: false
    });
  });
});
