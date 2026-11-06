import { describe, expect, it } from "vitest";
import { resolveQueryTarget } from "./queryTarget";

/**
 * "这次查哪几张表"的判据。
 *
 * 现场（2026-11-01）：用户只说「来源!A1:B5」，参数里没绑表，执行器就把四张已授权
 * 表全读了一遍 —— 结果卡列出四张表、三张各报一条缺字段，还把不完整说成"扫描已达
 * 上限"。这里钉住那条底线：**授权集合是上限，不是本次查询范围**；多张表却没绑定时
 * 必须 refuse，绝不退回"全部已授权表"。
 */
const FOUR = ["来源", "结果1", "结果2", "结果3"];

describe("目标表判定", () => {
  it("区域带表名就是绑定，其余表不进本次查询", () => {
    expect(
      resolveQueryTarget({
        rangeSheet: "来源",
        candidates: FOUR,
        activeWorksheet: "结果1"
      })
    ).toEqual({ sheets: ["来源"], issue: null, issueKind: null });
    // 与 `sheets` 一致时也不吵：两条路说的是同一张表。
    expect(
      resolveQueryTarget({
        sheets: ["来源"],
        rangeSheet: "来源",
        candidates: FOUR
      })
    ).toEqual({ sheets: ["来源"], issue: null, issueKind: null });
  });

  it("多张可查表但没有绑定：拒绝，并给出可选表名与三种写法", () => {
    const target = resolveQueryTarget({ candidates: FOUR });
    expect(target.sheets).toEqual([]);
    expect(target.issueKind).toBe("unbound");
    expect(target.issue).toContain("授权范围只是可访问的上限");
    for (const name of FOUR) expect(target.issue).toContain(`「${name}」`);
    // 三种写法都要能在提示里看到，模型才改得对。
    expect(target.issue).toContain("sheets");
    expect(target.issue).toContain("来源!A1:B5");
    expect(target.issue).toContain('"active"');
  });

  it("显式跨表：列全了就照列，顺序与重复由调用方决定后被规范化", () => {
    expect(
      resolveQueryTarget({ sheets: ["来源", "结果1"], candidates: FOUR }).sheets
    ).toEqual(["来源", "结果1"]);
    // 重复名字去重，但不改顺序：账本与结果卡的来源顺序要可复现。
    expect(
      resolveQueryTarget({
        sheets: ["结果1", "来源", "结果1"],
        candidates: FOUR
      }).sheets
    ).toEqual(["结果1", "来源"]);
  });

  it("越权与自相矛盾都按拒绝处理，且各自有独立判别点", () => {
    const outside = resolveQueryTarget({ sheets: ["别的表"], candidates: FOUR });
    expect(outside.issueKind).toBe("outside");
    expect(outside.issue).toContain("不在可查询的范围内");

    const outsideRange = resolveQueryTarget({
      rangeSheet: "别的表",
      candidates: FOUR
    });
    expect(outsideRange.issueKind).toBe("outside");

    // 区域说 A、sheets 说 B：不能挑一个用，两个都不是"用户要的"。
    const conflict = resolveQueryTarget({
      sheets: ["结果1"],
      rangeSheet: "来源",
      candidates: FOUR
    });
    expect(conflict.issueKind).toBe("conflict");
    expect(conflict.issue).toContain("不一致");

    // scope=active 与指向别处的区域同样矛盾。
    const activeConflict = resolveQueryTarget({
      scope: "active",
      rangeSheet: "来源",
      candidates: FOUR,
      activeWorksheet: "结果1"
    });
    expect(activeConflict.issueKind).toBe("conflict");
  });

  it("scope=active 只查当前表", () => {
    expect(
      resolveQueryTarget({
        scope: "active",
        candidates: FOUR,
        activeWorksheet: "结果1"
      })
    ).toEqual({ sheets: ["结果1"], issue: null, issueKind: null });
    // 当前表不在可查范围内 ⇒ 越权，而不是"那就都查一遍"。
    expect(
      resolveQueryTarget({
        scope: "active",
        candidates: ["来源"],
        activeWorksheet: "结果1"
      }).issueKind
    ).toBe("outside");
  });

  it("只有一张可查表时不必强绑：没有歧义", () => {
    expect(
      resolveQueryTarget({ candidates: ["来源"], activeWorksheet: "来源" })
    ).toEqual({ sheets: ["来源"], issue: null, issueKind: null });
  });

  it("没有可查表时明确说没有，而不是返回空清单让下游读空气", () => {
    const target = resolveQueryTarget({ candidates: [] });
    expect(target.issueKind).toBe("empty");
  });
});
