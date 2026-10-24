import { describe, expect, it } from "vitest";
import type { WpsObjectIdentity } from "./wpsObjectIdentity";
import { wpsConfirmationTargetKey, wpsConfirmationTargetLabel } from "./wpsClarificationTargetConfirmation";

const target = {
  workbookName: "样本.xlsx", workbookFullName: "C:\\Private\\样本.xlsx", workbookPath: "C:\\Private",
  workbookRef: {}, sheets: [{ index: 1, name: "样本", id: "code-1", ref: {} }]
} as WpsObjectIdentity;

describe("显式澄清范围确认的内存绑定（不是身份同一性证明）", () => {
  it("新包装对象不改变勾选所指的元数据，显示标签不泄露路径", () => {
    expect(wpsConfirmationTargetKey({ ...target, workbookRef: {}, sheets: [{ ...target.sheets[0], ref: {} }] } as WpsObjectIdentity))
      .toBe(wpsConfirmationTargetKey(target));
    expect(wpsConfirmationTargetLabel(target)).toBe("样本.xlsx · 样本");
    expect(wpsConfirmationTargetLabel(target)).not.toContain("Private");
  });
  it("不同完整路径不能沿用勾选", () => {
    expect(wpsConfirmationTargetKey({ ...target, workbookFullName: "D:\\Other\\样本.xlsx" })).not.toBe(wpsConfirmationTargetKey(target));
  });
  it("工作表身份、名称、索引变化都不能沿用勾选", () => {
    for (const patch of [{ id: "code-2" }, { name: "另一张" }, { index: 2 }]) {
      expect(wpsConfirmationTargetKey({ ...target, sheets: [{ ...target.sheets[0], ...patch }] })).not.toBe(wpsConfirmationTargetKey(target));
    }
  });
  it("缺少保存路径或授权工作表时不提供绑定", () => {
    expect(wpsConfirmationTargetKey({ ...target, workbookFullName: "样本.xlsx" })).toBeNull();
    expect(wpsConfirmationTargetKey({ ...target, sheets: [] })).toBeNull();
  });
  /*
   * 绑定键 = 工作簿身份 + **授权集合**。集合的身份不含"先点了哪一张"。
   * 与 Excel 侧同一份判据（共享层 `authorizedTargetEvidenceKey`）：
   * 两端都必须与点选顺序无关，否则同一份授权会产出两个键，
   * 分析前置检查把它读成"证据冲突"。
   */
  it("同一组工作表换个点选顺序，绑定键不变；成员变化仍然换键", () => {
    const sheet1 = { ...target.sheets[0], index: 1, name: "Sheet1", id: "code-1" };
    const sheet10 = { ...target.sheets[0], index: 10, name: "Sheet10", id: "code-10" };
    const sheet11 = { ...target.sheets[0], index: 11, name: "Sheet11", id: "code-11" };
    const picked = wpsConfirmationTargetKey({ ...target, sheets: [sheet11, sheet1, sheet10] } as WpsObjectIdentity);
    const readBack = wpsConfirmationTargetKey({ ...target, sheets: [sheet1, sheet10, sheet11] } as WpsObjectIdentity);
    expect(picked).toBe(readBack);
    expect(wpsConfirmationTargetKey({ ...target, sheets: [sheet1, sheet10] } as WpsObjectIdentity))
      .not.toBe(picked);
  });
});
