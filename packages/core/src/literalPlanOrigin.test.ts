import { describe, expect, it } from "vitest";
import { isTrustedLiteralPlan, stampLiteralPlanOrigin } from "./literalPlanOrigin";
import { judgePlanSourceGate } from "./planSourceAnchor";

const literal = () => ({ id: "literal-1", actions: [{ type: "writeValues", sheet: "目标", range: "A1", values: [["完成"]] }] });
describe("trusted literal compiler provenance", () => {
  it("permits an exact literal without a data anchor, including immutable spreads", () => {
    const plan = {...stampLiteralPlanOrigin(literal(), "literal_range_edit")};
    expect(isTrustedLiteralPlan(plan)).toBe(true);
    expect(judgePlanSourceGate({plan, anchor:null, authorizedSheets:["目标"], liveFingerprint:null})).toEqual({verdict:"ok",basis:"literal"});
  });
  it("does not exempt ordinary computed writes or a JSON-authored origin", () => {
    const plan = {...literal(), planOrigin:"literal_range_edit"};
    expect(isTrustedLiteralPlan(plan)).toBe(false);
    expect(judgePlanSourceGate({plan,anchor:null,authorizedSheets:["目标"],liveFingerprint:null}).verdict).toBe("refused");
    expect(isTrustedLiteralPlan(stampLiteralPlanOrigin(literal(), undefined))).toBe(false);
  });
  it("cannot survive history serialization or altered values/targets", () => {
    const plan = stampLiteralPlanOrigin(literal(), "literal_range_edit");
    expect(isTrustedLiteralPlan(JSON.parse(JSON.stringify(plan)))).toBe(false);
    expect(isTrustedLiteralPlan({...plan, actions:[{...plan.actions[0],sheet:"未授权表"}]})).toBe(false);
    plan.actions[0].values[0][0]="other";
    expect(isTrustedLiteralPlan(plan)).toBe(false);
  });
  it("does not attest result references, mixed actions or computed matrices", () => {
    for (const plan of [
      {...literal(),localResultRef:"res_old"},
      {...literal(),actions:[...literal().actions,...literal().actions]},
      {...literal(),actions:[{...literal().actions[0],values:[[222],[500]]}]}
    ]) expect(isTrustedLiteralPlan(stampLiteralPlanOrigin(plan,"literal_range_edit"))).toBe(false);
  });
});
