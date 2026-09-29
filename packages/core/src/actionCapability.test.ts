/**
 * 共享夹具驱动的行为锁。
 *
 * 用例来自 `config/action_capability_cases.json`，Python 侧
 * `server/tests/test_action_capability_shared_cases.py` 加载同一份文件、
 * 断言同一批期望值。**改这里的期望值必须改夹具，改夹具会让两侧同时生效** ——
 * 单边修改实现会立刻有一侧变红。
 *
 * 本文件不重复 Python 侧已有的红线用例，它只负责"两份实现是否一致"，
 * 以及 TS 侧独有的宿主适配器接线（`config/capabilities.json` 的导出视图）。
 */

import { describe, expect, it } from "vitest";

import cases from "../../../config/action_capability_cases.json";
import {
  actionCapabilityKeys,
  intersectActionCapabilities,
  isActionAllowedByKeys,
  matchesVersionPrefix,
  parseHostId,
  resolveAllowedActionTypes,
  resolveHostActionBound,
  type CapabilityMatrix,
  type HostActionTable
} from "./actionCapability";

const MATRIX = cases.matrix as unknown as CapabilityMatrix;

describe("共享夹具：版本前缀", () => {
  it.each(cases.versionPrefixCases)(
    "$version 对前缀 $prefix → $expected",
    ({ version, prefix, expected }) => {
      expect(matchesVersionPrefix(version, prefix)).toBe(expected);
    }
  );
});

describe("共享夹具：宿主标识", () => {
  it.each(cases.hostIdCases)("$input → $expected", ({ input, expected }) => {
    expect(parseHostId(input)).toBe(expected);
  });
});

describe("共享夹具：动作上界", () => {
  it.each(cases.boundCases)(
    "$hostId @ $hostVersion 的上界",
    ({ hostId, hostVersion, expectedActions }) => {
      const bound = resolveHostActionBound(MATRIX, hostId, hostVersion);
      expect(Object.keys(bound).sort()).toEqual([...expectedActions].sort());
    }
  );
});

describe("共享夹具：上界 ∩ 自报", () => {
  it.each(cases.intersectionCases)("$name", ({ hostId, hostVersion, claimed, expectedAllowed }) => {
    const allowed = resolveAllowedActionTypes(
      MATRIX,
      hostId,
      hostVersion,
      claimed as string[] | null
    );
    expect([...allowed].sort()).toEqual([...expectedAllowed].sort());
  });
});

describe("共享夹具：动作键", () => {
  const allowed = new Set(["clearRange:contents", "writeValues"]);

  it.each(cases.actionKeyCases)(
    "$actionType ($variant) → $expected",
    ({ actionType, variant, expected }) => {
      expect(
        isActionAllowedByKeys(
          allowed,
          actionType,
          variant as string | null
        )
      ).toBe(expected);
    }
  );
});

describe("包装函数与底层判定一致", () => {
  it("actionCapabilityKeys 产出的键与实际判定一致", () => {
    for (const actionType of ["writeValues", "clearRange"]) {
      const keys = actionCapabilityKeys(
        actionType,
        actionType === "clearRange" ? "contents" : null
      );
      const set = new Set(keys);
      expect(
        isActionAllowedByKeys(set, actionType, actionType === "clearRange" ? "contents" : null)
      ).toBe(true);
    }
  });

  it("intersectActionCapabilities 与 resolveAllowedActionTypes 对同一上界等价", () => {
    const bound: HostActionTable = resolveHostActionBound(MATRIX, "wps-et", "12.0");
    const claimed = ["writeValues", "createChart"];
    expect([...intersectActionCapabilities(bound, claimed)].sort()).toEqual(
      [...resolveAllowedActionTypes(MATRIX, "wps-et", "12.0", claimed)].sort()
    );
  });
});

describe("红线：自报只能收窄", () => {
  const bound = resolveHostActionBound(MATRIX, "wps-et", "12.0");
  const supported = Object.entries(bound)
    .filter(([, capability]) => capability.state === "supported")
    .map(([action]) => action);

  it("自报'全都有'时结果仍只是上界里 supported 的部分", () => {
    const everything = [
      ...Object.keys(bound),
      "createPivotTable",
      "addImage",
      "deleteWorksheet",
      "setHyperlink"
    ];
    expect([...intersectActionCapabilities(bound, everything)].sort()).toEqual(
      [...supported].sort()
    );
  });

  it("自报含未经验收动作时不放开", () => {
    expect(bound.writeFormulas.state).toBe("unverified");
    expect(bound.createChart.state).toBe("unsupported");
    expect([
      ...intersectActionCapabilities(bound, ["writeFormulas", "createChart"])
    ]).toEqual([]);
  });

  it("未知宿主得到空上界与空交集", () => {
    for (const host of [null, "", "office-online", "wps", "Excel-Office-Web"]) {
      expect(resolveHostActionBound(MATRIX, host, "12.0")).toEqual({});
      expect(
        resolveAllowedActionTypes(MATRIX, host, "12.0", ["writeValues"]).size
      ).toBe(0);
    }
  });

  it("clearRange 出现不等于全部变体放开", () => {
    const allowed = new Set(["clearRange:contents"]);
    expect(isActionAllowedByKeys(allowed, "clearRange", "contents")).toBe(true);
    for (const variant of ["formats", "all", "hyperlinks"]) {
      expect(isActionAllowedByKeys(allowed, "clearRange", variant)).toBe(false);
    }
  });
});
