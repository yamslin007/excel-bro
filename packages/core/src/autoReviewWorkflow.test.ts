import { describe, expect, it } from "vitest";
import {
  type AutoReviewState,
  type AutoReviewIdentity,
  type AutoReviewAvailability,
  shouldAutoGenerateReview,
  markReviewGenerated,
  isReviewInvalidated,
  invalidateCurrentReview,
  resetAutoReviewState,
  serializeAuthorizationKey
} from "./autoReviewWorkflow";

describe("shouldAutoGenerateReview", () => {
  const identity: AutoReviewIdentity = {
    planId: "plan-1",
    workbookIdentity: "workbook-1",
    authorizationKey: '["Sheet1","Sheet2"]'
  };

  const available: AutoReviewAvailability = {
    available: true,
    reason: null
  };

  it("首次满足条件时应该生成", () => {
    const state: AutoReviewState = {
      generated: new Set(),
      current: null
    };

    expect(shouldAutoGenerateReview(state, identity, available)).toBe(true);
  });

  it("已生成过同一 planId 时不重复生成", () => {
    const state: AutoReviewState = {
      generated: new Set(["plan-1"]),
      current: null
    };

    expect(shouldAutoGenerateReview(state, identity, available)).toBe(false);
  });

  it("宿主不支持时不生成", () => {
    const state: AutoReviewState = {
      generated: new Set(),
      current: null
    };

    const unavailable: AutoReviewAvailability = {
      available: false,
      reason: "当前宿主不支持副本生成"
    };

    expect(shouldAutoGenerateReview(state, identity, unavailable)).toBe(false);
  });
});

describe("markReviewGenerated", () => {
  it("记录已生成的 planId", () => {
    const state: AutoReviewState = {
      generated: new Set(),
      current: null
    };

    const identity: AutoReviewIdentity = {
      planId: "plan-1",
      workbookIdentity: "workbook-1",
      authorizationKey: '["Sheet1"]'
    };

    const newState = markReviewGenerated(state, identity);

    expect(newState.generated.has("plan-1")).toBe(true);
    expect(newState.current).toEqual(identity);
  });

  it("不改变原状态对象", () => {
    const state: AutoReviewState = {
      generated: new Set(),
      current: null
    };

    const identity: AutoReviewIdentity = {
      planId: "plan-1",
      workbookIdentity: "workbook-1",
      authorizationKey: '["Sheet1"]'
    };

    markReviewGenerated(state, identity);

    expect(state.generated.size).toBe(0);
    expect(state.current).toBe(null);
  });
});

describe("isReviewInvalidated", () => {
  const original: AutoReviewIdentity = {
    planId: "plan-1",
    workbookIdentity: "workbook-1",
    authorizationKey: '["Sheet1","Sheet2"]'
  };

  it("planId 变化导致失效", () => {
    const state: AutoReviewState = {
      generated: new Set(["plan-1"]),
      current: original
    };

    const newIdentity: AutoReviewIdentity = {
      ...original,
      planId: "plan-2"
    };

    expect(isReviewInvalidated(state, newIdentity)).toBe(true);
  });

  it("workbookIdentity 变化导致失效", () => {
    const state: AutoReviewState = {
      generated: new Set(["plan-1"]),
      current: original
    };

    const newIdentity: AutoReviewIdentity = {
      ...original,
      workbookIdentity: "workbook-2"
    };

    expect(isReviewInvalidated(state, newIdentity)).toBe(true);
  });

  it("authorizationKey 变化导致失效", () => {
    const state: AutoReviewState = {
      generated: new Set(["plan-1"]),
      current: original
    };

    const newIdentity: AutoReviewIdentity = {
      ...original,
      authorizationKey: '["Sheet1","Sheet3"]'
    };

    expect(isReviewInvalidated(state, newIdentity)).toBe(true);
  });

  it("所有字段相同时不失效", () => {
    const state: AutoReviewState = {
      generated: new Set(["plan-1"]),
      current: original
    };

    const sameIdentity: AutoReviewIdentity = { ...original };

    expect(isReviewInvalidated(state, sameIdentity)).toBe(false);
  });

  it("没有当前审阅时返回 false", () => {
    const state: AutoReviewState = {
      generated: new Set(),
      current: null
    };

    expect(isReviewInvalidated(state, original)).toBe(false);
  });
});

describe("invalidateCurrentReview", () => {
  it("清除当前审阅但保留生成记录", () => {
    const state: AutoReviewState = {
      generated: new Set(["plan-1", "plan-2"]),
      current: {
        planId: "plan-2",
        workbookIdentity: "workbook-1",
        authorizationKey: '["Sheet1"]'
      }
    };

    const newState = invalidateCurrentReview(state);

    expect(newState.generated.size).toBe(2);
    expect(newState.generated.has("plan-1")).toBe(true);
    expect(newState.generated.has("plan-2")).toBe(true);
    expect(newState.current).toBe(null);
  });
});

describe("resetAutoReviewState", () => {
  it("清除所有状态", () => {
    const state = resetAutoReviewState();

    expect(state.generated.size).toBe(0);
    expect(state.current).toBe(null);
  });
});

describe("serializeAuthorizationKey", () => {
  it("序列化授权工作表列表", () => {
    const authorization = ["Sheet2", "Sheet1", "Sheet3"];
    const key = serializeAuthorizationKey(authorization);

    expect(key).toBe('["Sheet1","Sheet2","Sheet3"]');
  });

  it("相同工作表不同顺序产生相同键", () => {
    const key1 = serializeAuthorizationKey(["A", "B", "C"]);
    const key2 = serializeAuthorizationKey(["C", "A", "B"]);

    expect(key1).toBe(key2);
  });

  it("不同工作表产生不同键", () => {
    const key1 = serializeAuthorizationKey(["A", "B"]);
    const key2 = serializeAuthorizationKey(["A", "C"]);

    expect(key1).not.toBe(key2);
  });
});
