import { describe, expect, it } from "vitest";
import {
  createOfficeKeyValueStore,
  officeHostRequestFields,
  officeReportedActionTypes,
  resolveOfficeCapability
} from "./officeHostAdapter";
import { OFFICE_EXECUTED_ACTION_TYPES } from "./officeActions";

function memoryStorage(): Storage {
  const map = new Map<string, string>();
  return {
    get length() {
      return map.size;
    },
    clear: () => map.clear(),
    getItem: (key) => (map.has(key) ? map.get(key)! : null),
    key: (index) => [...map.keys()][index] ?? null,
    removeItem: (key) => void map.delete(key),
    setItem: (key, value) => void map.set(key, value)
  };
}

describe("createOfficeKeyValueStore", () => {
  it("透传 localStorage", () => {
    const store = createOfficeKeyValueStore({ localStorage: memoryStorage() });
    expect(store.getItem("k")).toBeNull();
    store.setItem("k", "v");
    expect(store.getItem("k")).toBe("v");
    store.removeItem("k");
    expect(store.getItem("k")).toBeNull();
  });

  it("localStorage 不可用时读取走默认值，但写入必须抛错", () => {
    // 读不到走默认值是可接受的降级；写入静默失败不行 ——
    // 用户会以为设置保存成功了。
    const store = createOfficeKeyValueStore({});
    expect(store.getItem("k")).toBeNull();
    expect(() => store.setItem("k", "v")).toThrow("无法保存设置");
    expect(() => store.removeItem("k")).not.toThrow();
  });
});

describe("resolveOfficeCapability", () => {
  it("已实现的动作报 supported", () => {
    expect(resolveOfficeCapability("writeValues")).toEqual({ state: "supported" });
    expect(resolveOfficeCapability("createChart")).toEqual({ state: "supported" });
  });

  it("空动作类型报 unsupported", () => {
    expect(resolveOfficeCapability("  ").state).toBe("unsupported");
  });
});

describe("officeReportedActionTypes", () => {
  it("未提供清单时报空集（不声明即不做）", () => {
    expect(officeReportedActionTypes(undefined)).toEqual([]);
  });

  it("透传并保持顺序", () => {
    expect(officeReportedActionTypes(["b", "a"])).toEqual(["b", "a"]);
  });

  it("去重、去空白，但不改写动作名", () => {
    expect(officeReportedActionTypes(["a", "a", " b ", "", "  "])).toEqual([
      "a",
      "b"
    ]);
  });

  it("实际清单可直接作为自报内容，不会掺入空值", () => {
    const reported = officeReportedActionTypes(OFFICE_EXECUTED_ACTION_TYPES);
    expect(reported).toHaveLength(OFFICE_EXECUTED_ACTION_TYPES.length);
    expect(reported.every((entry) => entry.trim() === entry && entry)).toBe(true);
  });
});

describe("officeHostRequestFields", () => {
  it("不在 Excel 内时不声明宿主（文件夹/浏览器预览不得借 Office 名义放宽）", () => {
    expect(officeHostRequestFields(false, "16.0.18025.20000")).toEqual({
      hostId: null,
      hostVersion: null,
      selfReportedActions: []
    });
  });

  it("在 Excel 内时报宿主标识与版本，并带上执行清单", () => {
    const fields = officeHostRequestFields(true, "16.0.18025.20000");
    expect(fields.hostId).toBe("excel-office");
    expect(fields.hostVersion).toBe("16.0.18025.20000");
    expect(fields.selfReportedActions).toEqual([
      ...OFFICE_EXECUTED_ACTION_TYPES
    ]);
  });

  it("拿不到版本时仍报宿主，但版本为 null —— 服务端上界为空，是少做而非多做", () => {
    const fields = officeHostRequestFields(true, null);
    expect(fields.hostId).toBe("excel-office");
    expect(fields.hostVersion).toBeNull();
  });

  it("版本是空白串时归一成 null，不留下能被前缀匹配命中的假版本", () => {
    expect(officeHostRequestFields(true, "   ").hostVersion).toBeNull();
    expect(officeHostRequestFields(true, " 16.0 ").hostVersion).toBe("16.0");
  });
});
