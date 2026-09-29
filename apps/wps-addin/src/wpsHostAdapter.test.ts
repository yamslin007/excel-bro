import { describe, expect, it } from "vitest";
import type { WpsApplication, WpsPluginStorage } from "./wpsHost";
import type { WpsReadAuthorization } from "./wpsWorkbookReader";
import {
  createWpsHostAdapter,
  createWpsKeyValueStore,
  resolveWpsCapability
} from "./wpsHostAdapter";

function pluginStorage(withRemove = true): WpsPluginStorage & { map: Map<string, unknown> } {
  const map = new Map<string, unknown>();
  const storage: WpsPluginStorage & { map: Map<string, unknown> } = {
    map,
    getItem: (key) => (map.has(key) ? map.get(key) : undefined),
    setItem: (key, value) => {
      map.set(key, value);
    }
  };
  if (withRemove) storage.removeItem = (key) => void map.delete(key);
  return storage;
}

// 传 null 表示宿主没报版本。不能用 undefined —— 那会命中默认参数值。
function application(version: string | null = "12.0"): WpsApplication {
  return {
    Version: version ?? undefined,
    PluginStorage: pluginStorage(),
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  };
}

describe("createWpsKeyValueStore", () => {
  it("非字符串返回折成 null，而不是 String() 化", () => {
    const raw = pluginStorage();
    const store = createWpsKeyValueStore(raw);
    raw.map.set("num", 42);
    raw.map.set("obj", { a: 1 });
    raw.map.set("undef", undefined);
    // String(undefined) === "undefined" 会被当成一个真实存过的值传下去。
    expect(store.getItem("num")).toBeNull();
    expect(store.getItem("obj")).toBeNull();
    expect(store.getItem("undef")).toBeNull();
    expect(store.getItem("missing")).toBeNull();
  });

  it("字符串原样返回，空串视为不存在", () => {
    const raw = pluginStorage();
    const store = createWpsKeyValueStore(raw);
    store.setItem("k", "v");
    expect(store.getItem("k")).toBe("v");
    raw.map.set("empty", "");
    expect(store.getItem("empty")).toBeNull();
  });

  it("宿主未提供 removeItem 时退化为写空串，读侧仍报不存在", () => {
    const raw = pluginStorage(false);
    const store = createWpsKeyValueStore(raw);
    store.setItem("k", "v");
    store.removeItem("k");
    expect(raw.map.get("k")).toBe("");
    expect(store.getItem("k")).toBeNull();
  });
});

describe("resolveWpsCapability", () => {
  it("白名单内 + 开发验收通道 → supported", () => {
    expect(resolveWpsCapability("writeValues", "12.0")).toEqual({ state: "supported" });
    expect(resolveWpsCapability("writeFormulas", "12.0")).toEqual({ state: "supported" });
  });

  it("白名单内但版本未验收 → unverified，且说明带版本号", () => {
    const result = resolveWpsCapability("writeValues", "11.0.2");
    expect(result.state).toBe("unverified");
    expect(result.state === "unverified" && result.reason).toContain("11.0.2");
  });

  it("实机返回的 12.0 在开发模式可规划白名单动作", () => {
    expect(resolveWpsCapability("writeValues", "12.0")).toEqual({ state: "supported" });
  });

  it("版本缺失 → unverified，不当成已验收", () => {
    expect(resolveWpsCapability("writeValues", null).state).toBe("unverified");
  });

  it("白名单外的动作 → unverified 而非 unsupported", () => {
    // 配置只记录"哪些动作已验收"，不记录宿主 API 是否具备。
    // 报 unsupported 等于断言"WPS 没有这个 API"，拿不出证据。
    const result = resolveWpsCapability("createChart", "12.0");
    expect(result.state).toBe("unverified");
    expect(result.state === "unverified" && result.reason).toContain("createChart");
  });

  it("带冒号限定的白名单项按动作前缀匹配", () => {
    // 配置里是 "clearRange:contents"。
    expect(resolveWpsCapability("clearRange", "12.0")).toEqual({ state: "supported" });
  });
});

describe("createWpsHostAdapter", () => {
  it("原语读写前复核工作簿身份，拒绝同名表的另一份工作簿", async () => {
    const app = application();
    const sheet = {
      Name: "明细",
      Range: () => ({ Rows: { Count: 1 }, Columns: { Count: 1 }, Value2: "旧值", Formula: "旧值" })
    };
    const workbook = {
      Name: "原文件.xlsx",
      FullName: "C:\\Data\\原文件.xlsx",
      Path: "C:\\Data",
      Worksheets: { Count: 1, Item: () => sheet }
    };
    app.ActiveWorkbook = workbook;
    const authorization: WpsReadAuthorization = {
      workbookName: workbook.Name,
      workbookFullName: workbook.FullName,
      workbookPath: workbook.Path,
      workbookRef: workbook,
      sheets: [{ index: 1, name: "明细" }]
    };
    const adapter = createWpsHostAdapter(app, authorization);
    workbook.FullName = "C:\\Other\\原文件.xlsx";

    await expect(adapter.runBatch((ops) => ops.readValues({ name: "明细" }, "A1")))
      .rejects.toThrow("工作簿身份发生变化");
  });

  it("未授权时 locator 为 undefined，且拒绝执行任何原语", async () => {
    const adapter = createWpsHostAdapter(application());
    expect(adapter.id).toBe("wps-et");
    expect(adapter.locator()).toBeUndefined();
    await expect(adapter.runBatch(async () => "不该跑到这里")).rejects.toThrow(
      "尚未完成工作表授权"
    );
  });

  it("hostVersion 给出完整版本串，不截断到前缀", () => {
    // 能力矩阵按前缀匹配，但截断会让实机记录失去精确版本号。
    expect(createWpsHostAdapter(application("12.1.0.16120")).hostVersion()).toBe(
      "12.1.0.16120"
    );
    expect(createWpsHostAdapter(application(null)).hostVersion()).toBeNull();
  });

  it("storage 经过强制转换层", () => {
    const app = application();
    const adapter = createWpsHostAdapter(app);
    (app.PluginStorage as unknown as { map: Map<string, unknown> }).map.set("k", 7);
    expect(adapter.storage.getItem("k")).toBeNull();
  });
});
