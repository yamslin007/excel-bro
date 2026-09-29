// 折叠规则必须与 excel.ts:3766 逐字符一致，否则已落盘账本里的旧记录
// 会因身份对不上而失去续跑与撤销入口。这里按当时的实际行为立断言。
import { describe, expect, it } from "vitest";

import {
  locatorFromOfficeIdentity,
  officeIdentityFromLocator,
  officeIdentityFromUrl,
  sameLocator
} from "./locator";

describe("officeIdentityFromUrl", () => {
  it("加 url: 前缀、反斜杠转正斜杠、转小写", () => {
    expect(officeIdentityFromUrl("C:\\Users\\Jia\\Book1.xlsx")).toBe(
      "url:c:/users/jia/book1.xlsx"
    );
  });

  it("先 trim 再归一", () => {
    expect(officeIdentityFromUrl("  https://Example.com/A.xlsx  ")).toBe(
      "url:https://example.com/a.xlsx"
    );
  });

  it("空串与纯空白返回 null，不生成只有前缀的身份", () => {
    expect(officeIdentityFromUrl("")).toBeNull();
    expect(officeIdentityFromUrl("   ")).toBeNull();
  });
});

describe("officeIdentityFromLocator", () => {
  it("office-url 折叠为出网字符串", () => {
    expect(
      officeIdentityFromLocator({ kind: "office-url", value: "C:\\A\\B.xlsx" })
    ).toBe("url:c:/a/b.xlsx");
  });

  it("非 Office 通道抛错，不返回能通过服务端校验的伪身份", () => {
    expect(() =>
      officeIdentityFromLocator({ kind: "wps-path", value: "C:\\A\\B.xlsx" })
    ).toThrow(/仅 Office 通道/);
    expect(() =>
      officeIdentityFromLocator({ kind: "folder-relative", value: "a/b.xlsx" })
    ).toThrow(/仅 Office 通道/);
  });
});

describe("locatorFromOfficeIdentity", () => {
  it("解析旧账本记录", () => {
    expect(locatorFromOfficeIdentity("url:c:/a/b.xlsx")).toEqual({
      kind: "office-url",
      value: "c:/a/b.xlsx"
    });
  });

  it("缺前缀返回 null，不猜成 office-url", () => {
    expect(locatorFromOfficeIdentity("c:/a/b.xlsx")).toBeNull();
    expect(locatorFromOfficeIdentity("")).toBeNull();
  });

  it("只有前缀返回 null", () => {
    expect(locatorFromOfficeIdentity("url:")).toBeNull();
  });

  it("值里再含 url: 时只剥掉首个前缀", () => {
    // 用 slice(prefix.length) 而非 replace，后者在此处同样只替换首个匹配，
    // 但一旦有人改成 replaceAll 就会静默截断路径，故在此锁住。
    expect(locatorFromOfficeIdentity("url:https://h/url:x.xlsx")).toEqual({
      kind: "office-url",
      value: "https://h/url:x.xlsx"
    });
  });

  it("往返不改变已归一的身份", () => {
    const identity = "url:c:/users/jia/book1.xlsx";
    const locator = locatorFromOfficeIdentity(identity);
    expect(locator).not.toBeNull();
    expect(officeIdentityFromLocator(locator!)).toBe(identity);
  });
});

describe("sameLocator", () => {
  it("kind 不同一律不等，即便 value 字符串相同", () => {
    expect(
      sameLocator(
        { kind: "office-url", value: "C:\\A\\B.xlsx" },
        { kind: "wps-path", value: "C:\\A\\B.xlsx" }
      )
    ).toBe(false);
  });

  it("office-url 按折叠后比对", () => {
    expect(
      sameLocator(
        { kind: "office-url", value: "C:\\A\\B.xlsx" },
        { kind: "office-url", value: "c:/a/b.xlsx" }
      )
    ).toBe(true);
  });

  it("wps-path 大小写与分隔符不敏感", () => {
    expect(
      sameLocator(
        { kind: "wps-path", value: "C:\\A\\B.xlsx" },
        { kind: "wps-path", value: "c:/a/b.xlsx" }
      )
    ).toBe(true);
  });

  it("folder-relative 原样比对，不做大小写归一", () => {
    expect(
      sameLocator(
        { kind: "folder-relative", value: "A/B.xlsx" },
        { kind: "folder-relative", value: "a/b.xlsx" }
      )
    ).toBe(false);
    expect(
      sameLocator(
        { kind: "folder-relative", value: "a/b.xlsx" },
        { kind: "folder-relative", value: "a/b.xlsx" }
      )
    ).toBe(true);
  });
});
