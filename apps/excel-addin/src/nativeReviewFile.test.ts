import { describe, expect, it } from "vitest";
import { nativeColumnName, nativeReviewSeed } from "./nativeReviewFile";
import { readSeedEntries as unzipSeed } from "./nativeReviewTestHelpers";
describe("bounded local OOXML seed", () => {
  it("preserves absolute coordinates and cell types, without evaluating formula-like text", () => {
    const files = unzipSeed(nativeReviewSeed("Preview", 7, 28, [["LOCAL_PRIVATE_SENTINEL", "=1+2", false, 99, null, ""]], 100000));
    expect(Object.keys(files)).toHaveLength(5);
    const sheet = files["xl/worksheets/sheet1.xml"];
    expect(sheet).toContain('r="AB7" t="inlineStr"');
    expect(sheet).toContain("LOCAL_PRIVATE_SENTINEL");
    expect(sheet).toContain('r="AC7" t="inlineStr"');
    expect(sheet).toContain('r="AD7" t="b"><v>0</v>');
    expect(sheet).toContain('r="AE7"><v>99</v>');
    expect(sheet).not.toContain('<f>');
    expect(sheet).not.toContain('r="AF7"');
    expect(sheet).toContain('r="AG7" t="inlineStr"');
  });
  it("escapes XML, enforces the byte cap and refuses illegal input", () => {
    expect(unzipSeed(nativeReviewSeed("A&B", 1, 1, [["<&\""]], 100000))["xl/worksheets/sheet1.xml"]).toContain("&lt;&amp;&quot;");
    expect(() => nativeReviewSeed("Preview", 1, 1, [["x"]], 10)).toThrow("限额");
    expect(() => nativeReviewSeed("Preview", 1, 1, [[Infinity]], 10000)).toThrow("非有限");
    expect(() => nativeReviewSeed("Preview", 1, 1, [["\u0001"]], 10000)).toThrow("控制字符");
    expect(nativeColumnName(16384)).toBe("XFD");
  });
});
