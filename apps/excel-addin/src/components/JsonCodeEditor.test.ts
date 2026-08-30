import { describe, expect, it } from "vitest";
import { highlightJsonText, validateJsonText } from "./JsonCodeEditor";

describe("validateJsonText", () => {
  it("accepts valid JSON", () => {
    const result = validateJsonText('{"title": "测试", "value": 1000}');
    expect(result.ok).toBe(true);
  });

  it("reports line and column for a missing comma", () => {
    const result = validateJsonText('{\n  "a": 1\n  "b": 2\n}');
    expect(result.ok).toBe(false);
    expect(result.line).toBe(3);
  });

  it("reports line and column for a trailing comma", () => {
    const result = validateJsonText('{\n  "a": 1,\n}');
    expect(result.ok).toBe(false);
    expect(result.line).toBe(3);
  });
});

describe("highlightJsonText", () => {
  it("wraps keys, strings and numbers in token spans", () => {
    const html = highlightJsonText('{"title": "测试", "value": 1000}');
    expect(html).toContain('<span class="tk-key">"title"</span>');
    expect(html).toContain('<span class="tk-str">"测试"</span>');
    expect(html).toContain('<span class="tk-num">1000</span>');
  });

  it("colors booleans and null separately", () => {
    const html = highlightJsonText('{"ok": true, "x": null}');
    expect(html).toContain('<span class="tk-bool">true</span>');
    expect(html).toContain('<span class="tk-null">null</span>');
  });

  it("escapes HTML inside string values", () => {
    const html = highlightJsonText('{"note": "<b>&</b>"}');
    expect(html).not.toContain("<b>");
    expect(html).toContain("&lt;b&gt;&amp;&lt;/b&gt;");
  });
});
