import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

function css(name: string): string {
  return readFileSync(new URL(name, import.meta.url), "utf8");
}

describe("typography system", () => {
  it("defines the Fluent type ramp and Office font stack", () => {
    const source = css("./styles.css");

    expect(source).toContain('--font-size-caption: 12px');
    expect(source).toContain('--font-size-body: 14px');
    expect(source).toContain('--font-size-subtitle: 16px');
    expect(source).toContain('--font-size-title: 20px');
    expect(source).toContain('"Segoe UI Variable Text"');
    expect(source).not.toContain("Inter,");
  });

  it("does not reintroduce literal text sizes below 12px", () => {
    for (const name of ["./styles.css", "./focus.css"]) {
      const source = css(name);
      const declarations = [
        ...source.matchAll(/(?:^|\s)font-size:\s*([0-9.]+)px/gm),
        ...source.matchAll(/(?:^|\s)font:\s*[^;\n]*?\s([0-9.]+)px(?:\s*\/|\/)/gm)
      ];
      const undersized = declarations
        .map((match) => Number(match[1]))
        .filter((value) => value < 12);

      expect(undersized, name).toEqual([]);
    }
  });
});
