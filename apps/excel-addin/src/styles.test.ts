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

describe("folder selection footer", () => {
  it("keeps the selection state on one line and moves the persistent hint into help", () => {
    const source = css("./styles.css");
    const component = css("./App.tsx");

    expect(source).toContain(".folder-selection-summary {");
    expect(source).toContain(".folder-selection-help summary {");
    expect(source).toContain(".folder-selection-help > span {");
    const helpRule =
      source.match(/\.folder-selection-help \{([^}]*)\}/s)?.[1] ?? "";
    const helpNoteRule =
      source.match(/\.folder-selection-help > span \{([^}]*)\}/s)?.[1] ?? "";
    expect(helpRule).toContain("position: static");
    expect(helpNoteRule).toContain("left: var(--scope-popover-padding)");
    expect(helpNoteRule).toContain(
      "max-width: calc(100% - (var(--scope-popover-padding) * 2))"
    );
    expect(helpNoteRule).not.toContain("right: 0");
    expect(helpNoteRule).not.toContain("100vw");
    expect(source).toContain("min-height: 68px");
    expect(source).toContain("--scope-popover-padding: var(--spacing-md)");
    expect(source).toContain(".scope-popover-scroll {");
    expect(source).toContain("flex-direction: column");
    expect(source).toContain("flex: 1 1 auto");
    expect(source).toContain("background: var(--color-background)");
    const footerRule = source.match(/\.folder-selection-footer \{([^}]*)\}/s)?.[1] ?? "";
    expect(footerRule).toContain("position: relative");
    expect(footerRule).toContain("flex: 0 0 auto");
    expect(footerRule).not.toContain("position: sticky");
    expect(footerRule).not.toContain("margin-bottom");
    expect(source).not.toContain(".folder-selection-footer small {");
    expect(component).toContain('<div className="scope-popover-scroll">');
    expect(component).toContain('details className="folder-selection-help"');
    expect(component).toContain('aria-label="查看工作表选择保留说明"');
    expect(component).not.toContain("<small>切换文件不会丢失已勾选的工作表</small>");
  });
});

describe("composer shortcut hint", () => {
  it("keeps the concise hint inside the composer toolbar", () => {
    const source = css("./styles.css");
    const component = css("./App.tsx");

    expect(component).toContain('<span\n                className="composer-shortcut-hint"');
    expect(component).toContain(
      "Enter 发送 · Shift + Enter 换行 · 可粘贴截图"
    );
    expect(component).not.toContain("写入操作会先预览");
    expect(source).toContain(".composer-shortcut-hint {");
    expect(source).toContain("grid-template-columns: auto minmax(0, 1fr) auto");
    expect(source).not.toContain(".composer > span {");
  });
});

describe("first-run guide theme", () => {
  it("uses the active theme palette instead of a fixed green card", () => {
    const source = css("./styles.css");
    const widenStart = source.indexOf(".widen-pane-guide {");
    const widenEnd = source.indexOf("@keyframes widen-guide-in", widenStart);
    const widenSection = source.slice(widenStart, widenEnd);
    const pulseStart = source.indexOf("@keyframes model-guide-pulse");
    const pulseEnd = source.indexOf(".first-model-guide {", pulseStart);
    const pulseSection = source.slice(pulseStart, pulseEnd);

    expect(widenSection).toContain(
      "border: 1px solid var(--color-border-strong)"
    );
    expect(widenSection).toContain("var(--color-background)");
    expect(widenSection).toContain("var(--color-text-primary)");
    expect(widenSection).toContain("var(--color-text-secondary)");
    expect(widenSection).toContain("var(--color-brand-text)");
    expect(widenSection).not.toMatch(
      /#bcd8c5|#fbfefc|#f7fcf8|#254b32|#6e7e73|#d8e3db|#66746b/
    );
    expect(pulseSection).toContain("var(--color-brand-text)");
    expect(pulseSection).not.toContain("rgba(52, 153, 84");
  });
});
