import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

function css(name: string): string {
  return readFileSync(new URL(name, import.meta.url), "utf8");
}

/*
 * 跨端读取另一个 workspace 的样式表。
 *
 * 两个坑：① `new URL("../x", import.meta.url)` 会把开头的 `..` 归一化掉、少退一层；
 * ② 用 `path.resolve(here, "../x")` 在 Git Bash 下同样会被吞掉一层（实测解析成
 * `excel-addin/wps-addin/...`）。`join` 逐段拼接才稳定。
 */
const HERE = dirname(fileURLToPath(import.meta.url));

function wpsStylesheet(): string {
  return readFileSync(join(HERE, "..", "..", "wps-addin", "src", "styles.css"), "utf8");
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

    expect(component).toMatch(
      /<span\r?\n\s+className="composer-shortcut-hint"/
    );
    expect(component).toContain(
      "Enter 发送 · Shift + Enter 换行 · 可粘贴截图"
    );
    expect(component).not.toContain("写入操作会先预览");
    expect(source).toContain(".composer-shortcut-hint {");
    expect(source).toContain("grid-template-columns: auto minmax(0, 1fr) auto");
    expect(source).not.toContain(".composer > span {");
  });
});

describe("model connection feedback", () => {
  it("keeps connection test feedback inside the editor before formula settings", () => {
    const component = css("./App.tsx");
    const actionsIndex = component.indexOf('className="connection-editor-actions"');
    const feedbackIndex = component.indexOf("connection-test-feedback");
    const formulaIndex = component.indexOf(
      "<strong>/function 公式模型</strong>",
      feedbackIndex
    );

    expect(actionsIndex).toBeGreaterThan(-1);
    expect(feedbackIndex).toBeGreaterThan(actionsIndex);
    expect(formulaIndex).toBeGreaterThan(feedbackIndex);
    expect(component).toContain('aria-live={');
  });

  it("uses the active theme palette throughout model settings", () => {
    const source = css("./styles.css");
    const settingsStart = source.indexOf(".settings-loading {");
    const settingsEnd = source.indexOf(".history-drawer {", settingsStart);
    const settingsSection = source.slice(settingsStart, settingsEnd);

    expect(settingsStart).toBeGreaterThan(-1);
    expect(settingsEnd).toBeGreaterThan(settingsStart);
    expect(settingsSection).toContain("var(--color-background)");
    expect(settingsSection).toContain("var(--color-surface)");
    expect(settingsSection).toContain("var(--color-text-primary)");
    expect(settingsSection).toContain("var(--color-text-secondary)");
    expect(settingsSection).toContain("var(--color-brand-bg)");
    expect(settingsSection).toContain("var(--color-brand-text)");
    expect(settingsSection).toContain("var(--color-border)");
    expect(settingsSection).not.toMatch(
      /#728077|#398557|#cfe0d3|#f2faf4|#dff1e4|#277246|#294a34|#6f7e74|#dce5de|#31483a|#2a4032|#78837c|#d4ded6|#5f9f74|#376448|#f4f8f5|#718078|#8dbb9b|#f2f9f4|#dcefe2|#2e4637|#367049|#66736a|#fff4e9|#9a6030/i
    );
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

describe("header popup layering", () => {
  it("keeps both hosts' header labels accessible without taking layout space", () => {
    for (const name of ["./App.tsx", "../../wps-addin/src/WpsAgentPanel.tsx"]) {
      const component = css(name);
      expect(component).toContain("<ModelPickerIcon />");
      expect(component).toContain('aria-label="新对话"');
      expect(component).toContain('aria-label="我的工具"');
      expect(component).toContain('aria-label="备份与恢复"');
      expect(component).not.toContain('className="header-button labeled-header-button');
      expect(component).not.toContain('<span>新对话</span>');
      expect(component).not.toContain('<span>备份</span>');
    }
    const source = css("./styles.css");
    const compact = source.slice(source.indexOf("/* Compact icon-only header:"));
    expect(compact).toContain("width: 32px");
    expect(compact).toContain("height: 32px");
    expect(compact).toContain("justify-content: center");
    expect(compact).toContain("align-content: center");
    expect(compact).toContain(".chat-header .model-picker-trigger > svg");
    expect(compact).toContain(".chat-header .folder-backup-entry > b");
    expect(compact).toContain("position: absolute");
  });

  it("raises the header stacking context while a model or more menu is open", () => {
    const source = css("./styles.css");
    const component = css("./App.tsx");

    expect(source).toMatch(
      /\.chat-header\.menu-open\s*\{[\s\S]*?z-index:\s*50;/
    );
    expect(component).toContain(
      'className={`chat-header${modelMenuOpen || moreMenuOpen ? " menu-open" : ""}`}'
    );
  });
});

describe("narrow task pane", () => {
  it("constrains grid tracks rather than clipping an oversized layout", () => {
    const source = css("./styles.css");
    expect(source.match(/\.chat-shell \{([^}]*)\}/s)?.[1]).toContain(
      "grid-template-columns: minmax(0, 1fr)"
    );
    expect(source).toContain("/* Task-pane width containment */");
    expect(source).not.toContain('grid-template-areas: "mark brand model" "actions actions actions"');
    const containment = source.slice(source.indexOf("/* Task-pane width containment */"));
    expect(containment).toMatch(/\.chat-header\s*\{[^}]*display: flex;[^}]*flex-wrap: wrap;/s);
    expect(containment).toContain("min-width: min-content");
    expect(css("../../wps-addin/src/styles.css")).not.toMatch(/\.chat-header \.brand-copy\s*\{\s*display: none;/);
    expect(source).toContain(".chat-header .header-actions");
    expect(source).toContain("overflow-wrap: anywhere");
  });

  it("finishes real request preparation before consuming remote progress", () => {
    const component = css("./App.tsx");
    const start = component.indexOf("async function analyzeConfirmedIntent(");
    const end = component.indexOf("markServerOnline();", start);
    const planning = component.slice(start, end);
    expect(planning).not.toContain('"正在生成操作预览"');
    expect(planning).toContain('"正在准备规划请求"');
    expect(planning.indexOf('completeCurrentStep("已准备规划请求"')).toBeGreaterThan(0);
    expect(planning.indexOf('completeCurrentStep("已准备规划请求"')).toBeLessThan(
      planning.indexOf("await streamAssistantResponse(")
    );
  });
});

describe("natural organic style", () => {
  function rootTokens(): string {
    const source = css("./styles.css");
    const start = source.indexOf(":root {");
    const end = source.indexOf("font-family:", start);
    return source.slice(start, end);
  }

  it("keeps the original neutral green palette rather than a warm earth-tone one", () => {
    const tokens = rootTokens();

    expect(tokens).toContain("--color-background-subtle: #ffffff");
    expect(tokens).toContain("--color-text-primary: #37352f");
    expect(tokens).toContain("--color-brand-text: #2f6b47");
    expect(tokens).toContain("--color-border-strong: #d3d2cf");
    expect(tokens).not.toContain("#faf6f1");
    expect(tokens).not.toContain("#5a6b46");
    expect(tokens).not.toContain("#3d3229");
    expect(tokens).not.toContain("#cdbfa9");
  });

  it("uses Apple-style radii tokens with a size-layered scale, not one global value", () => {
    const source = css("./styles.css");
    const tokens = rootTokens();

    // 分层尺度：微件 6 / 紧凑控件 8 / 卡片·常规控件 12 / 面板·弹窗 16。
    // 关键负向：**不得**再退回"所有小控件都用同一个 10px"的机械统一。
    expect(tokens).toMatch(/--radius-xs:\s*6px/);
    expect(tokens).toMatch(/--radius-sm:\s*8px/);
    expect(tokens).toMatch(/--radius-md:\s*12px/);
    expect(tokens).toMatch(/--radius-lg:\s*16px/);
    expect(tokens).not.toMatch(/--radius-sm:\s*10px/);
    // 只有一组名字：不要再冒 `--radius-control` 之类的同值别名。
    expect(tokens).not.toContain("--radius-control:");
    expect(tokens).not.toContain("--radius-card:");
    expect(tokens).not.toContain("--radius-dialog:");
    expect(tokens).not.toContain("--paper-grain:");
    expect(source).not.toContain("feTurbulence");
    expect(source).not.toMatch(/body\s*\{[^}]*background-image:/s);
  });

  it("keeps the original warning and danger colors", () => {
    const tokens = rootTokens();

    expect(tokens).toContain("--color-warning-text: #805f21");
    expect(tokens).toContain("--color-danger-text: #842029");
    expect(tokens).not.toContain("--color-warning-text: #8a5a12");
    expect(tokens).not.toContain("--color-danger-text: #8c2f22");
  });

  it("recognises headings via the serif stack and never pure black", () => {
    const source = css("./styles.css");

    expect(source).toContain("--font-family-serif:");
    expect(source).toMatch(/h1,\s*h2,\s*h3\s*\{[^}]*font-family:\s*var\(--font-family-serif\)/s);
    expect(rootTokens()).not.toContain("#000000");
    expect(rootTokens()).not.toContain("#000;");
  });

  it("shapes the confirmation dialog with Apple-style corners on the original palette", () => {
    const source = css("./styles.css");
    const start = source.indexOf(".authorized-target-overlay {");
    // 切片要一直覆盖到按钮区（控件圆角用在按钮上），否则下面的正向断言
    // 会因为"根本没读到那段"而假绿。
    const section = source.slice(start, source.indexOf(".authorized-target-actions button.primary {"));

    expect(start).toBeGreaterThan(-1);
    // 面板 16 / 内层卡片 12 / 图标与按钮 8：按**元素尺度**分层，
    // 不是"外层减内边距算出同心值"（实测 padding 20px 时那会算出负数）。
    expect(section).toContain("border-radius: var(--radius-lg)");
    expect(section).toContain("border-radius: var(--radius-md)");
    expect(section).toContain("border-radius: var(--radius-sm)");
    // 连续圆角（超椭圆）作为渐进增强出现。
    expect(section).toContain("corner-shape: squircle");
    // 负向：不对称的 `a b c d` 圆角在这个弹窗里没有正当理由 —— 不许复活。
    expect(section).not.toMatch(/border-radius:\s*\d+px\s+\d+px\s+\d+px\s+\d+px/);
    // 负向：不再有把所有按钮做成两端半圆的胶囊。
    expect(section).not.toContain("var(--radius-full");
    expect(section).not.toContain("var(--paper-grain)");
    expect(section).toContain("prefers-reduced-motion");
    expect(section).not.toMatch(/#(?:faf6f1|5a6b46|fffdf9|3d3229|cdbfa9)/);

    // 点击后不得出现"黑边"：弹窗出现时会主动 `focus()` 自身（键盘陷阱 + Escape），
    // `<section tabIndex={-1}>` 一旦拿到程序化焦点，引擎默认的 `:focus-visible`
    // 描边就会在白卡外面套一圈黑框。`.authorized-target-dialog` 必须显式清掉
    // 轮廓。注意段内**先**出现 `@media` 里的同名动画规则，取主规则要用
    // `lastIndexOf`，否则断言会读到动画块而恒假绿。
    const dialogStart = section.lastIndexOf(".authorized-target-dialog {");
    const dialogRule = section.slice(dialogStart, section.indexOf("}", dialogStart));
    expect(dialogRule).toContain("position: relative");
    expect(dialogRule).toContain("outline: none");
  });

  it("expands and collapses the manual sheet picker with a motion-safe height transition", () => {
    const source = css("./styles.css");
    const ruleStart = source.indexOf(".scope-picker-reveal {");
    expect(ruleStart).toBeGreaterThan(-1);
    const baseRule = source.slice(ruleStart, source.indexOf("}", ruleStart));
    // 收起态是"真源"：0fr + 透明 + 不可见。
    // visibility:hidden 还负责把收起态的控件移出 Tab 序（不能只靠 0 高度）。
    expect(baseRule).toContain("grid-template-rows: 0fr");
    expect(baseRule).toContain("opacity: 0");
    expect(baseRule).toContain("visibility: hidden");

    const openStart = source.indexOf(".scope-picker-reveal.open {");
    expect(openStart).toBeGreaterThan(-1);
    const openRule = source.slice(openStart, source.indexOf("}", openStart));
    expect(openRule).toContain("grid-template-rows: 1fr");
    expect(openRule).toContain("opacity: 1");
    expect(openRule).toContain("visibility: visible");

    // 过渡**只在**媒体块里声明：默认（含 prefers-reduced-motion: reduce）瞬时收放。
    expect(baseRule).not.toContain("transition");
    expect(openRule).not.toContain("transition");

    const mediaStart = source.indexOf("@media (prefers-reduced-motion: no-preference) {", ruleStart);
    expect(mediaStart).toBeGreaterThan(-1);
    const mediaBody = source.slice(mediaStart, source.indexOf("\n}", mediaStart));
    // 两个方向都必须有过渡 —— 只给其中一个方向，收起时就会"啪"地消失（本次修复点）。
    expect(mediaBody).toContain(".scope-picker-reveal {");
    expect(mediaBody).toContain(".scope-picker-reveal.open {");
    expect([...mediaBody.matchAll(/transition:/g)].length).toBe(2);
    expect([...mediaBody.matchAll(/var\(--motion-picker-duration\)/g)].length).toBeGreaterThanOrEqual(5);
    // 时长/曲线必须走 token（"太快/太慢"只改一处），不在规则里写死毫秒数。
    expect(baseRule).not.toMatch(/\d+ms/);
    expect(openRule).not.toMatch(/\d+ms/);
    expect(source).toMatch(/--motion-picker-duration:\s*\d+ms/);
    expect(source).toMatch(/--motion-picker-ease:\s*cubic-bezier\(/);
    // 下限写进契约：260ms + 强 ease-out 的初版被反馈"太快、像瞬间出现"，
    // 收放不要退回 300ms 以下。
    const duration = Number(source.match(/--motion-picker-duration:\s*(\d+)ms/)![1]);
    expect(duration).toBeGreaterThanOrEqual(300);
  });

  it("spins the refresh glyph on click, motion-gated", () => {
    const source = css("./styles.css");
    expect(source).toContain("@keyframes scope-refresh-spin");
    expect(source).toContain("transform: rotate(360deg)");

    // 基础规则默认不转：动画只在 `.spinning`（点击时挂上）+ 媒体块里。
    const glyphStart = source.indexOf(".scope-refresh-glyph {");
    expect(glyphStart).toBeGreaterThan(-1);
    const glyphRule = source.slice(glyphStart, source.indexOf("}", glyphStart));
    expect(glyphRule).toContain("display: inline-block");
    expect(glyphRule).not.toContain("animation");

    const mediaStart = source.indexOf("@media (prefers-reduced-motion: no-preference) {", glyphStart);
    expect(mediaStart).toBeGreaterThan(-1);
    const mediaBody = source.slice(mediaStart, source.indexOf("\n}", mediaStart));
    expect(mediaBody).toContain(".scope-refresh-glyph.spinning");
    // 时长走 token（"转得快慢"只改一处）。
    expect(mediaBody).toContain("animation: scope-refresh-spin var(--motion-refresh-duration)");
    expect(source).toMatch(/--motion-refresh-duration:\s*\d+ms/);
  });

  it("routes every brand-family colour through theme tokens", () => {
    // 「全面换肤」：品牌绿系（绿/青绿色相且有饱和度）一律走 `--color-brand-*`。
    // 否则切到暖米色 / 石板蓝 / 鼠尾草绿预设时，界面仍是一整片默认绿。
    // 只允许三处保留字面量：`:root` token 定义、主题预设色板（本身就是"预览色"）、
    // 代码编辑器深色皮肤（不是主题色）。`var(...)` 的兜底值不算——它不阻止主题生效。
    const hsl = (hex: string) => {
      const raw = hex.replace("#", "");
      const full = raw.length === 3 ? raw.split("").map((c) => c + c).join("") : raw;
      if (full.length !== 6) return null;
      const [r, g, b] = [0, 2, 4].map((i) => parseInt(full.slice(i, i + 2), 16) / 255);
      const max = Math.max(r, g, b);
      const min = Math.min(r, g, b);
      const lightness = (max + min) / 2;
      const saturation = max === min
        ? 0
        : (max - min) / (lightness > 0.5 ? 2 - max - min : max + min);
      let hue = 0;
      if (max !== min) {
        if (max === r) hue = ((g - b) / (max - min) + (g < b ? 6 : 0)) * 60;
        else if (max === g) hue = ((b - r) / (max - min) + 2) * 60;
        else hue = ((r - g) / (max - min) + 4) * 60;
      }
      return { hue, saturation: saturation * 100 };
    };
    const isBrandFamily = (hex: string) => {
      const c = hsl(hex);
      return Boolean(c && c.saturation >= 12 && c.hue >= 95 && c.hue <= 185);
    };

    for (const sheet of [css("./styles.css"), wpsStylesheet()]) {
      const scoped = sheet
        .replace(/:root\s*\{[^}]*\}/g, "")
        .replace(
          /\.(?:preset-(?:default|warm-orange|calm-blue|vivid-green)|json-editor|tool-dsl)[^{}]*\{[^{}]*\}/g,
          ""
        )
        .replace(/var\([^()]*\)/g, "");
      const offenders = [...new Set([...scoped.matchAll(/#[0-9a-fA-F]{3,8}\b/g)].map((m) => m[0].toLowerCase()))]
        .filter(isBrandFamily);
      expect(offenders, `未走主题 token 的品牌色：${offenders.join(", ")}`).toEqual([]);
    }
  });

  it("keeps the Apple-style radius scale global: no stray literal radii, exceptions declared", () => {
    const source = css("./styles.css");
    const wps = wpsStylesheet();
    const exceptions: { selector: string; reason: string }[] = JSON.parse(
      css("./radiusExceptions.json")
    ).allow;
    const allowedSelectors = new Set(exceptions.map((entry) => entry.selector));

    for (const sheet of [source, wps]) {
      // 1) 圆角必须走 token（或 0 / 50% 这类结构性取值）。散落的字面量 px 会让
      //    "改一处圆角"变成"改几十处"，这正是这次改造要根除的东西。
      const literal = [...sheet.matchAll(/border-radius:\s*(\d+px)[^;]*;/g)].map((m) => m[1]);
      expect(literal).toEqual([]);

      // 2) 不对称/单侧圆角只允许出现在**登记过的结构性组件**上；
      //    未登记的规则出现四值圆角即视为"未经定义的圆角值"，直接失败。
      //    注意必须把 `0` 也算作合法分量 —— 贴边容器的典型写法就是
      //    `r r 0 0`，漏掉 0 会让最常见的单侧情形整类溜过。
      const CORNER = "(?:var\\(--radius-[a-z]+\\)|\\d+px|0)";
      for (const rule of sheet.matchAll(/([^{}]+)\{([^{}]*)\}/g)) {
        const body = rule[2];
        const asymmetric = body.match(
          new RegExp(`border-radius:\\s*${CORNER}\\s+${CORNER}\\s+${CORNER}\\s+${CORNER}`)
        );
        if (!asymmetric) {
          continue;
        }
        const selectorText = rule[1].replace(/\/\*[\s\S]*?\*\//g, "").trim();
        const declared = [...allowedSelectors].some((allowed) =>
          selectorText.split(",").some((part) => part.trim() === allowed)
        );
        expect(declared, `未登记的对称圆角例外：${selectorText}`).toBe(true);
      }

      // 3) 连续圆角是整套界面的默认形状语言，不是某一个组件的特例。
      expect(sheet).toContain("corner-shape: squircle");
    }

    // 4) 反向：真圆元素（圆点/圆形头像/色板）必须走 --radius-full 且**不带** squircle。
    //    给 50% 半径加超椭圆是自相矛盾的。
    for (const sheet of [source, wps]) {
      for (const selector of [".brand-mark", ".eb-brand-mark", ".probe-logo", ".preset-swatch", ".plan-editor-avatar"]) {
        for (const rule of sheet.matchAll(/([^{}]+)\{([^{}]*)\}/g)) {
          const selectorText = rule[1].replace(/\/\*[\s\S]*?\*\//g, "").trim();
          if (!selectorText.split(",").some((part) => part.trim() === selector)) {
            continue;
          }
          if (!rule[2].includes("border-radius")) {
            continue;
          }
          expect(rule[2], selector).toContain("var(--radius-full)");
          expect(rule[2], selector).not.toContain("corner-shape");
        }
      }
    }

    // 5) 例外清单本身要有理由，否则会退化成一个"随便加"的后门。
    for (const entry of exceptions) {
      expect(entry.reason.length, entry.selector).toBeGreaterThan(8);
    }
  });

  it("defines radius tokens in exactly one place so WPS cannot drift from the baseline", () => {
    const source = css("./styles.css");
    const wps = wpsStylesheet();

    // WPS 端在 main.tsx 里先 import 共享基线（excel-addin 的 styles.css），
    // 再 import 自己的覆盖表。因此圆角 token **只能在共享基线里定义一次** ——
    // WPS 若再写一份，就变成"两个地方维护同一组值"，迟早漂移。
    const radiusDecl = /--radius-(?:xs|sm|md|lg|full)\s*:/g;
    expect([...wps.matchAll(radiusDecl)].length).toBe(0);

    // 反向锚点：共享基线里必须**确实**定义了这套 token（否则上面的 0 是假绿）。
    expect([...source.matchAll(radiusDecl)].length).toBeGreaterThanOrEqual(5);
    expect(source).toMatch(/--radius-xs:\s*6px/);
    expect(source).toMatch(/--radius-sm:\s*8px/);
    expect(source).toMatch(/--radius-md:\s*12px/);
    expect(source).toMatch(/--radius-lg:\s*16px/);
  });

  it("styles every scope action button so none falls back to a native <button>", () => {
    const source = css("./styles.css");
    const panel = css("../../wps-addin/src/WpsProbePanel.tsx");

    // 面板里同排出现的三个按钮类名都必须有样式规则。
    // 背景：`.scope-discard`（「取消」）此前**完全没有 CSS**，浏览器默认样式
    // 直接露出来（方角 + 原生描边），和「刷新 / 确认」不像一套 UI。
    for (const selector of [".scope-refresh", ".scope-confirm", ".scope-discard"]) {
      // TSX 里写的是 `className="scope-discard"`（无点），样式表里是 `.scope-discard`。
      expect(panel, `面板没有用到 ${selector}`).toContain(`"${selector.slice(1)}"`);
      // 同一选择器可能有多条规则（布局一条、外观一条），CSS 是叠加的 ——
      // 必须把所有匹配的规则体拼起来看，只取第一条会读到只有 flex 的那条而假红。
      const escaped = selector.replace(".", "\\.");
      const rule = [...source.matchAll(new RegExp(`${escaped}\\s*\\{([^}]*)\\}`, "g"))]
        .map((m) => m[1])
        .join("\n");
      expect(rule, `${selector} 缺少规则`).not.toBe("");
      // 统一 UI 的硬指标：圆角走 token 且开连续圆角，高度与其它两个按钮一致。
      expect(rule, selector).toContain("border-radius: var(--radius-md)");
      expect(rule, selector).toContain("corner-shape: squircle");
      expect(rule, selector).toContain("min-height: 32px");
      expect(rule, selector).toContain("var(--font-size-caption)");
    }

    // 次要按钮的可辨识度：取消**不能**长得像主按钮（实心品牌绿）。
    const discard = [...source.matchAll(/\.scope-discard\s*\{([^}]*)\}/g)]
      .map((m) => m[1])
      .join("\n");
    expect(discard).not.toContain("background: var(--color-brand-text)");
    expect(discard).toContain("var(--color-background)");
  });

  it("keeps the image entry icon-only and marks the file chip with a paperclip", () => {
    const source = css("./styles.css");
    const excelApp = css("./App.tsx");
    const wpsAgent = css("../../wps-addin/src/WpsAgentPanel.tsx");
    const wpsPanel = css("../../wps-addin/src/WpsProbePanel.tsx");

    // ① 图片入口：两端都必须是纯图标按钮，不再带「图片」文字标签。
    // 背景：按钮里的 `<span>图片</span>` 与图标并肩，在窄工具条里把按钮撑宽、
    // 且与旁边的文字提示抢视线。去掉文字后按钮要能独立表达语义（`aria-label` 仍在）。
    for (const [name, src] of [
      ["App.tsx", excelApp],
      ["WpsAgentPanel.tsx", wpsAgent]
    ] as const) {
      const marker = src.indexOf('className="attach-image-button"');
      expect(marker, `${name} 找不到图片入口`).toBeGreaterThan(-1);
      // 取该按钮起始标签到 `</button>` 的整段 JSX。
      const button = src.slice(marker, src.indexOf("</button>", marker));
      expect(button, `${name} 图片入口仍带「图片」文字`).not.toContain(">图片<");
      // 正向锚点：图标 svg 还在，否则上面的负向断言可能是把按钮整个删了。
      expect(button, `${name} 图片入口丢了图标`).toContain("<svg");
      // 无障碍名不能随文字一起消失。
      expect(button, `${name} 图片入口丢了 aria-label`).toContain('aria-label="添加图片"');
    }

    // ② 图标按钮的几何：正方形 + 无内边距，图标靠 flex 居中。
    const iconRule = [...source.matchAll(/\.attach-image-button\s*\{([^}]*)\}/g)]
      .map((m) => m[1])
      .join("\n");
    expect(iconRule).toContain("width: 31px");
    expect(iconRule).toContain("height: 31px");
    expect(iconRule).toContain("padding: 0");
    // 已无文字，`gap` 只会留下错位空隙 —— 不许回潮。
    expect(iconRule).not.toContain("gap:");

    // ③ 文件选择 chip：图标必须是**线性描边回形针 SVG**，而不是 emoji `▦`/`📎`。
    // emoji 是彩色位图字形，不吃 `color`，放在绿色底片上观感割裂、跨系统字形还不一致；
    // 描边 SVG 只走 `currentColor`，与按钮里那套图片图标同源。两端同一个 `.scope-trigger`，必须一致。
    for (const [name, src] of [
      ["App.tsx", excelApp],
      ["WpsProbePanel.tsx", wpsPanel]
    ] as const) {
      const marker = src.indexOf('className={`scope-trigger');
      expect(marker, `${name} 找不到范围 chip`).toBeGreaterThan(-1);
      const chip = src.slice(marker, src.indexOf("</button>", marker));
      expect(chip, `${name} 文件 chip 未用 SVG 图标`).toContain("<svg");
      expect(chip, `${name} 文件 chip 退回了 emoji 网格图标`).not.toContain("▦");
      expect(chip, `${name} 文件 chip 退回了 emoji 回形针`).not.toContain("📎");
      // 回形针的几何特征：三段重复的 6.9px 圆角折返线（大环 / 内环 / 收口）。
      const path = chip.match(/<path d="([^"]*)"\s*\/>/)?.[1] ?? "";
      expect(path, `${name} 回形针路径缺失`).toContain("19.2 12.1");
      expect(
        path.match(/6\.9/g)?.length ?? 0,
        `${name} 回形针路径不像三段折返`
      ).toBeGreaterThanOrEqual(3);
      // 描边渲染：`> i > svg` 必须关掉 fill、用 `currentColor` 描边，否则会渲染成实心黑块。
      const rule = [...source.matchAll(/\.scope-trigger\s*>\s*i\s*>\s*svg\s*\{([^}]*)\}/g)]
        .map((m) => m[1])
        .join("\n");
      expect(rule, "chip 图标缺少样式规则").not.toBe("");
      expect(rule).toContain("fill: none");
      expect(rule).toContain("stroke: currentColor");
    }
  });

  it("keeps the progress-dot contrast above the graphics threshold in every preset", () => {
    // 「有新进度」浮标 = **跟随主题的品牌浅色渐变圆**，字行走**主题色**：
    // 运行中是三颗圆点，跑完是锚图标，两者共用同一个前景令牌。
    //
    // 这条断言防的是这个组合最容易出的错：浅底 + 浅品牌色在中调色相的预设上会糊。
    // 实测过 —— 直接拿 `--color-brand-text` 当圆点，warm-orange 只有 2.58、
    // vivid-green 2.93，都过不了非文本 UI 图形 3:1 的门槛（WCAG 1.4.11）；
    // 而且 warm-orange 的品牌色压在**纯白**上也才 3.16，所以不是"底太深"能解释的，
    // 必须把前景色压深。
    //
    // **压深是显式令牌，不是 `color-mix()` 现算**。这是踩过的坑：`background:` 上
    // 写"降级行 + color-mix 行"看似安全，其实**救不了** —— 两条是同一个属性，
    // color-mix 那条赢在层叠、只是算不出值（引擎不支持 / 变量无效），于是
    // `background` 落到初始值 `transparent`，圆点**整个消失**，界面上表现为
    // "外圈正常、里面空的"。改成读令牌后没有任何运行时色彩函数参与，任何引擎
    // 都拿得到确定值；而且这个值可以被下面这些断言直接算。
    //
    // **读的都是真实数据源**，没有任何期望值写死：
    //   ① 渐变用了哪些**令牌** —— 从 styles.css 的规则本体里抽 `var(--color-*)`；
    //   ② 那些令牌在四个预设里各是什么**色值** —— 从 useTheme.ts 的色板里解析；
    //   ③ 两个字形各自引用的令牌名 —— 从 styles.css 的圆点 / 锚规则里解析；
    //   ④ 前景色的**色值** —— 从 `:root` 与四个预设的 `progressGlyph` 里解析。
    // 所以改预设、换令牌、改色值，都会让这条断言跟着重算；真的过不了就变红。
    const source = css("./styles.css");
    const hook = css("./hooks/useTheme.ts");
    const stripComments = (text: string) => text.replace(/\/\*[\s\S]*?\*\//g, "");

    // ---- ① 外圈：必须是"跟随主题的浅色渐变"，不能再退回写死的深色药丸 ----
    const buttonRules = [
      ...source.matchAll(/\.new-progress-button(?::[-a-z()]+)?\s*\{([^}]*)\}/gs)
    ].map((match) => stripComments(match[1]));
    expect(buttonRules.length, "找不到 .new-progress-button 规则").toBeGreaterThan(0);
    const buttonRule = buttonRules[0];

    expect(buttonRule, "外圈的渐变不再走品牌令牌了").toContain("var(--color-brand-bg)");
    // 两个端点都必须是品牌令牌：起点曾经是 `--color-background-subtle`，而它在
    // default 预设下就是**纯白 `#ffffff`**，42px 的圆上有一半是白的，整个控件
    // 读成"白圆圈"（用户原话："你这个怎么一个白的了"）。
    expect(buttonRule, "外圈的渐变不再走品牌令牌了").toContain("var(--color-brand-hover)");

    // ---- ①b 外圈**不许画边**（2026-10-12 用户裁决）----
    //
    // 两次原话：「去掉这个圆形边缘线条」，随后「去掉边框颜色，改为透明即可」。
    // 曾经是 `border: 1px solid var(--color-border-strong)`：在浅色渐变外面再套一圈
    // 灰线，42px 的小圆上那一圈比圆本身还抢眼，读成"一个描了边的灰圈"。圆的轮廓
    // 现在由渐变本身 + `--shadow-md` 的投影撑（那是"影子"，不是"线"）。
    //
    // **判据是"值会不会真的画出一条线"，不是"有没有出现 `border` 这个词"。**
    // 这个区分是本轮用血换的：作者把边框声明删干净之后，浏览器 UA 样式接手了
    // `<button>`（Chrome 默认 `2px outset ButtonBorder`），浅绿圆外面冒出一圈**黑环**
    // —— 于是必须显式写一条"关掉边框"的声明把它按回去。那条声明**不画线**，
    // 必须放行；只有真会画出线的（带 `solid`/`dashed`/… 或非零宽度）才拦。
    // 所以按**属性名 + 值**一起判，而不是只看属性名：
    //   · 属性名：`border` 与 `border-*`（除 `*radius`，那是圆角形状必须留）；
    //     `outline` 不算（那是 `:focus-visible` 的键盘聚焦环，要留）。
    //   · 值：只有 `none` / `hidden` / `0` 这些才是"一个像素都不画"。
    // 变异 M4（把 `border: 1px solid var(--color-border-strong)` 画回来）打这一条。
    const IS_BORDER_PROPERTY = (property: string) =>
      property === "border" || (property.startsWith("border-") && !property.includes("radius"));
    const TURNS_THE_EDGE_OFF = (value: string) =>
      /^(?:none|hidden|0|0px|0em|0rem|0pt)$/.test(value.trim());
    for (const rule of buttonRules) {
      const paintedEdges = [...rule.matchAll(/([a-z-]+)\s*:\s*([^;}]+)[;}]/g)]
        .filter(([, property, value]) => IS_BORDER_PROPERTY(property) && !TURNS_THE_EDGE_OFF(value))
        .map(([, property, value]) => `${property}: ${value.trim()}`);
      expect(
        paintedEdges,
        `浮标外圈不许再画边（用户已裁决去掉那圈边缘线），但 ` +
          `${rule.trim().split("\n")[0]} 里有：${paintedEdges.join("、")}`
      ).toEqual([]);
    }
    // 负向锚点：**凡是上色的声明，值里一个 `#`/`rgb(` 都不许有**。曾经的
    // `background: #1f1f1f` 与 `border: 1px solid rgba(255,255,255,0.14)` 都是
    // "为深底写死"的写法 —— 不随主题走，换个预设就露馅。
    //
    // 两个容易漏的地方：
    //   ① 必须扫**整条值**而不是只看开头。把 hex 藏进渐变里
    //      （`linear-gradient(145deg, #1f1f1f, var(--color-brand-bg))`）同样绕过主题，
    //      只匹配 `background:` 后面紧跟的字符就会漏。
    //   ② **不能只扫 `background`**。边框曾是另一处被写死的地方（为深底准备的
    //      白色描边），只盯背景会把它整个漏掉 —— 所以 `border*` 继续留在名单里，
    //      万一以后又加回边（先得改上面 ①b 那条裁决），它也必须走令牌。
    //      `box-shadow`/`outline`/`color` 一并纳入，口径统一成"颜色只准走令牌"。
    //      `stroke`/`fill` 是 SVG 字形的上色入口（锚图标现在只走 `fill` —— 它是
    //      实心剪影；`stroke` 留着是为了别的 svg 字形，漏了它们等于字形可以随便
    //      写死颜色）。
    const COLOR_BEARING =
      /^(background|border|border-color|border-block|border-block-color|border-inline|border-inline-color|outline|outline-color|color|box-shadow|stroke|fill)$/;
    for (const rule of buttonRules) {
      for (const declaration of rule.matchAll(/([a-z-]+)\s*:\s*([^;]+);/g)) {
        if (!COLOR_BEARING.test(declaration[1])) continue;
        expect(
          /#|rgba?\(|hsla?\(/.test(declaration[2]),
          `外圈的 ${declaration[1]} 里出现了字面色值（应全部走主题令牌）：\n` +
            `  ${declaration[1]}: ${declaration[2].trim()}`
        ).toBe(false);
      }
    }

    // 所有被当作背景端点的令牌里，**最暗的那个**才是圆点对比度的最坏情况
    // （悬停/按下会把渐变整体压深，所以不能只算常态）。
    const backgroundTokens = new Set(
      buttonRules
        .flatMap((rule) => [...rule.matchAll(/background:\s*([^;]+);/g)].map((m) => m[1]))
        .flatMap((value) => [...value.matchAll(/var\((--color-[a-z-]+)\)/g)].map((m) => m[1]))
    );
    expect(backgroundTokens.size, "外圈背景没有引用任何主题令牌").toBeGreaterThan(1);

    // ---- ② 圆点：先降级再 color-mix ----
    const dotRule = stripComments(
      source.match(/\.new-progress-button \.progress-dots span \{([^}]*)\}/s)?.[1] ?? ""
    );
    expect(dotRule, "找不到圆点规则").not.toBe("");
    // ---- ② 圆点：**只准是一条令牌声明** ----
    //
    // 这里刻意**禁止** `color-mix()`。原因不是"怕老引擎"，而是"降级写法在同一个
    // 属性上是假的"：`background: var(--x)` 与 `background: color-mix(...)` 是同一个
    // 属性，后者赢在层叠、只是**算不出值**（引擎不支持、或里面的变量无效），
    // 于是 `background` 落到初始值 `transparent` —— 圆点整个消失，而不是"退回
    // 浅一点的主题色"。用户看到的现象正是"外圈正常、里面空的"。
    expect(
      dotRule,
      "圆点不要用 color-mix() 现算：同属性上的降级行救不了\"算不出值\"，" +
        "无效时 background 会落到 transparent、圆点直接消失。请改读显式令牌。"
    ).not.toContain("color-mix(");

    const dotRef = dotRule.match(/background:\s*var\((--color-[a-z-]+)\);/);
    expect(dotRef, "圆点的背景应当是一条 var(--color-*) 令牌声明").not.toBeNull();
    const dotToken = dotRef![1];
    expect(dotToken, "圆点应当以「进度浮标前景」专用令牌取色").toBe("--color-progress-glyph");

    // 圆点规则里也不许出现字面色值（口径与上面的外圈一致）。
    for (const declaration of dotRule.matchAll(/([a-z-]+)\s*:\s*([^;]+);/g)) {
      if (!COLOR_BEARING.test(declaration[1])) continue;
      expect(
        /#|rgba?\(|hsla?\(/.test(declaration[2]),
        `圆点的 ${declaration[1]} 里出现了字面色值：\n  ${declaration[1]}: ${declaration[2].trim()}`
      ).toBe(false);
    }

    // ---- ②b 锚图标（EB 停下后的字形）：必须与圆点**同一个令牌** ----
    //
    // 两个字形是同一个控件、同一层浅色底的两个状态，换个字形不该换种颜色；
    // 而且上面那四组 3:1 结论是按这个令牌算的，换成别的令牌就得重算。
    //
    // 判别点与"是不是令牌"分开：M6（换成 `--color-brand-text`，是令牌但不是同一个）
    // 打前半句，M7（写成与令牌**同值**的 `#265639`）打后半句 —— 同值也不行，
    // 必须走引用，否则主题一换就漂移。
    //
    // 取色入口是 **`fill` 而不是 `stroke`**：锚已经换成**实心剪影**（2026-10-12
    // 用户裁决），形状由 svg 里的 `d`/几何决定、CSS 只负责填色。所以下面除了
    // 盯 `fill` 走令牌，还要**反向确认没有任何描边声明** —— 挂着 `stroke` 会让
    // 实心形的边缘再糊一圈，改回描边就等于把换字形的裁决撤了。
    const anchorRule = stripComments(
      source.match(/\.new-progress-button \.progress-anchor-icon \{([^}]*)\}/s)?.[1] ?? ""
    );
    expect(anchorRule, "找不到锚图标规则").not.toBe("");
    const anchorDeclarations = [...anchorRule.matchAll(/([a-z-]+)\s*:\s*([^;}]+)[;}]/g)];
    const anchorRef = anchorRule.match(/fill:\s*var\((--color-[a-z-]+)\);/);
    expect(anchorRef, "锚图标的填充应当是一条 var(--color-*) 令牌声明").not.toBeNull();
    expect(
      anchorRef![1],
      "锚图标必须与圆点用同一个前景令牌（同一控件、同一底色，只换字形）"
    ).toBe(dotToken);
    // 实心剪影**不吃描边**：`fill: none` 会把它掏空、任何 `stroke*` 都是多余的糊边。
    expect(anchorRule, "锚已是实心剪影，不该再关掉填充").not.toContain("fill: none");
    expect(
      anchorDeclarations.map(([, property]) => property).filter((property) => property.startsWith("stroke")),
      "锚已是实心剪影，形状由 path 的 d 决定，不该再出现描边声明"
    ).toEqual([]);
    for (const declaration of anchorDeclarations) {
      if (!COLOR_BEARING.test(declaration[1])) continue;
      expect(
        /#|rgba?\(|hsla?\(/.test(declaration[2]),
        `锚图标的 ${declaration[1]} 里出现了字面色值：\n  ${declaration[1]}: ${declaration[2].trim()}`
      ).toBe(false);
    }

    // ---- ③ 色板：按预设解析出每个令牌的色值 ----
    const presetsBlock = hook.match(/const PRESETS[^{]*\{([\s\S]*?)\n\};/)?.[1] ?? "";
    expect(presetsBlock, "useTheme.ts 里找不到 PRESETS").not.toBe("");

    const palettes: Record<string, Record<string, string>> = {};
    // 顶层预设键可能是裸标识符（default）或带引号（"warm-orange"），按行首顶层
    // 缩进切块 —— 不能按 `{` 全局切，否则内层每个令牌都会被当成一个预设。
    const blocks = [...presetsBlock.matchAll(/^ {2}(?:"([^"]+)"|([A-Za-z_$][\w$]*)):\s*\{/gm)];
    blocks.forEach((block, index) => {
      const name = block[1] ?? block[2];
      const start = block.index! + block[0].length;
      const end = index + 1 < blocks.length ? blocks[index + 1].index! : presetsBlock.length;
      const tokens: Record<string, string> = {};
      for (const token of presetsBlock
        .slice(start, end)
        .matchAll(/([A-Za-z_$][\w$]*):\s*"(#[0-9a-fA-F]{6})"/g)) {
        tokens[token[1]] = token[2];
      }
      palettes[name] = tokens;
    });
    const presetNames = Object.keys(palettes);
    expect(presetNames, "预设色板数量变了，请同步核对").toHaveLength(4);

    const luminance = (hex: string) => {
      const channels = [0, 2, 4].map((i) => parseInt(hex.slice(1 + i, 3 + i), 16) / 255);
      const linear = channels.map((v) =>
        v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4
      );
      return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2];
    };
    const contrast = (a: string, b: string) => {
      const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x);
      return (hi + 0.05) / (lo + 0.05);
    };
    const channelsOf = (hex: string) =>
      [0, 2, 4].map((i) => parseInt(hex.slice(1 + i, 3 + i), 16));

    // `:root` 里也有一份 `--color-progress-glyph`，必须在 default 预设下与 JS 那份
    // **一致**：面板启动前、以及独立窗口（`poc-dialog`）里只有 CSS 这条生效。
    // 两处漂移过一次就会是"面板里正常、弹窗里看不见"这种最难查的现象。
    const rootBlock = source.slice(source.indexOf(":root {"));
    const rootGlyph = new RegExp(`--color-progress-glyph:\\s*(#[0-9a-fA-F]{6})`).exec(rootBlock)?.[1];
    expect(rootGlyph, `:root 里找不到 ${dotToken}`).toBeTruthy();
    expect(
      rootGlyph,
      ":root 的前景色与 default 预设的 progressGlyph 不一致（两处必须同改）"
    ).toBe(palettes.default?.progressGlyph);

    for (const preset of presetNames) {
      const tokens = palettes[preset];
      const fieldOf = (token: string) =>
        token.replace(/^--color-/, "").replace(/-([a-z])/g, (_, c: string) => c.toUpperCase());

      const stops: { token: string; hex: string }[] = [];
      for (const token of backgroundTokens) {
        const hex = tokens[fieldOf(token)];
        expect(hex, `预设 ${preset} 的色板里没有 ${token}（${fieldOf(token)}）`).toBeTruthy();
        stops.push({ token, hex });
      }
      const darkest = stops.reduce((a, b) => (luminance(b.hex) < luminance(a.hex) ? b : a));

      const brandText = tokens.brandText;
      expect(brandText, `预设 ${preset} 的色板里没有 brandText`).toBeTruthy();
      const dot = tokens.progressGlyph;
      expect(dot, `预设 ${preset} 的色板里没有 progressGlyph`).toBeTruthy();

      // (a) 前景色必须**比本主题色更深**。这一条直接编码了踩过的那个坑：图形若
      //     落到近白色（或压根没上色），压在品牌浅色上就是"看不见"，而界面上
      //     表现为"外圈正常、里面空的"。比主题色还浅的图形在浅底上不可能过 3:1。
      expect(
        luminance(dot),
        `预设 ${preset}：前景色 ${dot} 比主题色 ${brandText} 还浅 —— ` +
          "浅底上的图形只会更看不清，请压深"
      ).toBeLessThan(luminance(brandText));

      // (b) 必须是**同色相压深**，不是随便换个颜色：逐通道都与 brandText 成同一
      //     比例（实测 20% 压深下四组都落在 0.79–0.81）。这样"前景 = 主题色"这个
      //     用户可见的语义才成立，也挡住了"为了过对比度而改成黑/灰"这类退法。
      const base = channelsOf(brandText);
      const actual = channelsOf(dot);
      actual.forEach((value, index) => {
        const ratio = value / base[index];
        expect(
          ratio,
          `预设 ${preset}：前景色 ${dot} 的第 ${index + 1} 个通道不是 brandText ` +
            `${brandText} 的同比例压深（比值 ${ratio.toFixed(3)}，应在 0.75–0.85）`
        ).toBeGreaterThanOrEqual(0.75);
        expect(
          ratio,
          `预设 ${preset}：前景色 ${dot} 的第 ${index + 1} 个通道不是 brandText ` +
            `${brandText} 的同比例压深（比值 ${ratio.toFixed(3)}，应在 0.75–0.85）`
        ).toBeLessThanOrEqual(0.85);
      });

      // (c) 真正的门槛：压在**最暗**背景端点上的非文本 UI 图形对比度。
      const ratio = contrast(dot, darkest.hex);
      expect(
        ratio,
        `预设 ${preset}：前景色 ${dot} 压在最暗背景端点 ${darkest.token}=${darkest.hex} 上` +
          `只有 ${ratio.toFixed(2)}:1，低于 3:1`
      ).toBeGreaterThanOrEqual(3);
    }
  });

  it("ships only the original presets without a warm free-form one", () => {
    const hook = css("./hooks/useTheme.ts");
    const panel = css("./components/ThemePanel.tsx");
    const storage = css("./utils/imageStorage.ts");

    expect(storage).not.toContain('"natural-organic"');
    expect(hook).not.toContain('"natural-organic"');
    expect(hook).not.toContain("PAPER_GRAIN");
    expect(panel).not.toContain('id: "natural-organic"');
    expect(panel).toContain("默认 Notion");
  });
});
