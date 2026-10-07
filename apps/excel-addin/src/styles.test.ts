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
    // 背景：`.scope-discard`（放弃修改）此前**完全没有 CSS**，浏览器默认样式
    // 直接露出来（方角 + 原生描边），和「刷新 / 确认范围」不像一套 UI。
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

    // 次要按钮的可辨识度：放弃修改**不能**长得像主按钮（实心品牌绿）。
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
