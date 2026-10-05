import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { COMPOSER_RESIZE_LABEL } from "./hooks/useComposerResize";

/**
 * 输入框「上拉」的两端口径账（契约级，读源码）。
 *
 * Excel 侧的结果卡可以整体渲染测试，但 `App.tsx` 没有便宜的渲染入口；所以这里
 * 做的是**契约**断言：两端必须挂同一份实现、同一个无障碍文案，且都不许回退到
 * 浏览器原生的 `resize` 把手 —— 那条路在 WPS 的 WebView 里根本拖不动。
 * 界面行为的证据在 `WpsAgentPanel.test.tsx`（DOM）与 `useComposerResize.test.tsx`
 * （纯逻辑）里，这里只防"两端各写一套"和"又改回原生把手"。
 */
const appSource = readFileSync("src/App.tsx", "utf8");
const wpsSource = readFileSync("../wps-addin/src/WpsAgentPanel.tsx", "utf8");
const sharedStyles = readFileSync("src/styles.css", "utf8");
const wpsStyles = readFileSync("../wps-addin/src/styles.css", "utf8");

describe("输入框上拉：Excel 与 WPS 同一份实现", () => {
  it("两端都引共享 hook，谁都不许再自己实现一套", () => {
    expect(appSource).toContain('from "./hooks/useComposerResize"');
    expect(wpsSource).toContain("hooks/useComposerResize");
    // 原先内联在 App.tsx 里的那套自研处理器必须已经搬走 —— 留着就说明有两份。
    expect(appSource).not.toContain("function startComposerResize");
    expect(appSource).not.toContain("composerResizeRef");
    expect(wpsSource).not.toContain("function startComposerResize");
  });

  it("两端都把共享 props 挂在把手上，文案来自共享常量", () => {
    expect(COMPOSER_RESIZE_LABEL).toContain("拖动");
    expect(COMPOSER_RESIZE_LABEL).toContain("双击");
    // 正向锚点：确实用上了共享 props。
    expect(appSource).toContain("{...composerResizeHandleProps}");
    expect(wpsSource).toContain("{...composerResizeHandleProps}");
    // 负向：两端的源码里都不许再出现硬编码的把手说明。
    expect(appSource).not.toContain("上下拖动调整输入框高度");
    expect(wpsSource).not.toContain("上下拖动调整输入框高度");
  });

  it("把手在输入框**外面**的 region 上，不抢文本区的光标", () => {
    for (const source of [appSource, wpsSource]) {
      // Excel 侧 region 的 class 是模板字符串（按拖图状态拼 `dragging-image`），
      // 所以这里只找类名本身，不要求 `className="..."` 的形式。
      const region = source.indexOf("composer-input-region");
      const handle = source.indexOf('className="composer-resize-handle"');
      const box = source.indexOf('className="composer-box"');
      expect(region).toBeGreaterThan(-1);
      expect(handle).toBeGreaterThan(region);
      expect(box).toBeGreaterThan(handle);
    }
  });

  it("WPS 侧不再用原生 resize 把手，高度改由 hook 驱动", () => {
    // 共享基线：composer 的 textarea 一直是 resize: none（高度只由 JS 写）。
    expect(sharedStyles).toMatch(/\.composer textarea\s*\{[^}]*resize:\s*none;/);
    const wpsComposerBlock =
      wpsStyles.match(/\.wps-chat-shell \.composer-box textarea\s*\{[^}]*\}/)?.[0] ?? "";
    expect(wpsComposerBlock).not.toBe("");
    // 反向锚点：`resize: vertical` 正是真机里拖不动的那条路。
    expect(wpsComposerBlock).not.toContain("resize: vertical");
    // 下限必须与 hook 的 `minHeight` 一致，否则"拖到底"和实际渲染高度对不上。
    expect(wpsComposerBlock).toContain("min-height: 56px");
    expect(wpsSource).toContain("minHeight: 56");
  });
});
