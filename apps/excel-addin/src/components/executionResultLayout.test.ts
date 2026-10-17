// @vitest-environment jsdom
/**
 * 窄窗格（300px / 430px）不溢出的**静态守卫**。
 *
 * 边界：jsdom 没有真实布局引擎，**测不出**实际像素溢出。这里锁的是"防溢出的 CSS 规则本身"
 * —— 目的是让后续改动**不能静默删掉**这些规则。真实视觉核对仍必须在浏览器内完成
 * （见 `docs/DEVELOPMENT.md` 的分档说明），本文件**不是**视觉证据。
 */
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

const CSS = readFileSync(
  resolve(__dirname, "executionResult.css"),
  "utf8"
);

/** 去掉注释后再匹配，避免注释里的同类字样造成"假绿"。 */
const code = CSS.replace(/\/\*[\s\S]*?\*\//g, "");

function ruleFor(selector: string): string {
  const index = code.indexOf(selector);
  if (index < 0) return "";
  const open = code.indexOf("{", index);
  const close = code.indexOf("}", open);
  return code.slice(open + 1, close);
}

describe("执行结果卡 · 窄窗格防溢出规则", () => {
  it("卡片本体允许长内容断行且可被压缩（不靠内容撑宽）", () => {
    const card = ruleFor(".verification-card.execution-result-card");
    expect(card).toContain("overflow-wrap: anywhere");
    expect(card).toContain("min-width: 0");
    expect(card).toContain("box-sizing: border-box");
  });

  it("标题长句可换行：heading 与 strong 都不保留 min-width:auto", () => {
    expect(ruleFor(".execution-result-heading")).toContain("min-width: 0");
    expect(ruleFor(".execution-result-heading > strong")).toContain(
      "overflow-wrap: anywhere"
    );
  });

  it("异常信息列表：收紧缩进 + 允许长句断行", () => {
    expect(ruleFor(".execution-result-card ul")).toContain("padding-left");
    expect(ruleFor(".execution-result-card li")).toContain(
      "overflow-wrap: anywhere"
    );
  });

  it("长按钮文案不横向溢出：可换行、可断词、不超出容器", () => {
    const button = ruleFor(".execution-recovery-footer button");
    expect(button).toContain("white-space: normal");
    expect(button).toContain("overflow-wrap: anywhere");
    expect(button).toContain("max-width: 100%");
  });

  it("430px 及以下：整块纵向堆叠，说明与按钮各占一行", () => {
    const media = code.slice(code.indexOf("@media (max-width: 430px)"));
    expect(media).toContain("flex-direction: column");
    expect(media).toContain("align-items: stretch");
    expect(media).toContain("width: 100%");
    expect(media).toContain("overflow-wrap: anywhere");
  });
});
