/**
 * 定向测试：Excel 数据范围面板的**选项行**与**块间距**。
 *
 * 四项口径（= 用户提出的四条）：
 *  ① 列表行内不再显示「可分析 · 修改仍需预览确认」这一串 —— 它由 `ScopeAdmissionBadge`
 *     渲染，两处行内调用点（工作簿列表 / 工作目录列表）一起下线；授权状态改由列表
 *     下方的 `ScopeAdmissionSummary` 汇总。这一条用 **App.tsx 源码**取 evidence。
 *  ② 行内勾选框对齐 WPS 的观感：空框 20px、1px 描边、**不填色**（填了会与行底色
 *     `--color-surface` 重合而整个隐进列表里），勾上才整块换成品牌绿 + 白色勾。
 *     并堵住根因：`.scope-mode-list i`（25px + `--color-surface-hover`）是**后代**
 *     选择器，会连行内 `.sheet-checkbox` 一起吃 —— 实测那把勾选框撑成 25px 灰方块
 *     （rgb(241,240,238)）。收窄成直接子代后,勾选框只认 `.sheet-checkbox` 一份声明。
 *  ③ 所有「一行一个选项」的框共用 `--scope-option-height`（模式卡 / 工作表行 /
 *     工作目录行），不再各自写 38px / 57px。
 *  ④ 面板里块与块的间距 Excel/WPS 共用 `--scope-popover-block-gap`：Excel 的正文
 *     多包了一层 `fieldset.scope-selection-fields`（提交期间整体 disabled），
 *     外层 `.scope-popover-scroll` 的 `gap` 因此落空、间距实测为 0。
 *
 * 为什么钉源码/CSS 而不是挂 DOM：`App.tsx` 的内联 JSX 没法单独挂，整个 `<App/>`
 * 要拉 100+ 行 mock（见 `excelWriteEndToEnd.test.tsx`），代价与收益不成比例。
 * WPS 侧的**真实 DOM** 判据在 `apps/wps-addin/src/wpsScopeOptionRow.test.tsx`；
 * **真实像素**（行高、块间距）由 `scripts/check_scope_option_rows.mjs` 在无头浏览器里
 * 用**真实样式表**渲染测量。**真机 Excel / WPS 的像素验收未执行。**
 */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const here = dirname(fileURLToPath(import.meta.url));
const styles = readFileSync(join(here, "styles.css"), "utf8");
const app = readFileSync(join(here, "App.tsx"), "utf8");
const admission = readFileSync(join(here, "components", "ScopeAdmissionStatus.tsx"), "utf8");

/**
 * 剥掉注释再断言"某条规则/某段文案不存在"。
 *
 * 负向判据必须只针对**代码**：注释里要引用选择器名、要引用那句被投诉的文案来讲
 * 道理，若把注释也算进去，"不许出现这串字"就会变成"不许在源码里解释这串字"，
 * 逼着注释写不清前因后果。剥掉注释后判据一点没弱 —— 真在 JSX 里手写一个
 * 同名 `<small>` 依然会被抓住（那是代码，不是注释）。
 * `{/* … *\/}` 这种 JSX 注释是块注释，一并剥掉；`//` 整行注释也剥。
 */
function withoutComments(source: string): string {
  return source
    .replace(/\/\*[\s\S]*?\*\//g, "")
    .split(/\r?\n/)
    .filter((line) => !line.trimStart().startsWith("//"))
    .join("\n");
}

/** 只留代码：负向/正向判据都该打在代码上，注释不能替它作证也不能把它带红。 */
const appCode = withoutComments(app);

/** 取一条规则的声明块（`{`/`}` 定界，断言不跨规则）。 */
function ruleBody(source: string, selector: string): string {
  const start = source.indexOf(`${selector} {`);
  if (start < 0) throw new Error(`styles.css 里找不到规则：${selector}`);
  const end = source.indexOf("}", start);
  if (end < 0) throw new Error(`规则没有闭合：${selector}`);
  return source.slice(start, end);
}

/** 读一个自定义属性的值。 */
function tokenValue(source: string, name: string): string | undefined {
  const hit = withoutComments(source).match(new RegExp(`${name}\\s*:\\s*([^;]+);`));
  return hit?.[1]?.trim();
}

describe("① 列表行内不显示授权状态文案", () => {
  it("App.tsx 里不再渲染行内徽标，但两处列表、表名与汇总行都还在", () => {
    // 负向判据：整份 App.tsx 的**代码**里没有任何行内徽标（两个调用点 + import
    // 一起），连"手写一个同名的 <small>"都不行 —— 文案本身才是被投诉的东西。
    expect(appCode).not.toContain("ScopeAdmissionBadge");
    expect(appCode).not.toContain("scope-admission-badge");
    expect(appCode).not.toContain("修改仍需预览确认");
    expect(withoutComments(admission)).not.toContain("scope-admission-badge");

    // 正向锚点：删掉的是"徽标"，不是"列表"——两处列表与行内表名必须还在。
    // 同样只看代码：注释里提一句类名不该让这条锚点算数。
    expect(appCode).toContain('className="workbook-sheet-list"');
    expect(appCode).toContain('className="folder-sheet-list"');
    expect(appCode.split("<strong>{sheet.name}</strong>").length - 1).toBe(2);

    // 授权状态没有消失：仍然有且仅有汇总行渲染它。
    expect(appCode).toContain("<ScopeAdmissionSummary");
    expect(admission).toContain("SCOPE_ADMISSION_MESSAGES[reason]");
  });

  it("样式表里行内徽标的三条配色规则随组件一起下线，汇总仍要的 error 色保留", () => {
    const css = withoutComments(styles);
    expect(css).not.toContain(".scope-admission-badge");
    // 汇总行里的 `.scope-admission-error` 还在用（ScopeAdmissionSummary 的明细 span）。
    expect(ruleBody(styles, ".scope-admission-error")).toContain(
      "color: var(--color-danger-text)"
    );
    expect(admission).toContain("scope-admission-error");
  });
});

describe("② 行内勾选框：对齐 WPS 观感，且不再被模式图标规则吃掉", () => {
  it("把模式图标的规则收窄到直接子代，勾选框不再继承 25px 灰底", () => {
    const css = withoutComments(styles);

    // 根因：后代选择器 `.scope-mode-list i` 会命中深层的行内勾选框。
    expect(css).not.toContain(".scope-mode-list i {");

    // 正向锚点：模式图标自己的两条规则必须还在（收窄不等于删掉）。
    const icon = ruleBody(styles, ".scope-mode-list > button > i");
    expect(icon).toContain("width: 25px");
    expect(icon).toContain("background: var(--color-surface-hover)");
    expect(ruleBody(styles, ".scope-mode-list > button.selected > i")).toContain(
      "background: var(--color-brand-hover)"
    );
  });

  it("勾选框只有一份声明：空框 20px/1px 不填色，勾上换品牌绿+白勾", () => {
    const box = ruleBody(styles, ".sheet-checkbox");
    expect(box).toContain("width: 20px");
    expect(box).toContain("height: 20px");
    expect(box).toContain("flex: 0 0 20px");
    expect(box).toContain("border: 1px solid var(--color-border-strong)");
    // 不填色：填了会与行底色重合，"空框"整个隐进列表里（原来就是这个问题）。
    expect(box).toContain("background: transparent");
    expect(box).not.toContain("background: var(--color-surface)");

    const picked = ruleBody(styles, "button.selected > .sheet-checkbox");
    expect(picked).toContain("border-color: var(--color-brand-text)");
    expect(picked).toContain("background: var(--color-brand-text)");
    expect(picked).toContain("color: var(--color-background)");

    // 两处列表的行内勾选框都用这一个类（工作簿列表 + 工作目录列表）。
    expect(appCode.split('className="sheet-checkbox"').length - 1).toBe(2);
  });
});

describe("③ 选项行共用一个行高（= 工作目录卡的高度）", () => {
  it("token 就是模式卡的固有高度，三处选项行都读它", () => {
    expect(tokenValue(styles, "--scope-option-height")).toBe("59px");

    for (const selector of [
      ".scope-mode-list > button",
      ".workbook-sheet-list > button",
      ".folder-sheet-list > button"
    ]) {
      expect(ruleBody(styles, selector), selector).toContain(
        "min-height: var(--scope-option-height)"
      );
    }
  });
});

describe("④ 面板块间距：Excel 与 WPS 用同一个值", () => {
  it("间距 token 定义在面板上，外层滚动容器与内层 fieldset 都读它", () => {
    expect(tokenValue(ruleBody(styles, ".scope-popover"), "--scope-popover-block-gap")).toBe(
      "10px"
    );

    // 外层：WPS 的面板正文本来就是它的 grid 子项。
    const scroll = ruleBody(styles, ".scope-popover-scroll");
    expect(scroll).toContain("gap: var(--scope-popover-block-gap)");
    expect(scroll).not.toContain("gap: 10px");

    // 内层：Excel 多包的那层 fieldset 也是 grid 容器，否则它的 gap 落空（实测 0）。
    const fields = ruleBody(styles, ".scope-selection-fields");
    expect(fields).toContain("display: grid");
    expect(fields).toContain("gap: var(--scope-popover-block-gap)");
    // fieldset 的默认边框/内边距必须先归零，否则会多出一圈框。
    expect(fields).toContain("border: 0");
    expect(fields).toContain("padding: 0");
  });
});
