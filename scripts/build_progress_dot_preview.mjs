/**
 * 从**真实源文件**抽取进度浮标（运行中三点 / 停下后锚）动效，生成一份自包含预览页。
 *
 * 关键点：这段 CSS 不是手抄的，而是从 apps/excel-addin/src/styles.css 里按
 * 区块切出来的原文；四个预设的令牌取自 apps/excel-addin/src/hooks/useTheme.ts
 * 的 PRESETS；**两个字形的 DOM 也从 `components/ScrollToLatestButton.tsx` 里
 * 反解出来**（类名、锚的圆与三段路径逐字取自组件，不在这里重写一遍）。
 * 所以预览页里看到的，就是面板里会跑的那套规则与那套标记 —— 避免"预览好看、
 * 实机不对"这种自欺。
 *
 * 页面用 `?sheet=running|idle` 只留一行、并隐去预览页自己的外壳（标题、控件、说明）：
 * 于是截图脚本拿到的就是**一张卡片、左上角对齐**，裁剪框只要一个固定矩形，
 * 不会因为标题文案换行、说明变长而整体位移。
 *
 * 用法：node scripts/build_progress_dot_preview.mjs
 */
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const STYLES = resolve(REPO, "apps", "excel-addin", "src", "styles.css");
const THEME = resolve(REPO, "apps", "excel-addin", "src", "hooks", "useTheme.ts");
const COMPONENT = resolve(
  REPO,
  "apps",
  "excel-addin",
  "src",
  "components",
  "ScrollToLatestButton.tsx"
);
const OUT_DIR = resolve(REPO, "outputs", "progress-dot-preview");
const OUT = resolve(OUT_DIR, "index.html");

const styles = readFileSync(STYLES, "utf8");

/** 切出 `:root` 令牌块（含注释），全部令牌都依赖它。 */
const rootBlock = styles.slice(styles.indexOf(":root {"), styles.indexOf("\n}", styles.indexOf(":root {")) + 2);

/**
 * 切出浮标那一整块：从 `.new-progress-button {` 到 `.pet-float {` 之前。
 * 这一块里含按钮本体、两点/锚两条字形规则、字形切换关键帧、reduced-motion 媒体块。
 */
const start = styles.indexOf(".new-progress-button {");
const end = styles.indexOf(".pet-float {");
if (start < 0 || end < 0 || end <= start) {
  throw new Error("无法从 styles.css 定位浮标区块，请检查结构是否变动");
}
const pillBlock = styles.slice(start, end);

/**
 * 从 useTheme.ts 的 PRESETS 里读出四个预设（真值，不是常量副本）。
 *
 * 不能用一个"整个对象"的大正则：`"warm-orange"` / `"calm-blue"` 是**带引号**的键，
 * `(\w[\w-]*)` 匹配不到引号开头的。改成先定位 PRESETS 块，再按顶层键的 `{` 切段。
 */
const theme = readFileSync(THEME, "utf8");
const blockStart = theme.indexOf("const PRESETS");
const blockEnd = theme.indexOf("\n};", blockStart);
if (blockStart < 0 || blockEnd < 0) {
  throw new Error("无法从 useTheme.ts 定位 PRESETS 块");
}
const block = theme.slice(blockStart, blockEnd);

const WANTED = ["default", "warm-orange", "calm-blue", "vivid-green"];
const keyHits = [...block.matchAll(/\n\s{2}"?([\w-]+)"?:\s*\{/g)].map((m) => ({
  name: m[1],
  from: m.index
}));
const presets = {};
for (let i = 0; i < keyHits.length; i += 1) {
  const { name, from } = keyHits[i];
  if (!WANTED.includes(name)) continue;
  const to = i + 1 < keyHits.length ? keyHits[i + 1].from : block.length;
  const body = block.slice(from, to);
  const pick = (key) => new RegExp(`${key}:\\s*"(#[0-9a-fA-F]{6})"`).exec(body)?.[1];
  // 阴影值不是色值，得单独取原文 —— 浮标现在用 `--shadow-md`，要和卡片一致。
  const pickRaw = (key) => new RegExp(`${key}:\\s*"([^"]*)"`).exec(body)?.[1];
  presets[name] = {
    background: pick("background"),
    backgroundSubtle: pick("backgroundSubtle"),
    surface: pick("surface"),
    textPrimary: pick("textPrimary"),
    textSecondary: pick("textSecondary"),
    brandBg: pick("brandBg"),
    brandText: pick("brandText"),
    brandHover: pick("brandHover"),
    progressGlyph: pick("progressGlyph"),
    border: pick("border"),
    borderStrong: pick("borderStrong"),
    shadowMd: pickRaw("shadowMd")
  };
}
const presetNames = Object.keys(presets);
if (presetNames.length !== 4) {
  throw new Error(`预设解析数量不对：${presetNames.length}（期望 4）→ ${presetNames.join(",")}`);
}
const missing = presetNames.filter((name) =>
  Object.values(presets[name]).some((value) => !value)
);
if (missing.length > 0) {
  throw new Error(`预设字段解析不完整：${missing.join(", ")}`);
}

/**
 * 从 `ScrollToLatestButton.tsx` **反解**两个字形的标记。
 *
 * 不在这里重写一份：组件里的类名与锚的整段标记是唯一来源，改了组件预览就跟着改，
 * 不存在"预览里的锚和产品里的锚不是同一个锚"。解不出来就当场报错，不静默降级。
 *
 * 锚字形**整段照搬**，不再逐段（圆 / 三段 path）反解：实心剪影锚由
 * `<path fillRule>` + `<rect>` + 两个 `<circle>` + 两条 `<path>` 组成，逐段解既脆
 * （属性顺序、`/>` 前的空格一变就崩）又没必要 —— 预览页要的就是与产品**一模一样**
 * 的那段 SVG。这里只做三处改写：`className` → `class`、`fillRule` → `fill-rule`
 * （这两个是 React 的驼峰写法，静态 HTML 里不生效），以及去掉 JSX 注释（花括号
 * 包起来的块注释不是合法 HTML，留着会当文本渲染出来）。
 */
const component = readFileSync(COMPONENT, "utf8");
const dotClass = /className="(progress-dots)"/.exec(component)?.[1];
const anchorClass = /className="(progress-anchor-icon)"/.exec(component)?.[1];
const anchorMarkup = (() => {
  const block = /<svg\s+className="progress-anchor-icon"[\s\S]*?<\/svg>/.exec(component)?.[0];
  if (!block) return null;
  return block
    .replace(/\{\/\*[\s\S]*?\*\/\}/g, "")
    .replace(/className=/g, "class=")
    .replace(/fillRule=/g, "fill-rule=");
})();
if (!dotClass || !anchorClass || !anchorMarkup) {
  throw new Error(
    "无法从 ScrollToLatestButton.tsx 反解字形标记：" +
      `dotClass=${dotClass} anchorClass=${anchorClass} anchor=${Boolean(anchorMarkup)}`
  );
}

/**
 * 下面这些数字**全部从源文件算出来**，不在脚本里重写一遍常量：
 *   - 前景色的色值   ← `:root` 与 useTheme.ts 四个预设里的 `--color-progress-glyph`
 *   - 外圈用到的令牌 ← styles.css 里**按钮自身规则体**（含 `:hover`/`:active`）的 `background`
 *   - 令牌的色值     ← useTheme.ts 的四个预设
 * 预览页因此不会和实现各说各话：改了设计，页面上的数字跟着变。
 *
 * 前景色**不再**从 `color-mix()` 现算 —— 那是踩过的坑：同一个属性上的 color-mix
 * 一旦算不出值，`background` 会落到 `transparent`、图形整个消失（界面上就是
 * "外圈正常、里面空的"）。现在是显式令牌，脚本直接读它。
 */
const rootGlyph = /--color-progress-glyph:\s*(#[0-9a-fA-F]{6})/.exec(
  styles.slice(styles.indexOf(":root {"))
)?.[1];
if (!rootGlyph) {
  throw new Error("styles.css 的 :root 里没有 --color-progress-glyph");
}

/**
 * 收令牌必须**只取按钮自身的规则体**（`.new-progress-button` 与它的 `:hover`/`:active`），
 * 不能扫整个浮标区块：区块里还夹着圆点 / 锚两条字形规则，它们的取色声明里含
 * `var(--color-progress-glyph)` —— 那是**前景图形自己的颜色**，一旦混进来就会被当成
 * "渐变的一个端点"，表格里的"最暗端"会直接变成前景色，算出一个既非外圈、
 * 又毫无意义的对比度。这里和 `styles.test.ts` 的取法保持一致。
 */
const buttonRules = [
  ...pillBlock.matchAll(/\.new-progress-button(?::[-a-z()]+)?\s*\{([^}]*)\}/gs)
].map((match) => match[1]);
if (buttonRules.length < 2) {
  throw new Error(
    `按钮自身规则块只切到 ${buttonRules.length} 个，浮标区块的结构可能变了（至少要含常态 + :hover）`
  );
}

const GLYPH_TOKENS = ["--color-progress-glyph", "--color-brand-text"];
const backgroundTokens = [
  ...new Set(
    buttonRules
      .flatMap((rule) => [...rule.matchAll(/background:\s*([^;]+);/g)].map((m) => m[1]))
      .flatMap((value) => [...value.matchAll(/var\((--color-[a-z-]+)\)/g)].map((x) => x[1]))
  )
];
const leaked = backgroundTokens.filter((token) => GLYPH_TOKENS.includes(token));
if (leaked.length > 0) {
  throw new Error(
    `外圈令牌里混进了 ${leaked.join("、")}（前景图形自己的颜色）—— ` +
      "说明按钮规则体的切法把字形规则也吃进来了"
  );
}
const fieldOf = (token) =>
  token.replace(/^--color-/, "").replace(/-([a-z])/g, (_, c) => c.toUpperCase());

const luminance = (hex) => {
  const channels = [0, 2, 4].map((i) => parseInt(hex.slice(1 + i, 3 + i), 16) / 255);
  const linear = channels.map((v) =>
    v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4
  );
  return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2];
};
const contrast = (a, b) => {
  const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x);
  return (hi + 0.05) / (lo + 0.05);
};
const channelsOf = (hex) => [0, 2, 4].map((i) => parseInt(hex.slice(1 + i, 3 + i), 16));

const rows = presetNames.map((name) => {
  const palette = presets[name];
  const stops = backgroundTokens.map((token) => [token, palette[fieldOf(token)]]);
  const darkest = stops.reduce((a, b) => (luminance(b[1]) < luminance(a[1]) ? b : a));
  const glyph = palette.progressGlyph;
  // 前景色相对主题色的压深比例（逐通道一致，取平均只为在表格里显示一行字）。
  const deepen =
    channelsOf(glyph).reduce((sum, value, i) => sum + value / channelsOf(palette.brandText)[i], 0) / 3;
  return {
    name,
    palette,
    stops,
    darkest,
    glyph,
    deepen,
    ratio: contrast(glyph, darkest[1])
  };
});
const worstRow = rows.reduce((a, b) => (b.ratio < a.ratio ? b : a));
// `:root` 那份（面板启动前 / 独立窗口里只有 CSS 生效）必须与 default 预设一致。
const rootMatchesDefault = rootGlyph === presets.default.progressGlyph;

const dotsButton = `        <button type="button" class="new-progress-button" data-mode="running" aria-label="查看最新执行进度" title="查看最新进度">
          <span class="${dotClass}" aria-hidden="true"><span></span><span></span><span></span></span>
        </button>`;
// 组件里那段锚标记自带 JSX 缩进，这里按行去缩进后重新对齐，只为预览页源码好读。
const anchorReindented = anchorMarkup
  .split("\n")
  .map((line) => `          ${line.trim()}`)
  .join("\n");
const anchorButton = `        <button type="button" class="new-progress-button" data-mode="idle" aria-label="回到底部查看最新内容" title="回到底部">
${anchorReindented}
        </button>`;

const html = `<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>进度浮标动效预览（取自真实样式表）</title>
<style>
${rootBlock}

/* 预览页自身的外壳，与产品样式无关 */
body {
  margin: 0;
  font: 14px/1.6 "Segoe UI Variable Text", "Segoe UI", system-ui, sans-serif;
  color: var(--color-text-primary);
  background: #f5f5f4;
}
.wrap { max-width: 980px; margin: 0 auto; padding: 28px 20px 64px; }
h1 { font-size: 20px; margin: 0 0 6px; }
.lede { color: var(--color-text-secondary); margin: 0 0 24px; }
.controls {
  display: flex; flex-wrap: wrap; gap: 18px; align-items: flex-end;
  padding: 14px 16px; margin-bottom: 20px;
  background: #fff; border: 1px solid var(--color-border);
  border-radius: var(--radius-md);
}
.controls label { display: flex; flex-direction: column; gap: 6px; font-size: 12px; color: var(--color-text-secondary); }
.controls select, .controls input[type="range"] { font: inherit; }
.row { display: flex; flex-wrap: wrap; gap: 20px; }
/* 第二行留竖向间距。 */
.row + .row { margin-top: 20px; }
.card {
  flex: 1 1 320px; background: #fff; border: 1px solid var(--color-border);
  border-radius: var(--radius-md); overflow: hidden;
}
.card h2 { font-size: 13px; font-weight: 600; margin: 0; padding: 10px 14px; border-bottom: 1px solid var(--color-border); color: var(--color-text-secondary); }

/* ---- 截图模式：?sheet=running|idle ----
   只留一张「真实尺寸」卡片、贴在 (0,0)，预览页自己的外壳全部隐去。
   卡片高度因此是 **确定的** 190px 面板 + 2px 边框 = 192px，裁剪框与文案无关；
   卡片自己的 h2 也一并隐去（状态由小样的标签说明），否则它的行高也会进来。 */
body[data-sheet] .head, body[data-sheet] .note { display: none; }
body[data-sheet] .card.zoom { display: none; }
body[data-sheet] .card h2 { display: none; }
body[data-sheet="running"] .row.idle, body[data-sheet="idle"] .row.running { display: none; }
body[data-sheet] .wrap { margin: 0; max-width: 520px; padding: 0; }
body[data-sheet] { background: #fff; }
body[data-sheet] .row + .row { margin-top: 0; }
body[data-sheet] .card { border-radius: 0; }
/* 面板底：模拟真实聊天区，浮标贴底居中 */
.panel {
  position: relative; height: 190px;
  background: var(--color-background-subtle);
  overflow: hidden;
}
.ghost {
  padding: 14px; font-size: 12px; color: var(--color-text-tertiary);
}
.ghost p { margin: 0 0 10px; }
.bubble {
  max-width: 78%; padding: 8px 10px; margin-bottom: 8px;
  background: var(--color-surface); border: 1px solid var(--color-border);
  border-radius: var(--radius-md); font-size: 12px; color: var(--color-text-secondary);
}
.bubble.me { margin-left: auto; background: var(--color-brand-bg); color: var(--color-brand-text); }

/* 放大观察：只放大几何与动效，颜色不变 */
.panel.zoom .new-progress-button { transform: translateX(-50%) scale(2.6); }
.panel.zoom .new-progress-button:hover,
.panel.zoom .new-progress-button:active { transform: translateX(-50%) scale(2.6); }

.note { margin-top: 22px; font-size: 12px; color: var(--color-text-secondary); }
.note code { background: var(--color-surface-hover); padding: 1px 5px; border-radius: 4px; }
table { border-collapse: collapse; font-size: 12px; margin-top: 12px; width: 100%; }
th, td { text-align: left; padding: 6px 10px; border-bottom: 1px solid var(--color-border); }
th { color: var(--color-text-secondary); font-weight: 600; }
.swatch { display: inline-block; width: 11px; height: 11px; border-radius: 50%; vertical-align: -1px; margin-right: 6px; border: 1px solid rgba(0,0,0,.12); }

${pillBlock}
</style>
</head>
<body>
<div class="wrap">
  <div class="head">
  <h1>「有新进度」浮标：运行中三点 / 停下后锚</h1>
  <p class="lede">下面这段样式是从 <code>apps/excel-addin/src/styles.css</code> 原文抽取的，四个预设的令牌取自 <code>hooks/useTheme.ts</code>，两个字形的标记反解自 <code>components/ScrollToLatestButton.tsx</code>。改动这里就等于改动真实面板。</p>

  <div class="controls">
    <label>主题预设
      <select id="preset"></select>
    </label>
    <label>播放速度 <span id="speedVal">1.00×</span>
      <input id="speed" type="range" min="0.15" max="1" step="0.05" value="1">
    </label>
    <label>暂停（看聚合那一拍；只影响三点）
      <input id="pause" type="checkbox">
    </label>
  </div>
  </div>

  <div class="row running">
    <div class="card">
      <h2>运行中 · 真实尺寸（42px 浮标，贴面板底部居中）</h2>
      <div class="panel" id="panel1">
        <div class="ghost">
          <div class="bubble">已读取「订单表」12 行、4 列。</div>
          <div class="bubble me">按会员编号左连接，汇总订单金额。</div>
          <div class="bubble">正在准备写入预览…</div>
        </div>
${dotsButton}
      </div>
    </div>

    <div class="card zoom">
      <h2>运行中 · 放大 2.6×（看聚散与旋转）</h2>
      <div class="panel zoom" id="panel2">
        <div class="ghost">
          <div class="bubble">三颗点散开 → 收成一颗 → 再散开；整体同时匀速自转。</div>
        </div>
${dotsButton}
      </div>
    </div>
  </div>

  <div class="row idle">
    <div class="card">
      <h2>已停下 · 真实尺寸（同一控件，换成锚）</h2>
      <div class="panel">
        <div class="ghost">
          <div class="bubble">已读取「订单表」12 行、4 列。</div>
          <div class="bubble me">按会员编号左连接，汇总订单金额。</div>
          <div class="bubble">正在准备写入预览…</div>
        </div>
${anchorButton}
      </div>
    </div>

    <div class="card zoom">
      <h2>已停下 · 放大 2.6×（看实心剪影锚的轮廓）</h2>
      <div class="panel zoom">
        <div class="ghost">
          <div class="bubble">同一层浅色底、同一个前景令牌，只换字形：锚是静止的。</div>
        </div>
${anchorButton}
      </div>
    </div>
  </div>

  <div class="note">
    <p><strong>两个字形、一个控件、一种取色。</strong>运行中是三颗点（会动，表达"下面还在长东西"）；跑完换成锚（不动，表达"东西停在那儿、点我下去"）。两者共用前景令牌 <code>--color-progress-glyph</code> —— 同一层浅色底，换个字形不该换种颜色。状态标在按钮的 <code>data-mode="running|idle"</code> 上。</p>
    <p><strong>聚合那一拍为什么读得出来。</strong>半径在 45%–55% 收到 <code>0</code>，三颗点同心重叠 —— 因为同色，最大的 8px 完全盖住 5 / 6px 那两颗，视野里就是<strong>一颗</strong>圆点。这一刻外层自转看不见（绕自身中心转一个点等于没转），所以"合成一颗"不会被旋转打散。</p>
    <p><strong>外圈是"卡片那种浅绿"，不是某一种绿。</strong>用的是 <code>.activity-card</code> 同族令牌（${backgroundTokens.map((t) => `<code>${t}</code>`).join(" → ")}），四个预设各自给出本主题的浅色调。外圈背景里<strong>一个写死色值都没有</strong> —— 这就是"跟随主题"的含义，也是它以前是固定 <code>#1f1f1f</code> 时最大的问题。</p>
    <p><strong>前景色为什么是"压深过的主题色"。</strong>直接拿主题色压在这条渐变上，四个预设里有两个过不了非文本 UI 图形 <strong>3:1</strong> 的门槛（WCAG 1.4.11）：warm-orange 只有 2.58、vivid-green 2.93。而且这不是"渐变太深"造成的 —— warm-orange 的品牌色压在<strong>纯白</strong>上也才 3.16，它本身就是中调橙，把底色调得更浅也救不回来。所以做法是同色相<strong>压深</strong>：只降明度、不动色相。锚用同一个令牌，所以下面这四组对比度对它同样成立。</p>
    <p><strong>压深做成显式令牌，不用 <code>color-mix()</code> 现算。</strong>这不是"怕老引擎"那么简单 —— <code>background:</code> 上写"降级行 + color-mix 行"<strong>救不了</strong>：两条是同一个属性，color-mix 那条赢在层叠、只是<strong>算不出值</strong>，于是 <code>background</code> 落到初始值 <code>transparent</code>，图形<strong>整个消失</strong>，界面上就是"外圈正常、里面空的"。现在读 <code>--color-progress-glyph</code>，任何引擎都拿得到确定值。</p>
    <table>
      <tr><th>预设</th><th>外圈渐变最暗端</th><th>前景图形</th><th>对比度</th></tr>
      ${rows
        .map(
          (row) =>
            `<tr><td>${row.name}</td><td><span class="swatch" style="background:${row.darkest[1]}"></span><code>${row.darkest[1]}</code><br><small>${row.darkest[0]}</small></td><td><span class="swatch" style="background:${row.glyph}"></span><code>${row.glyph}</code><br><small>${row.palette.brandText} × ${(row.deepen * 100).toFixed(0)}%</small></td><td><strong>${row.ratio.toFixed(2)}:1</strong></td></tr>`
        )
        .join("\n      ")}
    </table>
    <p>最坏的是 <strong>${worstRow.name}</strong>，${worstRow.ratio.toFixed(2)}:1，仍过门槛。这里取的是渐变里<strong>最暗的那个端点</strong>而不是常态端点：悬停/按下会把渐变整体压深，而三颗点是在半径 13px 上环绕的、会扫到圆边，所以最坏情况不是圆心那一点。</p>
    <p>这条不变量由 <code>apps/excel-addin/src/styles.test.ts</code> 守卫（改令牌、改色板、改压深比例、把锚换成别的取色都会重算），判别力由 <code>scripts/run_progress_dot_contrast_mutations.mjs</code> 的七个反向变异证明。</p>
  </div>
</div>

<script>
const PRESETS = ${JSON.stringify(presets, null, 2)};

const root = document.documentElement.style;
const select = document.getElementById("preset");

function applyPreset(name) {
  const p = PRESETS[name];
  root.setProperty("--color-background", p.background);
  root.setProperty("--color-background-subtle", p.backgroundSubtle);
  root.setProperty("--color-surface", p.surface);
  root.setProperty("--color-text-primary", p.textPrimary);
  root.setProperty("--color-text-secondary", p.textSecondary);
  root.setProperty("--color-brand-bg", p.brandBg);
  root.setProperty("--color-brand-text", p.brandText);
  root.setProperty("--color-brand-hover", p.brandHover);
  root.setProperty("--color-progress-glyph", p.progressGlyph);
  root.setProperty("--color-border", p.border);
  /* 浮标外圈不画边（用户裁决）：作者侧那条边框声明已删，并显式写了 border: none
     把浏览器 UA 给 <button> 的默认边框按回去（不写它浅绿圆外会冒黑环）。外圈的
     边界靠渐变本身 + --shadow-md —— 所以这条投影必须跟着预设走，否则预览里的
     浮标会用默认预设的灰影假象。
     --color-border-strong 本身已不被浮标引用，仍一并设上：它属于主题令牌集，
     哪天页面又用它，不至于停在默认预设上。 */
  root.setProperty("--color-border-strong", p.borderStrong);
  root.setProperty("--shadow-md", p.shadowMd);
}

for (const name of Object.keys(PRESETS)) {
  const option = document.createElement("option");
  option.value = name;
  option.textContent = name;
  select.append(option);
}
select.addEventListener("change", () => applyPreset(select.value));
/* 查询串参数：preset=<name> 选预设，sheet=running|idle 进截图模式（只留一张卡片）。 */
const params = new URLSearchParams(location.search);
const wanted = params.get("preset");
const initial = wanted && PRESETS[wanted] ? wanted : Object.keys(PRESETS)[0];
select.value = initial;
applyPreset(initial);

const sheet = params.get("sheet");
const inSheetMode = sheet === "running" || sheet === "idle";
if (inSheetMode) document.body.dataset.sheet = sheet;
/* 截图模式下把三点**钉在"散开"那一拍**。
   不钉的话，截到哪一拍取决于 --virtual-time-budget 落在 3.2s 周期的哪里：预算
   1500ms 正好落在 45%–55% 的聚合段，小样里就是"一颗点"，看不出三点的位形。
   钉的时候**不能把延迟写成 0s**：.progress-dots 上有两条动画，延迟 0 会把字形
   切换淡入停在它的 0%（opacity: 0），三点直接看不见。负延迟要**刚过淡入时长**
   （160ms）又不能太深 —— 半径动画 0% 散开、45% 聚合，延迟给到 1s 时它已经走到
   31%、三颗点几乎并成一团。250ms 落在 7.8%，三颗点仍是散开的。 */
const FREEZE_DELAY = "-250ms";
if (inSheetMode) {
  for (const el of document.querySelectorAll(".progress-dots, .progress-dots span")) {
    el.style.animationPlayState = "paused";
    el.style.animationDelay = FREEZE_DELAY;
  }
}

/* 速度：等比缩放到原来的 1/speed 倍。
   注意 .progress-dots 上有**两条**动画（自转 + 字形切换淡入），animation-duration
   是**列表**：只写一个值会把它同时套到两条上，160ms 的淡入被拉成 3.2s，截图里
   三颗点会一直是半透明的。所以这里必须按同样的顺序写两个值。 */
const BASE = { spin: 3.2, dot: 3.2, swap: 0.16 };
const speed = document.getElementById("speed");
const speedVal = document.getElementById("speedVal");
function applySpeed() {
  const rate = Number(speed.value);
  speedVal.textContent = rate.toFixed(2) + "×";
  const factor = 1 / rate;
  for (const el of document.querySelectorAll(".progress-dots")) {
    el.style.animationDuration = BASE.spin * factor + "s, " + BASE.swap + "s";
  }
  for (const el of document.querySelectorAll(".progress-dots span")) {
    el.style.animationDuration = BASE.dot * factor + "s";
  }
}
speed.addEventListener("input", applySpeed);
applySpeed();

/* 暂停：用同一个负延迟把画面钉在"完全聚合"那一拍（周期的 50%）。
   同样按列表顺序写：自转给 -1.6s，淡入已经跑完、给 0 即可。 */
const pause = document.getElementById("pause");
pause.addEventListener("change", () => {
  for (const el of document.querySelectorAll(".progress-dots, .progress-dots span")) {
    el.style.animationPlayState = pause.checked ? "paused" : "";
    el.style.animationDelay = pause.checked
      ? (el.classList.contains("progress-dots") ? "-1.6s, 0s" : "-1.6s")
      : "";
  }
});
</script>
</body>
</html>
`;

mkdirSync(OUT_DIR, { recursive: true });
writeFileSync(OUT, html, "utf8");
console.log(`已生成 ${OUT}`);
console.log(`预设：${presetNames.join(", ")}`);
console.log(`字形标记取自组件：.${dotClass} / .${anchorClass}（锚整段照搬，${anchorMarkup.split("\n").length} 行）`);
console.log(`片段长度：:root=${rootBlock.length} 字符，浮标块=${pillBlock.length} 字符`);
// 这两处若漂移，会表现成"面板里正常、独立窗口（poc-dialog）里图形看不见"。
console.log(
  rootMatchesDefault
    ? `:root 的前景色 ${rootGlyph} 与 default 预设一致 ✓`
    : `⚠ styles.css 的 :root (${rootGlyph}) 与 useTheme.ts 的 default (${presets.default.progressGlyph}) 不一致 —— 两处必须同改`
);
