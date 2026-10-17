/**
 * 「回到底部」浮标的**滚动行为**演示页 —— 自包含 HTML，可以直接用鼠标滚。
 *
 * 为什么要有这一页：`preset-sheet.png` 只能证明"两种字形长什么样"，证明不了
 * "什么时候出现"。用户报的正是后者（"往上滑没有出现锚"），而静态图看不见行为。
 *
 * 和 `build_progress_dot_preview.mjs` 一样，这里**不重写一份实现**：
 *   - 判据阈值 `DEFAULT_BOTTOM_THRESHOLD` 与"只看位置"这条规则 **从 hook 源码抽**，
 *     抽不到或形状变了就当场抛错 —— 谁把口径退回旧写法，这一页就直接构建失败，
 *     不会继续演一个错的行为；
 *   - 浮标样式原文从 `styles.css` 切出，四个预设令牌取自 `useTheme.ts`；
 *   - 两个字形的标记从 `ScrollToLatestButton.tsx` 反解。
 *
 * 页面顶部明确标注：这是**逻辑复刻演示，不是实机证据**。
 *
 * 用法：node scripts/build_scroll_hint_demo.mjs
 */
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const STYLES = resolve(REPO, "apps", "excel-addin", "src", "styles.css");
const THEME = resolve(REPO, "apps", "excel-addin", "src", "hooks", "useTheme.ts");
const FOLLOW_HOOK = resolve(REPO, "apps", "excel-addin", "src", "hooks", "useMessageAutoFollowRuntime.tsx");
const COMPONENT = resolve(REPO, "apps", "excel-addin", "src", "components", "ScrollToLatestButton.tsx");
const OUT_DIR = resolve(REPO, "outputs", "progress-dot-preview");
const OUT = resolve(OUT_DIR, "scroll-hint-demo.html");

const styles = readFileSync(STYLES, "utf8");

const rootBlock = styles.slice(styles.indexOf(":root {"), styles.indexOf("\n}", styles.indexOf(":root {")) + 2);

const start = styles.indexOf(".new-progress-button {");
const end = styles.indexOf(".pet-float {");
if (start < 0 || end < 0 || end <= start) {
  throw new Error("无法从 styles.css 定位浮标区块，请检查结构是否变动");
}
const pillBlock = styles.slice(start, end);

// ---- 判据：阈值 + 公式 + "只看位置"这条规则，全部从 hook 源码抽 ----
const hook = readFileSync(FOLLOW_HOOK, "utf8");
const threshold = /const DEFAULT_BOTTOM_THRESHOLD = (\d+);/.exec(hook)?.[1];
if (!threshold) {
  throw new Error("无法从 useMessageAutoFollowRuntime.tsx 抽到 DEFAULT_BOTTOM_THRESHOLD");
}
const nearFormula =
  /Math\.max\(\s*0,\s*metrics\.scrollHeight - metrics\.scrollTop - metrics\.clientHeight\s*\)\s*<=\s*threshold/.test(
    hook
  );
if (!nearFormula) {
  throw new Error(
    "`isNearScrollBottom` 的公式形状变了（期望 max(0, scrollHeight - scrollTop - clientHeight) <= threshold）—— " +
      "这页演示的是它，请同步更新后再生成"
  );
}
/**
 * 这一条是本次修的**本体**：滚上去当场置真，**不要求**期间有新内容。
 * 曾经是"离底 + 有新内容"两个条件同时成立（`hasUnseenProgress`），
 * 于是往上翻历史时浮标根本不出现。
 */
if (!hook.includes("setAwayFromBottom(!following)")) {
  throw new Error(
    "hook 里没有 `setAwayFromBottom(!following)` —— 可见性判据可能被退回旧口径（要求「有新内容」），" +
      "这页演示的正是「只看位置」，拒绝生成"
  );
}
if (!hook.includes("awayFromBottom")) {
  throw new Error("hook 没有导出 `awayFromBottom`，调用点/演示文案都需要同步");
}

// ---- 预设令牌 ----
const theme = readFileSync(THEME, "utf8");
const blockStart = theme.indexOf("const PRESETS");
const blockEnd = theme.indexOf("\n};", blockStart);
if (blockStart < 0 || blockEnd < 0) {
  throw new Error("无法从 useTheme.ts 定位 PRESETS 块");
}
const block = theme.slice(blockStart, blockEnd);
const WANTED = ["default", "warm-orange", "calm-blue", "vivid-green"];
const keyHits = [...block.matchAll(/\n\s{2}"?([\w-]+)"?:\s*\{/g)].map((m) => ({ name: m[1], from: m.index }));
const presets = {};
for (let i = 0; i < keyHits.length; i += 1) {
  const { name, from } = keyHits[i];
  if (!WANTED.includes(name)) continue;
  const to = i + 1 < keyHits.length ? keyHits[i + 1].from : block.length;
  const body = block.slice(from, to);
  const pick = (key) => new RegExp(`${key}:\\s*"(#[0-9a-fA-F]{6})"`).exec(body)?.[1];
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
  throw new Error(`预设解析数量不对：${presetNames.length}（期望 4）`);
}
if (presetNames.some((name) => Object.values(presets[name]).some((value) => !value))) {
  throw new Error(`预设字段解析不完整：${presetNames.join(", ")}`);
}

// ---- 字形标记：从组件反解 ----
// 锚**整段照搬**（含 `<path fillRule>` / `<rect>` / 两个 `<circle>` / 两条 `<path>`），
// 不再逐段解：实心剪影锚的段数与属性顺序都会变，逐段解既脆又没必要。只做三处改写：
// `className`/`fillRule` 这两个 React 驼峰属性改成 HTML 名（否则静态 HTML 里不生效）、
// 去掉 JSX 注释（花括号包起来的块注释不是合法 HTML）。解不出来当场报错，不静默降级。
const component = readFileSync(COMPONENT, "utf8");
const dotClass = /className="(progress-dots)"/.exec(component)?.[1];
const anchorClass = /className="(progress-anchor-icon)"/.exec(component)?.[1];
const anchorMarkup = (() => {
  const block = /<svg\s+className="progress-anchor-icon"[\s\S]*?<\/svg>/.exec(component)?.[0];
  if (!block) return null;
  return block
    .replace(/\{\/\*[\s\S]*?\*\/\}/g, "")
    .replace(/\s+/g, " ")
    .replace(/className=/g, "class=")
    .replace(/fillRule=/g, "fill-rule=")
    .trim();
})();
if (!dotClass || !anchorClass || !anchorMarkup) {
  throw new Error(
    "无法从 ScrollToLatestButton.tsx 反解字形标记：" +
      `dotClass=${dotClass} anchorClass=${anchorClass} anchor=${Boolean(anchorMarkup)}`
  );
}

const dotMarkup = `<span class="${dotClass}" aria-hidden="true"><span></span><span></span><span></span></span>`;

/**
 * 页面自己的脚本用**字符串拼接**写：
 * 整页是一个 JS 模板字面量，页内若出现反引号会提前闭合、出现 ${ 会被外层插值。
 * 这里只在真正需要注入的值上用 ${}，其余一律 + 拼。
 */
const html = `<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>「回到底部」浮标 · 滚动行为演示</title>
<style>
${rootBlock}

/* ---- 演示页外壳（与产品样式无关） ---- */
body {
  margin: 0;
  font: 14px/1.6 "Segoe UI Variable Text", "Segoe UI", system-ui, sans-serif;
  color: var(--color-text-primary);
  background: #f5f5f4;
}
.wrap { max-width: 940px; margin: 0 auto; padding: 26px 20px 56px; }
h1 { font-size: 19px; margin: 0 0 6px; }
.lede { color: var(--color-text-secondary); margin: 0 0 14px; }
.warn {
  margin: 0 0 20px; padding: 10px 14px; font-size: 12.5px;
  background: #fff7e6; border: 1px solid #f0d9a8; border-radius: var(--radius-md);
}
.demo { display: flex; gap: 22px; align-items: flex-start; flex-wrap: wrap; }
.panel {
  position: relative; width: 380px; height: 520px; flex: 0 0 auto;
  border: 1px solid var(--color-border-strong); border-radius: var(--radius-md);
  overflow: hidden; background: var(--color-background-subtle);
}
.side { flex: 1 1 300px; min-width: 280px; }
.box {
  background: #fff; border: 1px solid var(--color-border);
  border-radius: var(--radius-md); padding: 14px 16px;
}
.box + .box { margin-top: 14px; }
.box h2 { font-size: 12.5px; font-weight: 600; margin: 0 0 10px; color: var(--color-text-secondary); }
.btns { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 12px; }
.btns button {
  font: inherit; padding: 6px 12px; cursor: pointer;
  border: 1px solid var(--color-border-strong); border-radius: var(--radius-md);
  background: #fff; color: var(--color-text-primary);
}
.btns button[data-primary] {
  background: var(--color-brand-bg); border-color: var(--color-brand-text); color: var(--color-brand-text);
}
.btns button:disabled { opacity: .45; cursor: default; }
label.field { display: flex; flex-direction: column; gap: 5px; font-size: 12px; color: var(--color-text-secondary); }
table { border-collapse: collapse; width: 100%; font-size: 12.5px; }
th, td { text-align: left; padding: 5px 8px; border-bottom: 1px solid var(--color-border); }
th { color: var(--color-text-secondary); font-weight: 600; width: 42%; }
td b { font-weight: 700; }
.on { color: #1a7f4b; }
.off { color: #98a2a0; }
code { background: var(--color-surface-hover, #f1f1f0); padding: 1px 5px; border-radius: 4px; }
.tips { margin: 0; padding-left: 18px; font-size: 12.5px; color: var(--color-text-secondary); }
.tips li { margin-bottom: 5px; }

/* 消息流：与面板同族的滚动容器 */
.stream {
  height: 100%; min-height: 0; overflow-x: hidden; overflow-y: auto;
  padding: 16px 14px 4px; background: var(--color-background-subtle);
}
.bubble {
  max-width: 84%; padding: 8px 11px; margin-bottom: 8px;
  background: var(--color-surface); border: 1px solid var(--color-border);
  border-radius: var(--radius-md); font-size: 12.5px; color: var(--color-text-primary);
}
.bubble.me { margin-left: auto; background: var(--color-brand-bg); color: var(--color-brand-text); border-color: transparent; }

pre#selftest {
  margin: 20px 0 0; padding: 12px 14px; font-size: 12.5px; line-height: 1.7;
  background: #10231a; color: #d6f5e3; border-radius: var(--radius-md); white-space: pre-wrap;
}

${pillBlock}
</style>
</head>
<body>
<div class="wrap">
  <h1>「回到底部」浮标：向上滑就出现，跑着是三点、停下是锚</h1>
  <p class="lede">判据抄自 <code>apps/excel-addin/src/hooks/useMessageAutoFollowRuntime.tsx</code>，样式原文抽自 <code>src/styles.css</code>，字形标记反解自 <code>src/components/ScrollToLatestButton.tsx</code>。</p>
  <p class="warn"><strong>这是逻辑复刻演示，不是实机证据。</strong>它只用来把"什么时候出现、出现的是哪个字形"演清楚；真实验收仍在 Excel / WPS 面板里，用真身跑一遍。</p>

  <div class="demo">
    <div class="panel">
      <div class="stream" id="stream"></div>
    </div>

    <div class="side">
      <div class="box">
        <div class="btns">
          <button type="button" data-primary id="run">开始运行（EB 在跑）</button>
          <button type="button" id="reset">重置会话</button>
        </div>
        <label class="field">主题预设
          <select id="preset"></select>
        </label>
      </div>

      <div class="box">
        <h2>实时判据（与 hook 同一套）</h2>
        <table>
          <tr><th>距底部</th><td><b id="rDist">0</b> px</td></tr>
          <tr><th>阈值 <code>DEFAULT_BOTTOM_THRESHOLD</code></th><td><b>${threshold}</b> px</td></tr>
          <tr><th><code>awayFromBottom</code></th><td><b id="rAway">false</b></td></tr>
          <tr><th><code>running</code>（本轮是否在跑）</th><td><b id="rRunning">false</b></td></tr>
          <tr><th>浮标</th><td><b id="rVis">不显示</b></td></tr>
          <tr><th><code>data-mode</code></th><td><b id="rMode">—</b></td></tr>
        </table>
      </div>

      <div class="box">
        <h2>试这三步</h2>
        <ul class="tips">
          <li>页面刚打开时贴着底部、浮标不出现 —— 这是对的，下面没有东西可看。</li>
          <li>点「开始运行」，然后在对话里<strong>往上滑</strong>：浮标<strong>当场</strong>出现，且是<strong>三点</strong>（还在跑）。</li>
          <li>再点一次按钮让本轮停下：同一个位置换成<strong>锚</strong>。滚回底部，浮标自动消失。</li>
        </ul>
      </div>
    </div>
  </div>

  <pre id="selftest" hidden></pre>
</div>

<script>
var THRESHOLD = ${threshold};
var DOT = ${JSON.stringify(dotMarkup)};
var ANCHOR = ${JSON.stringify(anchorMarkup)};

var stream = document.getElementById("stream");
var panel = document.querySelector(".panel");
var runBtn = document.getElementById("run");
var resetBtn = document.getElementById("reset");

var running = false;
var awayFromBottom = false;
var timer = null;
var step = 0;
/*
 * 浮标是**条件渲染**的（和组件一样：不满足条件就根本不在 DOM 里），所以这里挂着
 * 当前那个按钮节点，而不是给它加个 hidden 类 —— .new-progress-button 自带
 * display:inline-flex，作者样式会盖掉 [hidden] 的 display:none，藏不住。
 */
var hint = null;
var renderedGlyph = null;

var SEED = [
  { me: true, text: "按会员编号左连接，汇总订单金额。" },
  { me: false, text: "已读取「订单表」12 行、4 列。" },
  { me: false, text: "正在判断任务类型…" },
  { me: true, text: "顺便把空值那一列去掉。" },
  { me: false, text: "读到 3 个空白格，需要你确认怎么处理。" },
  { me: true, text: "留空就行，不要写公式。" },
  { me: false, text: "已按「保留原值」记录该决定。" },
  { me: false, text: "正在核对写入目标：订单表!A1:D13。" },
  { me: true, text: "结果放到新表里，别动原表。" },
  { me: false, text: "已生成预览：12 行 × 4 列，未写入任何单元格。" },
  { me: false, text: "等待你确认后才会落到宿主。" },
  { me: true, text: "好，先确认这份摘要。" },
  { me: false, text: "来源指纹已记录，写入前会再核对一次。" },
  { me: true, text: "上次那张表的分组键要不要一起带上？" },
  { me: false, text: "已把「会员编号」作为分组键记进本轮上下文。" },
  { me: false, text: "如果这一列有空值，我会先停下来问你，不自己填。" },
  { me: true, text: "行，遇到空白就来问。" },
  { me: false, text: "已记录：空白一律先问，不推断。" },
  { me: false, text: "撤销快照会按写入范围逐块记录，方便整块回退。" },
  { me: false, text: "以上都是本次会话的既有内容。" },
  { me: false, text: "往上滑，你会看到右下角浮标出现。" }
];
var STEPS = [
  "准备写入预览…",
  "已生成 12 行结果，等待确认。",
  "正在核对写入目标…",
  "写入完成，开始回读…",
  "回读一致。",
  "已生成撤销快照。"
];

/* 与 hook 里的 isNearScrollBottom 同一套：max(0, scrollHeight - scrollTop - clientHeight) <= 阈值 */
function isNearScrollBottom(el) {
  return Math.max(0, el.scrollHeight - el.scrollTop - el.clientHeight) <= THRESHOLD;
}
function distanceFromBottom(el) {
  return Math.max(0, el.scrollHeight - el.scrollTop - el.clientHeight);
}

function setText(id, value, good) {
  var el = document.getElementById(id);
  el.textContent = value;
  el.className = good === undefined ? "" : (good ? "on" : "off");
}

function buildHint() {
  var btn = document.createElement("button");
  btn.type = "button";
  btn.className = "new-progress-button";
  btn.addEventListener("click", function () {
    stream.scrollTop = stream.scrollHeight;
    awayFromBottom = false;
    render();
  });
  return btn;
}

function render() {
  /* 可见性只看位置；字形只由 running 决定 —— 与 ScrollToLatestButton 一致。 */
  if (awayFromBottom) {
    if (!hint) {
      hint = buildHint();
      renderedGlyph = null;
      panel.appendChild(hint);
    }
    hint.dataset.mode = running ? "running" : "idle";
    hint.setAttribute("aria-label", running ? "查看最新执行进度" : "回到底部查看最新内容");
    hint.setAttribute("title", running ? "查看最新进度" : "回到底部");
    if (renderedGlyph !== running) {
      hint.innerHTML = running ? DOT : ANCHOR;
      renderedGlyph = running;
    }
  } else if (hint) {
    hint.remove();
    hint = null;
    renderedGlyph = null;
  }

  var dist = Math.round(distanceFromBottom(stream));
  setText("rDist", String(dist));
  setText("rAway", String(awayFromBottom), awayFromBottom);
  setText("rRunning", String(running), running);
  setText("rVis", awayFromBottom ? "显示" : "不显示", awayFromBottom);
  setText("rMode", hint ? hint.dataset.mode : "—", hint ? running : undefined);
  runBtn.textContent = running ? "停止运行（EB 结束）" : "开始运行（EB 在跑）";
}

/* 滚一下就重算，**不要求**期间有新内容 —— 这正是修掉的那条。 */
stream.addEventListener("scroll", function () {
  awayFromBottom = !isNearScrollBottom(stream);
  render();
});

function appendStep() {
  if (step >= STEPS.length) { stopRun(); return; }
  var div = document.createElement("div");
  div.className = "bubble";
  div.textContent = STEPS[step];
  stream.appendChild(div);
  step += 1;
  if (!awayFromBottom) stream.scrollTop = stream.scrollHeight;
  render();
}

function startRun() {
  running = true;
  step = 0;
  if (!awayFromBottom) stream.scrollTop = stream.scrollHeight;
  render();
  timer = setInterval(appendStep, 700);
}

function stopRun() {
  running = false;
  if (timer) clearInterval(timer);
  timer = null;
  render();
}

runBtn.addEventListener("click", function () {
  if (running) stopRun(); else startRun();
});

function seedStream() {
  stream.innerHTML = "";
  for (var i = 0; i < SEED.length; i += 1) {
    var div = document.createElement("div");
    div.className = SEED[i].me ? "bubble me" : "bubble";
    div.textContent = SEED[i].text;
    stream.appendChild(div);
  }
  awayFromBottom = false;
  stream.scrollTop = stream.scrollHeight;
  render();
}
resetBtn.addEventListener("click", function () {
  stopRun();
  seedStream();
});

/* ---- 预设切换：令牌跟着走，和面板里的 useTheme 一样 ---- */
var PRESETS = ${JSON.stringify(presets, null, 2)};
var rootStyle = document.documentElement.style;
var select = document.getElementById("preset");
function applyPreset(name) {
  var p = PRESETS[name];
  rootStyle.setProperty("--color-background", p.background);
  rootStyle.setProperty("--color-background-subtle", p.backgroundSubtle);
  rootStyle.setProperty("--color-surface", p.surface);
  rootStyle.setProperty("--color-text-primary", p.textPrimary);
  rootStyle.setProperty("--color-text-secondary", p.textSecondary);
  rootStyle.setProperty("--color-brand-bg", p.brandBg);
  rootStyle.setProperty("--color-brand-text", p.brandText);
  rootStyle.setProperty("--color-brand-hover", p.brandHover);
  rootStyle.setProperty("--color-progress-glyph", p.progressGlyph);
  rootStyle.setProperty("--color-border", p.border);
  rootStyle.setProperty("--color-border-strong", p.borderStrong);
  rootStyle.setProperty("--shadow-md", p.shadowMd);
}
for (var i = 0; i < Object.keys(PRESETS).length; i += 1) {
  var name = Object.keys(PRESETS)[i];
  var option = document.createElement("option");
  option.value = name;
  option.textContent = name;
  select.append(option);
}
select.addEventListener("change", function () { applyPreset(select.value); });
var wanted = new URLSearchParams(location.search).get("preset");
var initial = wanted && PRESETS[wanted] ? wanted : Object.keys(PRESETS)[0];
select.value = initial;
applyPreset(initial);

seedStream();

/*
 * ---- 初始状态参数：?state=away|running ----
 * 只为截图/直达用：away = 一进来就已经往上滑开（浮标是锚），
 * running = 一进来就"正在跑 + 已滑开"（浮标是三点）。
 * 不带参数就是贴底、不显示浮标 —— 也别忘了这一态，它才是默认。
 */
var STATE = new URLSearchParams(location.search).get("state");
if (STATE === "away" || STATE === "running") {
  stream.scrollTop = 0;
  awayFromBottom = true;
  if (STATE === "running") running = true;
  render();
}

/*
 * ---- 自检模式：?selftest=1 ----
 * 把上面那套交互按顺序跑一遍，把结论写进 #selftest 与 document.title。
 * 这样构建完可以直接用无头 Chrome 验一次（--dump-dom），不至于把没跑过的页面交出去。
 * 只用同步动作，不用定时器，结论是确定的。
 */
var SELFTEST = new URLSearchParams(location.search).get("selftest") === "1";
function runSelftest() {
  var lines = [];
  var failures = 0;
  function check(name, cond, extra) {
    if (!cond) failures += 1;
    lines.push((cond ? "PASS " : "FAIL ") + name + (extra ? " :: " + extra : ""));
  }

  check(
    "前置：内容高于视口（否则这条演示无从谈起）",
    stream.scrollHeight > stream.clientHeight + THRESHOLD,
    "scrollHeight=" + stream.scrollHeight + " clientHeight=" + stream.clientHeight
  );
  check("初始贴底时不显示浮标", !hint);

  /* 滚到顶，确保**明确**离开底部 —— 滚 200px 可能还在 120px 阈值以内，
     那本来就是"贴底"，浮标不该出现（这条判据是位置，不是"滚过没有"）。 */
  stream.scrollTop = 0;
  stream.dispatchEvent(new Event("scroll"));
  check(
    "往上滑后浮标出现（不要求有新内容）",
    !!hint && distanceFromBottom(stream) > THRESHOLD,
    "距底部 " + Math.round(distanceFromBottom(stream)) + "px > " + THRESHOLD + "px"
  );
  check(
    "此时是「已停下」那一态：锚",
    !!hint && hint.dataset.mode === "idle" && !!hint.querySelector(".progress-anchor-icon") && !hint.querySelector(".progress-dots")
  );

  startRun();
  check(
    "运行中换成三点",
    !!hint &&
      hint.dataset.mode === "running" &&
      hint.querySelectorAll(".progress-dots > span").length === 3 &&
      !hint.querySelector(".progress-anchor-icon")
  );
  stopRun();
  check(
    "停下后换回锚（同一个位置）",
    !!hint && hint.dataset.mode === "idle" && !!hint.querySelector(".progress-anchor-icon")
  );

  /*
   * 浮标是 absolute 定位在面板里的，面板又 overflow:hidden —— 一旦定位参照搞错，
   * 它会整颗贴在底边被裁掉半截。这条把"完整落在面板内"量出来，不靠肉眼。
   */
  var btnRect = hint.getBoundingClientRect();
  var panelRect = panel.getBoundingClientRect();
  check(
    "浮标完整落在面板内、且离底边有余量",
    btnRect.bottom <= panelRect.bottom - 6 &&
      btnRect.top >= panelRect.top &&
      btnRect.left >= panelRect.left &&
      btnRect.right <= panelRect.right,
    "按钮下沿 " +
      Math.round(panelRect.bottom - btnRect.bottom) +
      "px 高于面板下沿；左右各留 " +
      Math.round(btnRect.left - panelRect.left) +
      "/" +
      Math.round(panelRect.right - btnRect.right) +
      "px"
  );

  hint.dispatchEvent(new MouseEvent("click", { bubbles: true }));
  check("点一下就回到底部并收起浮标", !hint && distanceFromBottom(stream) <= THRESHOLD);

  var pre = document.getElementById("selftest");
  pre.hidden = false;
  pre.textContent = lines.join("\\n");
  document.title = (failures === 0 ? "SELFTEST PASS " : "SELFTEST FAIL ") + (lines.length - failures) + "/" + lines.length;
}
if (SELFTEST) runSelftest();
</script>
</body>
</html>
`;

mkdirSync(OUT_DIR, { recursive: true });
writeFileSync(OUT, html, "utf8");
console.log(`已生成 ${OUT}`);
console.log(`判据取自 hook：阈值=${threshold}px，公式形状校验通过，"只看位置"规则校验通过`);
console.log(`字形标记取自组件：.${dotClass} / .${anchorClass}（锚整段照搬）`);
