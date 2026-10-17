// 离线实测：数据范围弹层的**动作条是不是固定页脚**（2026-10-16）。
//
// 为什么需要它：jsdom 没有布局引擎，`vitest` 只能钉结构（动作条是不是滚动容器的兄弟，
// 见 apps/excel-addin/src/scopeActionBarLayer.test.tsx、apps/wps-addin/src/wpsScopeDraft.test.tsx），
// 量不出"内容滚起来它到底跟不跟着走"。而用户提出的正是这件事：
// Excel 的「完成」按钮和 WPS 的「确认」按钮**不在同一层**。
//
// 层的差别只有一个判别点：动作条是不是 `.scope-popover-scroll` 的兄弟。
//   · 是（修复后）：`.scope-popover` 是 column flex + `overflow: hidden`，
//     滚动容器 `flex: 1 1 auto; min-height: 0; overflow-y: auto` ⇒ 动作条被钉在底部，
//     内容滚动时**它一动不动**（实测位移 0px）。
//   · 不是（修复前）：动作条只是滚动内容末尾的一行 ⇒ 滚动时它跟着走，
//     面板一打开（scrollTop=0）它在视口**外面**，根本看不见。
//
// 所以判据是两根数：`动作条在内容滚动前后的位移`（固定 0px / 缺陷版随内容移动）
// 与 `动作条下沿到面板下沿的距离`（固定 0px / 缺陷版在面板下方很远）。
// 再加一根**防假绿锚点**：同一份夹具里标题的位移必须 > 100px ——
// 否则"动作条位移 0"可能只是因为容器压根没滚动。
//
// `--legacy` 把两端的结构都改回修复前的写法（动作条放进滚动容器），
// 用来证明**这条测量本身有判别力** —— 不红就是测量没接在那根神经上。
//
// 三处夹具，各自的证据强度不同，报告时不许混：
//   ① Excel：`App.tsx` 的内联 JSX 没法单独挂，用**逐字照抄同一段结构 +
//      只加载 `apps/excel-addin/src/styles.css`** 的夹具页（夹具结论只对
//      "样式表 × 这段结构" 成立，不等于实机 Excel 证据）；
//   ② WPS 夹具：同上，照抄 `WpsProbePanel.tsx` 的结构；
//   ③ WPS 真实组件：挂**真实的 `WpsProbePanel`**（同一份组件 + 同一份样式表），
//      只把面板高度夹到 300px 制造滚动 —— 这一格是真实件证据，不是夹具。
//
// Usage: node scripts/check_scope_action_bar_layer.mjs [path/to/playwright-core] [--legacy] [--shots] [--demo]
import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { createServer } from "node:http";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { build } from "esbuild";

const require = createRequire(import.meta.url);
const argModule = process.argv.slice(2).find((value) => !value.startsWith("--"));
/** 只有在真要驱动浏览器时才 require：缺驱动时错误指向"装驱动"，不是"页面坏了"。 */
let playwright = null;
function chromiumDriver() {
  if (!playwright) ({ chromium: playwright } = require(argModule || "playwright-core"));
  return playwright;
}
const legacy = process.argv.includes("--legacy");
const shots = process.argv.includes("--shots");
const root = path.resolve(import.meta.dirname, "..");
const outDir = path.join(root, "outputs", "scope-action-bar-layer-2026-10-16");
const DEMO_PATH = path.join(outDir, "action-bar-layer-demo.html");

const excelCss = await readFile(path.join(root, "apps/excel-addin/src/styles.css"), "utf8");
const wpsCss = await readFile(path.join(root, "apps/wps-addin/src/styles.css"), "utf8");
/** 只量最终排版：`.scope-picker-reveal` 展开有 360ms 过渡，量中间帧会得到假绿。 */
const noMotion = "\n*, *::before, *::after { transition: none !important; animation: none !important; }\n";

/*
 * 把面板高度夹到 320px，制造"内容比容器高"的局面。
 * 真机上这个高度是 `min(560px, calc(100vh - 124px))`，同样会滚 —— 夹小只是为了
 * 让夹具**一定**滚起来。这条被下面 `overflow > 60` 的断言守住：
 * 没滚起来的夹具会让"动作条位移 0"变成一句废话。
 */
const CLAMP = `
.lane { width: 412px; }
.lane > h2 { font-size: 13px; margin: 0 0 6px; }
.lane > .sub { color: #6b7b72; font-size: 12px; margin: 0 0 8px; line-height: 1.6; }
.lane .scope-popover { position: static; right: auto; bottom: auto; left: auto; height: 320px; max-height: 320px; }
.lane .scope-popover-scroll { overflow-y: scroll; }`;

// ── 夹具：逐字照抄两端弹层结构 ──────────────────────────────────────
const excelSheetRows = (names, picked) => names.map((name, index) =>
  `<button class="${index < picked ? "selected" : ""}" aria-pressed="${index < picked}">` +
  `<i class="sheet-checkbox" aria-hidden="true">${index < picked ? "✓" : ""}</i>` +
  `<span><strong>${name}</strong><small>${10 + index} 行 · 4 列</small></span></button>`).join("");

const EXCEL_NAMES = ["格式验收", "Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5", "Sheet6", "Sheet7"];

function excelFixture(isLegacy, withError = false) {
  const bar = `<div class="scope-action-buttons">
            <button class="scope-refresh"><i class="scope-refresh-glyph" aria-hidden="true">↻</i> 刷新</button>
            <button class="scope-confirm">确认</button>
            <button type="button" class="scope-discard">取消</button>
          </div>`;
  const error = `<p role="alert" class="authorized-target-error">请先确认当前对话的数据范围；需求仍保留，尚未读取字段。</p>`;
  return `
<section class="scope-popover" aria-label="选择数据范围">
  <div class="scope-popover-scroll">
    <fieldset class="scope-selection-fields">
      <div class="scope-popover-header">
        <div><strong>这次对话使用哪些数据？</strong><span>默认跟随你当前正在查看的工作表</span></div>
        <button aria-label="关闭数据范围选择">×</button>
      </div>
      <div class="workbook-identity">
        <i aria-hidden="true">XLSX</i>
        <span><small>当前文件</small><strong title="评分表.xlsx">评分表.xlsx</strong>
        <em>当前工作表 Sheet2</em></span>
      </div>
      <div class="scope-controls">
        <fieldset class="scope-selection-fields">
          <div class="scope-mode-list">
            <button><i aria-hidden="true">◎</i><span><strong>跟随当前</strong><small>发送时自动使用正在查看的工作表</small></span></button>
            <button class="selected"><i aria-hidden="true">☷</i><span><strong>选择多个</strong><small>用于跨表查询、比较或汇总</small></span></button>
            <div class="scope-picker-reveal open">
              <div class="workbook-sheet-picker">
                <div class="sheet-search-row"><input aria-label="搜索工作表" placeholder="搜索工作表..."><span>已选 2/8</span></div>
                <div class="sheet-picker-toolbar"><button>全选搜索结果</button><button>清空</button></div>
                <div class="workbook-sheet-list">${excelSheetRows(EXCEL_NAMES, 2)}</div>
              </div>
            </div>
            <button><i aria-hidden="true">⌑</i><span><strong>工作目录</strong><small>批量处理多个本地工作簿</small></span></button>
          </div>
        </fieldset>
        <div class="scope-admission-summary" role="status"><span>已选 2 张表 · 可分析</span></div>
        ${isLegacy ? `${bar}${withError ? error : ""}` : ""}
      </div>
    </fieldset>
  </div>
  ${isLegacy ? "" : `${bar}${withError ? error : ""}`}
</section>`;
}

const WPS_NAMES = ["订单", "明细", "汇总", "Sheet4", "Sheet5", "Sheet6"];
const wpsSheetRows = (names, picked) => names.map((name, index) =>
  `<button type="button" class="${index < picked ? "selected" : ""}" aria-pressed="${index < picked}">` +
  `<i class="sheet-checkbox" aria-hidden="true">${index < picked ? "✓" : ""}</i>` +
  `<span>${name}</span>${index === 0 ? "<small>当前表</small>" : ""}</button>`).join("");

function wpsFixture(isLegacy, withError = false) {
  const actions = `<div class="wps-scope-actions">
            <button type="button" class="scope-refresh"><i class="scope-refresh-glyph" aria-hidden="true">↻</i> 刷新</button>
            <button type="button" class="scope-confirm">确认</button>
            <button type="button" class="scope-discard">取消</button>
          </div>`;
  const error = `<p role="alert" class="authorized-target-error">请先确认当前对话的数据范围。</p>`;
  return `
<section class="scope-popover wps-scope-popover" aria-label="选择数据范围">
  <div class="scope-popover-scroll">
    <div class="scope-popover-header">
      <div><strong>这次对话使用哪些数据？</strong><span>默认跟随你当前正在查看的工作表</span></div>
      <button type="button" aria-label="关闭数据范围">×</button>
    </div>
    <div class="wps-current-file-card">
      <span aria-hidden="true">XLSX</span>
      <div><small>当前文件</small><strong>选项行测量.xlsx</strong><em>当前工作表 订单</em></div>
    </div>
    <div class="scope-controls">
      <fieldset class="scope-selection-fields">
        <div class="scope-mode-list">
          <button type="button"><i aria-hidden="true">⊙</i><span><strong>跟随当前</strong><small>发送时自动使用正在查看的工作表</small></span></button>
          <button type="button" class="selected"><i aria-hidden="true">☷</i><span><strong>选择多个</strong><small>用于跨表查询、比较或汇总</small></span></button>
          <div class="scope-picker-reveal open">
            <div class="wps-manual-sheet-picker">
              <div class="wps-sheet-search-row"><input aria-label="搜索工作表" placeholder="搜索工作表…"><span>已选 2/6</span></div>
              <div class="sheet-picker-toolbar wps-sheet-picker-toolbar"><button type="button">全选搜索结果</button><button type="button">清空</button></div>
              <div class="sheet-picker-options wps-sheet-picker-options">${wpsSheetRows(WPS_NAMES, 2)}</div>
            </div>
          </div>
          <button type="button" disabled><i aria-hidden="true">▫</i><span><strong>工作目录</strong><small>WPS 当前仅支持打开的工作簿</small></span></button>
        </div>
      </fieldset>
      <div class="scope-admission-summary" role="status"><span>已选 2 张表 · 可分析</span></div>
      ${isLegacy ? `${actions}${withError ? error : ""}` : ""}
    </div>
  </div>
  ${isLegacy ? "" : `${actions}${withError ? error : ""}`}
</section>`;
}

/** 一个变体一页：两条泳道（Excel / WPS），结构同源。 */
/** 一个变体一页：三条泳道。
 *
 * 前两条是**稳态**（弹层里没有错误条）—— 固定页脚的下沿应当正好贴住面板下沿，
 * 所以"贴底距离 == 面板下描边"这句话只在没有其他兄弟节点时成立；
 * 第三条带错误条，专门证明**错误条坐在页脚下面时页脚依然是钉住的**（位移仍然 0）。
 */
function fixturePage(isLegacy) {
  const label = isLegacy ? "修复前（动作条在滚动容器里）" : "修复后（动作条是固定页脚）";
  return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<title>数据范围弹层 · 动作条的层（${label}）</title>
<style>${excelCss}${wpsCss}${noMotion}${CLAMP}</style>
<style>body { margin: 0; padding: 18px; background: #f4f6f5; display: flex; gap: 22px;
  align-items: flex-start; font-family: "Segoe UI", "Microsoft YaHei", sans-serif; }</style>
</head><body>
<div class="lane" data-lane="excel"><h2>Excel · ${label}</h2>${excelFixture(isLegacy, false)}</div>
<div class="lane" data-lane="wps"><h2>WPS · ${label}</h2>${wpsFixture(isLegacy)}</div>
<div class="lane" data-lane="excel-error"><h2>Excel（带错误条）· ${label}</h2>${excelFixture(isLegacy, true)}</div>
</body></html>`;
}

// ── `--demo`：修复前 / 修复后并排的可交互页面 ──────────────────────
//
// 数字由上面的断言负责，这张页面负责"眼睛"。不需要浏览器，也不跑构建。
function buildDemoPage() {
  const excelFixed = excelFixture(false);
  const excelLegacy = excelFixture(true);
  return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<title>数据范围弹层 · 动作条的层（修复前 / 修复后）</title>
<style>${excelCss}${wpsCss}${noMotion}${CLAMP}</style>
<style>
  body { margin: 0; padding: 22px; background: #f4f6f5; color: #27352e;
    font-family: "Segoe UI", "Microsoft YaHei", sans-serif; }
  h1 { font-size: 16px; margin: 0 0 6px; }
  .note { color: #6b7b72; font-size: 12px; line-height: 1.8; margin: 0 0 18px; }
  .note code { background: #e8eeea; border-radius: 3px; padding: 0 3px; }
  .grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 18px; align-items: start; }
  .panel { background: #fff; border: 1px solid #dfe6e1; border-radius: 12px; padding: 12px 12px 14px; }
  .panel h2 { font-size: 13px; margin: 0 0 4px; }
  .panel .tag { font-size: 11px; font-weight: 700; border-radius: 999px; padding: 1px 8px; margin-left: 6px; }
  .tag.before { background: #fdecec; color: #a5413a; }
  .tag.after { background: #e6f2e9; color: #2f6b47; }
  .panel .sub { color: #6b7b72; font-size: 12px; margin: 0 0 10px; line-height: 1.6; }
  .stage { background: #fbfcfb; border: 1px dashed #dfe6e1; border-radius: 8px; padding: 8px; }
  .lane .scope-popover { height: 300px; max-height: 300px; }
</style></head><body>
<h1>数据范围弹层：动作条所在的层</h1>
<p class="note">样式表与两个任务窗格是同一份（<code>apps/*/src/styles.css</code>），结构逐字照抄
<code>App.tsx</code> / <code>WpsProbePanel.tsx</code>。两侧面板都被夹到 <b>300px</b> 高，
以便"内容比容器高"这件事看得见（真机上是 <code>min(560px, calc(100vh - 124px))</code>）。<br>
<b>修复前</b>：动作条是滚动内容的一部分，面板一打开它在视口外面（图中面板底部那条是内容）。
<b>修复后</b>：动作条是滚动容器的兄弟，被 <code>flex: 0 0 auto</code> 钉在面板底部，
内容滚它不动。<br>
无头浏览器实测见 <code>scripts/check_scope_action_bar_layer.mjs</code>
（含 <code>--legacy</code> 反向对照：固定版实测位移 <b>0px</b>、缺陷版随内容移动）。</p>
<div class="grid">
  <div class="panel">
    <h2>Excel<span class="tag before">修复前</span></h2>
    <p class="sub">动作条嵌在 <code>scroll → fieldset → scope-controls</code> 里面，
    跟着内容滚；<code>flex: 0 0 auto</code> 挡不住容器整体滚动。</p>
    <div class="stage"><div class="lane">${excelLegacy}</div></div>
  </div>
  <div class="panel">
    <h2>Excel<span class="tag after">修复后</span></h2>
    <p class="sub">动作条是 <code>.scope-popover-scroll</code> 的兄弟，与 WPS 同层，
    也读同一份共享声明；按钮为 刷新 / 确认 / 取消。</p>
    <div class="stage"><div class="lane">${excelFixed}</div></div>
  </div>
</div>
</body></html>`;
}

async function writeDemoPage() {
  await mkdir(path.dirname(DEMO_PATH), { recursive: true });
  await writeFile(DEMO_PATH, buildDemoPage(), "utf8");
  return DEMO_PATH;
}

if (process.argv.includes("--demo")) {
  console.log(`demo → ${await writeDemoPage()}`);
  process.exit(0);
}

// ── WPS 真实组件（结构不是夹具，只夹高度制造滚动）────────────────────
const stubChatPanel = {
  name: "stub-wps-chat-panel",
  setup(plugin) {
    plugin.onResolve({ filter: /^\.\/WpsAgentPanel$/ }, () => ({ path: "wps-chat-panel-stub", namespace: "stub" }));
    plugin.onLoad({ filter: /.*/, namespace: "stub" }, () => ({
      loader: "tsx",
      resolveDir: path.join(root, "apps/wps-addin/src"),
      contents: `
import { useEffect } from "react";
export default function WpsAgentPanelStub({ scopeDock, onScopeConfirmationReady }) {
  useEffect(() => {
    onScopeConfirmationReady?.(async () => true);
    return () => onScopeConfirmationReady?.(null);
  }, [onScopeConfirmationReady]);
  return <div className="wps-chat-shell">{scopeDock}</div>;
}`
    }));
  }
};

const bundledWps = await build({
  stdin: {
    contents: `
import React from 'react';
import { createRoot } from 'react-dom/client';
import WpsProbePanel from './WpsProbePanel';
createRoot(document.getElementById('root')).render(<WpsProbePanel />);`,
    resolveDir: path.join(root, "apps/wps-addin/src"),
    loader: "tsx"
  },
  bundle: true, write: false, format: "iife", platform: "browser", jsx: "automatic",
  alias: { "@excel-bro/core": path.join(root, "packages/core/src") },
  loader: { ".png": "dataurl", ".svg": "dataurl" },
  define: { "import.meta.env": "{}", "process.env.NODE_ENV": '"production"' },
  plugins: [stubChatPanel]
});

/** 宿主替身：`getWpsApplication()` 读的就是 `window.Application`。 */
const wpsHostStub = `
window.Application = (() => {
  const makeSheet = (name, index) => ({ Name: name, get Index() { return index + 1; } });
  const sheets = [makeSheet("订单", 0), makeSheet("明细", 1)];
  return {
    Version: "12.0",
    ActiveWorkbook: {
      Name: "动作条测量.xlsx",
      FullName: "C:\\\\数据\\\\动作条测量.xlsx",
      Worksheets: { Count: sheets.length, Item: (index) => sheets[Number(index) - 1] },
      ActiveSheet: sheets[0]
    },
    ActiveSheet: sheets[0],
    PluginStorage: { getItem: () => undefined, setItem: () => undefined },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  };
})();
window.fetch = async (input) => {
  const url = String(typeof input === "string" ? input : input?.url ?? "");
  const body = url.includes("/api/models") ? { models: [], defaultModelId: null } : { status: "ok", configured: false };
  return new Response(JSON.stringify(body), { status: 200, headers: { "Content-Type": "application/json" } });
};
`;

const realWpsHtml = `<!doctype html><html><head><meta charset="utf-8"><style>${excelCss}${wpsCss}${noMotion}
.scope-popover { position: static !important; right: auto; bottom: auto; left: auto;
  height: 300px; max-height: 300px; }
.scope-popover-scroll { overflow-y: scroll; }
</style></head>
<body><div id="root"></div><script>${wpsHostStub}</script><script>${bundledWps.outputFiles[0].text}</script></body></html>`;

// ── 页面侧测量 ────────────────────────────────────────────────────
/**
 * 量一条泳道：动作条在滚动容器里的**归属**，以及内容滚动前后它自己的位移。
 *
 * 两根数缺一不可：
 *  · `contained` = 动作条在不在这条泳道的滚动容器里（层的判别点）；
 *  · `barShift`  = 把滚动容器从顶滚到底，动作条自己的 top 变了多少
 *                  （固定页脚 0px；在滚动容器里则等于滚动量）。
 * 再加一根**防假绿锚点** `headerShift`：同一个容器里标题的位移必须 > 100px，
 * 证明它真的滚了 —— 否则"动作条位移 0"只是因为容器压根没滚动。
 */
const measureLane = ({ lane, bar }) => {
  const round = (value) => Math.round(value * 100) / 100;
  const scope = document.querySelector(`[data-lane="${lane}"]`);
  const popover = scope.querySelector(".scope-popover");
  const scroll = scope.querySelector(".scope-popover-scroll");
  const actions = scope.querySelector(bar);
  const header = scope.querySelector(".scope-popover-header");
  scroll.scrollTop = 0;
  const popoverBox = popover.getBoundingClientRect();
  const barBox = actions.getBoundingClientRect();
  const headerTop0 = header.getBoundingClientRect().top;
  const atTop = {
    barTop: barBox.top,
    // 动作条下沿到面板下沿：固定页脚只差一条 1px 描边；在滚动容器里时它是内容末尾，
    // 会远在面板下方（负数）。
    flushBottom: round(popoverBox.bottom - barBox.bottom),
    visibleInPopover: barBox.top >= popoverBox.top - 0.5 && barBox.bottom <= popoverBox.bottom + 0.5
  };
  scroll.scrollTop = scroll.scrollHeight;
  const headerShift = round(headerTop0 - header.getBoundingClientRect().top);
  const barShift = round(actions.getBoundingClientRect().top - atTop.barTop);
  scroll.scrollTop = 0;
  return {
    contained: scroll.contains(actions),
    parentIsPopover: actions.parentElement === popover,
    popoverBorderBottom: round(parseFloat(getComputedStyle(popover).borderBottomWidth) || 0),
    overflow: Math.round(scroll.scrollHeight - scroll.clientHeight),
    flushBottomAtTop: atTop.flushBottom,
    visibleInPopoverAtTop: atTop.visibleInPopover,
    barShift,
    headerShift
  };
};

/** 真实 WPS 面板：纯测量（点开「选择多个」由调用方负责，且已等到列表出现）。 */
const measureRealWps = () => {
  const round = (value) => Math.round(value * 100) / 100;
  const popover = document.querySelector(".scope-popover");
  const scroll = popover.querySelector(".scope-popover-scroll");
  const actions = popover.querySelector(".wps-scope-actions");
  if (!actions) return { actionsFound: false };
  const header = popover.querySelector(".scope-popover-header");
  scroll.scrollTop = 0;
  const popoverBox = popover.getBoundingClientRect();
  const barBox = actions.getBoundingClientRect();
  const barTop0 = barBox.top;
  const flushBottomAtTop = round(popoverBox.bottom - barBox.bottom);
  const visibleInPopoverAtTop = barBox.top >= popoverBox.top - 0.5 && barBox.bottom <= popoverBox.bottom + 0.5;
  const headerTop0 = header.getBoundingClientRect().top;
  scroll.scrollTop = scroll.scrollHeight;
  const barShift = round(actions.getBoundingClientRect().top - barTop0);
  const headerShift = round(headerTop0 - header.getBoundingClientRect().top);
  scroll.scrollTop = 0;
  return {
    actionsFound: true,
    contained: scroll.contains(actions),
    parentIsPopover: actions.parentElement === popover,
    popoverBorderBottom: round(parseFloat(getComputedStyle(popover).borderBottomWidth) || 0),
    overflow: Math.round(scroll.scrollHeight - scroll.clientHeight),
    flushBottomAtTop,
    visibleInPopoverAtTop,
    barShift,
    headerShift
  };
};

// ── 跑起来 ────────────────────────────────────────────────────────
const serve = async (html) => {
  const server = createServer((_req, res) => {
    res.setHeader("Content-Type", "text/html; charset=utf-8");
    res.end(html);
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  return { server, url: `http://127.0.0.1:${server.address().port}` };
};

let browser;
try {
  const chromium = chromiumDriver();
  browser = await chromium.launch({ channel: "msedge", headless: true }).catch(() =>
    chromium.launch({ channel: "chrome", headless: true }));
  if (shots) await mkdir(outDir, { recursive: true });

  const fixture = await serve(fixturePage(legacy));
  const page = await browser.newPage({ viewport: { width: 1340, height: 620 } });
  await page.goto(fixture.url);
  await page.waitForFunction(() => document.querySelectorAll("[data-lane]").length === 3);
  const excel = await page.evaluate(measureLane, { lane: "excel", bar: ".scope-action-buttons" });
  const wps = await page.evaluate(measureLane, { lane: "wps", bar: ".wps-scope-actions" });
  const excelError = await page.evaluate(measureLane, { lane: "excel-error", bar: ".scope-action-buttons" });
  if (shots) {
    await page.screenshot({ path: path.join(outDir, `action-bar-${legacy ? "legacy" : "fixed"}.png`), fullPage: true });
  }

  const real = await serve(realWpsHtml);
  const realPage = await browser.newPage({ viewport: { width: 420, height: 900 } });
  await realPage.goto(real.url);
  await realPage.locator(".scope-trigger").waitFor({ state: "attached", timeout: 15000 });
  await realPage.evaluate(() => document.querySelector(".scope-trigger").click());
  // 展开「选择多个」，否则面板太短、压根滚不起来（那会让"不位移"变成假绿）。
  await realPage.evaluate(() => {
    const target = [...document.querySelectorAll(".scope-mode-list button")]
      .find((button) => (button.textContent ?? "").includes("选择多个"));
    target.click();
  });
  await realPage.locator(".wps-sheet-picker-options button").first().waitFor({ state: "attached", timeout: 15000 });
  await realPage.waitForFunction(() => {
    const pane = document.querySelector(".scope-popover-scroll");
    return pane && pane.scrollHeight - pane.clientHeight > 60;
  }, undefined, { timeout: 15000 });
  const realWps = await realPage.evaluate(measureRealWps);

  const report = [];
  const check = (name, actual, expected) => {
    report.push({ name, actual, expected });
    if (legacy) {
      assert.notEqual(actual, expected,
        `--legacy 下「${name}」应当与修复后不同（实测 ${actual}，期望不等 ${expected}），否则这条测量没有判别力`);
    } else {
      assert.equal(actual, expected, `${name}：实测 ${actual}px，期望 ${expected}px`);
    }
  };

  // ⓪ 防假绿锚点：夹具真的滚起来了（三条泳道各一条）。
  for (const [label, lane] of [["Excel", excel], ["WPS", wps], ["Excel（带错误条）", excelError]]) {
    assert.ok(lane.overflow > 60, `${label} 夹具没有溢出（${lane.overflow}px），"动作条不位移"会是假绿`);
    assert.ok(lane.headerShift > 100,
      `${label} 夹具的标题在滚动时只位移了 ${lane.headerShift}px，说明容器其实没滚动`);
  }
  // 真实 WPS 组件**不是** `--legacy` 的反向对照：它就是仓库里那段代码，不受夹结构影响。
  // 它在两种模式下都必须保持"修复后"的形状 —— 这恰好反证了夹具 legacy 的差别确实来自结构。
  assert.ok(realWps.actionsFound, "真实 WPS 面板里找不到 .wps-scope-actions");
  assert.ok(realWps.overflow > 60,
    `真实 WPS 面板没有溢出（overflow=${realWps.overflow}），量不到滚动行为`);
  assert.ok(realWps.headerShift > 100,
    `真实 WPS 面板的标题在滚动时只位移了 ${realWps.headerShift}px，说明容器其实没滚动`);
  assert.equal(realWps.contained, false, "真实 WPS 动作条不该在滚动容器里");
  assert.equal(realWps.parentIsPopover, true, "真实 WPS 动作条的直接父节点必须是弹层");
  assert.equal(realWps.barShift, 0, `真实 WPS 动作条随内容滚动了 ${realWps.barShift}px`);
  assert.equal(realWps.visibleInPopoverAtTop, true, "真实 WPS 打开时动作条应当已经在面板可视区内");
  assert.equal(realWps.flushBottomAtTop, realWps.popoverBorderBottom,
    `真实 WPS 打开时动作条没有贴住面板下沿（实测 ${realWps.flushBottomAtTop}px）`);

  // ① 层：动作条是滚动容器的**兄弟**。
  check("Excel 动作条不在滚动容器里", excel.contained, false);
  check("Excel 动作条的直接父节点是弹层", excel.parentIsPopover, true);
  check("WPS 动作条不在滚动容器里", wps.contained, false);
  check("WPS 动作条的直接父节点是弹层", wps.parentIsPopover, true);

  // ② 像素：内容滚动时动作条**一动不动**（这是"固定页脚"的可证伪定义）。
  check("Excel 动作条随内容滚动位移", excel.barShift, 0);
  check("WPS 动作条随内容滚动位移", wps.barShift, 0);
  // 错误条坐在页脚下面时，页脚依然钉住。
  check("Excel（带错误条）动作条随内容滚动位移", excelError.barShift, 0);

  // ③ 像素：面板一打开，动作条就贴在面板下沿（不是在内容末尾等你去滚）。
  //    期望值取面板自己的 1px 下描边（内容盒底沿 = 边框盒底沿 - 描边宽）。
  check("Excel 打开时动作条下沿到面板下沿", excel.flushBottomAtTop, excel.popoverBorderBottom);
  check("WPS 打开时动作条下沿到面板下沿", wps.flushBottomAtTop, wps.popoverBorderBottom);
  check("Excel 打开时动作条在面板可视区内", excel.visibleInPopoverAtTop, true);
  check("WPS 打开时动作条在面板可视区内", wps.visibleInPopoverAtTop, true);
  check("Excel（带错误条）打开时动作条在面板可视区内", excelError.visibleInPopoverAtTop, true);

  console.log(legacy ? "legacy（修复前结构）实测：" : "实测：");
  for (const [label, lane] of [["Excel 夹具", excel], ["WPS 夹具", wps],
    ["Excel 夹具（带错误条）", excelError], ["WPS 真实组件", realWps]]) {
    console.log(`  ${label}：动作条在滚动容器里=${lane.contained}，` +
      `溢出 ${lane.overflow}px，滚动时动作条位移 ${lane.barShift}px（标题位移 ${lane.headerShift}px），` +
      `打开时贴底距离 ${lane.flushBottomAtTop}px（面板下描边 ${lane.popoverBorderBottom}px，` +
      `在可视区内=${lane.visibleInPopoverAtTop}）`);
  }
  console.log("  注：WPS 真实组件不受 --legacy 影响（它就是仓库里的那段代码），两种模式下都应保持修复后的形状。");
  for (const row of report) {
    console.log(`  ${legacy ? "✗ 应当不等" : "✓"} ${row.name}：实测 ${row.actual}，期望 ${row.expected}`);
  }
  console.log(legacy
    ? "结论：修复前的结构在每一项上都能被这条测量判红 —— 测量有判别力。"
    : "结论：两端动作条都是固定页脚：内容滚动时位移 0px，打开时就贴住面板下沿。");
  if (shots) console.log(`截图 → ${outDir}`);

  await fixture.server.close();
  await real.server.close();
} finally {
  await browser?.close();
}
