// 离线实测：数据范围面板里「选项行」的行高与「块与块之间」的间距。
//
// 为什么需要它：jsdom 没有布局引擎，`vitest` 里的用例只能钉结构/样式契约
// （apps/wps-addin/src/wpsScopeOptionRow.test.tsx、apps/excel-addin/src/scopeOptionRow.test.tsx），
// 量不出像素。而这一轮的四条口径里有三条是**纯几何**的：
//   ② 行内勾选框的尺寸/填充（被 `.scope-mode-list i` 那条后代选择器泄漏吃成 25px 灰方块）
//   ③ 行高要跟模式卡（跟随当前 / 选择多个 / 工作目录）**一样高**
//   ④ Excel 面板里块与块的间距要跟 WPS 一样（WPS 实测 10px，Excel 被外层 fieldset 吃掉 = 0）
//
// 两部分，各自用**自己那一端的真实样式表**渲染：
//   · WPS：挂**真实的 `WpsProbePanel`**（与任务窗格同一份组件 + 同一份样式表）；
//   · Excel：`App.tsx` 的内联 JSX 没法单独挂，这里用**逐字照抄 App.tsx 的同一段结构 + 只加载
//     `apps/excel-addin/src/styles.css`** 的夹具页。夹具页的结论只对"样式表 × 这段结构"成立，
//     不等于实机 Excel 证据。
//
// `--legacy` 把两端都改回修复前的写法（行高 38px / 后代选择器泄漏 / fieldset 是块盒），
// 用来证明**这条测量本身有判别力** —— 不红就是测量没接在那根神经上。
//
// Usage: node scripts/check_scope_option_rows.mjs [path/to/playwright-core] [--legacy] [--shots]
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

const excelCss = await readFile(path.join(root, "apps/excel-addin/src/styles.css"), "utf8");
const wpsCss = await readFile(path.join(root, "apps/wps-addin/src/styles.css"), "utf8");
/** 只量最终排版：`.scope-picker-reveal` 展开有 360ms 过渡，量中间帧会得到假绿。 */
const noMotion = "\n*, *::before, *::after { transition: none !important; animation: none !important; }\n";

/** 修复前的写法（反向对照用）。 */
const LEGACY_EXCEL = `
.scope-selection-fields { display: block; gap: 0; }
/*
 * Excel 的工作表行**原本没有** min-height（git HEAD 的 .workbook-sheet-list > button
 * 只有 display/width/align-items/gap/padding），行高由两行内容撑出来。
 * 这里不能写 38px —— 那是 **WPS** 的值；写进来就是编一个从未存在的"修复前"，
 * 虽然数值上都被内容高度盖过、结果碰巧一样，但反向对照的可信度就是这么丢的。
 * 用 auto 才是"没有这条声明"。
 */
.workbook-sheet-list > button { min-height: auto; }
.scope-mode-list i {
  display: grid; width: 25px; height: 25px; flex: 0 0 25px; place-items: center;
  border-radius: var(--radius-xs); background: var(--color-surface-hover); font-style: normal;
}`;
/*
 * WPS 侧只有"行高"这一项被改过（它是块间距那一项的**参照**，本来就对）。
 * 但为了让块间距这条测量在**两边都**被证明接在 `gap` 上，这里把 WPS 的
 * 面板间距也一起打进 0：如果它不跟着动，说明这条测量只是碰巧等于 10。
 */
const LEGACY_WPS = `
.wps-sheet-picker-options > button { min-height: 38px; }
.scope-popover-scroll { gap: 0; }`;

// ── WPS：挂真实面板 ────────────────────────────────────────────────
/** 聊天面板换替身：要测的是范围面板，不把整个聊天面板拖起来。 */
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
      Name: "选项行测量.xlsx",
      FullName: "C:\\\\数据\\\\选项行测量.xlsx",
      Worksheets: { Count: sheets.length, Item: (index) => sheets[Number(index) - 1] },
      ActiveSheet: sheets[0]
    },
    ActiveSheet: sheets[0],
    PluginStorage: { getItem: () => undefined, setItem: () => undefined },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  };
})();
// 面板会去问模型目录；给一条干净的空目录，免得等网络超时。
window.fetch = async (input) => {
  const url = String(typeof input === "string" ? input : input?.url ?? "");
  const body = url.includes("/api/models") ? { models: [], defaultModelId: null } : { status: "ok", configured: false };
  return new Response(JSON.stringify(body), { status: 200, headers: { "Content-Type": "application/json" } });
};
`;

// ── Excel：逐字照抄 App.tsx 的同一段结构 + 只加载 Excel 的样式表 ──────────
/** 与 App.tsx:10141-10293 / 10245-10268 / 10296-10323 同结构、同文案。 */
function excelRow(name, caption, selected, isLegacy) {
  return `<button class="${selected ? "selected" : ""}" aria-pressed="${selected}">` +
    `<i class="sheet-checkbox" aria-hidden="true">${selected ? "✓" : ""}</i>` +
    `<span><strong>${name}</strong><small>${caption}</small>` +
    (selected && isLegacy ? `<small class="scope-admission-badge">可分析 · 修改仍需预览确认</small>` : "") +
    `</span></button>`;
}
const excelFixture = (isLegacy = false) => `
<div class="fixture">
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
              <button class="selected"><i aria-hidden="true">◎</i><span><strong>跟随当前</strong><small>发送时自动使用正在查看的工作表</small></span></button>
              <button><i aria-hidden="true">☷</i><span><strong>选择多个</strong><small>用于跨表查询、比较或汇总</small></span></button>
              <div class="scope-picker-reveal open">
                <div class="workbook-sheet-picker">
                  <div class="sheet-search-row"><input aria-label="搜索工作表" placeholder="搜索工作表..."><span>已选 3/3</span></div>
                  <div class="sheet-picker-toolbar"><button>全选搜索结果</button><button>清空</button></div>
                  <div class="workbook-sheet-list">
                    ${excelRow("格式验收", "当前工作表 · 12 行 · 4 列", true, isLegacy)}
                    ${excelRow("Sheet1", "8 行 · 3 列", true, isLegacy)}
                    ${excelRow("Sheet2", "6 行 · 3 列", false, isLegacy)}
                  </div>
                </div>
              </div>
              <button><i aria-hidden="true">⌑</i><span><strong>工作目录</strong><small>批量处理多个本地工作簿</small></span></button>
            </div>
          </fieldset>
          <div class="scope-admission-summary" role="status"><span>已选 2 张表 · 可分析</span></div>
          <div class="scope-action-buttons">
            <button class="scope-refresh"><i class="scope-refresh-glyph" aria-hidden="true">↻</i> 刷新</button>
            <button class="scope-confirm">完成</button>
          </div>
        </div>
      </fieldset>
    </div>
  </section>
</div>`;

// ── `--demo`：修复前 / 修复后并排的可交互页面（真样式表 + 同一段结构）──
//
// 数字由上面的断言负责，这张页面负责"眼睛"：无头浏览器量的是数值，
// 它画的是形状。`--demo` 不需要浏览器，也不跑构建。
//
// 抽成函数是为了让 `--shots` 能顺手截它一张（`demo-preview.png`）：
// 图要是手工产的，页面一改图就过期 —— 过期的手工图比没有图更误导。
const DEMO_PATH = path.join(root, "outputs", "scope-option-rows-2026-10-16", "scope-option-rows-demo.html");

function buildDemoPage() {
  const legacyScope = `
/* ── 修复前的形状，只在 .legacy 作用域里（要展示的就是"以前长什么样"）── */
.legacy .scope-admission-badge {
  display: block; margin-top: 3px; white-space: normal; overflow-wrap: anywhere;
  font-size: var(--font-size-caption); line-height: 1.4; color: var(--color-brand-text);
}
.legacy.wps-sheet-picker-options > button { min-height: 38px; }
.legacy.wps-sheet-picker-options > button i {
  display: grid; width: 20px; height: 20px; place-items: center;
  border: 1px solid var(--color-border-strong); border-radius: var(--radius-xs);
  color: var(--color-brand-text); font-style: normal;
}
.legacy.wps-sheet-picker-options > button.selected i {
  border-color: var(--color-brand-text); background: var(--color-brand-text);
  color: var(--color-background);
}
.legacy .wps-sheet-option-label { display: flex; min-width: 0; align-items: center; gap: 6px; }
.legacy .wps-sheet-option-name {
  flex: 1 1 auto; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
/* Excel：后代选择器把行内勾选框吃成 25px 灰方块 + 行**没有** min-height
   （行高由内容撑出，所以我写成 auto 而不是某个编出来的像素值）
   + fieldset 是块盒（块间距塌成 0） */
.legacy .scope-mode-list i {
  display: grid; width: 25px; height: 25px; flex: 0 0 25px; place-items: center;
  border-radius: var(--radius-xs); background: var(--color-surface-hover); font-style: normal;
}
.legacy .workbook-sheet-list > button { min-height: auto; }
.legacy .scope-selection-fields { display: block; gap: 0; }`;

  const demoSheetNames = ["格式验收", "Sheet1", "Sheet2", "Sheet3"];
  const badgeText = "可分析 · 修改仍需预览确认";
  const wpsRows = (isLegacy) => demoSheetNames.map((name, index) => {
    const picked = index < 2;
    const active = index === 0;
    const checkbox = `<i${isLegacy ? "" : ' class="sheet-checkbox"'}>${picked ? "✓" : ""}</i>`;
    const label = isLegacy
      ? `<span class="wps-sheet-option-label"><span class="wps-sheet-option-name">${name}</span>` +
        (picked ? `<small class="scope-admission-badge">${badgeText}</small>` : "") + `</span>`
      : `<span>${name}</span>`;
    return `<button type="button" class="${picked ? "selected" : ""}" aria-pressed="${picked}">` +
      `${checkbox}${label}${active ? "<small>当前表</small>" : ""}</button>`;
  }).join("");

  const demoPage = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<title>数据范围面板 · 选项行与块间距（修复前 / 修复后）</title>
<style>${excelCss}${wpsCss}</style>
<style>${legacyScope}</style>
<style>
  body { margin: 0; padding: 22px; background: #f4f6f5; color: #27352e;
    font-family: "Segoe UI", "Microsoft YaHei", sans-serif; }
  h1 { font-size: 16px; margin: 0 0 6px; }
  .note { color: #6b7b72; font-size: 12px; line-height: 1.7; margin: 0 0 18px; }
  .note code { background: #e8eeea; border-radius: 3px; padding: 0 3px; }
  .grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 18px; align-items: start; }
  .panel { background: #fff; border: 1px solid #dfe6e1; border-radius: 12px; padding: 12px 12px 14px; }
  .panel h2 { font-size: 13px; margin: 0 0 4px; }
  .panel .tag { font-size: 11px; font-weight: 700; border-radius: 999px; padding: 1px 8px; margin-left: 6px; }
  .tag.before { background: #fdecec; color: #a5413a; }
  .tag.after { background: #e6f2e9; color: #2f6b47; }
  .panel .sub { color: #6b7b72; font-size: 12px; margin: 0 0 10px; line-height: 1.6; }
  /* 舞台钉死在**判据脚本同一个宽度**（420px 视口 - 夹具 12px 内边距 ×2 = 396px）：
     行内两行说明在别的宽度下会折行，同一份夹具量出的行高就不一样
     （实测 700px 宽时 legacy 行量到 76.5px、396px 宽才是 57px）。
     对照页自己报的数必须和 --legacy 判据报的数逐字一致，否则读者会以为是两回事。 */
  .stage { width: 412px; box-sizing: border-box;
    background: #fbfcfb; border: 1px dashed #dfe6e1; border-radius: 8px; padding: 8px; }
  .stage .wps-sheet-picker-options { max-height: none; }
  .stage .scope-popover { position: static; max-height: none; }
  .stage .wps-current-file-card > span { width: 42px; height: 42px; }
  .measured { margin: 10px 0 0; font-size: 12px; line-height: 1.8; color: #27352e; }
  .measured b { color: #2f6b47; }
  .measured .bad { color: #a5413a; }
</style></head><body>
<h1>数据范围面板：选项行与块间距</h1>
<p class="note">样式表与两个任务窗格是同一份（<code>apps/*/src/styles.css</code>），行结构与
<code>WpsProbePanel.tsx</code> / <code>App.tsx</code> 逐字一致；「修复前」一侧的旧规则只在
<code>.legacy</code> 作用域里生效。四个舞台都钉在 <b>396px</b> 内容宽 —— 与判据脚本
420px 视口下的面板同宽，两处报出的行高因此是同一个数（换宽度会因折行而不同）。<br>
无头浏览器实测见 <code>scripts/check_scope_option_rows.mjs</code>（含 <code>--legacy</code> 反向对照）。</p>
<div class="grid">
  <div class="panel">
    <h2>WPS 工作表列表<span class="tag before">修复前</span></h2>
    <p class="sub">行高 38px（比「跟随当前 / 选择多个 / 工作目录」那几张卡矮一截）；行内还挂着
    「可分析 · 修改仍需预览确认」这句授权文案 —— 它就是用户要求去掉的那几个字。
    （更早一轮里这句话会另起一行、把行撑到 55.8px；那一项上一轮已改成同行盒，
    所以这里量到的 38px 是**本轮改前**的形状。）</p>
    <div class="stage"><div class="wps-sheet-picker-options sheet-picker-options legacy" data-rows>${wpsRows(true)}</div></div>
    <p class="measured" data-measure="rows:legacy"></p>
  </div>
  <div class="panel">
    <h2>WPS 工作表列表<span class="tag after">修复后</span></h2>
    <p class="sub">行内只有勾选框 + 表名（授权文案一个字都不在行里）；行高与
    「跟随当前 / 选择多个 / 工作目录」同一档（59px），勾选不再改变行高。
    整次选择的授权状态仍看得见 —— 在列表**下方**那一条汇总里。</p>
    <div class="stage"><div class="wps-sheet-picker-options sheet-picker-options" data-rows>${wpsRows(false)}</div></div>
    <p class="measured" data-measure="rows:fixed"></p>
  </div>
  <div class="panel">
    <h2>Excel 数据范围面板<span class="tag before">修复前</span></h2>
    <p class="sub">行内同时起着两件事：勾选框被模式图标的规则吃成 25px 灰方块；
    勾选行还多出一行「可分析 · 修改仍需预览确认」。「当前文件」卡与「跟随当前」之间
    只隔两条描边（间距 0）。</p>
    <div class="stage legacy" data-popover>${excelFixture(true)}</div>
    <p class="measured" data-measure="popover:legacy"></p>
  </div>
  <div class="panel">
    <h2>Excel 数据范围面板<span class="tag after">修复后</span></h2>
    <p class="sub">行内只剩勾选框 + 表名；勾选框 20px、不填色；块与块之间 10px，与 WPS 同一档。</p>
    <div class="stage" data-popover>${excelFixture(false)}</div>
    <p class="measured" data-measure="popover:fixed"></p>
  </div>
</div>
<script>
  var round = function (value) { return Math.round(value * 10) / 10; };
  var lines = function (pairs) {
    return pairs.map(function (pair) { return pair[0] + " <b>" + pair[1] + "</b>"; }).join("　·　");
  };
  Array.prototype.forEach.call(document.querySelectorAll("[data-measure]"), function (out) {
    var parts = out.dataset.measure.split(":");
    var kind = parts[0], scope = out.closest(".panel");
    var rows = scope.querySelector("[data-rows]");
    // 「空框」只看**未勾选**的行：勾选行的勾选框本来就该是实心品牌绿。
    var boxes = scope.querySelectorAll(".wps-sheet-picker-options > button > i, .workbook-sheet-list > button > i");
    var checkbox = boxes[boxes.length - 1];
    var box = checkbox ? checkbox.getBoundingClientRect() : null;
    if (kind === "rows") {
      var heights = Array.prototype.map.call(scope.querySelectorAll("[data-rows] > button"), function (button) {
        return round(button.getBoundingClientRect().height);
      });
      out.innerHTML = lines([
        ["行高（勾选/未勾选）", heights.join(" / ") + "px"],
        ["未勾选行的勾选框", box ? round(box.width) + "×" + round(box.height) + "px" : "—"],
        ["空框填充", checkbox ? getComputedStyle(checkbox).backgroundColor : "—"]
      ]);
    } else {
      var mode = scope.querySelector(".scope-mode-list > button");
      var card = scope.querySelector(".workbook-identity");
      var header = scope.querySelector(".scope-popover-header");
      /*
       * 行高一律量**未勾选**的那一行（最后一颗）。原因：这一页的「修复前」一侧
       * 把被投诉的授权文案也画出来了（那是**结构**差异），勾选行因此比未勾选行高。
       * 行高这条数值要和 --legacy 判据对齐，而那一边控制的是**样式**
       * （夹具本身不带文案）—— 拿带文案的行去比就是把两种机制混在一根数上。
       * 文案的差异由上面看得见的形状负责，不由这根数负责。
       */
      var sheetRows = scope.querySelectorAll(".workbook-sheet-list > button");
      var lastRow = sheetRows[sheetRows.length - 1];
      out.innerHTML = lines([
        ["标题→卡片", round(card.getBoundingClientRect().top - header.getBoundingClientRect().bottom) + "px"],
        ["卡片→模式卡", round(mode.getBoundingClientRect().top - card.getBoundingClientRect().bottom) + "px"],
        ["工作表行高（未勾选行）", round(lastRow.getBoundingClientRect().height) + "px"],
        ["勾选框（未勾选行）", box ? round(box.width) + "×" + round(box.height) + "px" : "—"]
      ]);
    }
  });
</script>
</body></html>`;
  return demoPage;
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

const wpsHtml = `<!doctype html><html><head><meta charset="utf-8"><style>${excelCss}${wpsCss}${noMotion}${legacy ? LEGACY_WPS : ""}</style></head>
<body><div id="root"></div><script>${wpsHostStub}</script><script>${bundledWps.outputFiles[0].text}</script></body></html>`;

/*
 * `.scope-popover` 在真机上贴着输入框往上弹（`bottom: calc(100% + 7px)`）。
 * 夹具页里没有输入框，就把它改成从顶部开始、并且**不设 max-height** ——
 * 要量的是块间距，不能被滚动的裁切盒影响，截图也要能看全。
 */
const excelHtml = `<!doctype html><html><head><meta charset="utf-8"><style>${excelCss}
.fixture { position: relative; min-height: 900px; padding: 12px; background: var(--color-background-subtle); }
.fixture .scope-popover { position: static; max-height: none; }
</style><style>${noMotion}${legacy ? LEGACY_EXCEL : ""}</style></head>
<body>${excelFixture(false)}</body></html>`;

// ── 页面侧测量 ────────────────────────────────────────────────────
/** 面板里点开「选择多个」，等列表真的有空间（不然量什么都是假绿）。 */
const openWpsPicker = async () => {
  const settle = () => new Promise((resolve) => setTimeout(resolve, 25));
  const waitFor = async (probe, label) => {
    for (let attempt = 0; attempt < 120; attempt += 1) {
      const value = probe();
      if (value) return value;
      await settle();
    }
    throw new Error(`等不到「${label}」；当前可见文本：${document.body.innerText.slice(0, 200)}`);
  };
  const byText = (selector, text) => [...document.querySelectorAll(selector)]
    .find((element) => (element.textContent ?? "").includes(text));
  (await waitFor(() => document.querySelector(".scope-trigger"), "数据范围入口")).click();
  (await waitFor(() => byText(".scope-mode-list button", "选择多个"), "「选择多个」")).click();
  const list = await waitFor(() => document.querySelector(".wps-sheet-picker-options"), "工作表列表");
  await waitFor(() => byText(".wps-sheet-picker-options button", "明细"), "列表里的「明细」");
  const heights = [];
  for (let attempt = 0; attempt < 40; attempt += 1) {
    heights.push(list.getBoundingClientRect().height);
    if (heights.length >= 3 && heights.at(-1) === heights.at(-2) && heights.at(-1) === heights.at(-3)) break;
    await settle();
  }
  return { listHeight: list.getBoundingClientRect().height, rows: list.querySelectorAll("button").length };
};

/** WPS：量行高、勾选框、模式卡、块间距。 */
const measureWps = ({ selected }) => {
  // 页面函数在浏览器里跑，看不到 node 侧的 helper —— 每个函数自带一份。
  const round = (value) => Math.round(value * 100) / 100;
  const list = document.querySelector(".wps-sheet-picker-options");
  const row = [...list.querySelectorAll("button")]
    .find((button) => (button.textContent ?? "").includes("明细"));
  const box = row.querySelector(":scope > i");
  const modeList = document.querySelector(".scope-controls .scope-mode-list");
  const modeBoxes = [...modeList.querySelectorAll(":scope > button")];
  const card = document.querySelector(".wps-current-file-card");
  const groupBox = modeBoxes.at(-1);
  const computed = getComputedStyle(box);
  return {
    rowHeight: round(row.getBoundingClientRect().height),
    rowPressed: row.getAttribute("aria-pressed"),
    rowHasBadge: Boolean(row.querySelector(".scope-admission-badge")),
    rowText: row.textContent ?? "",
    checkbox: {
      width: round(box.getBoundingClientRect().width),
      height: round(box.getBoundingClientRect().height),
      background: computed.backgroundColor,
      borderWidth: computed.borderTopWidth
    },
    modeBoxHeights: modeBoxes.map((element) => round(element.getBoundingClientRect().height)),
    groupBoxHeight: round(groupBox.getBoundingClientRect().height),
    cardToModeGap: round(modeBoxes[0].getBoundingClientRect().top - card.getBoundingClientRect().bottom),
    selected
  };
};

/** WPS：勾选/取消一张表（React 提交是异步的，等 DOM 真的变了再量）。 */
const toggleWps = async ({ name, expect }) => {
  const row = () => [...document.querySelectorAll(".wps-sheet-picker-options button")]
    .find((button) => (button.textContent ?? "").includes(name));
  row().click();
  for (let attempt = 0; attempt < 60; attempt += 1) {
    if (row()?.getAttribute("aria-pressed") === expect) return true;
    await new Promise((resolve) => setTimeout(resolve, 25));
  }
  throw new Error(`「${name}」的勾选态没有变成 ${expect}`);
};

/** Excel 夹具：量同一批几何量。 */
const measureExcelFixture = () => {
  const round = (value) => Math.round(value * 100) / 100;
  const modeBoxes = [...document.querySelectorAll(".scope-mode-list > button")];
  const rows = [...document.querySelectorAll(".workbook-sheet-list > button")];
  // 「空框」的样子只能看**未勾选**的行（夹具里最后一行是未勾选的）。
  const box = rows.at(-1).querySelector(":scope > i");
  const computed = getComputedStyle(box);
  const header = document.querySelector(".scope-popover-header");
  const card = document.querySelector(".workbook-identity");
  return {
    modeBoxHeights: modeBoxes.map((element) => round(element.getBoundingClientRect().height)),
    groupBoxHeight: round(modeBoxes.at(-1).getBoundingClientRect().height),
    rowHeights: rows.map((element) => round(element.getBoundingClientRect().height)),
    rowHasBadge: Boolean(rows.at(-1).querySelector(".scope-admission-badge")),
    rowText: rows.at(-1).textContent ?? "",
    checkbox: {
      width: round(box.getBoundingClientRect().width),
      height: round(box.getBoundingClientRect().height),
      background: computed.backgroundColor,
      borderWidth: computed.borderTopWidth
    },
    selectedCheckboxBackground: getComputedStyle(rows[0].querySelector(":scope > i")).backgroundColor,
    headerToCardGap: round(card.getBoundingClientRect().top - header.getBoundingClientRect().bottom),
    cardToModeGap: round(modeBoxes[0].getBoundingClientRect().top - card.getBoundingClientRect().bottom)
  };
};

// ── 跑起来 ────────────────────────────────────────────────────────
/**
 * 截图只截**面板自身**那一块。
 * 整页截会在两种情况下骗人：面板向上弹时跑到视口外（一片空白），
 * 或者被祖先的 `overflow: hidden` 裁掉一半。按 rect 裁切就没有这些歧义。
 */
const popoverClip = async (page) => page.evaluate(() => {
  const box = document.querySelector(".scope-popover").getBoundingClientRect();
  return { x: box.left, y: box.top, width: box.width, height: box.height };
});

const serve = async (html) => {
  const server = createServer((_req, res) => {
    res.setHeader("Content-Type", "text/html; charset=utf-8");
    res.end(html);
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  return { server, url: `http://127.0.0.1:${server.address().port}` };
};

let browser;
const shotDir = path.join(root, "outputs", "scope-option-rows-2026-10-16");
try {
  const chromium = chromiumDriver();
  browser = await chromium.launch({ channel: "msedge", headless: true }).catch(() =>
    chromium.launch({ channel: "chrome", headless: true }));
  if (shots) await mkdir(shotDir, { recursive: true });

  // ── ① WPS：真实面板 ──
  const wps = await serve(wpsHtml);
  const wpsPage = await browser.newPage({ viewport: { width: 420, height: 900 } });
  await wpsPage.goto(wps.url);
  await wpsPage.locator(".scope-trigger").waitFor({ state: "attached", timeout: 15000 });
  const opened = await wpsPage.evaluate(openWpsPicker);
  assert.ok(opened.listHeight > 60,
    `工作表列表被祖先压扁成 ${opened.listHeight}px，量不到真实行高（必须先有空间，否则"没变高"是假绿）`);

  const wpsOff = await wpsPage.evaluate(measureWps, { selected: false });
  await wpsPage.evaluate(toggleWps, { name: "明细", expect: "true" });
  const wpsOn = await wpsPage.evaluate(measureWps, { selected: true });

  assert.ok(!wpsOff.rowHasBadge && !wpsOn.rowHasBadge, "WPS 行内不该有授权徽标");
  assert.ok(!wpsOff.rowText.includes("可分析"), `WPS 行内不该出现「可分析」：${wpsOff.rowText}`);
  assert.ok(wpsOff.rowText.includes("明细"), "WPS 行内必须还有工作表名");
  assert.equal(wpsOn.checkbox.height, wpsOff.checkbox.height, "勾选改变了勾选框尺寸");
  for (const [label, value] of [["未勾选", wpsOff], ["已勾选", wpsOn]]) {
    assert.equal(value.checkbox.width, 20, `${label}时勾选框应当是 20px 宽`);
    assert.equal(value.checkbox.height, 20, `${label}时勾选框应当是 20px 高`);
  }
  // 空框**不填色**（填了会与行底色重合、"空框"整个隐进列表里）；勾上才整块换品牌绿。
  assert.equal(wpsOff.checkbox.background, "rgba(0, 0, 0, 0)",
    `未勾选时勾选框不该有填充：${wpsOff.checkbox.background}`);
  assert.equal(wpsOn.checkbox.background, "rgb(47, 107, 71)", "勾选后勾选框应当整块换成品牌绿");
  await wpsPage.evaluate(toggleWps, { name: "明细", expect: "false" });
  const wpsBack = await wpsPage.evaluate(measureWps, { selected: false });
  assert.equal(wpsBack.checkbox.background, "rgba(0, 0, 0, 0)", "取消勾选后应当回到不填色");
  assert.equal(wpsBack.rowHeight, wpsOff.rowHeight, "取消勾选后行高应当回到原值");
  if (shots) {
    /*
     * 面板在真机上贴着输入框往上弹（`bottom: calc(100% + 7px)`）；夹具页里没有输入框，
     * 它会跑到视口上方去，截出来是一片空白。**测量已经全部跑完**了，这里才把
     * 外壳撑出一个定位上下文，纯粹为了能把面板截进画面。加了也不影响已取到的数字。
     */
    await wpsPage.setViewportSize({ width: 420, height: 1000 });
    // 夹具里没有输入框，面板向上的那个定位会被祖先的 `overflow: hidden` 裁掉。
    // 截图时让它改回文档流。
    await wpsPage.addStyleTag({ content: ".scope-popover { position: static !important; max-height: none; }" });
    // 测量期间的点击会把焦点滚进视野，滚动位置留着会让截图只看到列表尾巴。
    await wpsPage.evaluate(() => {
      window.scrollTo(0, 0);
      const scroller = document.querySelector(".scope-popover-scroll");
      if (scroller) scroller.scrollTop = 0;
    });
    /*
     * WPS 面板在这一格里**不截图**：面板是「贴着输入框往上弹」的绝对定位元素，
     * 夹具页里没有输入框、它的定位基准还被祖先的 `overflow: hidden` 裁掉，
     * 截出来只剩最后一行（试过 static/裁剪/page 滚动，都会得到会骗人的图）。
     * 要"看"WPS 的行高请用 `--demo`（真样式表 + 同一段行结构，修复前/后并排）。
     */
  }

  // ── ② Excel 夹具（同一份真实样式表）──
  const excel = await serve(excelHtml);
  const excelPage = await browser.newPage({ viewport: { width: 420, height: 900 } });
  await excelPage.goto(excel.url);
  const excelM = await excelPage.evaluate(measureExcelFixture);
  if (shots) {
    const clip = await popoverClip(excelPage);
    await excelPage.screenshot({ path: path.join(shotDir, `excel-${legacy ? "legacy" : "fixed"}.png`), clip });
    console.log(`  Excel 截图：面板 ${Math.round(clip.width)}×${Math.round(clip.height)} @y=${Math.round(clip.y)}`);
    /*
     * 顺手重画对照页并整页截一张。图与页出自同一个 `buildDemoPage()`，
     * 不会出现"页面改了、图还是旧的"。
     *
     * 必须**等布局稳定**再截：`.scope-picker-reveal` 的 `0fr → 1fr` 展开有 360ms 过渡，
     * 早截会按当时更矮的文档高度出图 —— 上半页在、下半页空白，而且正是工作表列表
     * 那一段（本次改动的正主）被切掉。这里轮询到两次测量一致为止。
     */
    const demoUrl = `file:///${(await writeDemoPage()).split(path.sep).join("/")}`;
    const demoShot = await browser.newPage({ viewport: { width: 1180, height: 900 } });
    await demoShot.goto(demoUrl);
    await demoShot.evaluate(async () => {
      const height = () => {
        const popover = document.querySelector(".stage[data-popover] .scope-popover");
        return popover ? Math.round(popover.getBoundingClientRect().height) : 0;
      };
      let previous = -1;
      for (let attempt = 0; attempt < 60 && height() !== previous; attempt += 1) {
        previous = height();
        await new Promise((resolve) => setTimeout(resolve, 40));
      }
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    });
    /*
     * 高度不能读 `documentElement.scrollHeight`：这张页里带着**应用自己的样式表**，
     * `html, body { height: 100% }` 把 scrollHeight 钉死在视口高度上，
     * 照它设视口就会把下半页（工作表列表那一段）整段截掉。
     * 按四个面板的实际下沿算，才是内容真正需要的高度。
     */
    const pageHeight = await demoShot.evaluate(() => {
      const panels = [...document.querySelectorAll(".panel")];
      const bottom = panels.length
        ? Math.max(...panels.map((panel) => panel.getBoundingClientRect().bottom + window.scrollY))
        : 0;
      return Math.ceil(Math.max(bottom, document.body.scrollHeight) + 24);
    });
    await demoShot.setViewportSize({ width: 1180, height: pageHeight });
    await demoShot.screenshot({ path: path.join(shotDir, "demo-preview.png") });
    await demoShot.close();
    console.log(`  对照页截图：demo-preview.png（1180×${pageHeight}）`);
  }

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

  // ③ 行高统一：工作表行 == 模式卡（= 工作目录卡）高。
  check("WPS 工作表行高（未勾选）", wpsOff.rowHeight, wpsOff.groupBoxHeight);
  check("WPS 工作表行高（已勾选）", wpsOn.rowHeight, wpsOn.groupBoxHeight);
  check("Excel 工作表行高", excelM.rowHeights[0], excelM.groupBoxHeight);
  // ④ 块间距：两端同一批度量。
  check("WPS 卡片→模式卡间距", wpsOff.cardToModeGap, 10);
  check("Excel 卡片→模式卡间距", excelM.cardToModeGap, 10);
  check("Excel 标题→卡片间距", excelM.headerToCardGap, 10);
  // ② 勾选框不再被模式图标的规则吃掉（Excel 夹具侧同一批度量）。
  check("Excel 勾选框宽度", excelM.checkbox.width, 20);
  check("Excel 勾选框填充（未勾选）", excelM.checkbox.background, "rgba(0, 0, 0, 0)");
  // 夹具侧的行内文案（与 WPS 侧同一批度量：行里只该有表名与行数说明）。
  assert.ok(!excelM.rowHasBadge, "Excel 行内不该有授权徽标");
  assert.ok(!excelM.rowText.includes("可分析"), `Excel 行内不该出现「可分析」：${excelM.rowText}`);
  assert.equal(excelM.selectedCheckboxBackground, "rgb(47, 107, 71)", "勾选后勾选框应当整块换成品牌绿");

  console.log(legacy ? "legacy（修复前形状）实测：" : "实测：");
  console.log(`  WPS 列表：${opened.rows} 行，列表高 ${Math.round(opened.listHeight)}px`);
  console.log(`  WPS 模式卡高度 [${wpsOff.modeBoxHeights.join(", ")}]，工作目录卡 ${wpsOff.groupBoxHeight}px`);
  console.log(`  WPS 工作表行：未勾选 ${wpsOff.rowHeight}px → 已勾选 ${wpsOn.rowHeight}px`);
  console.log(`  WPS 勾选框：${wpsOff.checkbox.width}×${wpsOff.checkbox.height}，` +
    `未选填充 ${wpsOff.checkbox.background} → 已选填充 ${wpsOn.checkbox.background}，描边 ${wpsOff.checkbox.borderWidth}`);
  console.log(`  Excel 模式卡高度 [${excelM.modeBoxHeights.join(", ")}]，工作表行 ${excelM.rowHeights.join("/")}px`);
  console.log(`  Excel 勾选框：${excelM.checkbox.width}×${excelM.checkbox.height}，填充 ${excelM.checkbox.background}，描边 ${excelM.checkbox.borderWidth}`);
  console.log(`  Excel 间距：标题→卡片 ${excelM.headerToCardGap}px，卡片→模式卡 ${excelM.cardToModeGap}px`);
  for (const row of report) {
    console.log(`  ${legacy ? "✗ 应当不等" : "✓"} ${row.name}：实测 ${row.actual}，期望 ${row.expected}`);
  }
  console.log(legacy
    ? "结论：修复前的写法在每一项上都能被这条测量判红 —— 测量有判别力。"
    : "结论：两端选项行同高、块间距一致、勾选框尺寸与填充一致。");
  if (shots) console.log(`截图 → ${shotDir}`);

  await wps.server.close();
  await excel.server.close();
} finally {
  await browser?.close();
}
