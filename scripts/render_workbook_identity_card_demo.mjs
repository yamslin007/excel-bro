// 示意对照：数据范围弹层里「当前文件卡 + 数据范围触发器」的日期显示（2026-10-16 夜）。
//
// 这条口径走过两步，对照图要**两格反向对照**才讲得清：
//   ① 最初：「文件名中未识别到报表日期 · 当前工作表 Sheet2」——用户最早投诉的就是这句；
//   ② 上一版：那句否定式说明删了，但**识别本身还在** —— 文件名里有日期就照旧显示
//      「报表日期 2026-03-15」，而且**连收起来的触发器都带着**「2026-03-15 · 当前表 Sheet2」；
//   ③ 现在：「不需要 识别日期」—— `extractWorkbookDataPeriod` 与 `workbookDataPeriod` 整条删除，
//      卡片与触发器都只剩"哪个文件、哪张表"，**文件名里有没有日期都一样**。
//
// 为什么两个调用点都要画：只画卡片的话，"只把卡片上的日期删掉、触发器里还留着"这种
// 半截状态看起来完全正常 —— 上一版就是这样漏的。
//
// 这一页只做两件事：
//   ① 用 `apps/excel-addin/src/styles.css`（**真实样式表**）+ 从 `App.tsx` 逐字抄来的结构
//      （`div.workbook-identity` 与 `button.scope-trigger`），把四格并排画出来；
//   ② 在真浏览器里把每格的两处文案读回来复核，并截一张图。
//
// 诚实的边界：这是**夹具页 + 复印的结构**，不是实机 Excel 截图。它的证据等级等于
// `check_scope_*.mjs` 里那两条 Excel 夹具泳道。真正的取值路径由
// `apps/excel-addin/src/workbookIdentityCard.test.tsx`（挂真实 `App.tsx`）钉住，
// 反向变异见 `scripts/run_workbook_identity_card_mutations.mjs`。
//
// Usage: node scripts/render_workbook_identity_card_demo.mjs [path/to/playwright-core]
import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { createServer } from "node:http";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const require = createRequire(import.meta.url);
const argModule = process.argv.slice(2).find((value) => !value.startsWith("--"));
let playwright = null;
function chromiumDriver() {
  if (!playwright) ({ chromium: playwright } = require(argModule || "playwright-core"));
  return playwright;
}
const root = path.resolve(import.meta.dirname, "..");
const outDir = path.join(root, "outputs", "workbook-identity-card-2026-10-16");
const excelCss = await readFile(path.join(root, "apps/excel-addin/src/styles.css"), "utf8");

/** 当前文件卡 —— 结构逐字取自 `App.tsx` 的 `div.workbook-identity`。 */
const card = (name, em) => `
      <div class="workbook-identity">
        <i aria-hidden="true">XLSX</i>
        <span>
          <small>当前文件</small>
          <strong title="${name}">${name}</strong>
          <em>${em}</em>
        </span>
      </div>`;

/**
 * 收起来的触发器 —— 结构逐字取自 `App.tsx` 的 `button.scope-trigger`
 * （回形针图标 + `small` 文件名 + `|` + `strong` 文案 + `b` 箭头）。
 */
const trigger = (name, label) => `
      <div class="scope-entry-row">
        <button class="scope-trigger" aria-expanded="false">
          <i aria-hidden="true"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M19.2 12.1l-6.9 6.9a4.6 4.6 0 0 1-6.5-6.5l6.9-6.9a3 3 0 0 1 4.3 4.3l-6.9 6.9a1.4 1.4 0 0 1-2-2l6.2-6.2"/></svg></i>
          <span>
            <small title="${name}">${name}</small>
            <span class="scope-sep" aria-hidden="true">|</span>
            <strong>${label}</strong>
          </span>
          <b>⌃</b>
        </button>
      </div>`;

/** `scope-popover` 在真机上贴着输入框往上弹；夹具页里改成静态，才截得进画面。 */
const stage = (name, em, label) => `
    <div class="stage">
      <section class="scope-popover" aria-label="选择数据范围">${card(name, em)}</section>
      ${trigger(name, label)}
    </div>`;

const CELLS = [
  {
    key: "v1-missing",
    title: "最初 · 文件名里没有日期",
    tag: "before",
    sub: "用户最早投诉的那句：`评分表.xlsx` 本来就不含日期，卡片却回了一句「文件名中未识别到报表日期」。"
      + "（悬空的「·」也是同一处败笔。）触发器那一侧此时还看不出问题——文件名里没有日期。",
    stage: stage("评分表.xlsx", "文件名中未识别到报表日期 · 当前工作表 Sheet2", "当前表 Sheet2"),
    expect: {
      card: "文件名中未识别到报表日期 · 当前工作表 Sheet2",
      trigger: "当前表 Sheet2"
    }
  },
  {
    key: "v2-dated",
    title: "上一版 · 文件名里有日期",
    tag: "before",
    sub: "否定式说明删了，但**识别本身还在**：卡片显示「报表日期 2026-03-15」，"
      + "**触发器也带着**「2026-03-15 · 当前表 Sheet2」。这一格就是那个半截状态 —— "
      + "只把卡片上的日期删掉、触发器里还留着，看上去完全正常。",
    stage: stage("评分表-2026-03-15.xlsx", "报表日期 2026-03-15 · 当前工作表 Sheet2", "2026-03-15 · 当前表 Sheet2"),
    expect: {
      card: "报表日期 2026-03-15 · 当前工作表 Sheet2",
      trigger: "2026-03-15 · 当前表 Sheet2"
    }
  },
  {
    key: "v3-plain",
    title: "现在 · 文件名里没有日期",
    tag: "after",
    sub: "只有「哪个文件、哪张表」。与 WPS 的当前文件卡同口径 —— 那边本来就只有「当前工作表 X」。",
    stage: stage("评分表.xlsx", "当前工作表 Sheet2", "当前表 Sheet2"),
    expect: { card: "当前工作表 Sheet2", trigger: "当前表 Sheet2" }
  },
  {
    key: "v3-dated",
    title: "现在 · 文件名里有日期也不显示",
    tag: "after",
    sub: "「不需要 识别日期」= 识别这件事本身下线，不是「有则显示」。这一格是防假绿的锚点："
      + "若判据只测上一格，「把识别请回来、有日期就显示」的版本能全身而过。",
    stage: stage("评分表-2026-03-15.xlsx", "当前工作表 Sheet2", "当前表 Sheet2"),
    expect: { card: "当前工作表 Sheet2", trigger: "当前表 Sheet2" }
  }
];

const page = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<title>数据范围弹层 · 当前文件卡与触发器的日期显示（最初 / 上一版 / 现在）</title>
<style>${excelCss}</style>
<style>
  *, *::before, *::after { transition: none !important; animation: none !important; }
  body { margin: 0; padding: 22px; background: #f4f6f5; color: #27352e;
    font-family: "Segoe UI", "Microsoft YaHei", sans-serif; }
  h1 { font-size: 16px; margin: 0 0 6px; }
  .note { color: #6b7b72; font-size: 12px; line-height: 1.7; margin: 0 0 18px; }
  .note code { background: #e8eeea; border-radius: 3px; padding: 0 3px; }
  .grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; align-items: start; }
  .panel { background: #fff; border: 1px solid #dfe6e1; border-radius: 12px; padding: 12px 12px 14px; }
  .panel h2 { font-size: 13px; margin: 0 0 4px; }
  .panel .tag { font-size: 11px; font-weight: 700; border-radius: 999px; padding: 1px 8px; margin-left: 6px; }
  .tag.before { background: #fdecec; color: #a5413a; }
  .tag.after { background: #e6f2e9; color: #2f6b47; }
  .panel .sub { color: #6b7b72; font-size: 12px; margin: 0 0 10px; line-height: 1.6; }
  /* 舞台宽度与判据脚本同源（面板内容宽）；换宽度会改变折行，两处的数就不同源了。 */
  .stage { box-sizing: border-box; background: #fbfcfb; border: 1px dashed #dfe6e1;
    border-radius: 8px; padding: 10px; }
  .stage .scope-popover { position: static; max-height: none; }
  .readout { margin: 10px 0 0; font-size: 12px; line-height: 1.9; }
  .readout b { color: #2f6b47; }
  .readout .warn { color: #a5413a; }
</style></head><body>
<h1>数据范围弹层：当前文件卡与触发器里的日期</h1>
<p class="note">样式表是应用自己的那一份（<code>apps/excel-addin/src/styles.css</code>），
卡片与触发器的结构与 <code>App.tsx</code> 逐字一致，四格只有两处文案不同。<br>
<b>这是夹具页，不是实机 Excel 截图。</b>判据在
<code>apps/excel-addin/src/workbookIdentityCard.test.tsx</code>（挂真实 <code>App.tsx</code>，
卡片与触发器**两个调用点**都断言），反向变异见
<code>scripts/run_workbook_identity_card_mutations.mjs</code>。</p>
<div class="grid">
${CELLS.map((cell) => `  <div class="panel" data-cell="${cell.key}">
    <h2>${cell.title}<span class="tag ${cell.tag}">${cell.tag === "before" ? "旧" : "现在"}</span></h2>
    <p class="sub">${cell.sub}</p>
    ${cell.stage}
    <p class="readout" data-readout="${cell.key}"></p>
  </div>`).join("\n")}
</div>
<script>
  var cells = ${JSON.stringify(CELLS.map(({ key }) => key))};
  cells.forEach(function (key) {
    var panel = document.querySelector('[data-cell="' + key + '"]');
    var cardText = panel.querySelector(".workbook-identity em").textContent;
    var triggerText = panel.querySelector(".scope-trigger strong").textContent;
    var hasDate = /\\d{4}/.test(cardText + triggerText);
    panel.querySelector("[data-readout]").innerHTML =
      "卡片：" + (cardText.includes("报表日期") || cardText.includes("未识别到")
        ? '<span class="warn">「' + cardText + '」</span>' : '<b>「' + cardText + '」</b>') +
      "<br>触发器：" + (hasDate
        ? '<span class="warn">「' + triggerText + '」</span>' : '<b>「' + triggerText + '」</b>');
  });
</script>
</body></html>`;

await mkdir(outDir, { recursive: true });
const htmlPath = path.join(outDir, "workbook-identity-card-demo.html");
await writeFile(htmlPath, page, "utf8");

const server = createServer((_req, res) => {
  res.setHeader("Content-Type", "text/html; charset=utf-8");
  res.end(page);
});
await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
const url = `http://127.0.0.1:${server.address().port}`;

let browser;
try {
  const chromium = chromiumDriver();
  browser = await chromium.launch({ channel: "msedge", headless: true }).catch(() =>
    chromium.launch({ channel: "chrome", headless: true }));
  const view = await browser.newPage({ viewport: { width: 940, height: 900 } });
  await view.goto(url);
  await view.locator("[data-readout]").first().waitFor({ state: "attached" });

  const readBack = await view.evaluate(() => Object.fromEntries(
    [...document.querySelectorAll("[data-cell]")].map((panel) => [
      panel.dataset.cell,
      {
        card: panel.querySelector(".workbook-identity em").textContent,
        trigger: panel.querySelector(".scope-trigger strong").textContent
      }
    ])
  ));

  for (const cell of CELLS) {
    const actual = readBack[cell.key];
    assert.equal(actual.card, cell.expect.card,
      `${cell.key} 卡片的第二行与预期不符：${actual.card}`);
    assert.equal(actual.trigger, cell.expect.trigger,
      `${cell.key} 触发器的文案与预期不符：${actual.trigger}`);
    console.log(`  ✓ ${cell.title}：卡片「${actual.card}」· 触发器「${actual.trigger}」`);
  }
  // 反向对照在场：两格旧状态必须真的画出老样子，否则这张对照图没有判别力。
  assert.ok(readBack["v1-missing"].card.includes("未识别到"),
    "「最初」那一格必须画出用户投诉的那句否定式说明");
  assert.ok(readBack["v2-dated"].card.includes("报表日期")
    && /\d{4}/.test(readBack["v2-dated"].trigger),
    "「上一版」那一格必须同时画出卡片与触发器上的日期（这才是那个半截状态）");
  // 现在两格：两处都不许出现四位年份。
  for (const key of ["v3-plain", "v3-dated"]) {
    assert.ok(!/\d{4}/.test(readBack[key].card + readBack[key].trigger),
      `${key} 不该再出现任何日期：${readBack[key].card} / ${readBack[key].trigger}`);
  }
  console.log("  ✓ 反向对照：最初那格画出否定式说明；上一版那格卡片与触发器都带日期");
  console.log("  ✓ 现在两格（有日期 / 无日期）：卡片与触发器都不含四位年份");

  const height = await view.evaluate(() => Math.ceil(
    Math.max(document.body.scrollHeight,
      Math.max(...[...document.querySelectorAll(".panel")].map((panel) => panel.getBoundingClientRect().bottom))
    ) + 24
  ));
  await view.setViewportSize({ width: 940, height });
  const shotPath = path.join(outDir, "workbook-identity-card-demo.png");
  await view.screenshot({ path: shotPath });
  console.log(`  ✓ 对照图：${shotPath}（940×${height}）`);
  console.log(`  ✓ 对照页：${htmlPath}`);
} finally {
  await browser?.close();
  server.close();
}
