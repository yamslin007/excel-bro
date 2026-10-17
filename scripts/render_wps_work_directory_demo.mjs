// 示意对照：WPS 数据范围面板的**第三种来源 —— 工作目录**（2026-10-16 夜）。
//
// 这一轮把 WPS 上那颗 `disabled` 的「工作目录」卡按 Excel 的同一套口径打开：
// 选目录 → 文件/表**两级勾选** → 承认汇总 → **确认才提交**授权 → 草稿与取消 →
// 刷新增表不作废、换目录/目标消失即作废；发送侧在工作目录范围下**明确拒绝**，
// 绝不拿当前打开的 WPS 工作簿顶上。
//
// 对照要画三件事，缺一件这张图就没判别力：
//   ① 卡片：旧的是死按钮 + 副标题写着那句能力说明（用户截图里的样子）；
//   ② 勾选区：两级勾选、EB 遗留表**可读不可勾**、已选页脚、承认汇总；
//   ③ 发送侧：工作目录范围下零出站、零工作簿读取，给诚实提示并把输入留回框里 ——
//      外加触发器那一格（工作目录模式下**不许**出现当前工作簿名）。
//
// 这一页只做两件事：
//   ① 用 `apps/wps-addin/src/styles.css`（它自己 import 了 Excel 的样式表，与 `main.tsx`
//      的加载顺序一致）+ 从 `WpsProbePanel.tsx` 逐字抄来的结构，把五个场景画出来；
//   ② 在真浏览器里把每格的关键文案读回来复核，并截一张图。
//
// 诚实的边界：这是**夹具页 + 复印的结构**，不是实机 WPS 截图。它的证据等级等于那两个
// 夹具泳道。真正的取值路径由 `apps/wps-addin/src/wpsWorkDirectoryScope.test.tsx`（挂真实
// `WpsProbePanel`）与 `wpsWorkDirectorySendGate.test.tsx`（挂真实 `WpsAgentPanel`，量
// 宿主读写次数与出站请求）钉住，反向变异见 `scripts/run_wps_work_directory_mutations.mjs`。
//
// Usage: node scripts/render_wps_work_directory_demo.mjs [path/to/playwright-core]
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
const outDir = path.join(root, "outputs", "wps-work-directory-2026-10-16");
// 与 `apps/wps-addin/src/main.tsx` 同一个加载顺序：先 Excel 的样式表，再 WPS 的。
const excelCss = await readFile(path.join(root, "apps/excel-addin/src/styles.css"), "utf8");
const wpsCss = await readFile(path.join(root, "apps/wps-addin/src/styles.css"), "utf8");

/* ── 结构：全部逐字取自 `apps/wps-addin/src/WpsProbePanel.tsx` ───────────────── */

/** 三张来源卡。`folderDisabled` 与 `folderCaption` 是旧 / 现在唯一的差别。 */
const modeList = ({ folderDisabled, folderCaption, folderSelected }) => `
        <div class="scope-mode-list">
          <button type="button" class="${folderSelected ? "" : "selected"}">
            <i aria-hidden="true">⊙</i>
            <span><strong>跟随当前</strong><small>发送时自动使用正在查看的工作表</small></span>
          </button>
          <button type="button">
            <i aria-hidden="true">☷</i>
            <span><strong>选择多个</strong><small>用于跨表查询、比较或汇总</small></span>
          </button>
          <button type="button" class="${folderSelected ? "selected" : ""}"${folderDisabled ? " disabled" : ""}>
            <i aria-hidden="true">⌑</i>
            <span><strong>工作目录</strong><small>${folderCaption}</small></span>
          </button>
        </div>`;

const fileOption = ({ relativePath, selectedCount, selectableCount, active }) => `
                          <button type="button" class="folder-file-option${active ? " active" : ""}"
                            aria-pressed="${active}" role="listitem">
                            <i aria-hidden="true">X</i>
                            <span>
                              <strong title="${relativePath}">${relativePath}</strong>
                              <small>${selectableCount} 张工作表</small>
                            </span>
                            <b>${selectedCount}/${selectableCount}</b>
                          </button>`;

const sheetOption = ({ name, rowCount, columnCount, selected }) => `
                                  <button type="button" class="${selected ? "selected" : ""}"
                                    aria-pressed="${selected}">
                                    <i class="sheet-checkbox" aria-hidden="true">${selected ? "✓" : ""}</i>
                                    <span>
                                      <strong>${name}</strong>
                                      <small>${rowCount} 行 · ${columnCount} 列</small>
                                    </span>
                                  </button>`;

/** 工作目录勾选区（目录已就绪、勾了 `销售.xlsx / 明细` 一张）。 */
/** 弹层滚动容器里的内容（目录工具条 + 两级勾选）—— 页脚与承认汇总**不在**这里面。 */
const folderPicker = () => `
            <div class="folder-scope-picker">
              <div class="folder-toolbar">
                <button type="button" class="browse-folder-button">更换工作目录 · 数据</button>
                <button type="button" class="folder-refresh-button" title="重新扫描文件和工作表">↻ 刷新</button>
              </div>
              <p class="folder-save-hint">在其他工作簿新增或改名后，请先保存文件再刷新。</p>
              <div class="folder-refresh-notice has-new" role="status">发现 1 张新增工作表</div>
              <div class="folder-selection-filter">
                <span>2 个文件 · 2 张表</span>
                <button type="button">只看已选</button>
              </div>
              <div class="folder-master-detail folder-picker-view-sheets">
                <div class="folder-file-list" role="list" aria-label="工作簿列表">
${fileOption({ relativePath: "销售.xlsx", selectedCount: 1, selectableCount: 1, active: true })}
${fileOption({ relativePath: "库存.xlsx", selectedCount: 0, selectableCount: 1, active: false })}
                </div>
                <section class="folder-sheet-detail" aria-label="选择 销售.xlsx 中的工作表">
                  <div class="folder-sheet-header">
                    <button type="button" class="folder-back-button">← 文件</button>
                    <span>
                      <strong title="销售.xlsx">销售.xlsx</strong>
                      <small>已选 1/1</small>
                    </span>
                  </div>
                  <div class="sheet-search-row folder-sheet-search">
                    <input value="" placeholder="搜索当前文件的工作表..." aria-label="搜索当前文件的工作表" />
                    <span>1 项</span>
                  </div>
                  <div class="sheet-picker-toolbar">
                    <button type="button">全选本文件</button>
                    <button type="button">清空本文件</button>
                  </div>
                  <div class="folder-sheet-list">
${sheetOption({ name: "明细", rowCount: 8, columnCount: 3, selected: true })}
                  </div>
                </section>
              </div>
            </div>`;

/**
 * 工作目录侧的固定页脚 + 承认汇总 —— 与 `WpsProbePanel.tsx` 同构（也即与 Excel 同构）：
 * 计数在左、确认在右（靠共享样式表里的 `.folder-selection-footer .scope-confirm
 * { margin-left: auto }` 把确认贴到右边），承认汇总落在页脚**下方**。
 * 它作为 `popover(inner, extra)` 的 `extra` 渲染 ⇒ 是 `.scope-popover-scroll` 的**兄弟**、
 * 住在弹层上 —— 和真机一样，内容滚起来不会把页脚一起滚走。
 */
const folderFooter = () => `
            <div class="folder-selection-footer">
              <div class="folder-selection-summary">
                <strong>已选 1 个文件 · 1 张表</strong>
                <details class="folder-selection-help">
                  <summary aria-label="查看工作表选择保留说明" title="查看说明">ⓘ</summary>
                  <span role="note">切换文件不会丢失已勾选的工作表</span>
                </details>
              </div>
              <button type="button" class="scope-confirm">确认</button>
            </div>
            <div class="scope-admission-summary" role="status"><span>已选 1 张表 · 可分析</span></div>`;

/** 收起来的触发器。`name === null` = 工作目录模式（**不给** `<small>`）。 */
const trigger = ({ name, label }) => `
        <div class="scope-entry-row">
          <button class="scope-trigger" aria-expanded="false">
            <i aria-hidden="true"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M19.2 12.1l-6.9 6.9a4.6 4.6 0 0 1-6.5-6.5l6.9-6.9a3 3 0 0 1 4.3 4.3l-6.9 6.9a1.4 1.4 0 0 1-2-2l6.2-6.2"/></svg></i>
            <span>
              ${name === null ? "" : `<small title="${name}">${name}</small>
              <span class="scope-sep" aria-hidden="true">|</span>`}
              <strong>${label}</strong>
            </span>
            <b>⌃</b>
          </button>
        </div>`;

/** 发送侧：一条 system 消息 + 一个仍留着草稿的输入框（结构与 `WpsAgentPanel` 一致）。 */
const composerStage = (refusal, draft) => `
        <div class="message-stream-shell">
          <main class="message-stream">
            <article class="message-row system">
              <div class="message-content">
                <div class="message-author">系统</div>
                <p class="message-text" data-refusal>${refusal}</p>
              </div>
            </article>
          </main>
        </div>
        <footer class="composer">
          <div class="composer-box">
            <textarea aria-label="输入消息" data-draft>${draft}</textarea>
            <div class="composer-toolbar">
              <span class="composer-shortcut-hint">Enter 发送 · Shift + Enter 换行 · 可粘贴截图</span>
            </div>
          </div>
        </footer>`;

/** `scope-popover` 在真机上贴着输入框往上弹；夹具页里改成静态，才截得进画面。 */
const popover = (inner, extra = "") => `
        <section class="scope-popover wps-scope-popover" aria-label="选择数据范围">
          <div class="scope-popover-scroll">${inner}</div>${extra}
        </section>`;

/* 对照图里的发送拒绝文案 —— 逐字抄自 `packages/core/src/folderScope.ts`
   的 `workDirectorySendRefusal`（目录名「数据」、已选 3 张表）。
   真实的注入路径由 `wpsWorkDirectorySendGate.test.tsx` 挂真实面板钉住。 */
const REFUSAL =
  "这次对话的数据范围是工作目录「数据」（已选 3 张表）。工作目录的读取通道本轮还没接到这个面板 —— " +
  "这条消息没有发出，也没有读取任何文件；你当前打开的工作簿没有被读取或修改。" +
  "请先用上方的范围选择器切回「跟随当前」或「选择多个」，再发送。";

const CELLS = [
  {
    key: "before-card",
    title: "旧 · 「工作目录」是不可点的死按钮",
    tag: "before",
    sub:
      "用户截图里的样子：卡片灰掉，副标题写着那句能力说明「WPS 当前仅支持打开的工作簿」。" +
      "整条链路不存在 —— 既没有勾选，也没有授权。",
    stage: popover(modeList({
      folderDisabled: true,
      folderCaption: "WPS 当前仅支持打开的工作簿",
      folderSelected: false
    })),
    expect: { folderDisabled: true, folderCaption: "WPS 当前仅支持打开的工作簿" }
  },
  {
    key: "after-card",
    title: "现在 · 卡片放开，副标题对齐 Excel",
    tag: "after",
    sub:
      "三种来源都可以来回切。副标题与 Excel 逐字一致（`批量处理多个本地工作簿`），" +
      "那句能力说明**一个字都不在** —— 不是换个说法留着。",
    stage: popover(modeList({
      folderDisabled: false,
      folderCaption: "批量处理多个本地工作簿",
      folderSelected: true
    })),
    expect: { folderDisabled: false, folderCaption: "批量处理多个本地工作簿" }
  },
  {
    key: "after-picker",
    title: "现在 · 文件/表两级勾选 · 承认汇总 · 确认才提交",
    tag: "after",
    sub:
      "一级是文件、二级是表；`销售.xlsx` 里那张 `#EB_RULES` 能被扫描到但**不可勾选**" +
      "（可勾选名单与存在性名单是两个判据）。" +
      "勾选**本身不是授权** —— 页脚只是「已选几张」，要按「确认」才换来这次对话的范围。",
    stage: popover(folderPicker(), folderFooter()),
    expect: {
      filter: "2 个文件 · 2 张表",
      footer: "已选 1 个文件 · 1 张表",
      confirm: "确认",
      fileOne: "1/1",
      sheets: "明细",
      admission: "已选 1 张表 · 可分析"
    }
  },
  {
    key: "after-trigger",
    title: "现在 · 触发器只说工作目录（切回工作簿后照旧）",
    tag: "after",
    sub:
      "这一行是「这次用哪些数据」。工作目录模式下**不许**出现当前工作簿名 —— " +
      "写上它就等于在说「这次用它」，而工作目录模式恰恰不用它。切回工作簿来源后照旧显示工作簿与当前表。",
    stage: `${trigger({ name: null, label: "工作目录 数据 · 已选 1 张表" })}
        <div class="gap"></div>
        ${trigger({ name: "测试.xlsx", label: "当前表 订单" })}`,
    expect: {
      folderTriggerSmall: null,
      folderTrigger: "工作目录 数据 · 已选 1 张表",
      workbookTriggerSmall: "测试.xlsx",
      workbookTrigger: "当前表 订单"
    }
  },
  {
    key: "after-send",
    title: "现在 · 工作目录范围下发送被拦：零出站、零工作簿读取",
    tag: "after",
    sub:
      "工作目录的读取通道本轮还没接到这个面板，所以发送侧**明确拒绝**，而不是拿当前打开的" +
      "工作簿顶上。提示里说清是哪个目录、几张表，以及「什么都没读」；" +
      "输入框里的字**原样保留**，用户切个范围就能直接发。",
    stage: composerStage(REFUSAL, "把目录里的销售表按月份汇总"),
    expect: { refusalHas: ["工作目录「数据」（已选 3 张表）", "没有读取任何文件", "当前打开的工作簿没有被读取或修改"], draft: "把目录里的销售表按月份汇总" }
  }
];

const page = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<title>WPS 数据范围 · 工作目录（旧 / 现在）</title>
<style>${excelCss}</style>
<style>${wpsCss}</style>
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
  /* 舞台宽度固定：换宽度会改变折行，读回来的数就与判据不同源了。 */
  .stage { box-sizing: border-box; background: #fbfcfb; border: 1px dashed #dfe6e1;
    border-radius: 8px; padding: 10px; }
  /* 真机上弹层是浮层；夹具页里改成静态，才截得进画面。 */
  .stage .scope-popover { position: static; max-height: none; box-shadow: none; }
  .stage .gap { height: 10px; }
  .readout { margin: 10px 0 0; font-size: 12px; line-height: 1.9; }
  .readout b { color: #2f6b47; }
  .readout .warn { color: #a5413a; }
</style></head><body>
<h1>WPS 数据范围面板：第三种来源「工作目录」</h1>
<p class="note">样式表是应用自己的那两份（<code>apps/wps-addin/src/styles.css</code> 按 <code>main.tsx</code>
的顺序先加载 <code>apps/excel-addin/src/styles.css</code>），结构从 <code>WpsProbePanel.tsx</code> 逐字抄来。<br>
<b>这是夹具页，不是实机 WPS 截图。</b>判据在
<code>apps/wps-addin/src/wpsWorkDirectoryScope.test.tsx</code>、
<code>wpsWorkDirectorySendGate.test.tsx</code> 与 <code>packages/core/src/folderScope.test.ts</code>；
反向变异见 <code>scripts/run_wps_work_directory_mutations.mjs</code>。</p>
<div class="grid">
${CELLS.map((cell) => `  <div class="panel" data-cell="${cell.key}">
    <h2>${cell.title}<span class="tag ${cell.tag}">${cell.tag === "before" ? "旧" : "现在"}</span></h2>
    <p class="sub">${cell.sub}</p>
    <div class="stage">${cell.stage}</div>
    <p class="readout" data-readout="${cell.key}"></p>
  </div>`).join("\n")}
</div>
</body></html>`;

await mkdir(outDir, { recursive: true });
const htmlPath = path.join(outDir, "wps-work-directory-demo.html");
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
  const view = await browser.newPage({ viewport: { width: 1180, height: 1000 } });
  await view.goto(url);
  await view.locator("[data-cell]").first().waitFor({ state: "attached" });

  /* 把所有关键事实**从渲染结果里读回来**（读的是真样式表算过的 DOM，不是源码）。 */
  const readBack = await view.evaluate(() => {
    const text = (node) => (node ? node.textContent : null);
    const cardOf = (panel) => {
      const buttons = [...panel.querySelectorAll(".scope-mode-list button")];
      const folder = buttons.find((button) => button.textContent.includes("工作目录"));
      return {
        disabledCount: buttons.filter((button) => button.disabled).length,
        folderDisabled: folder.disabled,
        folderCaption: text(folder.querySelector("small")),
        folderSelected: folder.classList.contains("selected")
      };
    };
    const triggerOf = (element) => ({
      small: text(element.querySelector("small")),
      strong: text(element.querySelector("strong"))
    });
    return Object.fromEntries([...document.querySelectorAll("[data-cell]")].map((panel) => {
      const key = panel.dataset.cell;
      let value;
      if (key === "before-card" || key === "after-card") {
        value = cardOf(panel);
      } else if (key === "after-picker") {
        const scroll = panel.querySelector(".scope-popover-scroll");
        const footer = panel.querySelector(".folder-selection-footer");
        value = {
          filter: text(panel.querySelector(".folder-selection-filter span")),
          footer: text(panel.querySelector(".folder-selection-footer strong")),
          confirm: text(panel.querySelector(".folder-selection-footer button.scope-confirm")),
          // 页脚与承认汇总的「层」和真机一致：住在弹层上（`.scope-popover-scroll` 的兄弟），
          // 不在滚动容器里 —— 否则内容一滚，这两段就跟着滚走了。
          footerOutsideScroll: scroll.querySelector(".folder-selection-footer") === null,
          footerOnPopover: footer.parentElement.classList.contains("scope-popover"),
          admissionOutsideScroll: scroll.querySelector(".scope-admission-summary") === null,
          // 「取消」只在草稿与已提交授权不一致时出现。这一格画的是**刚确认过**的状态
          // （页脚计数就是已提交的那份），所以这里不该有取消 —— 断言钉的是"取消没被无条件画上"。
          // 真正"有草稿差异时出现取消"的判据在
          // `wpsWorkDirectoryScope.test.tsx` 与 Excel 的 `folderScopeSafety.test.tsx` 里。
          discardInFooter: panel.querySelector(".folder-selection-footer button.scope-discard") !== null,
          fileOne: text(
            [...panel.querySelectorAll(".folder-file-option")]
              .find((button) => button.textContent.includes("销售.xlsx"))
              .querySelector("b")
          ),
          sheets: [...panel.querySelectorAll(".folder-sheet-list button strong")]
            .map((node) => node.textContent).join("/"),
          admission: text(panel.querySelector(".scope-admission-summary")),
          // 只查**勾选区**：这一格的说明文字里本来就提到 `#EB_RULES`（在讲"可读不可勾"），
          // 拿整格 textContent 查会把说明文字也算进去，判据就变成恒真了。
          hasEbRule: panel.querySelector(".folder-master-detail").textContent.includes("#EB_RULES")
        };
      } else if (key === "after-trigger") {
        const [folderTrigger, workbookTrigger] = [...panel.querySelectorAll(".scope-trigger")];
        value = {
          folderTriggerSmall: triggerOf(folderTrigger).small,
          folderTrigger: triggerOf(folderTrigger).strong,
          workbookTriggerSmall: triggerOf(workbookTrigger).small,
          workbookTrigger: triggerOf(workbookTrigger).strong
        };
      } else {
        value = {
          refusal: text(panel.querySelector("[data-refusal]")),
          draft: panel.querySelector("textarea[data-draft]").value
        };
      }
      return [key, value];
    }));
  });

  for (const cell of CELLS) {
    const actual = readBack[cell.key];
    if (cell.key === "before-card" || cell.key === "after-card") {
      assert.equal(actual.folderDisabled, cell.expect.folderDisabled,
        `${cell.key} 的「工作目录」卡禁用状态不符：${actual.folderDisabled}`);
      assert.equal(actual.folderCaption, cell.expect.folderCaption,
        `${cell.key} 的卡片副标题不符：${actual.folderCaption}`);
      console.log(`  ✓ ${cell.title}：副标题「${actual.folderCaption}」，disabled=${actual.folderDisabled}`);
    } else if (cell.key === "after-picker") {
      assert.equal(actual.filter, cell.expect.filter, `筛选行不符：${actual.filter}`);
      assert.equal(actual.footer, cell.expect.footer, `页脚不符：${actual.footer}`);
      assert.equal(actual.confirm, cell.expect.confirm, `页脚里的确认按钮不符：${actual.confirm}`);
      assert.ok(actual.footerOutsideScroll, "页脚被卷进了滚动容器（应与真机一样是 .scope-popover-scroll 的兄弟）");
      assert.ok(actual.footerOnPopover, "页脚不直接住在弹层上");
      assert.ok(actual.admissionOutsideScroll, "承认汇总被卷进了滚动容器（它决定能不能提交，不能被滚走）");
      assert.equal(actual.discardInFooter, false, "夹具画的是刚确认过的状态，页脚里不该出现「取消」（它只在草稿与已提交不一致时出现）");
      assert.equal(actual.fileOne, cell.expect.fileOne, `销售.xlsx 的计数不符：${actual.fileOne}`);
      assert.equal(actual.sheets, cell.expect.sheets, `二级列表不符：${actual.sheets}`);
      assert.ok(actual.admission.includes(cell.expect.admission),
        `承认汇总不符：${actual.admission}`);
      console.log(`  ✓ ${cell.title}：${actual.filter}｜${actual.footer}｜页脚确认「${actual.confirm}」在弹层上｜二级「${actual.sheets}」｜${actual.admission}`);
    } else if (cell.key === "after-trigger") {
      assert.equal(actual.folderTriggerSmall, cell.expect.folderTriggerSmall,
        `工作目录模式下的触发器不该有 <small>：${actual.folderTriggerSmall}`);
      assert.equal(actual.folderTrigger, cell.expect.folderTrigger, `触发器文案不符：${actual.folderTrigger}`);
      assert.equal(actual.workbookTriggerSmall, cell.expect.workbookTriggerSmall,
        `工作簿模式的 <small> 不符：${actual.workbookTriggerSmall}`);
      assert.equal(actual.workbookTrigger, cell.expect.workbookTrigger,
        `工作簿模式的触发器文案不符：${actual.workbookTrigger}`);
      console.log(
        `  ✓ ${cell.title}：工作目录「${actual.folderTrigger}」（无文件名词）·` +
          `工作簿「${actual.workbookTriggerSmall}|${actual.workbookTrigger}」`
      );
    } else {
      for (const needle of cell.expect.refusalHas) {
        assert.ok(actual.refusal.includes(needle), `拒绝文案缺少「${needle}」：${actual.refusal}`);
      }
      assert.equal(actual.draft, cell.expect.draft, `输入框里的草稿没被原样保留：${actual.draft}`);
      console.log(
        `  ✓ ${cell.title}：提示含「${cell.expect.refusalHas.join("」「")}」，输入框仍是「${actual.draft}」`
      );
    }
  }

  /* 反向对照在场：旧那一格必须真的画出老样子，否则这张对照图没有判别力。 */
  const before = readBack["before-card"];
  assert.equal(before.folderDisabled, true, "「旧」那一格必须画出禁用的死按钮");
  assert.equal(before.folderCaption, "WPS 当前仅支持打开的工作簿",
    "「旧」那一格必须画出那句能力说明");
  assert.equal(before.folderSelected, false, "「旧」那一格的工作目录卡不该是选中态");
  const afterCard = readBack["after-card"];
  assert.equal(afterCard.disabledCount, 0, "「现在」那一格三张卡都不该禁用");
  assert.equal(afterCard.folderCaption.includes("WPS 当前仅支持"), false,
    "「现在」那一格不许再出现那句能力说明");
  console.log("  ✓ 反向对照：旧那格是禁用的死按钮 + 旧能力说明；现在那格三张卡都可点、文案已换");

  /* 可勾选名单与存在性名单是两个判据：夹具里那张 EB 遗留表**可读不可勾**。 */
  assert.equal(readBack["after-picker"].hasEbRule, false,
    "EB 系统遗留表 `#EB_RULES` 不该出现在可勾选列表里");
  console.log("  ✓ EB 系统遗留表不在可勾选列表里（能被扫描到，但不可勾）");

  const height = await view.evaluate(() => Math.ceil(
    Math.max(document.body.scrollHeight,
      Math.max(...[...document.querySelectorAll(".panel")].map((panel) => panel.getBoundingClientRect().bottom))
    ) + 24
  ));
  await view.setViewportSize({ width: 1180, height });
  const shotPath = path.join(outDir, "wps-work-directory-demo.png");
  await view.screenshot({ path: shotPath });
  console.log(`  ✓ 对照图：${shotPath}（1180×${height}）`);
  console.log(`  ✓ 对照页：${htmlPath}`);
} finally {
  await browser?.close();
  server.close();
}
