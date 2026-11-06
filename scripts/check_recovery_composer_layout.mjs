// 布局证据：恢复操作区（RecoveryComposer）在真实 Chromium 布局里到底长什么样。
//
// 背景：jsdom 能证明"接管时卡为 null、操作区存在、按钮文字正确"，但证明不了
// "看得见、按钮没被裁、长正文自己滚、320px 窄窗格下输入区还在视口里"。
// 2026-11-02 把核对/继续搬进输入框内部之后，最容易在窄窗格里出这种只靠 jsdom
// 看不出来的毛病（正文把动作行挤出屏幕 / 按钮被压成不可点 / 输入区被顶出去）。
//
// 所以这里用**真实样式表 + 真实 Chromium 布局引擎**把四件事钉死：
// 1. 待核对（reconcile）与可继续（resume）两个状态，动作行永远是正文的**兄弟**，
//    正文自己滚（`.recovery-composer-scroll`），主操作始终完整落在视口内；
// 2. 320 / 400 / 500 三档宽度下，主按钮都有非零尺寸、有可见文字、未被横向裁掉；
// 3. 正文足够长时（塞一段长位置清单），正文溢出、自己滚，动作行仍在原地；
// 4. 「另提需求」之后那枚紧凑待办入口（RecoveryPendingEntry）也在视口内、不压输入框。
//
// 实现走**原始 CDP**（`--remote-debugging-port` + WebSocket）：本机没有 Playwright，
// 而 Chrome 是 GUI 子系统程序，`--dump-dom` 的 stdout 在 Windows 上拿不到。
// CDP 逐条 `Runtime.evaluate` 取数，视口用 `Emulation.setDeviceMetricsOverride` 固定。
//
// 用法：
//   node scripts/check_recovery_composer_layout.mjs           # 320/400/500，输出截图
//   PROBE_BROWSER=<chromium 路径> node scripts/...             # 换浏览器
//   --keep 打印夹具页路径（outputs/probe/），供人工复核
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import { readFile } from "node:fs/promises";
import net from "node:net";
import path from "node:path";
import { build } from "esbuild";

const root = path.resolve(import.meta.dirname, "..");
const probeDir = path.join(root, "outputs", "probe");
const keep = process.argv.includes("--keep");

function findBrowser() {
  const candidates = [
    process.env.PROBE_BROWSER,
    "C:/Program Files/Google/Chrome/Application/chrome.exe",
    "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
    "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
    "C:/Program Files/Microsoft/Edge/Application/msedge.exe"
  ].filter(Boolean);
  const found = candidates.find((candidate) => existsSync(candidate));
  assert.ok(found, "找不到 Chromium 内核浏览器：用 PROBE_BROWSER=<可执行文件> 指定");
  return found;
}

function freePort() {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.on("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const { port } = server.address();
      server.close(() => resolve(port));
    });
  });
}

async function connectCdp(port) {
  const deadline = Date.now() + 25_000;
  let targets = [];
  while (Date.now() < deadline) {
    try {
      targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
      if (targets.some((target) => target.type === "page")) break;
    } catch {
      // 浏览器还没起来：继续等
    }
    await new Promise((resolve) => setTimeout(resolve, 200));
  }
  const page = targets.find((target) => target.type === "page");
  assert.ok(page?.webSocketDebuggerUrl, "浏览器没有可用的页面目标");
  const socket = new WebSocket(page.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => {
    socket.addEventListener("open", resolve, { once: true });
    socket.addEventListener("error", () => reject(new Error("CDP 连接失败")), { once: true });
  });
  let nextId = 0;
  const waiting = new Map();
  socket.addEventListener("message", (event) => {
    const message = JSON.parse(event.data);
    const settle = waiting.get(message.id);
    if (settle) {
      waiting.delete(message.id);
      settle(message);
    }
  });
  const send = (method, params = {}) =>
    new Promise((resolve) => {
      const id = (nextId += 1);
      waiting.set(id, resolve);
      socket.send(JSON.stringify({ id, method, params }));
    });
  const evaluate = async (expression) => {
    const reply = await send("Runtime.evaluate", {
      expression,
      returnByValue: true,
      awaitPromise: true
    });
    assert.equal(reply.result?.exceptionDetails, undefined,
      `页面脚本抛错：${reply.result?.exceptionDetails?.text ?? "unknown"}`);
    return reply.result?.result?.value;
  };
  return { send, evaluate, close: () => socket.close() };
}

/**
 * 夹具：把两个状态（待核对 / 可继续）并排各渲染一份 `RecoveryComposer`，
 * 用共享层 `recoveryComposerView(facts)` 派生文案（不是手写），加上真实样式表。
 *
 * 事实按真机 B 的形状（三条写公式动作、前两步已写入、第 3 步待续跑）：
 *  - reconcile：账本确认 1 步、第 2 步在途（pending）、剩 1 步；
 *  - resume：核对后账本 2 步、剩 1 步、下一条写 `结果3!B2`。
 *
 * 另加一份"超长位置清单"的 resume（5 个目标，正文折"等 5 处"），验证正文自滚。
 */
const source = `
import React from 'react';
import { createRoot } from 'react-dom/client';
import { RecoveryComposer } from './components/RecoveryComposer';
import { RecoveryPendingEntry } from './components/RecoveryComposer';
import { ResultReviewBar } from './components/ResultReviewBar';
import { recoveryComposerView, recoveryComposerEntry } from '@excel-bro/core/recoveryComposer';

const ACTIONS = [
  { sheet: '结果1', range: 'B2', type: 'writeFormulas' },
  { sheet: '结果2', range: 'B2', type: 'writeFormulas' },
  { sheet: '结果3', range: 'B2', type: 'writeFormulas' }
];

const reconcileFacts = {
  verdict: { kind: 'reconcile', actionIndex: 1, outcome: 'pending', remaining: 1 },
  nextActionIndex: 1,
  totalActions: 3,
  nextTarget: { sheet: '结果3', range: 'B2' },
  actions: ACTIONS
};
const resumeFacts = {
  verdict: { kind: 'resume', remaining: 1 },
  nextActionIndex: 2,
  totalActions: 3,
  nextTarget: { sheet: '结果3', range: 'B2' },
  actions: ACTIONS
};
// 超长正文：reconcile + detailsOpen，detail 塞一段很长的技术口径，
// 让滚动容器在窄窗格里真正溢出 —— 验证"正文自己滚、动作行是兄弟、按钮留在原地"。
const longFacts = {
  verdict: { kind: 'reconcile', actionIndex: 1, outcome: 'pending', remaining: 1 },
  nextActionIndex: 1,
  totalActions: 3,
  nextTarget: { sheet: '结果3', range: 'B2' },
  actions: [
    { sheet: '结果1', range: 'B2', type: 'writeFormulas' },
    { sheet: '结果2', range: 'B2', type: 'writeFormulas' },
    { sheet: '结果3', range: 'B2', type: 'writeFormulas' }
  ]
};

function Shell({ id, children, tall }) {
  return (
    <div className={"chat-shell" + (new URLSearchParams(location.search).get("host") === "wps" ? " wps-chat-shell" : "")} style={{ display: 'flex', flexDirection: 'column', height: '100vh' }}>
      <section className="message-stream" style={{ flex: '1 1 auto', overflowY: 'auto', padding: '12px' }}>
        <div className="message-row system">
          <div className="message-content"><p className="message-text">{id}</p></div>
        </div>
      </section>
      <footer className="composer">
        <div className={'composer-box ' + (id.startsWith('review') ? 'is-reviewing' : 'is-recovery') + (tall ? ' is-tall' : '')} data-probe={id}>
          {children}
        </div>
      </footer>
    </div>
  );
}

function Harness() {
  const mode = new URLSearchParams(location.search).get("pass") ?? "reconcile";
  if (mode.startsWith('review')) return <Shell id={mode}>
    <ResultReviewBar inComposer paneNeedsResize title="三处独立公式写入" state={{
      phase: 'reviewing', key: 'layout', executionId: 'layout', pending: false, comparing: true,
      comparisonHelp: {scope:'执行前静态快照；公式冻结，复杂对象不作为历史依据。',layout:'可手动排列窗口。'},
      comparisonSheets: Array.from({length:8}, (_,i) => '华东区域销售数据季度汇总表' + (i + 1)),
      notice: mode === 'review-long' ? '核对失败，请返回原工作簿后重试。'.repeat(40) : null
    }} onComparisonSheet={() => {}} onCompare={() => {}} onExitComparison={() => {}} onConfirm={() => {}} onDiscard={() => {}} />
  </Shell>;
  const views = {
    reconcile: recoveryComposerView(reconcileFacts),
    resume: recoveryComposerView(resumeFacts),
    long: {
      ...recoveryComposerView(longFacts),
      detail:
        "本次执行共 3 步，账本游标停在 2（前 1 步已确认写入）。" +
        "第 2 步在途：写回回执尚未核对，现场与写后结果是否一致待确认。" +
        "（这一段是为了把滚动容器撑到溢出而重复的技术口径。）".repeat(12)
    }
  };
  if (mode === "entry") {
    return (
      <Shell id="entry">
        <RecoveryPendingEntry entry={recoveryComposerEntry(resumeFacts)} busy={false} onReturn={() => {}} />
      </Shell>
    );
  }
  const view = views[mode] ?? views.reconcile;
  return (
    <Shell id={mode}>
      <RecoveryComposer
        view={view} busy={false} acting={null} detailsOpen={mode === "long"}
        dismiss={null}
        undo={mode === "resume" ? { state: 'available', detail: '撤销已写入的两步。' } : null}
        onPrimary={() => {}} onNewRequest={() => {}} onToggleDetails={() => {}}
      />
    </Shell>
  );
}

createRoot(document.getElementById('root')).render(<Harness />);
`;

mkdirSync(probeDir, { recursive: true });
const bundled = await build({
  stdin: { contents: source, resolveDir: path.join(root, "apps/excel-addin/src"), loader: "tsx" },
  bundle: true,
  write: false,
  format: "iife",
  platform: "browser",
  jsx: "automatic",
  alias: { "@excel-bro/core": path.join(root, "packages/core/src") },
  loader: { ".png": "dataurl", ".svg": "dataurl", ".css": "empty" },
  define: { "import.meta.env": "{}", "process.env.NODE_ENV": '"production"' }
});

/*
 * 样式按真实导入顺序拼：`main.tsx` 先 `import "./styles.css"`，组件再各自 import
 * 自己的 css（后到者覆盖）。直接用源文件，不生成任何副本 —— 量到的就是线上那份。
 */
const cssFiles = [
  "apps/excel-addin/src/styles.css",
  "apps/excel-addin/src/components/resultReviewComposer.css",
  "apps/excel-addin/src/components/recoveryComposer.css",
  ...(process.env.PROBE_HOST === "wps" ? ["apps/wps-addin/src/styles.css"] : []),
  "apps/excel-addin/src/appearance.css",
  "apps/excel-addin/src/conversation.css"
];
const css = (
  await Promise.all(cssFiles.map((relative) => readFile(path.join(root, relative), "utf8")))
).join("\n");

const pageHtml = `<!doctype html><meta charset="utf-8">
<title>recovery composer layout probe</title>
<style>html,body,#root{height:100%;margin:0;} body{background:var(--color-surface,#fafaf9);}
*{scroll-behavior:auto !important;}</style>
<style>${css}</style>
<div id="root"></div>
<script>${bundled.outputFiles[0].text}</script>
`;
const probePath = path.join(probeDir, "recovery_composer_layout_probe.html");
writeFileSync(probePath, pageHtml, "utf8");
const probeUrl = `file:///${probePath.replaceAll("\\", "/")}`;
const probeUrlFor = (pass) => `${probeUrl}?pass=${pass}&host=${process.env.PROBE_HOST ?? "excel"}`;

const browserPath = findBrowser();
const port = await freePort();
const child = spawn(
  browserPath,
  [
    "--headless=new",
    "--disable-gpu",
    "--no-first-run",
    "--no-default-browser-check",
    "--disable-extensions",
    "--allow-file-access-from-files",
    `--remote-debugging-port=${port}`,
    `--user-data-dir=${path.join(probeDir, "profile-recovery-composer")}`,
    probeUrl
  ],
  { stdio: "ignore" }
);

const RECT = `(node) => {
  if (!node) return null;
  const box = node.getBoundingClientRect();
  return { top: box.top, bottom: box.bottom, left: box.left, right: box.right,
    width: box.width, height: box.height };
}`;

const insideViewport = (box, viewport, label) => {
  assert.ok(box, `${label}: 没有布局盒子`);
  assert.ok(box.width > 0 && box.height > 0, `${label}: 尺寸为零`);
  assert.ok(box.top >= -0.5, `${label}: 上沿在视口外 (top=${box.top})`);
  assert.ok(box.bottom <= viewport.height + 0.5,
    `${label}: 下沿在视口外 (bottom=${box.bottom}, viewport=${viewport.height})`);
  assert.ok(box.left >= -0.5 && box.right <= viewport.width + 0.5,
    `${label}: 横向被裁 (left=${box.left}, right=${box.right}, viewport=${viewport.width})`);
};

const results = [];
let client;
try {
  client = await connectCdp(port);

  for (const [width, height] of [[320, 640], [400, 720], [500, 760]]) {
    await client.send("Emulation.setDeviceMetricsOverride", {
      width, height, deviceScaleFactor: 1, mobile: false
    });
    await client.send("Page.enable");
    const label = `${width}x${height}`;

    for (const mode of ['review', 'review-long']) {
      await client.send('Page.navigate', { url: probeUrlFor(mode) });
      const deadline = Date.now() + 25000;
      while (Date.now() < deadline) {
        if (await client.evaluate("Boolean(document.querySelector('.result-review-composer'))")) break;
        await new Promise(resolve => setTimeout(resolve, 120));
      }
      const measured = await client.evaluate(`(() => {
        const rect = ${RECT};
        const box = document.querySelector('.result-review-composer');
        const scroll = box.querySelector('.result-review-composer-scroll');
        const actions = box.querySelector('.result-review-actions');
        return { viewport: {width:innerWidth,height:innerHeight}, actions:rect(actions),
          buttons:[...actions.querySelectorAll('button')].map(rect),
          sibling:actions.parentElement === scroll.parentElement,
          overflow:scroll.scrollHeight > scroll.clientHeight || [...box.querySelectorAll('.result-review-feedback')].some(node => node.scrollHeight > node.clientHeight),
          detailsClosed:!box.querySelector('details').open };
      })()`);
      assert.equal(measured.sibling, true);
      assert.equal(measured.detailsClosed, true);
      assert.equal(measured.buttons.length, 3);
      insideViewport(measured.actions, measured.viewport, `${label}/${mode} actions`);
      measured.buttons.forEach(button => insideViewport(button, measured.viewport, `${label}/${mode} button`));
      if (mode === 'review-long') assert.equal(measured.overflow, true);
      results.push(`${label}/${mode}: 审阅按钮完整可见，详情折叠，错误正文独立滚动`);
      if (mode === 'review') {
        const searchLayout = await client.evaluate(`(() => {
          const input=document.querySelector('.review-sheet-search input');
          const setter=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set;
          setter.call(input,'华东'); input.dispatchEvent(new Event('input',{bubbles:true}));
          return true;
        })()`);
        assert.equal(searchLayout, true);
        await new Promise(resolve => setTimeout(resolve, 100));
        const clearLayout = await client.evaluate(`(() => {
          const i=document.querySelector('.review-sheet-search input').getBoundingClientRect();
          const b=document.querySelector('.review-sheet-search-clear').getBoundingClientRect();
          return {inside:b.left>=i.left && b.right<=i.right && b.top>=i.top && b.bottom<=i.bottom};
        })()`);
        assert.equal(clearLayout.inside, true, 'clear button must stay inside search field');
        await client.evaluate(`document.querySelector('.review-sheet-search-clear').click()`);
        await new Promise(resolve => setTimeout(resolve, 100));
        for (let page = 0; page < 3; page++) {
          const paging = await client.evaluate(`(() => {
            const picker = document.querySelector('.review-sheet-picker');
            const buttons = [...picker.querySelectorAll('.review-sheet-list button')];
            return {count:buttons.length, pickers:document.querySelectorAll('.review-sheet-picker').length,
              overflow:picker.scrollWidth > picker.clientWidth,
              label:picker.querySelector('.review-sheet-page-count').textContent};
          })()`);
          assert.equal(paging.count, page === 2 ? 2 : 3);
          assert.equal(paging.pickers, 1);
          assert.equal(paging.overflow, false);
          assert.ok(paging.label.includes(String(page + 1) + ' / 3'));
          if (page < 2) {
            await client.evaluate(`document.querySelector('.review-sheet-page > button:last-child').click()`);
            await new Promise(resolve => setTimeout(resolve, 100));
          }
        }
        results.push(`${label}: 八张长表名分页 3/3/2，无重复控件或横向溢出`);
      }
      const shot = await client.send('Page.captureScreenshot', {format:'png'});
      writeFileSync(path.join(probeDir, `${mode}-${width}x${height}.png`), Buffer.from(shot.result.data, 'base64'));
    }

    // 每个状态各自量一遍：动作行在视口内、主按钮有文字、正文是兄弟。
    for (const state of ["reconcile", "resume"]) {
      await client.send("Page.navigate", { url: probeUrlFor(state) });
      const deadline = Date.now() + 25_000;
      let ready = false;
      while (Date.now() < deadline) {
        ready = await client.evaluate(
          `Boolean(document.querySelector('.recovery-composer')) && ` +
          `Boolean(document.querySelector('.recovery-composer [data-recovery-primary]'))`
        );
        if (ready) break;
        await new Promise((resolve) => setTimeout(resolve, 120));
      }
      assert.ok(ready, `${label}/${state}: 没有渲染出恢复操作区`);

      const measured = await client.evaluate(`(() => {
        const rect = ${RECT};
        const box = document.querySelector('.recovery-composer');
        const scroll = box && box.querySelector('.recovery-composer-scroll');
        const actions = box && box.querySelector('.recovery-composer-actions');
        const primary = box && box.querySelector('[data-recovery-primary]');
        const headline = box && box.querySelector('.recovery-composer-headline');
        if (!box || !scroll || !actions || !primary) return { error: 'missing-node' };
        return {
          state: '${state}',
          viewport: { width: innerWidth, height: innerHeight },
          action: box.getAttribute('data-recovery-composer'),
          headline: headline ? headline.textContent.trim() : '',
          primaryLabel: primary.textContent.trim(),
          primaryCount: box.querySelectorAll('[data-recovery-primary]').length,
          actions: rect(actions),
          primary: rect(primary),
          scroll: rect(scroll),
          scrollOverflow: scroll.scrollHeight > scroll.clientHeight,
          scrollInActions: Boolean(actions.querySelector('.recovery-composer-scroll')),
          composerBox: rect(box),
          primaryIsSibling: primary.parentElement === actions
        };
      })()`);
      assert.equal(measured.error, undefined, `${label}/${state}: ${measured.error}`);
      assert.equal(measured.action, state, `${label}/${state}: 状态标记不对`);
      assert.ok(measured.primaryLabel.length > 0, `${label}/${state}: 主按钮没有文字`);
      assert.equal(measured.primaryCount, 1, `${label}/${state}: 主操作只许一颗`);
      assert.equal(measured.primaryIsSibling, true,
        `${label}/${state}: 主按钮不在动作行里（结构漂移）`);
      assert.equal(measured.scrollInActions, false,
        `${label}/${state}: 正文被塞进了动作行（不是兄弟，会一起被裁）`);
      insideViewport(measured.primary, measured.viewport, `${label}/${state} 主操作`);
      insideViewport(measured.actions, measured.viewport, `${label}/${state} 动作行`);
      results.push(
        `${label}/${state}: 主操作「${measured.primaryLabel}」在视口内 ` +
        `(bottom=${measured.primary.bottom.toFixed(1)})、动作行 bottom=${measured.actions.bottom.toFixed(1)}`
      );

      // 截图：待核对 + 可继续各一张（每档宽度各一张）。
      const shot = await client.send("Page.captureScreenshot", { format: "png" });
      writeFileSync(
        path.join(probeDir, `recovery-composer-${state}-${width}x${height}.png`),
        Buffer.from(shot.result.data, "base64")
      );
    }

    // 超长清单：正文必须自己滚，动作行仍在视口内。
    await client.send("Page.navigate", { url: probeUrlFor("long") });
    {
      const deadline = Date.now() + 25_000;
      let ready = false;
      while (Date.now() < deadline) {
        ready = await client.evaluate(
          `Boolean(document.querySelector('.recovery-composer [data-recovery-primary]'))`
        );
        if (ready) break;
        await new Promise((resolve) => setTimeout(resolve, 120));
      }
      assert.ok(ready, `${label}/long: 没有渲染出恢复操作区`);
    }
    const long = await client.evaluate(`(() => {
      const rect = ${RECT};
      const box = document.querySelector('.recovery-composer');
      const scroll = box && box.querySelector('.recovery-composer-scroll');
      const actions = box && box.querySelector('.recovery-composer-actions');
      const primary = box && box.querySelector('[data-recovery-primary]');
      if (!box || !scroll || !actions || !primary) return { error: 'missing-node' };
      return {
        viewport: { width: innerWidth, height: innerHeight },
        primary: rect(primary),
        actions: rect(actions),
        scrollOverflow: scroll.scrollHeight > scroll.clientHeight,
        scrollHeight: scroll.scrollHeight,
        clientHeight: scroll.clientHeight,
        composerHeight: box.getBoundingClientRect().height
      };
    })()`);
    assert.equal(long.error, undefined, `${label}/long: ${long.error}`);
    assert.ok(long.scrollOverflow, `${label}/long: 长清单没有溢出，证明不了"内部滚动"`);
    insideViewport(long.primary, long.viewport, `${label}/long 主操作`);
    insideViewport(long.actions, long.viewport, `${label}/long 动作行`);
    results.push(
      `${label}/long: 正文 ${long.scrollHeight}px 自己滚（框高 ${long.composerHeight}px）、主操作仍可见`
    );

    // 「另提需求」后的待办入口：在视口内、不压输入框（这里输入框就是 composer-box）。
    await client.send("Page.navigate", { url: probeUrlFor("entry") });
    {
      const deadline = Date.now() + 25_000;
      let ready = false;
      while (Date.now() < deadline) {
        ready = await client.evaluate(
          `Boolean(document.querySelector('.recovery-pending-entry-button'))`
        );
        if (ready) break;
        await new Promise((resolve) => setTimeout(resolve, 120));
      }
      assert.ok(ready, `${label}/entry: 没有渲染出待办入口`);
    }
    const entry = await client.evaluate(`(() => {
      const rect = ${RECT};
      const box = document.querySelector('[data-probe="entry"]');
      const btn = box && box.querySelector('.recovery-pending-entry-button');
      if (!box || !btn) return { error: 'missing-node' };
      return {
        viewport: { width: innerWidth, height: innerHeight },
        entry: rect(btn),
        entryText: btn.textContent.trim(),
        box: rect(box)
      };
    })()`);
    assert.equal(entry.error, undefined, `${label}/entry: ${entry.error}`);
    assert.ok(entry.entryText.includes("恢复待办"), `${label}/entry: 入口文案应带「恢复待办」`);
    insideViewport(entry.entry, entry.viewport, `${label}/entry 待办入口`);
    results.push(`${label}/entry: 待办入口「${entry.entryText}」在视口内 (top=${entry.entry.top.toFixed(1)})`);
  }
} finally {
  client?.close();
  child.kill();
}

console.log(results.join("\n"));
console.log(`夹具页：${probePath}${keep ? "" : "（加 --keep 保留供人工复核）"}`);
console.log(`截图目录：${probeDir}`);
