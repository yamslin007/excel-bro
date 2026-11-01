// 布局证据：**限高执行过程框**在真实布局里到底限没限、输入框还在不在视口里。
//
// 为什么必须有这个脚本：`TaskProgressTimeline.test.tsx` 能断言"限高变量写进了
// 样式属性"、"内容超出时有收起入口"，但 jsdom 没有布局 —— 它证明不了
// `max-height` 真的把 160px 生效了，也证明不了输入框没有被过程框挤出去。
// 所以这里用**真实样式表 + 真实 Chromium 布局引擎**量：
//
// 1. 进行中：`.task-progress-content-body` 的实际高度 ≤ 160（不是"设置了 160"）；
// 2. 内容比限高长时它自己滚（scrollHeight > clientHeight）；
// 3. 320x640 矮窗格下 footer（输入区）完整落在视口内；
// 4. 底部「收起执行过程」按钮可见、没有被裁掉；
// 5. 点「多显示几行」变高（否则那颗按钮是假的），footer 仍在视口内；
// 6. 正常结束：卡片收起成一行，**结果摘要仍在**（它在过程框之外）。
//
// 用法：
//   node scripts/check_process_panel_layout.mjs
//   PROBE_BROWSER=<chromium 路径> node scripts/check_process_panel_layout.mjs
//   --keep 保留夹具页供人工复核
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
 * 夹具：一段**长**过程（14 步，每步带说明与目标）+ 输入区。
 *
 * 刻意做成"真实形态"而不是最小结构：步骤少的时候 160px 根本不会触发，
 * 那样量出来的"没超"是假的。
 */
const source = `
import React, { useState } from 'react';
import { createRoot } from 'react-dom/client';
import { TaskProgressTimeline } from './components/TaskProgressTimeline';
import { ProcessResultSummary } from './components/ProcessResultSummary';

const steps = Array.from({ length: 14 }, (_, index) => ({
  id: 'step-' + index,
  label: '第 ' + (index + 1) + ' 步：读取并汇总已授权的工作表区域',
  elapsedMs: 420,
  phase: 'executing',
  status: index === 13 ? 'progress' : 'completed',
  note: '这一段说明会比较长，用来把过程框撑到超过限高，从而验证它自己滚而不是把输入框挤出去。',
  target: { sheet: '汇总', range: 'A' + (index + 1) + ':D' + (index + 1) },
  startdAt: undefined
}));

function Harness() {
  const [live, setLive] = useState(true);
  window.__setLive = (value) => setLive(value);
  const log = live
    ? { operationId: 'op-probe', totalMs: 12_000, steps }
    : { operationId: 'op-probe', totalMs: 12_000, outcome: 'completed', steps: steps.map((s) => ({ ...s, status: 'completed' })) };
  return <div className="chat-shell" style={{ display: 'flex', flexDirection: 'column', height: '100vh' }}>
    <section className="message-stream" style={{ flex: '1 1 auto', overflowY: 'auto', padding: '12px' }}>
      <div className="message-row">
        <div className="message-stack" style={{ minWidth: 0 }}>
          <p className="message-text">把汇总表按分类重算一遍。</p>
          <TaskProgressTimeline live={live} log={log} />
          <ProcessResultSummary input={{ live, phase: live ? 'executing' : 'reviewing',
            verification: { status: 'verified', total: 9, passed: 9, unverified: 0 } }} />
        </div>
      </div>
    </section>
    <footer className="composer">
      <div className="composer-box"><textarea aria-label="给 Excel Bro 发消息" rows={1} defaultValue="" /></div>
    </footer>
  </div>;
}

createRoot(document.getElementById('root')).render(<Harness />);
`;

mkdirSync(probeDir, { recursive: true });
const bundled = await build({
  stdin: { contents: source, resolveDir: path.join(root, "apps/excel-addin/src"), loader: "tsx" },
  bundle: true,
  write: false,
  format: "iife",
  // CSS 由夹具页的 <style> 注入真实样式表；JS 里的 import 只做空加载。
  loader: { ".css": "empty" },
  define: {
    "process.env.NODE_ENV": '"production"',
    // 端上有 `import.meta.env.VITE_*`：iife 里取不到，显式给一份静态值。
    "import.meta.env": JSON.stringify({ DEV: false, PROD: true, MODE: "production" })
  }
});
const css = await readFile(path.join(root, "apps/excel-addin/src/styles.css"), "utf8");
const html = `<!doctype html>
<html><head><meta charset="utf-8" />
<style>html,body,#root{height:100%;margin:0;} body{background:var(--color-surface,#fff);}
*{scroll-behavior:auto !important;}</style>
<style>${css}</style>
</head><body><div id="root"></div><script>${bundled.outputFiles[0].text}</script></body></html>`;
const fixturePath = path.join(probeDir, "process_panel_layout_probe.html");
writeFileSync(fixturePath, html, "utf8");

const browser = findBrowser();
const port = await freePort();
const child = spawn(browser, [
  `--remote-debugging-port=${port}`,
  "--user-data-dir=" + path.join(probeDir, "profile-process-panel"),
  "--no-first-run",
  "--no-default-browser-check",
  "--disable-extensions",
  "--headless=new",
  `file:///${fixturePath.replace(/\\/g, "/")}`
], { stdio: "ignore" });

const cdp = await connectCdp(port);
try {
  await cdp.send("Runtime.enable");
  await cdp.send("Page.enable");
  let ready = false;
  for (let attempt = 0; attempt < 100 && !ready; attempt += 1) {
    ready = await cdp.evaluate("Boolean(document.querySelector('.task-progress-timeline'))");
    if (!ready) await new Promise((resolve) => setTimeout(resolve, 100));
  }
  if (!ready) {
    const detail = await cdp.evaluate("document.body.innerHTML.slice(0, 800)");
    assert.fail("夹具没有渲染出过程框；页面内容是：" + detail);
  }

  const measure = () => cdp.evaluate(`(() => {
    const card = document.querySelector('.task-progress-timeline');
    const body = document.querySelector('.task-progress-content-body');
    const footer = document.querySelector('.composer');
    const bottom = document.querySelector('.task-progress-bottom-actions');
    const summary = document.querySelector('.process-result-summary');
    const rect = (node) => node ? node.getBoundingClientRect() : null;
    const f = rect(footer);
    return {
      expanded: card?.getAttribute('data-expanded'),
      folded: card?.getAttribute('data-folded'),
      bodyHeight: body ? Math.round(body.getBoundingClientRect().height) : null,
      bodyScrollHeight: body ? Math.round(body.scrollHeight) : null,
      bodyClientHeight: body ? Math.round(body.clientHeight) : null,
      bottomVisible: bottom ? rect(bottom).bottom <= window.innerHeight + 0.5 : null,
      footerTop: f ? Math.round(f.top) : null,
      footerBottom: f ? Math.round(f.bottom) : null,
      viewport: window.innerHeight,
      summaryVisible: summary ? rect(summary).height > 0 : false,
      enlarge: Boolean(document.querySelector('[data-process-enlarge]'))
    };
  })()`);

  for (const [width, height] of [[320, 640], [400, 720], [500, 760]]) {
    await cdp.send("Emulation.setDeviceMetricsOverride", {
      width, height, deviceScaleFactor: 1, mobile: false
    });
    // 尺寸变化后等一帧，别量到上一档的布局。
    await cdp.evaluate("new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)))");

    const liveState = await measure();
    assert.ok(liveState.bodyHeight !== null, `${width}x${height}: 找不到过程框正文`);
    assert.ok(liveState.bodyHeight <= 161,
      `${width}x${height}: 进行中过程框没有限高（实测 ${liveState.bodyHeight}px，上限 160px）`);
    assert.ok(liveState.bodyScrollHeight > liveState.bodyClientHeight,
      `${width}x${height}: 内容没有超出限高，这次测量证明不了"内部滚动"（scrollHeight=${liveState.bodyScrollHeight}）`);
    assert.ok(liveState.footerBottom <= liveState.viewport + 0.5,
      `${width}x${height}: 输入区被挤出视口（footer bottom=${liveState.footerBottom}, viewport=${liveState.viewport}）`);
    assert.ok(liveState.footerTop >= 0,
      `${width}x${height}: 输入区被顶到视口上方（top=${liveState.footerTop}）`);
    assert.equal(liveState.bottomVisible, true, `${width}x${height}: 底部收起入口被裁掉`);
    assert.ok(liveState.enlarge, `${width}x${height}: 内容超出限高却没有"多显示几行"`);

    // 放大之后必须真的更高，而且仍然不把输入区挤出去。
    await cdp.evaluate("document.querySelector('[data-process-enlarge]').click()");
    await cdp.evaluate("new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)))");
    const enlarged = await measure();
    assert.ok(enlarged.bodyHeight > liveState.bodyHeight,
      `${width}x${height}: 放大没有效果（${liveState.bodyHeight} → ${enlarged.bodyHeight}）`);
    assert.ok(enlarged.footerBottom <= enlarged.viewport + 0.5,
      `${width}x${height}: 放大后输入区被挤出视口`);

    // 还原 → 正常结束：卡片收起成一行，结果摘要仍在（它在过程框之外）。
    await cdp.evaluate("document.querySelector('[data-process-enlarge]').click()");
    await cdp.evaluate("window.__setLive(false)");
    await cdp.evaluate("new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)))");
    const settled = await measure();
    assert.equal(settled.expanded, "false", `${width}x${height}: 正常结束后没有收起`);
    assert.equal(settled.bodyHeight, null, `${width}x${height}: 收起后仍然渲染了详细动作区`);
    assert.equal(settled.summaryVisible, true, `${width}x${height}: 收起后结果摘要不见了`);
    assert.ok(settled.footerBottom <= settled.viewport + 0.5,
      `${width}x${height}: 收起后输入区不在视口内`);

    // 复位，进入下一档视口。
    await cdp.evaluate("window.__setLive(true)");
    console.log(
      `${width}x${height}: 过程框 ${liveState.bodyHeight}px（内容 ${liveState.bodyScrollHeight}px 自己滚）、` +
      `放大 ${enlarged.bodyHeight}px、输入区 ${settled.footerTop}→${settled.footerBottom} 仍在视口；` +
      `结束后收起、摘要仍在`
    );
  }
  console.log(`夹具页：${fixturePath}${keep ? "（已保留）" : ""}`);
} finally {
  cdp.close();
  child.kill();
}
