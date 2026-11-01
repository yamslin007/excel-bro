// 布局证据：「唯一的恢复入口」到底在不在视口里。
//
// 为什么必须有这个脚本：`excelWriteEndToEnd.test.tsx` 能断言"卡片不在 `.message-stream`
// 里"，但 jsdom 没有布局 —— 它证明不了"看得见"。2026-10-29 的真机现象恰恰是
// 「元素存在、按钮也在，但被滚出视口」。所以这里用**真实构建 + 真实样式表 +
// 真实 Chromium 布局引擎**把四件事钉死：
//
// 1. 24 轮长聊天、消息流滚到底之后，恢复入口的按钮矩形仍完整落在视口内；
// 2. 再追加一条日志并滚到底（模拟核对后的自动跟随）之后，按钮矩形**不动**；
// 3. 输入框仍在视口内且没被恢复卡压住；历史里那条记录没有按钮；
// 4. 灵敏度：把同一张卡按**旧位置**（塞回第一条消息里）再量一遍 —— 它必须被判成
//    "看不见"。这一半证明脚本不是永远点头。
//
// 实现走**原始 CDP**（`--remote-debugging-port` + WebSocket）：本机没有 Playwright，
// 而 Chrome 是 GUI 子系统程序，`--dump-dom` 的 stdout 在 Windows 上拿不到。
// CDP 逐条 `Runtime.evaluate` 取数，视口用 `Emulation.setDeviceMetricsOverride` 固定 ——
// 量到的是真实布局，不是算出来的。
//
// 用法：
//   node scripts/check_recovery_entry_viewport.mjs          # 320x640 / 400x720 / 500x760
//   PROBE_BROWSER=<chromium 路径> node scripts/...          # 换浏览器
//   --keep 打印夹具页路径（outputs/probe/），供人工复核
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import { readFile } from "node:fs/promises";
import net from "node:net";
import os from "node:os";
import path from "node:path";
import { build } from "esbuild";

const root = path.resolve(import.meta.dirname, "..");
const probeDir = path.join(root, "outputs", "probe");
const OPERATION_ID = "op-viewport-0001";
const HISTORY_KEY = "excel-bro.chat.v4";

/**
 * 一段**长**聊天记录 + 一条带恢复卡的消息。
 *
 * 恢复卡挂在第 2 条消息上（不是最后一条）：这正是真机的形态 —— 卡属于当时那一轮，
 * 之后又追加了很多日志。旧实现把按钮画在这条消息下面，滚到底自然就看不见了。
 */
function seedHistory() {
  const messages = [];
  for (let index = 0; index < 24; index += 1) {
    messages.push({
      id: `m-user-${index}`,
      role: "user",
      text: `第 ${index + 1} 轮：把汇总表里这一段按月份重算，并说明口径。`,
      createdAt: new Date(Date.UTC(2026, 9, 29, 1, index)).toISOString()
    });
    messages.push({
      id: `m-assistant-${index}`,
      role: "assistant",
      provider: "model",
      text: `第 ${index + 1} 轮已经核对完 ${index + 3} 行，口径与上一轮一致；这一步只读，没有写单元格。`,
      createdAt: new Date(Date.UTC(2026, 9, 29, 1, index, 30)).toISOString()
    });
  }
  messages[3] = {
    ...messages[3],
    officeExecutionId: OPERATION_ID,
    text: "发现一条未完成的 Office 写入（第 2 / 3 步）。",
    officeResume: {
      operationId: OPERATION_ID,
      expectedRevision: 3,
      verdict: { kind: "reconcile", actionIndex: 1, outcome: "pending", remaining: 1 },
      acting: null,
      nextActionIndex: 1,
      totalActions: 3
    }
  };
  return {
    activeConversationId: "conversation-viewport",
    conversations: [
      {
        id: "conversation-viewport",
        title: "视口证据",
        messages,
        createdAt: new Date(Date.UTC(2026, 9, 29, 1, 0)).toISOString(),
        updatedAt: new Date(Date.UTC(2026, 9, 29, 2, 0)).toISOString()
      }
    ]
  };
}

const RECT = `(node) => {
  if (!node) return null;
  const box = node.getBoundingClientRect();
  return { top: box.top, bottom: box.bottom, left: box.left, right: box.right,
    width: box.width, height: box.height };
}`;

/** ① 滚到底后的稳态。 */
const measureSettled = `(() => {
  const rect = ${RECT};
  const panel = document.querySelector(".office-recovery-panel");
  const stream = document.querySelector(".message-stream");
  const textarea = document.querySelector(".composer textarea");
  const record = document.querySelector(".office-recovery-record");
  if (!panel || !stream || !textarea) return { error: "missing-node" };
  /*
   * 入口 = 输入框右下角那颗**主按钮**（2026-11-01 起核对/继续搬到这里）。
   * 卡面上不再有同名按钮：panelButtons 必须为 0，否则"同一个动作两个入口"。
   */
  const button = document.querySelector(".composer .composer-primary");
  if (!button) return { error: "no-primary" };
  // 产品里消息流是 scroll-behavior: smooth：直接赋 scrollTop 会变成一段动画，
  // 量到的是动画中间态。量尺寸时先关掉它（只改本页内联样式，不动产品 CSS）。
  stream.style.scrollBehavior = "auto";
  stream.scrollTop = stream.scrollHeight;
  return {
    viewport: { width: innerWidth, height: innerHeight },
    panelInStream: Boolean(panel.closest(".message-stream")),
    panelInComposer: Boolean(panel.closest(".composer")),
    primaryLabel: button.getAttribute("aria-label"),
    primaryCount: document.querySelectorAll(".composer .composer-primary").length,
    panelButtons: document.querySelectorAll(".office-recovery-panel button").length,
    settled: rect(button),
    streamScrolled: stream.scrollTop,
    streamScrollHeight: stream.scrollHeight,
    streamClientHeight: stream.clientHeight,
    panel: rect(panel),
    textarea: rect(textarea),
    recordInStream: Boolean(record && record.closest(".message-stream")),
    recordButtons: record ? record.querySelectorAll("button").length : -1,
    recordText: record ? record.textContent : ""
  };
})()`;

/** ② 追加一条日志再滚到底：位置必须不动。 */
const measureAfterAppend = `(() => {
  const rect = ${RECT};
  const button = document.querySelector(".composer .composer-primary");
  const stream = document.querySelector(".message-stream");
  if (!button || !stream) return { error: "missing-node" };
  const extra = document.createElement("div");
  extra.className = "message-row system";
  extra.innerHTML = '<div class="message-content"><p class="message-text">已核对在途动作，未写单元格。</p></div>';
  stream.appendChild(extra);
  stream.style.scrollBehavior = "auto";
  stream.scrollTop = stream.scrollHeight;
  return { afterAppend: rect(button), rows: stream.querySelectorAll(".message-row").length };
})()`;

/**
 * ③ 灵敏度对照：把**卡片**塞回第一条消息（旧位置），再滚到底。
 *
 * 入口搬到输入框之后，能不能"被判成看不见"要看**卡片本体**：它是用户读原因、
 * 判进度的那一块。如果它也住在消息流里，滚到底就只剩一句"去输入框核对"而无从判断。
 * 对照组必须复现"看不见"，否则这个脚本没有否定的能力。
 */
const measureLegacy = `(() => {
  const rect = ${RECT};
  const panel = document.querySelector(".office-recovery-panel");
  const stream = document.querySelector(".message-stream");
  const firstContent = document.querySelector(".message-row .message-content");
  if (!panel || !stream || !firstContent) return { error: "missing-node" };
  firstContent.appendChild(panel);
  stream.style.scrollBehavior = "auto";
  stream.scrollTop = stream.scrollHeight;
  return { legacy: rect(panel), legacyInStream: Boolean(panel.closest(".message-stream")) };
})()`;

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

/** 最小 CDP 客户端：一条连接、请求-响应配对。 */
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

const source = `
import React from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
createRoot(document.getElementById('root')).render(<App />);
`;

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
 *
 * `focus.css` **故意不在列表里**：它属于另一个入口（`focus.tsx`，整窗"专注"页），
 * 里面把 `html/body/#root` 改成 `height:auto; min-width:720px`。混进来会让任务窗格
 * 变成整页文档流 —— 消息流于是永远"没有可滚内容"，量出来的位置毫无意义。
 */
const cssFiles = [
  "apps/excel-addin/src/styles.css",
  "apps/excel-addin/src/components/taskExecutionNotice.css",
  "apps/excel-addin/src/components/composerPrimaryButton.css",
  "apps/excel-addin/src/components/composerActionDock.css",
  "apps/excel-addin/src/components/executionResult.css",
  "apps/excel-addin/src/components/worksheetCheckpointCard.css"
];
const css = (
  await Promise.all(cssFiles.map((relative) => readFile(path.join(root, relative), "utf8")))
).join("\n");

const pageHtml = `<!doctype html><meta charset="utf-8">
<title>recovery entry viewport probe</title>
<style>${css}</style>
<div id="root"></div>
<script>
try {
  const tall = location.hash === "#tall" || location.search.includes("pass=tall");
  const history = ${JSON.stringify(seedHistory())};
  if (tall) {
    /*
     * 窄窗格里"卡片再长也不许把输入框挤出屏幕"：给一条**很长**的阻断原因，
     * 面板正文必须自己滚（.office-recovery-panel-body），动作区留在原地。
     */
    const message = history.conversations[0].messages[3];
    message.officeResume.verdict = {
      kind: "blocked",
      detail: "当前不是产生该断点的工作簿。" + "这一段是为了把卡片撑长而重复的原因说明。".repeat(14)
    };
    message.officeResume.nextActionIndex = 2;
  }
  localStorage.setItem(${JSON.stringify(HISTORY_KEY)}, JSON.stringify(history));
} catch (error) {}
</script>
<script>${bundled.outputFiles[0].text}</script>
`;
mkdirSync(probeDir, { recursive: true });
const probePath = path.join(probeDir, "recovery_entry_viewport_probe.html");
writeFileSync(probePath, pageHtml, "utf8");
const probeUrl = `file:///${probePath.replaceAll("\\", "/")}`;

const browserPath = findBrowser();
const port = await freePort();
const profileDir = path.join(os.tmpdir(), `eb-viewport-probe-${process.pid}`);
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
    `--user-data-dir=${profileDir}`,
    probeUrl
  ],
  { stdio: "ignore" }
);

const insideViewport = (box, viewport, label) => {
  assert.ok(box, `${label}: 按钮没有布局盒子`);
  assert.ok(box.width > 0 && box.height > 0, `${label}: 按钮尺寸为零`);
  assert.ok(box.top >= -0.5, `${label}: 按钮上沿在视口外 (top=${box.top})`);
  assert.ok(box.bottom <= viewport.height + 0.5,
    `${label}: 按钮下沿在视口外 (bottom=${box.bottom}, viewport=${viewport.height})`);
  assert.ok(box.left >= -0.5 && box.right <= viewport.width + 0.5,
    `${label}: 按钮横向被裁 (left=${box.left}, right=${box.right}, viewport=${viewport.width})`);
};

/** ④ 高卡片量尺：正文很长时，面板自己滚、动作与输入框都还在视口里。 */
const measureTall = `(() => {
  const rect = ${RECT};
  const panel = document.querySelector(".office-recovery-panel");
  const body = panel && panel.querySelector(".office-recovery-panel-body");
  const stream = document.querySelector(".message-stream");
  const textarea = document.querySelector(".composer textarea");
  const button = panel && panel.querySelector(".office-recovery-panel-actions button");
  if (!panel || !body || !stream || !textarea || !button) return { error: "missing-node" };
  stream.style.scrollBehavior = "auto";
  stream.scrollTop = stream.scrollHeight;
  return {
    viewport: { width: innerWidth, height: innerHeight },
    panel: rect(panel),
    panelInStream: Boolean(panel.closest(".message-stream")),
    bodyScrollHeight: body.scrollHeight,
    bodyClientHeight: body.clientHeight,
    button: rect(button),
    buttonText: button.textContent.trim(),
    textarea: rect(textarea)
  };
})()`;

const results = [];
let client;
try {
  client = await connectCdp(port);
  for (const [width, height] of [[320, 640], [400, 720], [500, 760]]) {
    // 视口必须是**我们指定的**：量到一个别的尺寸等于没量。
    await client.send("Emulation.setDeviceMetricsOverride", {
      width,
      height,
      deviceScaleFactor: 1,
      mobile: false
    });
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: probeUrl });
    const label = `${width}x${height}`;
    const deadline = Date.now() + 25_000;
    let ready = false;
    while (Date.now() < deadline) {
      ready = await client.evaluate(
        `Boolean(document.querySelector(".office-recovery-panel")) && Boolean(document.querySelector(".composer .composer-primary")) && Boolean(document.querySelector(".composer textarea"))`
      );
      if (ready) break;
      await new Promise((resolve) => setTimeout(resolve, 120));
    }
    assert.ok(ready, `${label}: 页面没有渲染出恢复入口（长历史没被读出来？）`);

    const settled = await client.evaluate(measureSettled);
    assert.equal(settled.error, undefined, `${label}: ${settled.error}`);
    assert.equal(settled.viewport.width, width, `${label}: 视口宽度没有被固定住（${settled.viewport.width}）`);
    assert.equal(settled.viewport.height, height, `${label}: 视口高度没有被固定住（${settled.viewport.height}）`);

    // 结构与唯一性
    assert.equal(settled.panelInStream, false, `${label}: 恢复入口不许住在滚动容器里`);
    assert.equal(settled.panelInComposer, true, `${label}: 恢复入口必须住在输入区`);
    assert.equal(settled.primaryLabel, "核对在途动作", `${label}: 唯一入口的文案`);
    assert.equal(settled.primaryCount, 1, `${label}: 主入口只许有一颗（否则同一个动作两个入口）`);
    assert.equal(settled.panelButtons, 0, `${label}: 卡面上不许再有同名按钮`);
    assert.ok(settled.streamScrolled > 0, `${label}: 消息流确实滚动了（否则"位置不变"是空的）`);
    assert.ok(settled.streamScrollHeight > settled.streamScrolled,
      `${label}: 消息流没有可滚内容（长聊天没生效，这条证据就空了）`);
    assert.ok(settled.streamScrolled >= settled.streamScrollHeight - settled.streamClientHeight - 1,
      `${label}: 没有真的滚到底 (scrollTop=${settled.streamScrolled}, max=${
        settled.streamScrollHeight - settled.streamClientHeight})`);

    // ① 长聊天滚到底：入口在视口内
    insideViewport(settled.settled, settled.viewport, `${label} 滚到底`);

    // 截图：给人看的那一份（`--shots`）。数值断言在上面，图只用于人工复核版式。
    if (process.argv.includes("--shots")) {
      const shot = await client.send("Page.captureScreenshot", { format: "png" });
      writeFileSync(
        path.join(probeDir, `recovery-entry-${width}x${height}.png`),
        Buffer.from(shot.result.data, "base64")
      );
    }

    // ② 追加日志再滚到底：位置不动
    const appended = await client.evaluate(measureAfterAppend);
    assert.equal(appended.error, undefined, `${label}: ${appended.error}`);
    assert.ok(Math.abs(appended.afterAppend.top - settled.settled.top) <= 0.5,
      `${label}: 追加日志改变了入口位置 (before=${settled.settled.top}, after=${appended.afterAppend.top})`);
    assert.ok(Math.abs(appended.afterAppend.left - settled.settled.left) <= 0.5,
      `${label}: 追加日志改变了入口水平位置`);
    insideViewport(appended.afterAppend, settled.viewport, `${label} 追加日志后`);

    // ③ 输入框仍在视口内，且没被恢复卡压住
    assert.ok(settled.textarea.top >= -0.5 && settled.textarea.bottom <= height + 0.5,
      `${label}: 输入框被挤出视口 (${JSON.stringify(settled.textarea)})`);
    assert.ok(settled.panel.bottom <= settled.textarea.top + 0.5,
      `${label}: 恢复卡压住了输入框 (panel bottom=${settled.panel.bottom}, textarea top=${settled.textarea.top})`);

    // ④ 历史里留记录、记录里没有按钮
    assert.equal(settled.recordInStream, true, `${label}: 执行记录应留在历史里`);
    assert.equal(settled.recordButtons, 0, `${label}: 执行记录里不许有按钮`);
    assert.ok(settled.recordText.includes("已完成 1 / 3 步"), `${label}: 记录应写明进度 (${settled.recordText})`);

    // ⑤ 灵敏度：旧位置（卡片塞回旧消息 + 滚到底）必须看不见
    const legacy = await client.evaluate(measureLegacy);
    assert.equal(legacy.error, undefined, `${label}: ${legacy.error}`);
    assert.equal(legacy.legacyInStream, true, `${label}: 灵敏度对照组没有真的把卡片塞回消息流`);
    assert.ok(!legacy.legacy || legacy.legacy.bottom <= 0.5,
      `${label}: 旧位置竟然还看得见 —— 这个脚本无法复现真机缺陷，不许用它宣称视口内 ` +
      `(bottom=${legacy.legacy?.bottom})`);

    results.push(
      `${label}: 入口在视口内 (top=${settled.settled.top.toFixed(1)}, bottom=${settled.settled.bottom.toFixed(1)})、` +
      `追加日志后不变 (top=${appended.afterAppend.top.toFixed(1)})、输入框 top=${settled.textarea.top.toFixed(1)}；` +
      `旧位置复现"看不见" (bottom=${legacy.legacy ? legacy.legacy.bottom.toFixed(1) : "n/a"})`
    );
  }

  /*
   * ④ 最高最窄的那一档：一段很长的阻断原因 + 320x640。
   *
   * 面板的 `max-height: min(46vh, 320px)` 与"正文自己滚、动作区是兄弟"这两条一起
   * 保证「卡片再长也不把输入框挤出屏幕」。这里要求三件事同时成立，缺一条都不算：
   * 面板被限高、正文确实溢出了（否则这条证据是空的）、按钮与输入框都在视口里。
   */
  {
    const width = 320;
    const height = 640;
    await client.send("Emulation.setDeviceMetricsOverride", {
      width,
      height,
      deviceScaleFactor: 1,
      mobile: false
    });
    /*
     * URL 要**真的变**：只改 hash 不会重新加载文档，而上一档的灵敏度对照把面板
     * 搬进过消息流 —— 不换文档就会量到上一次留下的 DOM。查询串换了就是新文档。
     */
    await client.send("Page.navigate", { url: `${probeUrl}?pass=tall#tall` });
    const deadline = Date.now() + 25_000;
    let ready = false;
    while (Date.now() < deadline) {
      ready = await client.evaluate(
        `Boolean(document.querySelector(".office-recovery-panel-actions button")) && Boolean(document.querySelector(".composer textarea"))`
      );
      if (ready) break;
      await new Promise((resolve) => setTimeout(resolve, 120));
    }
    assert.ok(ready, "320x640 高卡片: 页面没有渲染出恢复入口");
    if (process.argv.includes("--shots")) {
      const shot = await client.send("Page.captureScreenshot", { format: "png" });
      writeFileSync(
        path.join(probeDir, "recovery-entry-tall-320x640.png"),
        Buffer.from(shot.result.data, "base64")
      );
    }
    const tall = await client.evaluate(measureTall);
    assert.equal(tall.error, undefined, `320x640 高卡片: ${tall.error}`);
    assert.equal(tall.panelInStream, false, "320x640 高卡片: 恢复入口不许住在滚动容器里");
    assert.equal(tall.buttonText, "重新核对（只读）", "320x640 高卡片: 阻断时只有只读重核");
    assert.ok(tall.panel.height <= height * 0.46 + 2,
      `高卡片: 面板没有被限高 (height=${tall.panel.height})`);
    assert.ok(tall.bodyScrollHeight > tall.bodyClientHeight,
      `高卡片: 正文没有溢出，说明这条证据是空的 (scroll=${tall.bodyScrollHeight}, client=${tall.bodyClientHeight})`);
    insideViewport(tall.button, tall.viewport, "320x640 高卡片 动作");
    assert.ok(tall.textarea.top >= -0.5 && tall.textarea.bottom <= height + 0.5,
      `高卡片: 输入框被挤出视口 (${JSON.stringify(tall.textarea)})`);
    assert.ok(tall.panel.bottom <= tall.textarea.top + 0.5,
      `高卡片: 恢复卡压住了输入框 (panel bottom=${tall.panel.bottom}, textarea top=${tall.textarea.top})`);
    results.push(
      `320x640 高卡片: 面板限高生效 (height=${tall.panel.height.toFixed(1)}，正文 ${tall.bodyScrollHeight}→${tall.bodyClientHeight} 自己滚)、` +
      `动作 bottom=${tall.button.bottom.toFixed(1)}、输入框 top=${tall.textarea.top.toFixed(1)}`
    );
  }
} finally {
  client?.close();
  child.kill();
}

console.log(results.join("\n"));
console.log(`夹具页：${probePath}${process.argv.includes("--keep") ? "" : "（加 --keep 表示保留供人工复核）"}`);
