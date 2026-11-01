// 输入框内澄清：真实共享组件和双端样式，检验限高、自滚和发送入口。
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

// 夹具：真实澄清组件、长选项、补充文字和底部操作区。
const source = `
import React, { useState } from 'react';
import { createRoot } from 'react-dom/client';
import { ClarificationComposer } from './components/ClarificationComposer';
import { ComposerPrimaryButton } from './components/ComposerPrimaryButton';
function Harness() {
 const [selected, setSelected] = useState([]);
 return <div className="chat-shell"><header className="chat-header">Excel Bro</header>
 <main className="message-stream">聊天记录仍可见</main><footer className="composer">
 <div className="composer-box is-answering"><div className="clarification-answer-region">
 <ClarificationComposer question="等级依据哪一列（或哪几个字段）来判定？" options={Array.from({length:8},(_,i)=>({id:String(i),label:'选择规则 '+i,description:'按字段数值判断；保留长说明以检验窄窗格换行与内部滚动。',resolution:String(i)}))} selected={selected} multiple={false} freeText="" busy={false} onToggle={id=>setSelected([id])} onFieldChange={()=>{}} />
 </div>
 <textarea aria-label="其他" data-clarification-other="true" style={{height:200}} placeholder="其他" />
 <div className="composer-toolbar"><button type="button" className="clarification-bypass-toggle">取消本次任务</button><ComposerPrimaryButton action={{kind:"answer",label:"发送回答",title:"提交回答",glyph:"↑",invoke:"answer",disabled:false,busy:false}} onInvoke={()=>{}} /></div>
 </div></footer></div>;
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
const css = await readFile(path.join(root, "apps/excel-addin/src/styles.css"), "utf8") + await readFile(path.join(root, "apps/excel-addin/src/components/clarificationComposer.css"), "utf8");
const buttonCss = await readFile(path.join(root, "apps/excel-addin/src/components/composerPrimaryButton.css"), "utf8");
const html = `<!doctype html>
<html><head><meta charset="utf-8" />
<style>html,body,#root{height:100%;margin:0;} body{background:var(--color-surface,#fff);}
*{scroll-behavior:auto !important;}</style>
<style>${css}${buttonCss}</style>
</head><body><div id="root"></div><script>${bundled.outputFiles[0].text}</script></body></html>`;
const fixturePath = path.join(probeDir, "clarification_inbox_probe.html");
writeFileSync(fixturePath, html, "utf8");

const browser = findBrowser();
const port = await freePort();
const child = spawn(browser, [
  `--remote-debugging-port=${port}`,
  "--user-data-dir=" + path.join(probeDir, "profile-clarification-inbox"),
  "--no-first-run",
  "--no-default-browser-check",
  "--disable-extensions",
  "--headless=new",
  `file:///${fixturePath.replace(/\\/g, "/")}`
], { stdio: "ignore" });

const cdp = await connectCdp(port);
try {
 await cdp.send("Runtime.enable"); await cdp.send("Page.enable");
 for(let i=0;i<100;i++) { if(await cdp.evaluate("!!document.querySelector('.clarification-composer')")) break; await new Promise(r=>setTimeout(r,100)); }
 for(const host of ['excel','wps']) {
  if(host==='wps') {
   const css = await readFile(path.join(root,'apps/wps-addin/src/styles.css'),'utf8');
   await cdp.evaluate(`document.querySelector('.chat-shell').classList.add('wps-chat-shell');const style=document.createElement('style');style.textContent=${JSON.stringify(css)};document.head.append(style)`);
  }
  for(const [width,height] of [[320,640],[400,720],[500,760]]) {
   await cdp.send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});
   await cdp.evaluate('new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)))');
   const result=await cdp.evaluate(`(()=>{
    const region=document.querySelector('.clarification-choices-scroll'),box=document.querySelector('.composer-box'),button=document.querySelector('[aria-label="发送回答"]');
    const r=region.getBoundingClientRect(),b=button.getBoundingClientRect();
    const before=document.querySelector(".clarification-composer-heading").getBoundingClientRect().top; region.scrollTop=region.scrollHeight;
    const fixed=document.querySelector(".clarification-composer-heading").getBoundingClientRect().top===before;
    const cancel=document.querySelector(".clarification-bypass-toggle").getBoundingClientRect();
    const heading=document.querySelector(".clarification-composer-mode").getBoundingClientRect(),question=document.querySelector(".clarification-composer-question").getBoundingClientRect();
    return {fixed,aligned:Math.abs(cancel.top-b.top)<4,inline:question.left>heading.right,circle:getComputedStyle(document.querySelector(".clarification-options button"),"::before").content,other:document.querySelector("textarea").placeholder,inside:box.contains(region),scroll:region.scrollHeight>region.clientHeight,regionHeight:r.height,buttonBottom:b.bottom,buttonTop:b.top,width:document.documentElement.scrollWidth,textareaHeight:document.querySelector('textarea').getBoundingClientRect().height};
   })()`);
   assert.ok(result.inside && result.scroll,JSON.stringify(result));
   assert.ok(result.fixed && result.aligned && result.inline,JSON.stringify(result));
   assert.equal(result.circle,"none"); assert.equal(result.other,"其他");
   assert.ok(result.buttonTop>=0 && result.buttonBottom<=height,JSON.stringify(result));
   assert.ok(result.width<=width && result.textareaHeight<=73,JSON.stringify(result));
   console.log(host,width,height,result);
  }
 }
 await cdp.send('Page.captureScreenshot',{format:'png'}).then(r=>writeFileSync(path.join(probeDir,'clarification-inbox.png'),Buffer.from(r.result.data,'base64')));
} finally { cdp.close(); child.kill(); }
