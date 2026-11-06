// Offline browser check: transactional appearance settings and both reading surfaces.
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import net from 'node:net';

const root = path.resolve(import.meta.dirname, '..');
const out = path.join(root, 'outputs/probe/scope-overlay');
mkdirSync(out, { recursive: true });
const browser = [process.env.PROBE_BROWSER,
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'].filter(Boolean).find(existsSync);
assert.ok(browser, 'Set PROBE_BROWSER to a Chromium executable');


// Layout fixture uses production styles and host-specific worksheet row markup.
// Mounted application tests separately assert the production overlay ancestry.
const read = name => readFileSync(path.join(root, name), 'utf8');
const css = read('apps/excel-addin/src/styles.css');
const wpsCss = read('apps/wps-addin/src/styles.css');
const shared = read('apps/excel-addin/src/appearance.css') + read('apps/excel-addin/src/conversation.css');
const file = path.join(out, 'fixture.html');
const rows = host => Array.from({length:40}, (_,i) => `<button aria-pressed="true" class="selected"><i class="sheet-checkbox">✓</i><span><strong>工作表 ${i+1} · 销售汇总</strong>${host==='excel'?'<small>字段按需读取</small>':''}</span></button>`).join('');
const render = host => {
 const fileCard = host==='excel' ? `<div class="workbook-identity"><i>XLSX</i><span><small>当前文件</small><strong>需要核对的数据工作簿.xlsx</strong><em>当前工作表 来源</em></span></div>` : `<div class="wps-current-file-card"><span>XLSX</span><div><small>当前文件</small><strong>需要核对的数据工作簿.xlsx</strong><em>当前工作表 来源</em></div></div>`;
 const modes = `<div class="scope-controls"><fieldset class="scope-selection-fields"><div class="scope-mode-list"><button><i>⊙</i><span><strong>跟随当前</strong><small>发送时自动使用正在查看的工作表</small></span></button><button class="selected"><i>☷</i><span><strong>选择多个</strong><small>用于跨表查询、比较或汇总</small></span></button><div class="scope-picker-reveal open"><div class="${host==='excel'?'workbook-sheet-picker':'wps-manual-sheet-picker'}"><div class="${host==='excel'?'sheet-search-row':'wps-sheet-search-row'}"><input placeholder="搜索工作表…"><span>已选 40/40</span></div><div class="sheet-picker-toolbar"><button>全选搜索结果</button><button>清空</button></div><div class="${host==='excel'?'workbook-sheet-list':'sheet-picker-options wps-sheet-picker-options'}">${rows(host)}</div></div></div><button><i>▣</i><span><strong>工作目录</strong><small>选择文件中的工作表</small></span></button></div></fieldset></div>`;
 return `<main class="chat-shell chat-shell-headerless ${host==='wps'?'wps-chat-shell':''}"><div class="message-stream">对话内容</div><div class="scope-layer"><section class="scope-popover ${host==='wps'?'wps-scope-popover':''}"><div class="scope-popover-header"><div><strong>这次对话使用哪些数据？</strong><span class="scope-authorization-hint">勾选即授权，取消勾选即撤销</span></div><button aria-label="关闭数据范围">×</button></div><div class="scope-popover-scroll">${host==='excel'?'<fieldset class="scope-selection-fields">':''}${fileCard}${modes}${host==='excel'?'</fieldset>':''}</div><div class="${host==='excel'?'scope-action-buttons':'wps-scope-actions'}"><button class="scope-refresh">↻ 刷新</button><button class="scope-close">确认</button></div></section></div><footer class="composer"><div class="composer-box" style="height:var(--probe-composer-height)">输入中的多行草稿</div></footer></main>`;
};
writeFileSync(file,`<!doctype html><meta charset="utf-8"><style>${css}</style><style id="host-css"></style><style>${shared}</style><style>*,*::before,*::after{animation:none!important;transition:none!important}</style><div id="root"></div>`);
const port=await new Promise(resolve=>{const server=net.createServer();server.listen(0,'127.0.0.1',()=>{const p=server.address().port;server.close(()=>resolve(p));});});
const url=`file:///${file.replaceAll('\\','/')}`;
const child=spawn(browser,['--headless=new','--no-first-run','--disable-extensions',`--remote-debugging-port=${port}`,`--user-data-dir=${path.join(out,'profile')}`,url],{stdio:'ignore',windowsHide:true});
let socket;
try {
 let target;
 for(let i=0;i<100;i++){try{target=(await(await fetch(`http://127.0.0.1:${port}/json/list`)).json()).find(t=>t.type==='page');if(target)break;}catch{}await new Promise(r=>setTimeout(r,100));}
 assert.ok(target,'Browser startup timed out');
 socket=new WebSocket(target.webSocketDebuggerUrl);
 await new Promise((resolve,reject)=>{socket.addEventListener('open',resolve,{once:true});socket.addEventListener('error',reject,{once:true});});
 let id=0;const pending=new Map();
 socket.addEventListener('message',e=>{const reply=JSON.parse(e.data);if(pending.has(reply.id)){pending.get(reply.id)(reply);pending.delete(reply.id);}});
 const send=(method,params={})=>new Promise(resolve=>{pending.set(++id,resolve);socket.send(JSON.stringify({id,method,params}));});
 const evaluate=async expression=>{const reply=await send('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});assert.ok(!reply.result.exceptionDetails,JSON.stringify(reply));return reply.result.result.value;};
 const wait=async(expression)=>{for(let i=0;i<100;i++){if(await evaluate(expression))return;await new Promise(r=>setTimeout(r,50));}throw Error('Timed out: '+expression);};
 const frame=()=>evaluate('new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)))');
 const shot=async name=>{const reply=await send('Page.captureScreenshot',{format:'png'});writeFileSync(path.join(out,name+'.png'),Buffer.from(reply.result.data,'base64'));};

 await wait("!!document.querySelector('#host-css')");
 const results=[];
 for(const host of ['excel','wps']) {
  await evaluate(`document.querySelector('#host-css').textContent=${JSON.stringify(host==='wps'?wpsCss:'')};document.querySelector('#root').innerHTML=${JSON.stringify(render(host))}`);
  for(const [width,height] of [[320,480],[400,360],[740,760],[1280,900]]) {
   await send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});
   for(const composerHeight of [72,300]) {
    await evaluate(`document.documentElement.style.setProperty('--probe-composer-height','${composerHeight}px');document.querySelector('.scope-popover-scroll').scrollTop=0`);
    await frame();
    const geometry = await evaluate(`(()=>{
     const box=s=>{const e=document.querySelector(s),r=e.getBoundingClientRect();return {top:r.top,bottom:r.bottom,left:r.left,right:r.right,width:r.width,height:r.height}};
     const body=document.querySelector('.scope-popover-scroll');
     return {panel:box('.scope-popover'),header:box('.scope-popover-header'),footer:box('.scope-close'),body:box('.scope-popover-scroll'),overflow:body.scrollHeight-body.clientHeight,nested:[...body.querySelectorAll('*')].filter(e=>getComputedStyle(e).overflowY==='auto'&&e.scrollHeight>e.clientHeight+1).map(e=>e.className),pageOverflow:document.documentElement.scrollWidth>innerWidth};
    })()`);
    const label=`${host} ${width}x${height} composer=${composerHeight}`;
    assert.ok(geometry.panel.top>=0&&geometry.panel.bottom<=height,label+' panel clipped');
    assert.ok(geometry.header.top>=geometry.panel.top&&geometry.header.bottom<=geometry.body.top+1,label+' header overlaps body');
    assert.ok(geometry.footer.bottom<=height&&geometry.footer.top>=geometry.body.bottom-1,label+' footer overlaps body');
    assert.ok(geometry.body.height>70,label+' no reading space');
    assert.ok(geometry.overflow>300,label+' long-list fixture must scroll');
    assert.deepEqual(geometry.nested,[],label+' nested scrollbar');
    assert.equal(geometry.pageOverflow,false,label+' horizontal overflow');
    await evaluate(`document.querySelector('.scope-popover-scroll').scrollTop=100000`);await frame();
    const after=await evaluate(`(()=>{const e=document.querySelector('.scope-close'),r=e.getBoundingClientRect(),rows=document.querySelectorAll('.workbook-sheet-list > button,.wps-sheet-picker-options > button'),last=rows[rows.length-1].getBoundingClientRect(),body=document.querySelector('.scope-popover-scroll').getBoundingClientRect();return {footerTop:r.top,hit:e.contains(document.elementFromPoint(r.left+r.width/2,r.top+r.height/2)),lastVisible:last.top>=body.top&&last.bottom<=body.bottom+1}})()`);
    assert.equal(after.footerTop,geometry.footer.top,label+' footer moved');
    assert.ok(after.hit&&after.lastVisible,label+' confirm or last sheet unreachable');
    results.push({host,width,height,composerHeight,panel:geometry.panel,scrollHeight:geometry.body.height});
    if(composerHeight===300){await evaluate(`document.querySelector('.scope-popover-scroll').scrollTop=0`);await frame();await shot(`${host}-${width}x${height}`);}
   }
  }
 }
 writeFileSync(path.join(out,'results.json'),JSON.stringify(results,null,2));
 console.log(`PASS: ${results.length} host/viewport/composer combinations; full panel bounds, fixed header/footer, one scrollbar, last sheet and confirm reachable.`);
} finally {socket?.close();child.kill();}
