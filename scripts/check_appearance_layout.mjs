// Offline browser check: transactional appearance settings and both reading surfaces.
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import net from 'node:net';
import { build } from 'esbuild';

const root = path.resolve(import.meta.dirname, '..');
const out = path.join(root, 'outputs/probe/appearance');
mkdirSync(out, { recursive: true });
const browser = [process.env.PROBE_BROWSER,
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'].filter(Boolean).find(existsSync);
assert.ok(browser, 'Set PROBE_BROWSER to a Chromium executable');

const source = `
import React,{useState} from 'react';
import {createRoot} from 'react-dom/client';
import ThemePanel from './components/ThemePanel';
import TaskProgressTimeline from './components/TaskProgressTimeline';
import ProcessResultSummary from './components/ProcessResultSummary';
import {ExecutionResultCard} from './components/ExecutionResultCard';
import {ConversationMessage} from './components/ConversationMessage';
import ScrollToLatestButton from './components/ScrollToLatestButton';
import MessageCopyButton from './components/MessageCopyButton';
import {ClarificationMessageIntro} from './components/ClarificationMessageIntro';
import {RecoveryTerminalRecord} from './components/RecoveryTerminalRecord';
import {brandLogoUrl} from './brandLogo';
import {useTheme} from './hooks/useTheme';
import {saveThemeSnapshot,loadThemeSettings,loadBackgroundImage} from './utils/imageStorage';
import {APPEARANCE_DEFAULTS} from './utils/appearanceSettings';
const image = new Blob(['<svg xmlns="http://www.w3.org/2000/svg" width="1000" height="900"><defs><linearGradient id="g" x2="1" y2="1"><stop stop-color="#123c30"/><stop offset=".55" stop-color="#82bca0"/><stop offset="1" stop-color="#f8df9f"/></linearGradient></defs><rect width="1000" height="900" fill="url(#g)"/><g fill="#2d6953" opacity=".45"><ellipse cx="120" cy="150" rx="220" ry="100" transform="rotate(35 120 150)"/><ellipse cx="820" cy="600" rx="300" ry="140" transform="rotate(-30 820 600)"/></g><g fill="#fff" opacity=".17"><circle cx="750" cy="130" r="85"/><circle cx="600" cy="250" r="50"/><circle cx="300" cy="700" r="100"/></g></svg>'],{type:'image/svg+xml'});

window.storageRead=loadThemeSettings;
window.imageRead=loadBackgroundImage;
const log={operationId:'reading-probe',outcome:'completed',totalMs:24300,steps:Array.from({length:6},(_,i)=>({id:'step-'+i,label:['理解需求','核对授权范围','读取公式','检查引用','验证计算结果','整理检查结果'][i],phase:'planning',status:'completed',elapsedMs:4050,note:'仅检查已选择的工作表，未修改单元格。',detail:'读取结果与预期一致；详细信息保持清晰可读。'}))};
function Fixture(){const [running,setRunning]=useState(false);window.setProbeRunning=setRunning;const api=useTheme();const [open,setOpen]=useState(false);window.themeApi=api;window.openAppearance=()=>setOpen(true);return <div className={'chat-shell chat-shell-headerless '+(location.search?'wps-chat-shell':'')}>
<div className="message-stream-shell"><main className="message-stream">
<ConversationMessage role="user" author={false} executionResult={false} actions={<MessageCopyButton copied={false} onClick={()=>{}}/>}><p className="message-text">帮我检查已选择的三张工作表，看看公式引用是否一致。</p></ConversationMessage>
<ConversationMessage role="assistant"><TaskProgressTimeline log={log}/></ConversationMessage>
<ConversationMessage role="assistant" author={true} executionResult={false} actions={<MessageCopyButton copied={false} onClick={()=>{}}/>}><p className="message-text"><strong>三张表的公式引用一致</strong><br/>已核对结果1、结果2、结果3 的 B2。公式引用与预期相符，本次检查未修改工作表。</p><TaskProgressTimeline log={log}/></ConversationMessage>
<ConversationMessage role="assistant" author={true} executionResult={false}><section className="clarification-card cancelled"><ClarificationMessageIntro status="cancelled" summary="确认检查范围" question="下一次检查，需要包括空白单元格吗？" reason="本次澄清已取消，未修改工作表。" scopeLabel="仅讨论需求"/></section><RecoveryTerminalRecord kind="undone" title="已撤销本次执行" summary="这次写入已撤销，展开可查看执行记录。"><p>原执行结果已恢复。</p></RecoveryTerminalRecord></ConversationMessage>
<ConversationMessage role="system" author={false} executionResult={false}><p className="message-text">示例对话 · 仅用于视觉验收</p></ConversationMessage>
<ConversationMessage role="system" author={false} executionResult={true}><p className="message-text">示例写入已完成，等待你审阅结果。</p><ProcessResultSummary input={{live:false,phase:'reviewing',verification:{status:'verified',total:6,passed:6,unverified:0}}}/><ExecutionResultCard tone="verified" title="验证通过" summary="6/6 项符合预期"><div className="execution-recovery-footer"><span>审阅入口保持可用，本夹具不会操作工作簿。</span></div></ExecutionResultCard></ConversationMessage>
</main><ScrollToLatestButton running={running} onClick={()=>{const el=document.querySelector(".message-stream");el.scrollTop=el.scrollHeight;}}/></div>
<footer className="composer"><div className="appearance-sample-scope"><button className="scope-new-chat">＋ 新对话</button><button className="secondary-button">↶</button><button className="secondary-button">◷</button><span className="scope-trigger">工作簿.xlsx · 已选 3 张表</span></div><div className="composer-box"><textarea aria-label="输入消息" placeholder="描述你想查询、分析或修改的内容…"/><div className="composer-toolbar"><div className="composer-tools-left"><button className="composer-app-trigger" onClick={()=>setOpen(true)} aria-label="壁纸设置"><img src={brandLogoUrl} alt=""/></button></div><span className="composer-shortcut-hint">Enter 发送 · Shift + Enter 换行</span><div className="composer-tools-right"><button className="model-trigger">模型 ⌄</button><button className="composer-primary send-button" aria-label="发送">▲</button></div></div></div></footer>
<ThemePanel {...api} open={open} onClose={()=>setOpen(false)}/></div>}

let root;
window.remount=()=>{root?.unmount();root=createRoot(document.getElementById('root'));root.render(<Fixture/>)};
saveThemeSnapshot({...APPEARANCE_DEFAULTS,hasCustomBackground:true},image).then(()=>window.remount());
`;
const bundle=await build({stdin:{contents:source,resolveDir:path.join(root,'apps/excel-addin/src'),loader:'tsx'},bundle:true,loader:{'.png':'dataurl','.svg':'dataurl'},write:false,outfile:'fixture.js',format:'iife',define:{'process.env.NODE_ENV':'"production"','import.meta.env':'{}'}});
const css=readFileSync(path.join(root,'apps/excel-addin/src/styles.css'),'utf8');
const wpsCss=readFileSync(path.join(root,'apps/wps-addin/src/styles.css'),'utf8');
const scrollbars=readFileSync(path.join(root,'apps/excel-addin/src/chatScrollbars.css'),'utf8');
const componentCss=bundle.outputFiles.filter(f=>f.path.endsWith('.css')).map(f=>f.text).join('\n');
const shared=readFileSync(path.join(root,'apps/excel-addin/src/appearance.css'),'utf8')+readFileSync(path.join(root,'apps/excel-addin/src/conversation.css'),'utf8');
const file=path.join(out,'fixture.html');
writeFileSync(file,`<!doctype html><meta charset="utf-8"><style>${css}</style><style>${componentCss}</style><style id="host-css"></style><style>${scrollbars}</style><style>${shared}</style><div id="root"></div><script>${bundle.outputFiles[0].text}</script>`);
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
 const messageGeometry = new Map();
 for(const host of ['excel','wps']){
  await send('Page.navigate',{url:url+(host==='wps'?'?wps':'')});
  await wait('window.themeApi?.ready && !!window.themeApi.backgroundUrl');
  if(host==='wps')await evaluate(`document.getElementById('host-css').textContent=${JSON.stringify(wpsCss)}`);
  for(const [width,height] of [[320,640],[400,720],[740,760],[760,760],[1000,760],[1280,900],[320,480]]){
   await send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});
   await evaluate('window.openAppearance()'); await wait('window.themeApi.editing');
   await evaluate('window.themeApi.updateSettings({brightness:85,cardOpacity:88,cardBlur:12,blur:4})');await frame();
   const geometry=await evaluate(`(()=>{const d=document.querySelector('.appearance-dialog').getBoundingClientRect(),p=document.querySelector('.appearance-preview').getBoundingClientRect(),f=document.querySelector('.appearance-footer').getBoundingClientRect(),c=document.querySelector('.appearance-scroll'),b=document.querySelector('.appearance-apply').getBoundingClientRect();return {left:d.left,right:d.right,bottom:d.bottom,top:d.top,previewBottom:p.bottom,footerTop:f.top,footerBottom:f.bottom,scrollable:c.scrollHeight>c.clientHeight,hit:document.elementFromPoint(b.left+b.width/2,b.top+b.height/2)?.classList.contains('appearance-apply'),width:document.documentElement.scrollWidth};})()`);
   assert.ok(geometry.left>=0&&geometry.right<=width+1&&geometry.top>=0&&geometry.bottom<=height+1&&geometry.footerBottom<=height&&geometry.hit&&geometry.scrollable&&geometry.width<=width,JSON.stringify({host,width,geometry}));
   if(width<741)assert.ok(geometry.previewBottom<geometry.footerTop,'Preview hides controls');
   assert.equal((await evaluate('window.storageRead()')).brightness,100,'Preview saved without Apply');
   const live=await evaluate(`(()=>{const s=getComputedStyle(document.querySelector('.assistant .conversation-reading')),p=getComputedStyle(document.querySelector('.appearance-preview .conversation-reading'));return {live:s.backgroundColor,preview:p.backgroundColor,foreground:s.color};})()`);
   assert.equal(live.live,live.preview,JSON.stringify(live));
   assert.ok(await evaluate("getComputedStyle(document.querySelector('.user .message-text')).color!==getComputedStyle(document.querySelector('.assistant .message-text')).color"),'User and assistant text must be visually distinct');
   assert.ok(await evaluate("['.message-stream-shell','.message-stream','.appearance-preview-messages'].every(selector=>getComputedStyle(document.querySelector(selector)).backgroundColor==='rgba(0, 0, 0, 0)')"),'Wallpaper must remain visible through empty canvas and gaps');
   const frost=await evaluate(`(()=>{const s=getComputedStyle(document.querySelector('.composer'));return {background:s.backgroundColor,filter:s.backdropFilter,border:s.borderTopWidth};})()`);
   assert.ok(frost.background.endsWith('0.72)') && frost.filter.includes('blur(20px)') && frost.border==='1px',JSON.stringify({host,width,frost}));
   await shot(`${host}-${width}x${height}`);
   await evaluate("document.querySelector('.appearance-footer button').click()");await wait("!document.querySelector('.appearance-dialog')");
   assert.equal(await evaluate('window.themeApi.settings.brightness'),100);
   const reading=await evaluate(`(()=>{const main=document.querySelector('.message-stream'),box=document.querySelector('.composer-box').getBoundingClientRect(),send=document.querySelector('.send-button').getBoundingClientRect();return {width:document.documentElement.scrollWidth,scrollWidth:main.scrollWidth,clientWidth:main.clientWidth,bottom:box.bottom,sendRight:send.right};})()`);
   assert.ok(reading.width<=width&&reading.scrollWidth<=reading.clientWidth+1&&reading.bottom<=height+1&&reading.sendRight<=width,JSON.stringify({host,width,reading}));
   const layout=await evaluate(`(()=>{const stream=document.querySelector('.message-stream'),origin=stream.getBoundingClientRect();const selectors=['.message-row.user .message-content','.message-row.assistant .message-content','.task-progress-timeline','.clarification-card','.recovery-terminal-record'];return {padding:parseFloat(getComputedStyle(stream).paddingLeft),cards:selectors.map(selector=>{const el=document.querySelector(selector);if(!el)return {selector,missing:true};const r=el.getBoundingClientRect();return {selector,left:r.left-origin.left,top:r.top-origin.top+stream.scrollTop,width:r.width};})};})()`);
   const expectedGutter=width<=420?12:width>=760?Math.max(24,(width-800)/2):18;
   assert.equal(layout.padding,expectedGutter,`${host} ${width}: preserve Excel's responsive message gutter`);
   assert.ok(layout.cards.every(card=>!card.missing),JSON.stringify(layout));
   const copyPositions=await evaluate("[...document.querySelectorAll('.conversation-actions')].map(actions=>{const content=actions.parentElement.querySelector('.conversation-reading'),a=actions.getBoundingClientRect(),c=content.getBoundingClientRect(),b=actions.querySelector('button').getBoundingClientRect();return {outside:!content.contains(actions),gap:a.top-c.bottom,inside:b.left>=a.left&&b.right<=a.right};})");
   assert.ok(copyPositions.length===2&&copyPositions.every(c=>c.outside&&c.gap>=5&&c.inside),JSON.stringify({host,width,copyPositions}));

   const replyWidths=await evaluate("[...document.querySelectorAll('.message-stream .conversation-message:not(.user) > .conversation-reading')].map(el=>{const r=el.getBoundingClientRect(),row=el.parentElement.getBoundingClientRect();return {left:r.left,width:r.width,rowWidth:row.width};})");
   assert.ok(replyWidths.every(r=>Math.abs(r.left-replyWidths[0].left)<1&&Math.abs(r.width-replyWidths[0].width)<1&&Math.abs(r.width-r.rowWidth)<1),JSON.stringify({host,width,replyWidths}));
   assert.ok(await evaluate("(()=>{const el=document.querySelector('.assistant .conversation-reading'),s=getComputedStyle(el),p=getComputedStyle(el,'::before');return s.boxShadow==='none'&&s.maskImage==='none'&&(p.content==='none'||p.content==='normal');})()"),'No feathered overlays or glow may return');
   for(const running of [true,false]){
    await evaluate(`window.setProbeRunning(${running})`);await frame();
    const jump=await evaluate("(()=>{const b=document.querySelector('.new-progress-button'),r=b.getBoundingClientRect(),p=b.parentElement.getBoundingClientRect();return {delta:Math.abs((r.left+r.right-p.left-p.right)/2),inside:r.top>=p.top&&r.bottom<=p.bottom,mode:b.dataset.mode};})()");
    assert.ok(jump.delta<1&&jump.inside&&jump.mode===(running?'running':'idle'),JSON.stringify({host,width,running,jump}));
   }

   assert.ok(await evaluate("['.assistant .message-text','.task-progress-timeline','.clarification-card'].every(selector=>{const s=getComputedStyle(document.querySelector(selector));return s.borderTopWidth==='0px'&&s.boxShadow==='none'&&s.backgroundColor==='rgba(0, 0, 0, 0)';})"),'Conversation sections must not grow independent cards');
   if(host==='excel')messageGeometry.set(width+'x'+height,layout);
   else assert.deepEqual(layout,messageGeometry.get(width+'x'+height),`WPS message geometry differs from Excel at ${width}x${height}`);
   await evaluate("document.querySelector('.composer-box textarea').focus()");await frame();
   const focus=await evaluate(`(()=>{const input=document.querySelector('.composer-box textarea'),box=document.querySelector('.composer-box'),i=getComputedStyle(input),b=getComputedStyle(box),r=input.getBoundingClientRect(),outer=box.getBoundingClientRect();return {active:document.activeElement===input,outline:i.outlineStyle,shadow:b.boxShadow,border:b.borderColor,expectedBorder:getComputedStyle(document.documentElement).getPropertyValue('--color-border-strong').trim(),inside:r.left>=outer.left&&r.right<=outer.right&&r.top>=outer.top&&r.bottom<=outer.bottom};})()`);
   assert.ok(focus.active&&focus.outline==='none'&&focus.shadow==='none'&&focus.inside,JSON.stringify({host,width,focus}));
   await shot(`${host}-composer-focus-${width}x${height}`);
   await evaluate("document.querySelector('.composer-app-trigger').focus()");await frame();
   assert.notEqual(await evaluate("getComputedStyle(document.activeElement).outlineStyle"),'none','Keep keyboard focus on toolbar buttons');
   await evaluate('document.activeElement.blur()');
   await shot(`${host}-reading-${width}x${height}`);
   await evaluate("document.querySelector('.message-stream').scrollTop=document.querySelector('.message-stream').scrollHeight");await frame();
   await shot(`${host}-result-${width}x${height}`);
   await evaluate("document.querySelector('.message-stream').scrollTop=0");
  }
  await send('Emulation.setDeviceMetricsOverride',{width:1000,height:760,deviceScaleFactor:1,mobile:false});
  await evaluate('window.openAppearance()');await wait('window.themeApi.editing');
  await evaluate("document.querySelector('.appearance-position').scrollIntoView({block:'center'})");await frame();
  const crop=await evaluate("(()=>{const r=document.querySelector('.appearance-position').getBoundingClientRect();return {x:r.left,y:r.top,w:r.width,h:r.height};})()");
  await send('Input.dispatchMouseEvent',{type:'mousePressed',x:crop.x+crop.w*.25,y:crop.y+crop.h*.4,button:'left',clickCount:1});
  await send('Input.dispatchMouseEvent',{type:'mouseMoved',x:crop.x+crop.w*.75,y:crop.y+crop.h*.7,button:'left',buttons:1});
  await send('Input.dispatchMouseEvent',{type:'mouseReleased',x:crop.x+crop.w*.75,y:crop.y+crop.h*.7,button:'left',clickCount:1});await frame();
  assert.ok(await evaluate('window.themeApi.settings.positionX>65 && window.themeApi.settings.positionY>60'),'Pointer drag did not update wallpaper position');
  await evaluate(`(()=>{const original=IDBDatabase.prototype.transaction;IDBDatabase.prototype.transaction=function(stores,...args){const tx=original.call(this,stores,...args);if(Array.isArray(stores)&&stores.length===2){IDBDatabase.prototype.transaction=original;queueMicrotask(()=>tx.abort());}return tx;};window.themeApi.updateSettings({brightness:60});})()`);await frame();
  await evaluate("document.querySelector('.appearance-apply').click()");await wait("!!document.querySelector('.appearance-error')");
  assert.equal((await evaluate('window.storageRead()')).brightness,100,'Aborted transaction changed saved settings');
  assert.equal(await evaluate('window.imageRead().then(image=>image.size>0)'),true,'Aborted transaction lost the wallpaper');
  assert.ok(await evaluate('window.themeApi.editing'),'Failed save closed the preview');
  await evaluate("document.querySelector('.appearance-footer button').click()");await wait("!document.querySelector('.appearance-dialog')");
  await evaluate('window.openAppearance()');await wait('window.themeApi.editing');
  await evaluate("document.querySelectorAll('.appearance-tabs button')[2].click()");await frame();
  await evaluate("[...document.querySelectorAll('.appearance-segment button')].find(b=>b.textContent==='深色').click()");await frame();
  assert.equal(await evaluate('document.documentElement.dataset.ebAppearance'),'dark');
  await shot(`${host}-dark`);
  await evaluate('window.themeApi.updateSettings({cardOpacity:82,zoom:130})');await frame();
  await evaluate("document.querySelector('.appearance-apply').click()");await wait("!document.querySelector('.appearance-dialog')");
  assert.equal((await evaluate('window.storageRead()')).cardOpacity,82);
  await evaluate('window.remount()');await wait('window.themeApi.ready && window.themeApi.settings.zoom===130');
  assert.equal(await evaluate('document.documentElement.dataset.ebAppearance'),'dark');
  assert.equal(await evaluate('window.imageRead().then(image=>image.size>0)'),true);
  // Validate computed text against the final composited background, including
  // worst-case black/white wallpaper. A blur is never counted as contrast.
  await evaluate("document.querySelector('.task-progress-toggle').click()");await frame();
  for(const mode of ['light','dark']){
   for(const preset of ['default','warm-orange','calm-blue','vivid-green']){
    await evaluate('window.themeApi.beginEdit()');await frame();
    await evaluate(`window.themeApi.updateSettings({colorMode:${JSON.stringify(mode)},preset:${JSON.stringify(preset)},cardOpacity:55,accentColor:'#ffeeaa'})`);await frame();
    const audit=await evaluate(`(()=>{
     const rgb=s=>(s.match(/[\\d.]+/g)||[]).map(Number);
     const lum=c=>c.slice(0,3).map(v=>{v/=255;return v<=.04045?v/12.92:((v+.055)/1.055)**2.4;}).reduce((n,v,i)=>n+v*[.2126,.7152,.0722][i],0);
     const selectors=['.assistant .message-text','.clarification-reason','.clarification-question','.recovery-terminal-record strong','.recovery-terminal-summary','.activity-step-label','.activity-step-time','.activity-step-detail-text','.task-progress-toggle','.user .message-text','.process-result-summary [data-result-line=next]','.execution-result-summary','.execution-result-heading strong'];
     return selectors.map(selector=>{const el=document.querySelector(selector);if(!el)return {selector,missing:true};const cs=getComputedStyle(el),ink=rgb(cs.color);let minimum=100;
      for(const wall of [0,255]){let bg=[wall,wall,wall];const chain=[];for(let n=el;n;n=n.parentElement)chain.unshift(n);for(const n of chain){const styles=[getComputedStyle(n)];const before=getComputedStyle(n,'::before');if(before.content!=='none'&&before.content!=='normal')styles.push(before);for(const style of styles){const c=rgb(style.backgroundColor),a=c[3]??1;if(c.length>=3)bg=bg.map((v,i)=>c[i]*a+v*(1-a));}}const a=lum(ink),b=lum(bg);minimum=Math.min(minimum,(Math.max(a,b)+.05)/(Math.min(a,b)+.05));}
      return {selector,contrast:minimum,size:parseFloat(cs.fontSize),color:cs.color,parentBg:getComputedStyle(el.parentElement).backgroundColor,transition:getComputedStyle(el.parentElement).transition};
     });})()`);
    assert.ok(audit.every(item=>!item.missing&&item.contrast>=4.5&&item.size>=13),JSON.stringify({host,mode,preset,audit}));
    assert.equal(audit.find(item=>item.selector==='.user .message-text').color,'rgb(255, 255, 255)','User messages keep white text in every theme');
   }
   await evaluate("document.querySelector('.message-stream').scrollTop=0;document.querySelector('.task-progress-content-body').scrollTop=0");await frame();
   await shot(`${host}-reading-expanded-${mode}`);
  }
  await evaluate('window.themeApi.cancelEdit()');
  console.log(`${host}: shared message geometry, seven viewport sizes, transparent wallpaper canvas, pointer drag, preview/live styles, cancel, transaction abort, apply persisted remount and 8 theme contrast combinations passed`);
 }
} finally {socket?.close();child.kill();}
