// Real shared component + both host styles; no host or workbook access.
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import net from 'node:net';
import { build } from 'esbuild';

const root = path.resolve(import.meta.dirname, '..');
const out = path.join(root, 'outputs/probe/conversation-navigation');
mkdirSync(out, { recursive: true });
const browser = [process.env.PROBE_BROWSER,
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'].filter(Boolean).find(existsSync);
assert.ok(browser, 'Set PROBE_BROWSER to a Chromium executable');
const source = `
import React, {useState} from 'react';
import {createRoot} from 'react-dom/client';
import ConversationNavigation from './components/ConversationNavigation';
import {useMessageAutoFollow} from './hooks/useMessageAutoFollow';
function Harness(){
 const [count,setCount]=useState(60);
 const messages=Array.from({length:count},(_,i)=>({id:'u'+i,role:'user',text:'第 '+(i+1)+' 个问题：请解释这个计算结果和对应的统计方法。'}));
 const {containerRef,handleScroll,scrollToMessage}=useMessageAutoFollow(String(count),'fixture');
 window.appendTurn=()=>setCount(c=>c+1);
 return <div className="chat-shell"><header className="chat-header">Excel Bro</header>
 <div className="message-stream-shell has-conversation-navigation">
 <main ref={containerRef} className="message-stream" onScroll={handleScroll}>
 {messages.map(m=><article className="message-row assistant" key={m.id} data-message-id={m.id} style={{minHeight:180}}><div className="message-content"><span className="message-author">Excel Bro</span><p className="message-text">「三个独立公式写入（结果1/2/3 的 B2）」已执行 3 步并通过独立验证。已记录 3 项本次执行撤销数据，并已核对执行后状态。已登记动作级断点：关闭或重载任务窗格后可以从已完成动作之后继续；续跑前先核对真实单元格状态，只续跑本次登记的动作。超长连续文本必须完整显示：abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789</p><div className="verification-card"><strong>验证通过</strong><p>当前没有可用的撤销记录；这不表示执行已撤销。</p></div><div className="activity-log-card">执行结果卡片</div></div></article>)}
 </main><ConversationNavigation messages={messages} streamRef={containerRef} onNavigate={scrollToMessage}/></div>
 <footer className="composer"><div className="composer-box"><textarea placeholder="输入消息"/><div className="composer-toolbar"><button>发送</button></div></div></footer></div>;
}createRoot(document.getElementById('root')).render(<Harness/>);`;
const bundle = await build({stdin:{contents:source,resolveDir:path.join(root,'apps/excel-addin/src'),loader:'tsx'},bundle:true,write:false,format:'iife',loader:{'.css':'empty'},define:{'process.env.NODE_ENV':'"production"'}});
const css = readFileSync(path.join(root,'apps/excel-addin/src/styles.css'),'utf8') + readFileSync(path.join(root,'apps/excel-addin/src/components/ConversationNavigation.css'),'utf8') + readFileSync(path.join(root,'apps/excel-addin/src/chatScrollbars.css'),'utf8');
const file=path.join(out,'fixture.html');
writeFileSync(file,`<!doctype html><html><meta charset="utf-8"><style>${css}html,body,#root{height:100%;margin:0}</style><div id="root"></div><script>${bundle.outputFiles[0].text}</script></html>`);
const port=await new Promise(resolve=>{const s=net.createServer();s.listen(0,'127.0.0.1',()=>{const p=s.address().port;s.close(()=>resolve(p));});});
const child=spawn(browser,[`--remote-debugging-port=${port}`,`--user-data-dir=${path.join(out,'profile')}`,'--headless=new','--no-first-run','--disable-extensions',`file:///${file.replaceAll('\\','/')}`],{stdio:'ignore',windowsHide:true});
let socket;
try {
 let target;
 for(let i=0;i<100;i++){try{target=(await(await fetch(`http://127.0.0.1:${port}/json/list`)).json()).find(t=>t.type==='page');if(target)break;}catch{}await new Promise(r=>setTimeout(r,200));}
 assert.ok(target,'Browser startup timed out');
 socket=new WebSocket(target.webSocketDebuggerUrl);
 await new Promise((resolve,reject)=>{socket.addEventListener('open',resolve,{once:true});socket.addEventListener('error',reject,{once:true});});
 let id=0; const pending=new Map();
 socket.addEventListener('message',e=>{const r=JSON.parse(e.data);if(pending.has(r.id)){pending.get(r.id)(r);pending.delete(r.id);}});
 const send=(method,params={})=>new Promise(resolve=>{pending.set(++id,resolve);socket.send(JSON.stringify({id,method,params}));});
 const evaluate=async expression=>{const r=await send('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});assert.ok(!r.result.exceptionDetails,JSON.stringify(r));return r.result.result.value;};
 const settle=()=>evaluate('new Promise(r=>setTimeout(()=>requestAnimationFrame(()=>requestAnimationFrame(r)),360))');
 for(let i=0;i<100;i++){if(await evaluate("!!document.querySelector('nav')"))break;await new Promise(r=>setTimeout(r,100));}
 for(const host of ['excel','wps']){
  if(host==='wps')await evaluate(`document.querySelector('.chat-shell').classList.add('wps-chat-shell');const s=document.createElement('style');s.textContent=${JSON.stringify(readFileSync(path.join(root,'apps/wps-addin/src/styles.css'),'utf8'))};document.head.append(s)`);
  for(const [width,height] of [[320,640],[400,720],[500,760],[800,760]]){
   await send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});await settle();
   await evaluate("document.activeElement?.blur();document.querySelector('nav').dispatchEvent(new MouseEvent('mouseout',{bubbles:true,relatedTarget:document.body}))");await settle();
   const layout=await evaluate(`(()=>{const n=document.querySelector('nav').getBoundingClientRect(),s=document.querySelector('main').getBoundingClientRect(),c=document.querySelector('footer').getBoundingClientRect();return {railLeft:n.left,railRight:n.right,streamLeft:s.left,railBottom:n.bottom,composerTop:c.top,pageWidth:document.documentElement.scrollWidth};})()`);
   assert.ok(layout.railLeft>=0 && layout.streamLeft<=layout.railLeft && layout.railRight<=width-(host==='excel'?10:14) && layout.railBottom<=layout.composerTop && layout.pageWidth<=width,JSON.stringify(layout));
   const hitTest=await evaluate(`(()=>{const nav=document.querySelector('nav'), r=nav.getBoundingClientRect(),x=2,y=(r.top+r.bottom)/2;const on=!!document.elementFromPoint(x,y)?.closest('nav');nav.style.visibility='hidden';const off=!!document.elementFromPoint(x,y)?.closest('nav');nav.style.removeProperty('visibility');return {edgeX:x,withNavigation:on,withoutNavigation:off};})()`);
   assert.equal(hitTest.withNavigation,false,JSON.stringify(hitTest));
   assert.equal(hitTest.withoutNavigation,false,JSON.stringify(hitTest));
   console.log(host,width,'webview edge hit-test:',hitTest);
   const alignment=await evaluate(`(()=>{const shell=document.querySelector('.message-stream-shell').getBoundingClientRect(), article=document.querySelector('article').getBoundingClientRect();return {delta:Math.abs((article.left+article.right-shell.left-shell.right)/2),gutter:getComputedStyle(document.querySelector('main')).scrollbarGutter};})()`);
   assert.ok(host==='excel' ? alignment.delta>=7 && alignment.delta<=9 : alignment.delta<=2.5,JSON.stringify(alignment));
   const contentFit=await evaluate(`(()=>{const stream=document.querySelector('main'),shell=document.querySelector('.message-stream-shell').getBoundingClientRect(), content=document.querySelector('.message-content'),row=content.getBoundingClientRect(),card=content.querySelector('.activity-log-card').getBoundingClientRect(),nav=document.querySelector('nav').getBoundingClientRect();return {left:row.left-shell.left,right:shell.right-row.right,cardFits:card.right<=nav.left+0.5,overflow:stream.scrollWidth>stream.clientWidth,nav:nav.left};})()`);
   assert.ok(!contentFit.overflow && contentFit.cardFits,JSON.stringify(contentFit));
   const visibleContent=await evaluate(`(()=>{const nav=document.querySelector('nav').getBoundingClientRect(),text=document.querySelector('.message-text'),card=document.querySelector('.verification-card').getBoundingClientRect(), range=document.createRange();range.selectNodeContents(text);const lines=Array.from(range.getClientRects());return {cardRight:card.right,navLeft:nav.left,textRight:Math.max(...lines.map(r=>r.right)),textLeft:Math.min(...lines.map(r=>r.left)),cardLeft:card.left};})()`);
   assert.ok(visibleContent.textRight<=visibleContent.navLeft+0.5 && visibleContent.cardRight<=visibleContent.navLeft+0.5 && visibleContent.textLeft>=visibleContent.cardLeft-0.5,JSON.stringify(visibleContent));
   if(host==='excel') assert.ok(contentFit.left>=5.5 && contentFit.left<=6.5 && contentFit.right>=21.5 && contentFit.right<=22.5,JSON.stringify(contentFit));
   console.log(host,width,'message margins',contentFit);
   for(const scale of [1.25,1.5]){
    await send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:scale,mobile:false});await settle();
    const resized=await evaluate(`(()=>{const stream=document.querySelector('main'),nav=document.querySelector('nav').getBoundingClientRect(),card=document.querySelector('.verification-card').getBoundingClientRect(),text=document.querySelector('.message-text'),range=document.createRange();range.selectNodeContents(text);return {cardRight:card.right,navLeft:nav.left,textRight:Math.max(...Array.from(range.getClientRects(),r=>r.right)),overflow:stream.scrollWidth>stream.clientWidth};})()`);
    assert.ok(!resized.overflow && resized.cardRight<=resized.navLeft+1 && resized.textRight<=resized.navLeft+1,JSON.stringify({scale,...resized}));
   }
   await send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});await settle();
   await evaluate("document.querySelector('nav').dispatchEvent(new MouseEvent('mouseover',{bubbles:true}))");await settle();
   await evaluate("document.querySelector('nav button').click()");await settle();
   const top=await evaluate("document.querySelector('main').scrollTop");
   await evaluate('window.appendTurn()');await settle();
   assert.equal(await evaluate("document.querySelector('main').scrollTop"),top,'New content stole navigation position');
   assert.equal(await evaluate("document.querySelector('nav button[aria-current]').getAttribute('aria-label').startsWith('第 1 轮')"),true);
   await evaluate("document.querySelector('nav button').blur();document.querySelector('nav button').focus()");await settle();
   const preview=await evaluate("(()=>{const r=document.querySelector('.conversation-navigation').getBoundingClientRect();return {left:r.left,right:r.right};})()");
   assert.ok(preview.left>=0 && preview.right<=width,JSON.stringify(preview));
   const scrollTest=await evaluate(`(()=>{const list=document.querySelector('.conversation-navigation-ticks'),main=document.querySelector('main');const before=main.scrollTop;list.scrollTop+=180;return {overflow:list.scrollHeight>list.clientHeight,top:list.scrollTop,before,after:main.scrollTop,expanded:document.querySelector('nav').dataset.expanded};})()`);
   assert.ok(scrollTest.overflow && scrollTest.top>0 && scrollTest.before===scrollTest.after && scrollTest.expanded==='true',JSON.stringify(scrollTest));
   const shot=await send('Page.captureScreenshot',{format:'png'});writeFileSync(path.join(out,`${host}-${width}.png`),Buffer.from(shot.result.data,'base64'));
   console.log(host,width,height,layout,'navigation/preview/follow: passed');
  }
 }
} finally {socket?.close();child.kill();}
