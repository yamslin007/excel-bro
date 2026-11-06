// Offline browser check: shared new-conversation button + both scope-row structures.
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import net from 'node:net';
import { build } from 'esbuild';

const root = path.resolve(import.meta.dirname, '..');
const out = path.join(root, 'outputs/probe/scope-entry');
mkdirSync(out, { recursive: true });
const browser = [process.env.PROBE_BROWSER,
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'].filter(Boolean).find(existsSync);
assert.ok(browser, 'Set PROBE_BROWSER to a Chromium executable');
const source = `
import React, {useState} from 'react';
import {createRoot} from 'react-dom/client';
import {ConversationUtilityButtons} from './components/ConversationUtilityButtons';
import {NewConversationButton} from './components/NewConversationButton';
import {ModelPickerLabel} from './components/ModelPickerLabel';
import {ComposerPrimaryButton} from './components/ComposerPrimaryButton';
import {ComposerPendingBar} from './components/ComposerPendingBar';
import {ComposerAppMenu} from './components/ComposerAppMenu';
import {brandLogoUrl} from './brandLogo';
import PetCompanion from './PetCompanion';
function AppMenu(){
 const [open,setOpen]=useState(false);
 return <ComposerAppMenu logoSrc={brandLogoUrl} open={open} onToggle={()=>setOpen(!open)} onClose={()=>setOpen(false)}>{['工具','宠物','壁纸','EB 函数','专注窗口','加宽对话窗格'].map(name=><button key={name} type="button" role="menuitem" onClick={()=>setOpen(false)}><strong>{name}</strong><span>功能入口</span></button>)}</ComposerAppMenu>;
}
function ModelPicker(){
  const [model,setModel]=useState(null),[open,setOpen]=useState(false);
  return <div className="model-picker"><button type="button" className="model-picker-trigger" onClick={()=>setOpen(!open)}><ModelPickerLabel model={model}/></button>{open&&<div className="model-menu" role="menu"><button role="menuitemradio" onClick={()=>{setModel({id:'connection:fixture',modelId:'long-model-id-for-narrow-taskpane-preview-2026',label:'连接昵称'});setOpen(false);}}><i/><span>选择测试模型</span><b>✓</b></button><button data-deepseek role="menuitemradio" onClick={()=>{setModel({id:'connection:ds',modelId:'deepseek-v4-flash',label:'DS · deepseek-v4-flash'});setOpen(false);}}>DS · deepseek-v4-flash</button></div>}</div>;
}
const wps=new URLSearchParams(location.search).has('wps');
const picker=<button className="scope-trigger" onClick={()=>window.scopeClicks++} aria-label="选择数据范围"><i aria-hidden="true">＋</i><span><small title="公式来源证据验收-很长的工作簿文件名称.xlsx">公式来源证据验收-很长的工作簿文件名称.xlsx</small><span className="scope-sep">|</span><strong>当前表 结果3</strong></span><b>⌃</b></button>;
window.newChats=0;window.scopeClicks=0;window.backups=0;window.historyOpens=0;
// 常驻摘要用非 decision 类（默认收起），点击才展开 —— 与真实 pendingNodes 的默认姿态一致。
const pendingItems=[{id:'recovery',label:'恢复待办 · 还剩 1 步',kind:'recovery'},{id:'status',label:'服务状态',kind:'status'},{id:'notice',label:'历史记录提示',kind:'notice'}];
const notices=<>
  <p className="composer-service-status" role="status">服务已就绪</p>
  <p className="composer-service-status" role="status">历史记录已从云端恢复</p>
  <p className="composer-service-status" role="status" data-long>{'很长的一句状态提示，用来把详情撑高，验证底部容器先让位而不是把输入框顶出视口。'.repeat(8)}</p>
</>;
createRoot(document.getElementById('root')).render(<div className={'chat-shell chat-shell-headerless'+(wps?' wps-chat-shell':'')}>
<main className="message-stream">{Array.from({length:24},(_,i)=><div key={i} className={'conversation-message '+(i===23?'assistant':'system')} data-row={i}><div className="conversation-reading"><p className="message-text">第 {i+1} 条消息，用来把消息区撑到溢出。</p></div></div>)}<div id="last-row" style={{height:24}}>最后一条</div></main>
<footer className="composer"><ComposerPendingBar items={pendingItems}>{notices}</ComposerPendingBar><div className="composer-dock"><div className="scope-entry-row scope-entry-row-with-new-chat"><NewConversationButton disabled={false} onClick={()=>window.newChats++}/><ConversationUtilityButtons disabled={false} backupCount={2} onBackup={()=>window.backups++} onHistory={()=>window.historyOpens++}/>{wps?<div className="scope-entry-target"><section className="wps-scope-dock"><div className="scope-entry-row">{picker}</div></section></div>:picker}</div><div className="composer-input-region"><div className="composer-box"><textarea rows={1} placeholder="描述你想查询、分析或修改的内容…"/><div className="composer-toolbar"><div className="composer-tools-left"><AppMenu/></div><span className="composer-shortcut-hint">Enter 发送</span><div className="composer-tools-right"><ModelPicker/><ComposerPrimaryButton className="send-button composer-send-button" action={{kind:"send",glyph:"↑",label:"发送",title:"发送这条需求",disabled:false,busy:false,invoke:"send"}} onInvoke={()=>{window.sent=(window.sent||0)+1;}}/></div></div></div></div></div></footer><PetCompanion busy={false}/></div>);`;
const bundle=await build({stdin:{contents:source,resolveDir:path.join(root,'apps/excel-addin/src'),loader:'tsx'},bundle:true,loader:{'.png':'dataurl','.svg':'dataurl'},plugins:[{name:'fixture-css',setup(build){build.onLoad({filter:/\.css$/},()=>({contents:'',loader:'js'}));}}],write:false,format:'iife',define:{'process.env.NODE_ENV':'"production"'}});
// 样式按真实导入顺序拼（`main.tsx` → chatScrollbars → appearance → conversation；
// WPS 的 styles.css 插在 excel 的之后）。**新组件的 css 必须加进这里**，否则量到的不是线上那份。
const readCss=relative=>readFileSync(path.join(root,relative),'utf8');
const baseCss=readCss('apps/excel-addin/src/styles.css')
  +readCss('apps/excel-addin/src/components/composerPrimaryButton.css');
const lateCss=readCss('apps/excel-addin/src/chatScrollbars.css')
  +readCss('apps/excel-addin/src/appearance.css')
  +readCss('apps/excel-addin/src/conversation.css');
const wpsCss=readCss('apps/wps-addin/src/styles.css');
const file=path.join(out,'fixture.html');
writeFileSync(file,`<!doctype html><meta charset="utf-8"><style>html,body,#root{height:100%;margin:0;}*{scroll-behavior:auto !important;}</style><style id="base-style">${baseCss}</style><style id="wps-style"></style><style id="late-style">${lateCss}</style><div id="root"></div><script>${bundle.outputFiles[0].text}</script>`);
const port=await new Promise(resolve=>{const server=net.createServer();server.listen(0,'127.0.0.1',()=>{const p=server.address().port;server.close(()=>resolve(p));});});
const url=`file:///${file.replaceAll('\\','/')}`;
const child=spawn(browser,['--headless=new','--no-first-run','--disable-extensions',`--remote-debugging-port=${port}`,`--user-data-dir=${path.join(out,'profile')}`,url],{stdio:'ignore',windowsHide:true});
let socket;
try {
  let target;
  for(let i=0;i<100;i++){try{target=(await(await fetch(`http://127.0.0.1:${port}/json/list`)).json()).find(t=>t.type==='page');if(target)break;}catch{}await new Promise(r=>setTimeout(r,200));}
  assert.ok(target,'Browser startup timed out');
  socket=new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((resolve,reject)=>{socket.addEventListener('open',resolve,{once:true});socket.addEventListener('error',reject,{once:true});});
  let id=0;const pending=new Map();
  socket.addEventListener('message',e=>{const reply=JSON.parse(e.data);if(pending.has(reply.id)){pending.get(reply.id)(reply);pending.delete(reply.id);}});
  const send=(method,params={})=>new Promise(resolve=>{pending.set(++id,resolve);socket.send(JSON.stringify({id,method,params}));});
  const evaluate=async expression=>{const reply=await send('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});assert.ok(!reply.result.exceptionDetails,JSON.stringify(reply));return reply.result.result.value;};
  for(const host of ['excel','wps']){
    await send('Page.navigate',{url:url+(host==='wps'?'?wps':'')});
    let ready=false;
    for(let i=0;i<100;i++){if(await evaluate("!!document.querySelector('.scope-new-chat')")){ready=true;break;}await new Promise(r=>setTimeout(r,100));}
    assert.ok(ready,'Fixture failed to mount');
    if(host==='wps')await evaluate(`document.getElementById('wps-style').textContent=${JSON.stringify(wpsCss)}`);
    for(const width of [320,400,500]){
      await send('Emulation.setDeviceMetricsOverride',{width,height:640,deviceScaleFactor:1,mobile:false});
      await evaluate('new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)))');
      const measured=await evaluate(`(()=>{const button=document.querySelector('.scope-new-chat'),scope=document.querySelector('.scope-trigger'),a=button.getBoundingClientRect(),b=scope.getBoundingClientRect(),name=scope.querySelector('small');return {topDelta:Math.abs(a.top-b.top),heightDelta:Math.abs(a.height-b.height),left:a.left,buttonRight:a.right,scopeLeft:b.left,scopeRight:b.right,bottom:b.bottom,pageWidth:document.documentElement.scrollWidth,buttonFits:button.scrollWidth<=button.clientWidth,nameTruncated:name.scrollWidth>name.clientWidth};})()`);
      assert.ok(measured.topDelta<=1&&measured.heightDelta<=1,JSON.stringify(measured));
      assert.ok(measured.left>=0&&measured.buttonRight<measured.scopeLeft&&measured.scopeRight<=width&&measured.bottom<=640&&measured.pageWidth<=width&&measured.buttonFits,JSON.stringify(measured));
      if(width===320)assert.ok(measured.nameTruncated,'Long filename must truncate instead of moving the new-chat button');
      const utilities=await evaluate(`(()=>{const selectors=['.scope-new-chat','.folder-backup-entry','.history-entry','.scope-trigger'];return selectors.map(s=>{const r=document.querySelector(s).getBoundingClientRect();return {left:r.left,right:r.right,top:r.top};});})()`);
      assert.ok(utilities.every((r,i)=>i===0||(r.left>utilities[i-1].right&&Math.abs(r.top-utilities[0].top)<=1)),JSON.stringify(utilities));
      assert.equal(await evaluate("document.querySelector('.chat-header')"),null);
      assert.equal(await evaluate("document.querySelector('.message-stream').getBoundingClientRect().top"),0);
      // 底部两层结构：工具栏与输入框在 `.composer-dock` 里紧贴，常驻摘要整个在它上方。
      const dock=await evaluate(`(()=>{const dock=document.querySelector('.composer-dock'),pending=document.querySelector('.composer-pending'),row=dock.querySelector('.scope-entry-row'),region=dock.querySelector('.composer-input-region'),send=document.querySelector('.composer-primary');const kids=[...dock.children];const a=row.getBoundingClientRect(),b=region.getBoundingClientRect(),p=pending.getBoundingClientRect(),s=send.getBoundingClientRect();return {kids:kids.length,first:kids[0].className,second:kids[1].className,rowBottom:a.bottom,regionTop:b.top,pendingBottom:p.bottom,dockTop:dock.getBoundingClientRect().top,sendLeft:s.left,sendRight:s.right,sendTop:s.top,sendBottom:s.bottom,pageWidth:document.documentElement.scrollWidth};})()`);
      assert.equal(dock.kids,2,JSON.stringify(dock));
      assert.ok(dock.first.includes('scope-entry-row')&&dock.second.includes('composer-input-region'),JSON.stringify(dock));
      assert.ok(Math.abs(dock.rowBottom-dock.regionTop)<=1,`工具栏与输入框之间不许夹别的东西: ${JSON.stringify(dock)}`);
      assert.ok(dock.pendingBottom<=dock.dockTop+0.5,`常驻摘要必须在输入容器上方: ${JSON.stringify(dock)}`);
      assert.ok(dock.sendLeft>=0&&dock.sendRight<=width&&dock.sendTop>=0&&dock.sendBottom<=640,`${width}px 发送入口越界: ${JSON.stringify(dock)}`);
      assert.ok(dock.pageWidth<=width,`${width}px 横向溢出: ${JSON.stringify(dock)}`);
      // 撑开长待办详情后，输入区仍然完整可见（让位的是摘要，不是输入框）。
      await evaluate("document.querySelector('.composer-pending-summary').click();new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)))");
      const openState=await evaluate(`(()=>{const pending=document.querySelector('.composer-pending'),details=document.querySelector('.composer-pending-details'),send=document.querySelector('.composer-primary'),p=pending.getBoundingClientRect(),d=details.getBoundingClientRect(),s=send.getBoundingClientRect();return {open:pending.dataset.pendingOpen,summaryTop:p.top,detailsScrolls:details.scrollHeight>details.clientHeight,detailsBottom:d.bottom,sendTop:s.top,sendBottom:s.bottom,pageWidth:document.documentElement.scrollWidth};})()`);
      assert.equal(openState.open,'true',JSON.stringify(openState));
      assert.ok(openState.detailsScrolls,`长详情必须自己滚: ${JSON.stringify(openState)}`);
      assert.ok(openState.summaryTop>=0,`摘要行被顶出视口: ${JSON.stringify(openState)}`);
      assert.ok(openState.sendTop>=0&&openState.sendBottom<=640,`展开待办后发送入口越界: ${JSON.stringify(openState)}`);
      assert.ok(openState.pageWidth<=width,`展开待办后横向溢出: ${JSON.stringify(openState)}`);
      await evaluate("document.querySelector('.composer-pending-summary').click();new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)))");
      assert.equal(await evaluate("document.querySelector('.composer-pending').dataset.pendingOpen"),'false');
      // 消息区独立滚动：滚到底后最后一条不被底部容器盖住。
      const reach=await evaluate(`(()=>{const stream=document.querySelector('.message-stream');stream.scrollTop=stream.scrollHeight;const last=document.querySelector('#last-row').getBoundingClientRect(),box=stream.getBoundingClientRect();return {scrollable:stream.scrollHeight>stream.clientHeight,lastBottom:last.bottom,boxBottom:box.bottom};})()`);
      assert.ok(reach.scrollable,`消息区没有独立滚动: ${JSON.stringify(reach)}`);
      assert.ok(reach.lastBottom<=reach.boxBottom+0.5,`最后一条消息被底部盖住: ${JSON.stringify(reach)}`);
      await evaluate("document.querySelector('.folder-backup-entry').click();document.querySelector('.history-entry').click()");
      assert.equal(await evaluate('window.backups'),[320,400,500].indexOf(width)+1);
      assert.equal(await evaluate('window.historyOpens'),[320,400,500].indexOf(width)+1);
      await evaluate("document.querySelector('.scope-new-chat').click()");
      assert.equal(await evaluate('window.scopeClicks'),0,'New chat opened scope picker');
      if(process.argv.includes('--pet-only')) {
        const pet=await evaluate(`(()=>{const img=document.querySelector('.pet-brand-logo'),r=img.getBoundingClientRect();return {loaded:img.complete&&img.naturalWidth>0,width:r.width,height:r.height,left:r.left,right:r.right,top:r.top,bottom:r.bottom,old:!!document.querySelector('.pet-jelly')};})()`);
        assert.ok(pet.loaded&&!pet.old&&pet.width===pet.height&&pet.left>=0&&pet.right<=width&&pet.top>=0&&pet.bottom<=640,JSON.stringify(pet));
        await evaluate("document.querySelector('.pet-avatar').click();new Promise(r=>requestAnimationFrame(r))");
        assert.equal(await evaluate("document.querySelector('.pet-bubble').getAttribute('aria-label')"),'EB 小提示');
        const petShot=await send('Page.captureScreenshot',{format:'png'});
        writeFileSync(path.join(out,`${host}-pet-${width}.png`),Buffer.from(petShot.result.data,'base64'));
        await evaluate("document.querySelector('.pet-avatar').click();new Promise(r=>requestAnimationFrame(r))");
        console.log(`${host} ${width}px: EB logo loaded without distortion; tips toggle works`);
        continue;
      }
      if(width===320){
        assert.equal(await evaluate("document.querySelector('.model-picker-trigger').textContent"),'模型');
        await evaluate("document.querySelector('.model-picker-trigger').click();new Promise(r=>requestAnimationFrame(r))");
        await evaluate("document.querySelector('.model-menu button').click();new Promise(r=>requestAnimationFrame(r))");
        assert.equal(await evaluate("document.querySelector('.model-picker-trigger span').textContent"),'long-model-id-for-narrow-taskpane-preview-2026');
      }
      await evaluate("document.querySelector('.model-picker-trigger').click();new Promise(r=>requestAnimationFrame(r))");
      const menu=await evaluate(`(()=>{const r=document.querySelector('.model-menu').getBoundingClientRect(),b=document.querySelector('.model-picker-trigger').getBoundingClientRect();return {left:r.left,right:r.right,top:r.top,bottom:r.bottom,buttonTop:b.top,pageWidth:document.documentElement.scrollWidth};})()`);
      assert.ok(menu.left>=0&&menu.right<=width&&menu.top>=0&&menu.bottom<menu.buttonTop&&menu.pageWidth<=width,JSON.stringify(menu));
      const controls=await evaluate(`(()=>{const button=document.querySelector('.composer-primary'),arrow=button.querySelector('.composer-send-arrow'),model=document.querySelector('.model-picker-trigger'),a=button.getBoundingClientRect(),b=model.getBoundingClientRect(),style=getComputedStyle(arrow);return {width:a.width,height:a.height,left:a.left,right:a.right,modelRight:b.right,filled:style.fill!=='none'&&style.stroke==='none',symbol:!!model.querySelector('.model-picker-symbol'),chevron:!!model.querySelector('.model-picker-chevron'),visible:document.elementFromPoint(a.left+a.width/2,a.top+a.height/2)?.closest('.composer-primary')===button};})()`);
      assert.ok(controls.width===32&&controls.height===32&&controls.right<=width&&controls.modelRight<controls.left&&controls.filled&&controls.symbol&&controls.chevron&&controls.visible,JSON.stringify(controls));
      const shot=await send('Page.captureScreenshot',{format:'png'});
      writeFileSync(path.join(out,`${host}-${width}.png`),Buffer.from(shot.result.data,'base64'));
      await evaluate("document.querySelector('.model-picker-trigger').click();new Promise(r=>requestAnimationFrame(r))");
      const toolbarShot=await send('Page.captureScreenshot',{format:'png'});
      writeFileSync(path.join(out,`${host}-toolbar-${width}.png`),Buffer.from(toolbarShot.result.data,'base64'));
      await evaluate("document.querySelector('.model-picker-trigger').click();new Promise(r=>requestAnimationFrame(r))");
      await evaluate("document.querySelector('[data-deepseek]').click();new Promise(r=>requestAnimationFrame(r))");
      assert.equal(await evaluate("document.querySelector('.model-picker-name').textContent"),'deepseek-v4-flash');
      const brand=await evaluate(`(()=>{const e=document.querySelector('[data-model-brand]'),s=getComputedStyle(e),r=e.getBoundingClientRect();return {name:e.dataset.modelBrand,mask:s.maskImage,width:r.width,height:r.height,pageWidth:document.documentElement.scrollWidth};})()`);
      assert.ok(brand.name==='DeepSeek'&&brand.mask.includes('data:image/svg+xml')&&brand.width===18&&brand.height===18&&brand.pageWidth<=width,JSON.stringify(brand));
      const brandShot=await send('Page.captureScreenshot',{format:'png'});
      writeFileSync(path.join(out,`${host}-deepseek-${width}.png`),Buffer.from(brandShot.result.data,'base64'));
      await evaluate("document.querySelector('.composer-app-trigger').click();new Promise(r=>requestAnimationFrame(r))");
      const appMenu=await evaluate(`(()=>{const menu=document.querySelector('.composer-app-menu .view-menu'),r=menu.getBoundingClientRect(),b=document.querySelector('.composer-app-trigger').getBoundingClientRect();return {left:r.left,right:r.right,top:r.top,bottom:r.bottom,buttonTop:b.top,items:menu.querySelectorAll('[role=menuitem]').length,logoLoaded:document.querySelector('.composer-app-trigger img').naturalWidth>0};})()`);
      assert.ok(appMenu.left>=0&&appMenu.right<=width&&appMenu.top>=0&&appMenu.bottom<appMenu.buttonTop&&appMenu.items===6&&appMenu.logoLoaded,JSON.stringify(appMenu));
      const appShot=await send('Page.captureScreenshot',{format:'png'});
      writeFileSync(path.join(out,`${host}-eb-menu-${width}.png`),Buffer.from(appShot.result.data,'base64'));
      await evaluate("document.querySelector('.composer-app-trigger').dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}));new Promise(r=>requestAnimationFrame(r))");
      assert.equal(await evaluate("document.querySelector('.composer-app-menu .view-menu')"),null);
      console.log(`${host} ${width}px: same row, equal height, long filename fits; independent button action`);
    }
  }
} finally {socket?.close();child.kill();}
