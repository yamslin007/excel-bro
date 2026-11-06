import assert from 'node:assert/strict';
import {build} from 'esbuild';
import {createServer} from 'node:http';
import {spawn} from 'node:child_process';
import {existsSync, mkdirSync, writeFileSync} from 'node:fs';
import path from 'node:path';
import net from 'node:net';

const root = path.resolve(import.meta.dirname, '..');
const out = path.join(root, 'outputs/probe/review-before-store');
mkdirSync(out, {recursive:true});
const browser = ['C:/Program Files/Google/Chrome/Application/chrome.exe', 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'].find(existsSync);
assert.ok(browser);
const bundle = await build({entryPoints:[path.join(root,'apps/excel-addin/src/utils/reviewBeforeStore.ts')], bundle:true, write:false, format:'iife', globalName:'ReviewStore'});
const server = createServer((_req,res) => {res.setHeader('Content-Type','text/html');res.end(`<script>${bundle.outputFiles[0].text}</script>`);});
await new Promise(resolve => server.listen(0,'127.0.0.1',resolve));
const url = `http://127.0.0.1:${server.address().port}`;
const port = await new Promise(resolve=>{const s=net.createServer();s.listen(0,'127.0.0.1',()=>{const p=s.address().port;s.close(()=>resolve(p));});});
const child = spawn(browser,['--headless=new','--no-first-run','--disable-extensions',`--remote-debugging-port=${port}`,`--user-data-dir=${path.join(out,'profile')}`,url],{stdio:'ignore',windowsHide:true});
let socket;
try {
  let targets;
  for(let i=0;i<100;i++){try{targets=await (await fetch(`http://127.0.0.1:${port}/json`)).json();if(targets.some(t=>t.url.startsWith(url)))break;}catch{}await new Promise(r=>setTimeout(r,100));}
  const target=targets.find(t=>t.url.startsWith(url)); assert.ok(target);
  socket=new WebSocket(target.webSocketDebuggerUrl); await new Promise(r=>socket.addEventListener('open',r,{once:true}));
  let id=0;const pending=new Map();socket.addEventListener('message',event=>{const m=JSON.parse(event.data);if(m.id){pending.get(m.id)?.(m);pending.delete(m.id);}});
  const send=(method,params={})=>new Promise(resolve=>{pending.set(++id,resolve);socket.send(JSON.stringify({id,method,params}));});
  const evaluate=async expression=>{const result=await send('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});assert.ok(!result.result.exceptionDetails,JSON.stringify(result));return result.result.result.value;};
  const ready=async()=>{for(let i=0;i<100;i++){try{if(await evaluate('typeof ReviewStore === "object"'))return;}catch{}await new Promise(r=>setTimeout(r,50));}throw Error('Page not ready');};
  await ready();
  for(const host of ['office','wps']) {
    assert.equal(await evaluate(`ReviewStore.saveReviewBefore('${host}','original','execution',['Target'],{values:[['header',7]],formula:'=A1'})`),true);
    assert.equal(await evaluate(`ReviewStore.loadReviewBefore('${host}','other','execution',['Target'])`),null);
    assert.equal(await evaluate(`ReviewStore.loadReviewBefore('${host}','original','execution',['Other'])`),null);
  }
  await send('Page.reload');await new Promise(r=>setTimeout(r,250));await ready();
  for(const host of ['office','wps']) {
    assert.deepEqual(await evaluate(`ReviewStore.loadReviewBefore('${host}','original','execution',['Target'])`),{values:[['header',7]],formula:'=A1'});
    await evaluate(`ReviewStore.deleteReviewBefore('${host}','original','execution')`);
    assert.equal(await evaluate(`ReviewStore.loadReviewBefore('${host}','original','execution',['Target'])`),null);
  }
  writeFileSync(path.join(out,'results.json'),JSON.stringify({passed:true,hosts:['office','wps'],checks:['save','identity boundary','scope boundary','reload','delete']},null,2));
  console.log('PASS: actual browser IndexedDB, both hosts; reload preserves snapshots, wrong workbook/scope rejected, cleanup removes only bound entry.');
} finally {socket?.close();child.kill();server.close();}
