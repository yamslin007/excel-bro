// Offline rendered-layout regression. No real model, Office host or user storage.
// Usage: node scripts/check_taskpane_layout.mjs [path/to/playwright/package]
import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { build } from "esbuild";

const require = createRequire(import.meta.url);
const { chromium } = require(process.argv[2] || "playwright");
const root = path.resolve(import.meta.dirname, "..");
const label = "Layout-Test-Model-" + "long-name-".repeat(12);
const wps = process.argv.includes('--wps');
const source = `
import React from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import WpsAgentPanel from '../../wps-addin/src/WpsAgentPanel';
import { TaskProgressTimeline } from './components/TaskProgressTimeline';
import { LocalMappingResolution } from './components/LocalMappingResolution';
createRoot(document.getElementById('root')).render(${wps ? '<WpsAgentPanel selectedSheetIndexes={[]} />' : '<App />'});
window.mountLayoutFixture = () => {
  const fixture = document.createElement('div');
  fixture.className = 'message-row assistant layout-fixture';
  document.querySelector('.message-stream').appendChild(fixture);
  createRoot(fixture).render(<div className="message-content">
    <TaskProgressTimeline live log={{ totalMs: 12300, steps: [
      {id:'local', label:'已准备规划请求', elapsedMs:120, status:'completed', phase:'intent'},
      {id:'planning.understand', label:'已理解需求与授权范围', elapsedMs:0, status:'completed',
       phase:'planning', trace:{id:'trace', kind:'understanding', summary:'LongUnbrokenRange'.repeat(30),
       basis:'仅使用已授权元数据'}},
      {id:'planning.model.1', label:'正在生成操作方案', elapsedMs:1100, status:'started', phase:'planning'}
    ]}} />
    <section className="clarification-card">
      <div className="clarification-heading"><div><span>执行前确认</span><strong>本地映射需要你决定</strong></div>
        <em>{'LongSheetName'.repeat(10)}!C2:C6 → D2:D6</em></div>
      <p className="clarification-question">源区域有一个未覆盖值，请选择整格映射或拆分与优先顺序。</p>
      <LocalMappingResolution rule={{kind:'columnMapping', sourceSheet:'样本', targetSheet:'样本',
        sourceRange:'C2:C6', targetRange:'D2:D6', mapping:[{source:'甲',target:'高'},{source:'乙',target:'低'}]}}
        values={['甲，乙', 'LongSourceValue'.repeat(30)]} allowExactMapping busy={false}
        onResolve={async () => {}} onCancel={() => {}} />
    </section>
  </div>);
};`;
const bundled = await build({
  stdin: { contents: source, resolveDir: path.join(root, "apps/excel-addin/src"), loader: "tsx" },
  bundle: true, write: false, format: "iife", platform: "browser", jsx: "automatic",
  alias: { "@excel-bro/core": path.join(root, "packages/core/src") },
  define: { "import.meta.env": "{}", "process.env.NODE_ENV": '"production"' }
});
let css = await readFile(path.join(root, "apps/excel-addin/src/styles.css"), "utf8");
if (wps) css += await readFile(path.join(root, "apps/wps-addin/src/styles.css"), "utf8");
// Test-only mutation: do not edit production CSS to prove the overflow check.
if (process.argv.includes('--without-containment')) {
  css = css.split('/* Task-pane width containment */')[0]
    .replace('grid-template-columns: minmax(0, 1fr);', 'grid-template-columns: auto;');
}
if (process.argv.includes('--force-header-wrap')) {
  css += '.chat-header .header-actions { flex-basis: 100%; }';
}
const html = '<div id="root"></div><style>' + css + '</style><script>' + bundled.outputFiles[0].text + '</script>';
const server = createServer((_req, res) => { res.setHeader("Content-Type", "text/html; charset=utf-8"); res.end(html); });
await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
let browser;
async function assertFits(page, description) {
  const failures = await page.evaluate(() => {
    const selectors = ['.chat-shell','.chat-header','.brand-copy strong','.header-actions button','.model-picker-trigger','.model-menu',
      '.layout-fixture','.task-progress-timeline','.activity-step-body','.activity-step-trace',
      '.clarification-card','.local-mapping-resolution input','.local-mapping-resolution button','.composer'];
    return selectors.flatMap((selector) => Array.from(document.querySelectorAll(selector)).flatMap((element) => {
      const box = element.getBoundingClientRect();
      return box.width && (box.left < -1 || box.right > innerWidth + 1 || element.scrollWidth > element.clientWidth + 1)
        ? [{selector,left:box.left,right:box.right,scroll:element.scrollWidth,client:element.clientWidth}] : [];
    }));
  });
  assert.deepEqual(failures, [], `${description}: clipped content`);
}
try {
  browser = await chromium.launch({ channel: "msedge", headless: true });
  for (const width of [320, 350, 400, 430, 500, 620, 700, 900]) {
    const context = await browser.newContext({ viewport: { width, height: 1000 } });
    await context.route('http://127.0.0.1:8765/**', (route) => route.fulfill({
      contentType: 'application/json', headers: { 'Access-Control-Allow-Origin': '*' },
      body: JSON.stringify(route.request().url().endsWith('/api/models') ? {
        models:[{id:'connection:layout',label,provider:'model',available:true,supportsVision:false}],
        defaultModelId:'connection:layout'
      } : {status:'ok',configured:true,mode:'model',model:'layout'})
    }));
    const page = await context.newPage();
    page.on('pageerror', (error) => console.error('render error:', error.message));
    await page.goto(`http://127.0.0.1:${server.address().port}`);
    await page.locator('.brand-copy').waitFor({state:'attached'});
    try {
      await page.waitForFunction(() => document.querySelector('.model-picker-trigger')?.title.includes('Layout-Test-Model'), null, {timeout: 5000});
    } catch (error) {
      console.error('mock header:', await page.locator('.brand-copy').innerText());
      throw error;
    }
    await assertFits(page, `${width}px actual header`);
    assert.equal(await page.locator('.brand-copy').innerText(), 'Excel Bro', `${width}px no model line`);
    const headerRows = await page.locator('.chat-header').evaluate(header => {
      const centers = [...header.querySelectorAll('.brand-mark,.brand-copy,.model-picker-trigger,.header-actions button')]
        .map(item => { const box = item.getBoundingClientRect(); return box.top + box.height / 2; });
      return Math.max(...centers) - Math.min(...centers);
    });
    if (width >= 400) assert.ok(headerRows <= 1, `${width}px header wraps despite available space: ${headerRows}`);
    if (width === 320) assert.ok(headerRows > 1, '320px header should wrap instead of clipping or hiding controls');
    const directoryGuide = page.getByRole('button', { name: '以后再设置' });
    if (await directoryGuide.count()) await directoryGuide.click();
    for (const [selector, glyph] of [
      ['.model-picker-trigger', ''], ['.new-chat-entry', '＋'],
      ['.tools-entry', ''], ['.folder-backup-entry', '↶']
    ]) {
      const button = page.locator('.chat-header ' + selector);
      assert.equal(await button.textContent(), glyph, `${width}px ${selector}: no text label`);
      assert.ok(await button.getAttribute('aria-label'), `${selector}: accessible name`);
      assert.ok(await button.getAttribute('title'), `${selector}: hover help`);
      const box = await button.boundingBox();
      assert.equal(box.width, 32, `${width}px ${selector}: compact width`);
      assert.equal(box.height, 32, `${width}px ${selector}: compact height`);
    }
    const modelButton = page.getByRole('button', {name:'选择或添加模型', exact:true});
    const centerOffset = await modelButton.evaluate(button => {
      const outer = button.getBoundingClientRect();
      const icon = button.querySelector('svg').getBoundingClientRect();
      return {x:(icon.left+icon.right-outer.left-outer.right)/2,
        y:(icon.top+icon.bottom-outer.top-outer.bottom)/2};
    });
    assert.ok(Math.abs(centerOffset.x) <= 0.5 && Math.abs(centerOffset.y) <= 0.5,
      `${width}px model glyph off-center: ${JSON.stringify(centerOffset)}`);
    await modelButton.click();
    assert.equal(await modelButton.getAttribute('aria-expanded'), 'true');
    assert.ok((await page.getByRole('menuitemradio', {checked:true}).textContent()).includes(label),
      'Selected model remains verifiable in menu');
    await assertFits(page, `${width}px open model menu`);
    await page.getByRole('menuitemradio').first().click();
    assert.equal(await modelButton.getAttribute('aria-expanded'), 'false');
    await page.evaluate(() => window.mountLayoutFixture());
    await page.locator('.local-mapping-resolution').waitFor();
    for (const mode of ['exact', 'split']) {
      if (mode === 'split') await page.getByRole('button', { name: '改为说明拆分规则' }).click();
      await assertFits(page, `${width}px ${mode}`);
    }
    console.log(`${wps ? 'WPS' : 'Excel'} ${width}px: actual header/composer + shared timeline/clarification fit (exact/split)`);
    await context.close();
  }
} finally {
  await browser?.close();
  await new Promise((resolve) => server.close(resolve));
}
