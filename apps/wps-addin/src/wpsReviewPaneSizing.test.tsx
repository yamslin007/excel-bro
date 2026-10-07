import { act, StrictMode, type ReactElement } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { createWpsReviewPaneAdapter } from './wpsReviewPaneSizing';
import { useWpsReviewPaneRatio } from './useWpsReviewPaneRatio';
import { initializeRestoredTaskPaneWidth } from './ribbon';
import { REVIEW_PANE_LEASE_KEY, TASK_PANE_STORAGE_KEY } from './wpsReviewPaneLease';
import type { WpsApplication, WpsWorkbook, WpsTaskPane } from './wpsHost';
import type { WpsNativeComparison } from './wpsNativeComparison';
import capabilities from '../../../config/capabilities.json';

(globalThis as {IS_REACT_ACT_ENVIRONMENT?:boolean}).IS_REACT_ACT_ENVIRONMENT = true;
const interval = capabilities.reviewPaneSizing.reviewPaneMeasureIntervalMs;
const sentinel = 'PRIVATE_CELL_SENTINEL_83c';
let roots: Array<ReturnType<typeof createRoot>>;
beforeEach(() => {
  vi.useFakeTimers(); roots=[];
  Object.defineProperty(window,'innerWidth',{value:840,writable:true,configurable:true});
  Object.defineProperty(window,'devicePixelRatio',{value:1,writable:true,configurable:true});
  Object.defineProperty(document,'visibilityState',{value:'visible',configurable:true});
});
afterEach(() => {
  act(() => roots.forEach(root=>root.unmount()));
  delete window.Application; vi.useRealTimers(); vi.unstubAllGlobals(); document.body.innerHTML='';
});
function host(accept = false, scale = 1) {
  const storage = new Map<string, unknown>([[TASK_PANE_STORAGE_KEY, 810]]);
  const range = vi.fn(()=>{throw new Error('Sizing must not read cells');});
  const book = (Name:string, Path=''):WpsWorkbook=>({Name, Path,
    FullName:Path?`${Path}\\${Name}`:undefined,
    Worksheets:{Count:1,Item:()=>({Name:'Data',Range:range})}});
  const original = book('Original.xlsx','C:\\Data');
  const copy = book('Book1');
  let books = [original,copy]; let nativeWidth = 840 * scale;
  const widths:number[]=[];
  const pane:WpsTaskPane = {ID:810,Visible:true,DockPosition:2,
    get Width(){return nativeWidth;},set Width(width){
      nativeWidth=width!; widths.push(width!);
      if (accept) window.innerWidth=width!/scale;
    }};
  const app:WpsApplication={Hwnd:50,Width:900,ActiveWorkbook:original,
    Workbooks:{get Count(){return books.length;},Item:index=>books[Number(index)-1]},
    PluginStorage:{getItem:key=>storage.get(key),setItem:(key,value)=>{storage.set(key,value);},removeItem:key=>{storage.delete(key);}},
    CreateTaskPane:vi.fn(()=>pane), GetTaskPane:vi.fn(id=>id===810?pane:undefined)};
  window.Application=app;
  // Assert the private sentinel really exists locally before proving no outbound/persistence.
  const review={original,copy,copyIdentity:{name:'Book1',fullName:null,path:null},
    plan:{actions:[{values:[[sentinel]]}]}} as unknown as WpsNativeComparison;
  return {app,pane,review,storage,widths,range,removeCopy:()=>{books=[original];},
    duplicateOriginal:()=>{books=[original,{...original},copy];}};
}
function View({review}:{review:WpsNativeComparison|null}) {
  const needsResize=useWpsReviewPaneRatio(review);
  return <output>{needsResize?'manual guidance':'ready'}</output>;
}
function mount(element:ReactElement) {
  const container=document.createElement('div'); document.body.append(container);
  const root=createRoot(container); roots.push(root);
  act(()=>root.render(element));
  return {container,render:(next:ReactElement)=>act(()=>root.render(next)),unmount:()=>act(()=>root.unmount())};
}
const tick=(ms=interval)=>act(()=>vi.advanceTimersByTime(ms));
const resize=(width:number)=>act(()=>{window.innerWidth=width;window.dispatchEvent(new Event('resize'));});

it('measures this WPS application in points, not the iframe, and restores ordinary width on exit',()=>{
  const h=host(true); const view=mount(<View review={h.review}/>);
  expect(h.widths).toEqual([480]); expect(view.container.textContent).toBe('manual guidance');
  tick(); expect(view.container.textContent).toBe('ready');
  h.app.Width=1125; tick(); expect(h.widths).toEqual([480,600]);
  view.render(<View review={null}/>); expect(h.widths).toEqual([480,600,840]);
  tick(interval*3); expect(h.widths).toHaveLength(3);
  expect(h.storage.has(REVIEW_PANE_LEASE_KEY)).toBe(false);
});
it('calibrates native pane units rather than assuming they equal CSS pixels',()=>{
  const h=host(true,1.5); mount(<View review={h.review}/>);
  expect(h.widths).toEqual([720]); expect(window.innerWidth).toBe(480);
  tick(); expect(h.widths).toEqual([720]);
});
it('manual dragging wins, remains protected from ribbon migration, and new review resets sizing',()=>{
  const h=host(true); const view=mount(<View review={h.review}/>); tick();
  h.pane.Width=430; resize(430); tick();
  h.app.Width=1350; tick(interval*5);
  initializeRestoredTaskPaneWidth(); expect(h.pane.Width).toBe(430);
  expect(h.widths).toEqual([480,430]);
  view.render(<View review={{...h.review}}/>); expect(h.widths.slice(-2)).toEqual([840,720]);
});
it('a silently ignored width request does not claim successful sizing or loop',()=>{
  const h=host(false); const view=mount(<View review={h.review}/>); tick(interval*3);
  expect(h.widths).toEqual([480]); expect(view.container.textContent).toBe('manual guidance');
});
it('debounces resize, pauses hidden measurements, resumes on visibility, and cleans up',()=>{
  const h=host(true); const view=mount(<View review={h.review}/>); tick();
  h.app.Width=1125; resize(470); resize(460);
  tick(199); expect(h.widths).toEqual([480]); tick(1); expect(h.widths).toEqual([480,600]); tick();
  Object.defineProperty(document,'visibilityState',{value:'hidden',configurable:true});
  h.app.Width=1350; tick(interval*3); expect(h.widths).toEqual([480,600]);
  Object.defineProperty(document,'visibilityState',{value:'visible',configurable:true});
  act(()=>document.dispatchEvent(new Event('visibilitychange'))); tick(200);
  expect(h.widths).toEqual([480,600,720]);
  view.render(<View review={null}/>); const calls=h.widths.length;
  resize(500); tick(interval*3); expect(h.widths).toHaveLength(calls);
});
it('never reads cells, creates a pane, uploads or persists local plan values',()=>{
  const h=host(true); const fetch=vi.fn(()=>{throw new Error('No sizing network');}); vi.stubGlobal('fetch',fetch);
  expect(JSON.stringify(h.review.plan)).toContain(sentinel);
  mount(<View review={h.review}/>); tick();
  expect(h.range).not.toHaveBeenCalled(); expect(h.app.CreateTaskPane).not.toHaveBeenCalled();
  expect(fetch).not.toHaveBeenCalled();
  const record=String(h.storage.get(REVIEW_PANE_LEASE_KEY));
  expect(JSON.parse(record)).toEqual({paneId:'810',owner:expect.any(String),expiresAt:expect.any(Number)});
  expect(record).not.toContain(sentinel); expect(record).not.toContain('Original.xlsx');
});
it.each(['hostWidth','hwnd','books','dock','hiddenPane','noPane'] as const)('missing %s safely keeps manual guidance and performs zero width writes',kind=>{
  const h=host();
  if(kind==='hostWidth') h.app.Width=undefined;
  if(kind==='hwnd') h.app.Hwnd=undefined;
  if(kind==='books') h.app.Workbooks=undefined;
  if(kind==='dock') h.pane.DockPosition=1;
  if(kind==='hiddenPane') h.pane.Visible=false;
  if(kind==='noPane') h.app.GetTaskPane=()=>undefined;
  const view=mount(<View review={h.review}/>); tick(interval*4);
  expect(h.widths).toEqual([]); expect(view.container.textContent).toBe('manual guidance');
});
it.each(['closedCopy','ambiguous','savedCopy','otherBook','otherPane','otherHost'] as const)('stops on %s without guessing or resizing another pane',kind=>{
  const h=host(true); const view=mount(<View review={h.review}/>); tick();
  if(kind==='closedCopy') h.removeCopy();
  if(kind==='ambiguous') h.duplicateOriginal();
  if(kind==='savedCopy') h.review.copy.Path='C:\\Saved';
  if(kind==='otherBook') h.app.ActiveWorkbook={...h.review.original,Name:'Other.xlsx'};
  if(kind==='otherPane') h.storage.set(TASK_PANE_STORAGE_KEY,811);
  if(kind==='otherHost') h.app.Hwnd=99;
  h.app.Width=1125; tick();
  expect(h.widths).toEqual([480]); expect(view.container.textContent).toBe('manual guidance');
  view.render(<View review={null}/>);
  if(kind==='otherPane'||kind==='otherHost') expect(h.widths).toEqual([480]);
});
it('does not take width ownership from another live review',()=>{
  const h=host(); h.storage.set(REVIEW_PANE_LEASE_KEY,JSON.stringify({paneId:'810',owner:'other',expiresAt:Date.now()+6000}));
  const view=mount(<View review={h.review}/>); expect(h.widths).toEqual([]);
  view.render(<View review={null}/>); expect(h.widths).toEqual([]);
});
it('restores captured ordinary width after capability failure even if its marker expired',()=>{
  const h=host(true); const view=mount(<View review={h.review}/>); tick();
  h.app.Width=undefined; tick(interval*8); view.render(<View review={null}/>);
  expect(h.widths).toEqual([480,840]);
});
it('rechecks live scope immediately before setting native width',()=>{
  const h=host(); const adapter=createWpsReviewPaneAdapter(h.app,h.review);
  h.removeCopy(); expect(()=>adapter.resize(480,840)).toThrow(); expect(h.widths).toEqual([]);
});
it('does not restore into a disconnected application after the global host was replaced',()=>{
  const h=host(true); const view=mount(<View review={h.review}/>); tick();
  window.Application={...h.app,Hwnd:100}; tick(); view.render(<View review={null}/>);
  expect(h.widths).toEqual([480]);
});
it('releases sizing ownership without removeItem so the next review can start immediately',()=>{
  const h=host(true); h.app.PluginStorage.removeItem=undefined;
  const view=mount(<View review={h.review}/>); view.render(<View review={null}/>);
  view.render(<View review={{...h.review}}/>);
  expect(h.widths).toEqual([480,840,480]);
});
it('StrictMode effect replay neither leaks ownership nor loses the ordinary width',()=>{
  const h=host(true); const view=mount(<StrictMode><View review={h.review}/></StrictMode>);
  expect(h.widths).toEqual([480,840,480]);
  view.render(<StrictMode><View review={null}/></StrictMode>);
  expect(h.widths).toEqual([480,840,480,840]);
});
it('scale changes require manual adjustment/new review instead of using stale calibration',()=>{
  const h=host(true); const view=mount(<View review={h.review}/>); tick();
  window.devicePixelRatio=1.5; h.app.Width=1125; tick();
  expect(h.widths).toEqual([480]); expect(view.container.textContent).toBe('manual guidance');
});
