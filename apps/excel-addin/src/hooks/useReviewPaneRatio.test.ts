import { useSyncExternalStore } from 'react';
import { ResultReview } from '@excel-bro/core/resultReview';
// @vitest-environment jsdom
import { act, cleanup, renderHook } from '@testing-library/react';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { useReviewPaneRatio, reviewPaneTargetWidth, reviewLayoutPlan } from './useReviewPaneRatio';
import { useUIState, TASK_PANE_WIDTH_STORAGE_KEY } from './useUIState';
import { measureLocalExcelWindow } from '../excelWindowLayout';
import capabilities from '../../../../config/capabilities.json';
vi.mock('../excelWindowLayout', () => ({ measureLocalExcelWindow: vi.fn() }));
const scope = {workbookIdentity:'url:E:/Original.xlsx', original:{name:'Ledger'}, copySheetName:'EB_'+'a'.repeat(20)};
const interval = capabilities.reviewPaneSizing.reviewPaneMeasureIntervalMs;
/** 工作区宽度（AI 比例的唯一参照）。 */
let host: number, setWidth: ReturnType<typeof vi.fn>;
const measured = () => ({clientWidthCssPx: host - 400, workAreaWidthCssPx: host});
beforeEach(() => {
  vi.useFakeTimers(); localStorage.clear(); host = 1200;
  Object.defineProperty(window, 'innerWidth', {value:800,writable:true,configurable:true});
  setWidth = vi.fn();
  vi.mocked(measureLocalExcelWindow).mockReset().mockImplementation(async () => measured());
  vi.stubGlobal('Office', {onReady:(fn:()=>void)=>fn(), context:{requirements:{isSetSupported:()=>true}},
    extensionLifeCycle:{taskpane:{setWidth}}});
});
afterEach(() => {cleanup();vi.useRealTimers();vi.unstubAllGlobals();});
const tick = async (ms = interval) => act(async () => {await vi.advanceTimersByTimeAsync(ms);});
const resize = (width: number) => act(() => { window.innerWidth = width; window.dispatchEvent(new Event('resize')); });

it('derives the ratio from the whole monitor work area, never the original window or the iframe', async () => {
  const {result} = renderHook(()=>useReviewPaneRatio(scope)); await tick(0);
  // 工作区 1200 ⇒ 30% = 360。若误用原表窗口内容区(800) 会得到 240，误用 iframe(800) 也是 240。
  expect(setWidth).toHaveBeenCalledExactlyOnceWith(360);
  expect(setWidth).not.toHaveBeenCalledWith(240);
  expect(measureLocalExcelWindow).toHaveBeenCalledWith(scope.workbookIdentity,'Ledger',scope.copySheetName);
  expect(result.current).toBe(true); // requested != accepted
  resize(360); await tick(); expect(result.current).toBe(false);
  expect(setWidth).toHaveBeenCalledTimes(1); // own resize is not a loop
});
it('changing tabs in the same copy preserves the pane sizing session', async () => {
  const {rerender} = renderHook(({value}) => useReviewPaneRatio(value), {initialProps:{value:scope}});
  await tick(0); resize(360); await tick();
  const count = setWidth.mock.calls.length;
  rerender({value:{...scope, original:{name:'结果3'}}});
  await tick();
  expect(setWidth).toHaveBeenCalledTimes(count);
  expect(measureLocalExcelWindow).toHaveBeenLastCalledWith(scope.workbookIdentity, '结果3', scope.copySheetName);
});
it('follows host-only work-area changes even when the iframe stays fixed, debounces pane resize', async () => {
  renderHook(()=>useReviewPaneRatio(scope)); await tick(0); resize(360); await tick();
  host = 1500; await tick(); expect(setWidth).toHaveBeenLastCalledWith(450);
  resize(450); await tick(); setWidth.mockClear();
  host = 1000; resize(440); resize(430); resize(420);
  await tick(199); expect(setWidth).not.toHaveBeenCalled();
  await tick(1); expect(setWidth).toHaveBeenCalledExactlyOnceWith(320);
});
it('respects manual pane-only dragging for the rest of this review, resets for a new copy', async () => {
  const {rerender,result} = renderHook(({copy})=>useReviewPaneRatio({...scope,copySheetName:copy}),{initialProps:{copy:scope.copySheetName}});
  await tick(0); resize(360); await tick(); setWidth.mockClear();
  resize(430); await tick(); host=1600; await tick(interval*3);
  act(()=>window.dispatchEvent(new Event('focus'))); await tick();
  expect(setWidth).not.toHaveBeenCalled(); expect(result.current).toBe(false);
  rerender({copy:'EB_'+'b'.repeat(20)}); await tick(0); expect(setWidth).toHaveBeenCalledWith(480);
});
it('does not fight manual dragging before its requested width is acknowledged', async () => {
  renderHook(()=>useReviewPaneRatio(scope)); await tick(0); setWidth.mockClear();
  resize(450); await tick(); host=1600; await tick(); expect(setWidth).not.toHaveBeenCalled();
});
it('re-applies the default ratio after manual dragging only when the user asks to rearrange', async () => {
  const {rerender} = renderHook(({rearm})=>useReviewPaneRatio(scope,rearm),{initialProps:{rearm:0}});
  await tick(0); resize(360); await tick(); setWidth.mockClear();
  resize(430); await tick();
  host = 1600; await tick(interval*3);
  act(()=>window.dispatchEvent(new Event('focus'))); await tick();
  // 用户手动拖过 ⇒ 之后的工作区变化绝不自动纠正（这是"用户自己的布局"）。
  expect(setWidth).not.toHaveBeenCalled();
  rerender({rearm:1}); await tick(0);
  // 只有用户主动「重新排列」才清掉 manual 标记并重套默认比例。
  expect(setWidth).toHaveBeenCalledExactlyOnceWith(480);
});
it('does not re-arrange on ordinary re-renders, message updates or readback', async () => {
  const {rerender} = renderHook(({s})=>useReviewPaneRatio(s),{initialProps:{s:scope}});
  await tick(0); resize(360); await tick(); setWidth.mockClear();
  const calls = vi.mocked(measureLocalExcelWindow).mock.calls.length;
  // 新对象、同内容：依赖是基本类型，状态机必须继续活着而不是重建。
  rerender({s:{...scope}}); await tick(); await tick(); await tick();
  expect(vi.mocked(measureLocalExcelWindow).mock.calls.length).toBeGreaterThan(calls);
  expect(setWidth).not.toHaveBeenCalled();
});
it('preserves ordinary chat preference and restores it after review', async () => {
  localStorage.setItem(TASK_PANE_WIDTH_STORAGE_KEY,'720');
  const {rerender} = renderHook(({review})=>useUIState(review,review?scope:undefined), {initialProps:{review:true}});
  await tick(0); expect(setWidth).toHaveBeenCalledExactlyOnceWith(360); resize(360); await tick();
  resize(430); await tick(); expect(localStorage.getItem(TASK_PANE_WIDTH_STORAGE_KEY)).toBe('720');
  rerender({review:false}); expect(setWidth).toHaveBeenLastCalledWith(720);
  const calls=vi.mocked(measureLocalExcelWindow).mock.calls.length; await tick(interval*3);
  expect(vi.mocked(measureLocalExcelWindow).mock.calls).toHaveLength(calls);
});
it('stops safely on missing/changed window evidence rather than guessing from iframe', async () => {
  vi.mocked(measureLocalExcelWindow).mockRejectedValue(new Error('changed'));
  const {result}=renderHook(()=>useReviewPaneRatio(scope)); await tick(0); await tick(interval*3);
  expect(setWidth).not.toHaveBeenCalled(); expect(result.current).toBe(true);
  expect(measureLocalExcelWindow).toHaveBeenCalledTimes(1);
});
it('retains guidance if the host silently ignores setWidth', async () => {
  const {result}=renderHook(()=>useReviewPaneRatio(scope)); await tick(0); await tick(interval*3);
  expect(result.current).toBe(true); expect(setWidth).toHaveBeenCalledTimes(1);
});
it('does not retry a throwing host width API', async () => {
  setWidth.mockImplementation(()=>{throw new Error('unsupported runtime');});
  const {result}=renderHook(()=>useReviewPaneRatio(scope)); await tick(0); await tick(interval*3);
  expect(result.current).toBe(true); expect(setWidth).toHaveBeenCalledTimes(1);
});
it('does not overlap measurements while a native call is slow', async () => {
  let resolve!: (value:{clientWidthCssPx:number;workAreaWidthCssPx:number})=>void;
  vi.mocked(measureLocalExcelWindow).mockImplementation(()=>new Promise(r=>{resolve=r;}));
  renderHook(()=>useReviewPaneRatio(scope)); resize(700); await tick(interval*3);
  expect(measureLocalExcelWindow).toHaveBeenCalledTimes(1);
  await act(async()=>resolve({clientWidthCssPx:800,workAreaWidthCssPx:1200}));
  expect(setWidth).toHaveBeenCalledExactlyOnceWith(360);
});
it('keeps the manual-resize guidance when TaskPaneApi 1.1 is unavailable', async () => {
  Office.context.requirements.isSetSupported=()=>false;
  const {result}=renderHook(()=>useReviewPaneRatio(scope)); await tick();
  expect(result.current).toBe(true); expect(measureLocalExcelWindow).not.toHaveBeenCalled(); expect(setWidth).not.toHaveBeenCalled();
});
it('ignores an in-flight measurement after unmount', async () => {
  let resolve!: (value:{clientWidthCssPx:number;workAreaWidthCssPx:number})=>void;
  vi.mocked(measureLocalExcelWindow).mockImplementation(()=>new Promise(r=>{resolve=r;}));
  const {unmount}=renderHook(()=>useReviewPaneRatio(scope)); unmount();
  await act(async()=>resolve({clientWidthCssPx:800,workAreaWidthCssPx:1200})); await tick(interval*3);
  expect(setWidth).not.toHaveBeenCalled(); expect(measureLocalExcelWindow).toHaveBeenCalledTimes(1);
});
it('clamps the readable minimum to the Office half-window bound', () => {
  expect(reviewPaneTargetWidth(1200)).toBe(360);
  expect(reviewPaneTargetWidth(700)).toBe(320);
  expect(reviewPaneTargetWidth(500)).toBe(250); // cannot force 320 past Office maximum
  expect(reviewPaneTargetWidth(501)).toBe(250); // rounding must not exceed half
});
it('keeps 35/35/30 of the whole work area and never 30% of the right window', () => {
  const plan = reviewLayoutPlan(2000);
  expect(plan.copyWidth).toBe(700);
  expect(plan.originalWidth).toBe(1300);
  expect(plan.aiPaneWidth).toBe(600);
  expect(plan.originalSheetWidth).toBe(700);
  expect(plan.aiPaneWidth).not.toBe(Math.round(plan.originalWidth * 0.3)); // 390 would squeeze the original to 910
  expect(plan.degraded).toBe(false);
});

it('sizes the executed-result copy and restores normal chat width after leaving comparison', async () => {
  localStorage.setItem(TASK_PANE_WIDTH_STORAGE_KEY, '800');
  const review = new ResultReview<string, {executionId:string;verified:boolean;undoable:boolean}, string, typeof scope>();
  const execute = vi.fn(async () => ({executionId:'e',verified:true,undoable:true}));
  const {result} = renderHook(() => {
    const state = useSyncExternalStore(review.subscribe,review.getSnapshot,review.getSnapshot);
    return useUIState(state.comparing || (state.phase === 'reviewing' && state.pending),review.getComparisonCopy());
  });
  await act(async () => { await review.start('r', {
    isCurrent:()=>true,prepare:async()=> 'before',execute,mark:async()=> 'marker',clearMarker:async()=>{},
    showBefore:async()=>scope,closeBefore:async()=>null,restore:async()=>{},accepted:()=>{}
  }); });
  setWidth.mockClear();
  await act(async () => { await review.compare('r'); }); await tick(0);
  expect(setWidth).toHaveBeenCalledWith(360);
  resize(360); await tick(); setWidth.mockClear();
  act(()=>window.dispatchEvent(new Event('focus'))); await tick(0);
  expect(setWidth).not.toHaveBeenCalledWith(800);
  expect(localStorage.getItem(TASK_PANE_WIDTH_STORAGE_KEY)).toBe('800');
  expect(result.current.comparisonPaneNeedsResize).toBe(false);
  await act(async () => { await review.exitComparison('r'); }); await tick(0);
  expect(setWidth).toHaveBeenCalledWith(800);
  expect(execute).toHaveBeenCalledTimes(1);
});
