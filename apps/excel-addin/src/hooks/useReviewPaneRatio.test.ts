// @vitest-environment jsdom
import { act, cleanup, renderHook } from '@testing-library/react';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { useReviewPaneRatio, reviewPaneTargetWidth } from './useReviewPaneRatio';
import { useUIState, TASK_PANE_WIDTH_STORAGE_KEY } from './useUIState';
import { measureLocalExcelWindow } from '../excelWindowLayout';
import capabilities from '../../../../config/capabilities.json';
vi.mock('../excelWindowLayout', () => ({ measureLocalExcelWindow: vi.fn() }));
const scope = {workbookIdentity:'url:E:/Original.xlsx', original:{name:'Ledger'}, copySheetName:'EB_'+'a'.repeat(20)};
const interval = capabilities.reviewPaneSizing.reviewPaneMeasureIntervalMs;
let host: number, setWidth: ReturnType<typeof vi.fn>;
beforeEach(() => {
  vi.useFakeTimers(); localStorage.clear(); host = 1200;
  Object.defineProperty(window, 'innerWidth', {value:800,writable:true,configurable:true});
  setWidth = vi.fn();
  vi.mocked(measureLocalExcelWindow).mockReset().mockImplementation(async () => ({clientWidthCssPx:host}));
  vi.stubGlobal('Office', {onReady:(fn:()=>void)=>fn(), context:{requirements:{isSetSupported:()=>true}},
    extensionLifeCycle:{taskpane:{setWidth}}});
});
afterEach(() => {cleanup();vi.useRealTimers();vi.unstubAllGlobals();});
const tick = async (ms = interval) => act(async () => {await vi.advanceTimersByTimeAsync(ms);});
const resize = (width: number) => act(() => { window.innerWidth = width; window.dispatchEvent(new Event('resize')); });

it('uses the verified right host client width, not iframe or screen width', async () => {
  const {result} = renderHook(()=>useReviewPaneRatio(scope)); await tick(0);
  expect(setWidth).toHaveBeenCalledExactlyOnceWith(480);
  expect(measureLocalExcelWindow).toHaveBeenCalledWith(scope.workbookIdentity,'Ledger',scope.copySheetName);
  expect(result.current).toBe(true); // requested != accepted
  resize(480); await tick(); expect(result.current).toBe(false);
  expect(setWidth).toHaveBeenCalledTimes(1); // own resize is not a loop
});
it('follows host-only changes even when the iframe stays fixed, debounces pane resize', async () => {
  renderHook(()=>useReviewPaneRatio(scope)); await tick(0); resize(480); await tick();
  host = 1500; await tick(); expect(setWidth).toHaveBeenLastCalledWith(600);
  resize(600); await tick(); setWidth.mockClear();
  host = 1000; resize(590); resize(580); resize(570);
  await tick(199); expect(setWidth).not.toHaveBeenCalled();
  await tick(1); expect(setWidth).toHaveBeenCalledExactlyOnceWith(400);
});
it('respects manual pane-only dragging for the rest of this review, resets for a new copy', async () => {
  const {rerender,result} = renderHook(({copy})=>useReviewPaneRatio({...scope,copySheetName:copy}),{initialProps:{copy:scope.copySheetName}});
  await tick(0); resize(480); await tick(); setWidth.mockClear();
  resize(430); await tick(); host=1600; await tick(interval*3);
  act(()=>window.dispatchEvent(new Event('focus'))); await tick();
  expect(setWidth).not.toHaveBeenCalled(); expect(result.current).toBe(false);
  rerender({copy:'EB_'+'b'.repeat(20)}); await tick(0); expect(setWidth).toHaveBeenCalledWith(640);
});
it('does not fight manual dragging before its requested width is acknowledged', async () => {
  renderHook(()=>useReviewPaneRatio(scope)); await tick(0); setWidth.mockClear();
  resize(450); await tick(); host=1600; await tick(); expect(setWidth).not.toHaveBeenCalled();
});
it('preserves ordinary chat preference and restores it after review', async () => {
  localStorage.setItem(TASK_PANE_WIDTH_STORAGE_KEY,'720');
  const {rerender} = renderHook(({review})=>useUIState(review,review?scope:undefined), {initialProps:{review:true}});
  await tick(0); expect(setWidth).toHaveBeenCalledExactlyOnceWith(480); resize(480); await tick();
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
  let resolve!: (value:{clientWidthCssPx:number})=>void;
  vi.mocked(measureLocalExcelWindow).mockImplementation(()=>new Promise(r=>{resolve=r;}));
  renderHook(()=>useReviewPaneRatio(scope)); resize(700); await tick(interval*3);
  expect(measureLocalExcelWindow).toHaveBeenCalledTimes(1);
  await act(async()=>resolve({clientWidthCssPx:1200})); expect(setWidth).toHaveBeenCalledExactlyOnceWith(480);
});
it('does not measure or resize when TaskPaneApi is unavailable', async () => {
  Office.context.requirements.isSetSupported=()=>false;
  const {result}=renderHook(()=>useReviewPaneRatio(scope)); await tick();
  expect(result.current).toBe(true); expect(measureLocalExcelWindow).not.toHaveBeenCalled(); expect(setWidth).not.toHaveBeenCalled();
});
it('ignores an in-flight measurement after unmount', async () => {
  let resolve!: (value:{clientWidthCssPx:number})=>void;
  vi.mocked(measureLocalExcelWindow).mockImplementation(()=>new Promise(r=>{resolve=r;}));
  const {unmount}=renderHook(()=>useReviewPaneRatio(scope)); unmount();
  await act(async()=>resolve({clientWidthCssPx:1200})); await tick(interval*3);
  expect(setWidth).not.toHaveBeenCalled(); expect(measureLocalExcelWindow).toHaveBeenCalledTimes(1);
});
it('clamps readable minimum to Office half-window bound', () => {
  expect(reviewPaneTargetWidth(1200)).toBe(480);
  expect(reviewPaneTargetWidth(700)).toBe(320);
  expect(reviewPaneTargetWidth(500)).toBe(250); // cannot force 320 past Office maximum
  expect(reviewPaneTargetWidth(501)).toBe(250); // rounding must not exceed half
});
