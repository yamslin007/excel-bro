// @vitest-environment jsdom
/*
 * 「重新排列」的判别性测试。
 *
 * 关键判据：**只有**用户主动调用 `rearrange` 才会碰窗口；重渲染、消息更新、
 * 保存工具都不会。排列失败只写 feedback —— 审阅阶段与已完成的写入/撤销不被改写，
 * 也绝不把业务执行误报为失败（反之亦然）。
 */
import { act, cleanup, renderHook } from '@testing-library/react';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { useExcelNativeComparison } from './useExcelNativeComparison';
import { rearrangeLocalExcelWindows } from '../excelWindowLayout';
import type { AnalysisPlan } from '../contracts';
import type { ExcelNativeComparison } from '../excelNativeComparison';

const mocks = vi.hoisted(() => ({ stage: vi.fn(), guard: vi.fn(), close: vi.fn(), mode: vi.fn() }));
vi.mock('../excelNativeComparison', () => ({
  stageExcelNativeComparison: mocks.stage,
  assertExcelNativeComparisonUnchanged: mocks.guard,
  closeExcelNativeComparison: mocks.close,
  excelNativeReviewMode: mocks.mode
}));
vi.mock('../excelWindowLayout', () => ({ rearrangeLocalExcelWindows: vi.fn() }));

const plan: AnalysisPlan = {
  id: 'plan', title: 'Preview', summary: '', assumptions: [], warnings: [],
  actions: [{ type: 'writeValues', sheet: 'Ledger', range: 'C2', values: [[1]] }]
};
const session = {
  plan, authorization: ['Ledger'], changedCells: 1,
  workbookIdentity: 'url:file:///E:/Original.xlsx', originalWindowName: 'Original.xlsx',
  copyWindowName: 'Book1', copySheetName: 'EB_rearrange', copySheetId: 'copy',
  copyEvidence: 'proof', actionEvidence: 'proof', original: { name: 'Ledger' }
} as unknown as ExcelNativeComparison;

beforeEach(() => {
  vi.clearAllMocks();
  mocks.stage.mockResolvedValue(session);
  mocks.guard.mockResolvedValue(undefined);
  mocks.close.mockResolvedValue(true);
  mocks.mode.mockReturnValue('native');
  vi.mocked(rearrangeLocalExcelWindows).mockResolvedValue({
    status: 'arranged', layoutVersion: 3, copySide: 'left', originalSide: 'right', synchronousScrolling: 'enabled'
  });
});
afterEach(() => { cleanup(); vi.unstubAllGlobals(); });

async function staged() {
  const hook = renderHook(() => useExcelNativeComparison());
  await act(async () => { await hook.result.current.stage('m1', plan, ['Ledger']); });
  return hook;
}

it('rearranges through the local adapter with the live review scope, never a stale one', async () => {
  const { result } = await staged();
  expect(result.current.review?.session.windowLayout).toBeUndefined();
  let ok: boolean | undefined;
  await act(async () => { ok = await result.current.rearrange('m1'); });
  expect(rearrangeLocalExcelWindows).toHaveBeenCalledExactlyOnceWith('url:file:///E:/Original.xlsx', 'Ledger', 'EB_rearrange');
  expect(ok).toBe(true);
  expect(result.current.review?.session.windowLayout).toBe('arranged');
  expect(result.current.review?.session.synchronousScrolling).toBe('enabled');
  expect(result.current.review?.outcome).toBe('reviewing');
});

it('never rearranges on re-render, streaming or readback: only an explicit call moves windows', async () => {
  const { result, rerender } = await staged();
  rerender(); rerender();
  await act(async () => {});
  expect(rearrangeLocalExcelWindows).not.toHaveBeenCalled(); // 没有任何"顺手重排"
  await act(async () => { await result.current.rearrange('m1'); });
  expect(rearrangeLocalExcelWindows).toHaveBeenCalledOnce();
});

it('reports a failed arrangement as feedback without touching the review or the completed write', async () => {
  const { result } = await staged();
  const before = result.current.review;
  vi.mocked(rearrangeLocalExcelWindows).mockRejectedValue(new Error('未找到两个匹配的窗口'));
  let ok: boolean | undefined;
  await act(async () => { ok = await result.current.rearrange('m1'); });
  expect(ok).toBe(false);
  expect(result.current.feedback?.text).toContain('未找到两个匹配的窗口');
  // 排列失败 ≠ 业务失败：审阅仍在、阶段仍是 review、布局结论没被改写。
  expect(result.current.review).toBe(before);
  expect(result.current.review?.outcome).toBe('reviewing');
  expect(result.current.review?.session.windowLayout).toBeUndefined();
  expect(mocks.close).not.toHaveBeenCalled();
});

it('refuses to arrange when the card has no live review, without contacting the bridge', async () => {
  const { result } = renderHook(() => useExcelNativeComparison());
  let ok: boolean | undefined;
  await act(async () => { ok = await result.current.rearrange('m1'); });
  expect(ok).toBe(false);
  expect(rearrangeLocalExcelWindows).not.toHaveBeenCalled(); // 绝不移动无关窗口
  expect(result.current.feedback?.text).toContain('没有活跃审阅副本');
});

it('rejects a rearrange request aimed at a different, non-review message', async () => {
  const { result } = await staged();
  let ok: boolean | undefined;
  await act(async () => { ok = await result.current.rearrange('m2'); });
  expect(ok).toBe(false);
  expect(rearrangeLocalExcelWindows).not.toHaveBeenCalled();
  expect(result.current.review?.messageId).toBe('m1');
});
