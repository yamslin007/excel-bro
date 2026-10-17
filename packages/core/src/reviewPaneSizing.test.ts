import { expect, it } from 'vitest';
import capabilities from '../../../config/capabilities.json';
import { ReviewPaneSizing, reviewLayoutPlan, reviewPaneTargetWidth } from './reviewPaneSizing';

const limits = capabilities.reviewPaneSizing;

it('computes the 35/35/30 split from the whole usable work area', () => {
  const plan = reviewLayoutPlan(2000);
  // 副本 35% / 原表 35% / AI 30%，全部相对**整个工作区**。
  expect(plan.copyWidth).toBe(700);
  expect(plan.originalWidth).toBe(1300);
  expect(plan.originalSheetWidth).toBe(700);
  expect(plan.aiPaneWidth).toBe(600);
  expect(plan.aiPaneWidth).toBe(Math.round(2000 * limits.reviewAiWidthFraction));
  // 反向：AI 绝不是"右侧窗口的 30%"（那会得到 390）。
  expect(plan.aiPaneWidth).not.toBe(Math.round(plan.originalWidth * limits.reviewAiWidthFraction));
  expect(plan.degraded).toBe(false);
  // 两张表的可视宽度接近（副本 == 原表表格宽）。
  expect(Math.abs(plan.copyWidth - plan.originalSheetWidth)).toBeLessThanOrEqual(1);
});

it('scales with the work area and keeps the same three-way proportions', () => {
  for (const width of [1280, 1920, 2560, 3840]) {
    const plan = reviewLayoutPlan(width);
    expect(plan.copyWidth + plan.originalWidth).toBe(width);
    expect(plan.originalWidth).toBe(width - plan.copyWidth);
    expect(plan.degraded).toBe(false);
    expect(Math.abs(plan.copyWidth - plan.originalSheetWidth)).toBeLessThanOrEqual(1);
  }
});

it('degrades on a small work area instead of squeezing a pane out', () => {
  const plan = reviewLayoutPlan(600);
  expect(plan.degraded).toBe(true);
  // AI 仍然不小于最小宽，且表格不会被压成 0。
  expect(plan.aiPaneWidth).toBeGreaterThan(0);
  expect(plan.originalSheetWidth).toBeGreaterThan(0);
  expect(plan.copyWidth).toBeGreaterThan(0);
  expect(plan.aiPaneWidth).toBeLessThanOrEqual(Math.floor(plan.originalWidth / 2) + 1);
});

it('rejects unusable host widths rather than guessing a screen size', () => {
  for (const width of [0, -1, NaN, Infinity, limits.maxHostWidthCssPx + 1]) {
    expect(() => reviewPaneTargetWidth(width)).toThrow();
    expect(() => reviewLayoutPlan(width)).toThrow();
  }
});

it('uses the shared AI fraction against the work area, never the right window', () => {
  // 1200 → 30% = 360；700 → 30% = 210 但触发 320 下限；500/501 → 半窗口上限 250。
  expect([1200, 700, 500, 501].map(reviewPaneTargetWidth)).toEqual([360, 320, 250, 250]);
});

it('requires actual width evidence, not just a successful setter', () => {
  const sizing = new ReviewPaneSizing(800);
  expect(sizing.observe(1200, 800)).toMatchObject({ requestWidth: 360, needsResize: true });
  expect(sizing.observe(1200, 800)).toMatchObject({ requestWidth: undefined, needsResize: true });
  expect(sizing.observe(1200, 360)).toMatchObject({ requestWidth: undefined, needsResize: false, manual: false });
});

it('follows a host resize but never fights a pane-only manual resize', () => {
  const sizing = new ReviewPaneSizing(800);
  sizing.observe(1200, 800); sizing.observe(1200, 360);
  expect(sizing.observe(1500, 360).requestWidth).toBe(450);
  sizing.observe(1500, 450);
  expect(sizing.observe(1500, 430).manual).toBe(true);
  expect(sizing.observe(1800, 430)).toMatchObject({ manual: true, requestWidth: undefined, needsResize: false });
  expect(new ReviewPaneSizing(430).observe(1800, 430).requestWidth).toBe(540);
});

it('keeps the outer split single-sourced from the copy fraction', () => {
  // 服务端据此推导右侧原表窗口宽度（1 - 副本占宽），两端不得各写一份比例。
  expect(limits.reviewCopyWidthFraction).toBeGreaterThan(0);
  expect(limits.reviewCopyWidthFraction).toBeLessThan(1);
  expect(1 - limits.reviewCopyWidthFraction).toBeCloseTo(0.65, 10);
  expect(capabilities.excelWindowLayout).not.toHaveProperty('originalWidthFraction');
});
