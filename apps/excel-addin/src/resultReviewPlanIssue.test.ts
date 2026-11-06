import { describe, expect, it } from 'vitest';
import { resultReviewPlanIssue } from './excelNativeComparison';
import type { AnalysisPlan } from './contracts';

const plan = {
  id: 'review-multi-source', title: '结果', summary: '结果',
  sourceFingerprintSheets: ['来源甲', '来源乙', '汇总'],
  sourceDependencySheets: ['来源甲', '来源乙'],
  actions: [{ type: 'writeValues', sheet: '汇总', range: 'A1:B1', values: [['合计', 12]] }],
  assumptions: [], warnings: []
} as AnalysisPlan;
describe('result review authorization', () => {
  it('accepts multiple authorized sources and one output without changing the source plan', () => {
    const before = JSON.stringify(plan);
    expect(resultReviewPlanIssue(plan, ['来源甲', '来源乙', '汇总'])).toBeNull();
    expect(JSON.stringify(plan)).toBe(before);
  });
  it('refuses a source outside authorization', () => {
    expect(resultReviewPlanIssue(plan, ['来源甲', '汇总'])).not.toBeNull();
  });
  it('refuses an unauthorized output', () => {
    expect(resultReviewPlanIssue(plan, ['来源甲', '来源乙'])).not.toBeNull();
  });
  it('accepts multiple authorized output sheets', () => {
    expect(resultReviewPlanIssue({ ...plan, actions: [...plan.actions,
      { type: 'writeValues', sheet: '来源甲', range: 'A1:B1', values: [['a', 1]] }]
    }, ['来源甲', '来源乙', '汇总'])).toBeNull();
  });
});
