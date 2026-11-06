import { describe, expect, it } from 'vitest';
import { composerPendingSummary } from './composerPending';

describe('底部常驻事项的紧凑摘要', () => {
  it('没有事项就不画这一块', () => {
    const summary = composerPendingSummary([]);
    expect(summary.count).toBe(0);
    expect(summary.text).toBe('');
    expect(summary.defaultOpen).toBe(false);
  });

  it('决定类事项排在摘要最前，并默认展开', () => {
    const summary = composerPendingSummary([
      { id: 'status', label: '服务状态', kind: 'status' },
      { id: 'decision', label: '待审阅决定', kind: 'decision' }
    ]);
    expect(summary.labels[0]).toBe('待审阅决定');
    expect(summary.defaultOpen).toBe(true);
    expect(summary.text).toContain('待处理 2 项');
  });

  it('纯状态/通知默认收起，超过两条只报数量', () => {
    const summary = composerPendingSummary([
      { id: 'a', label: '授权范围', kind: 'status' },
      { id: 'b', label: '恢复核对', kind: 'recovery' },
      { id: 'c', label: '历史存储', kind: 'notice' },
      { id: 'd', label: '服务状态', kind: 'status' }
    ]);
    expect(summary.defaultOpen).toBe(false);
    expect(summary.text).toContain('恢复核对');
    expect(summary.text).toContain('等 4 项');
    expect(summary.labels).toHaveLength(2);
  });
});
