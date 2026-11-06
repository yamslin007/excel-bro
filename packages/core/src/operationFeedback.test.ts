import { describe, expect, it } from 'vitest';
import {
  compactOperationFeedback,
  copyCleanupSummary,
  dedupeReviewCopyCleanupIssues,
  groupOperationFeedback,
  operationFeedbackAnchorIds,
  type OperationFeedbackMessage,
  type ReviewCopyCleanupIssue
} from './operationFeedback';

const feedback = (
  executionId: string,
  kind: 'confirmed' | 'discarded' | 'undone' | 'copy_cleanup_failed',
  text: string
) => ({ executionId, kind, text });

function rows(...messages: OperationFeedbackMessage[]) { return messages; }

describe('操作反馈归属（不同 executionId 不许误合并）', () => {
  it('短反馈挂到同一 executionId 的执行记录下，自己不再单独占一行', () => {
    const grouping = groupOperationFeedback(rows(
      { id: 'a', verification: { executionId: 'exec-a' } },
      { id: 'f1', feedback: feedback('exec-a', 'discarded', '已放弃并恢复：主表已按本次恢复点还原。') },
      { id: 'f2', feedback: feedback('exec-a', 'copy_cleanup_failed', '展示副本未清理') }
    ));
    expect([...grouping.byAnchor.keys()]).toEqual(['a']);
    expect(grouping.byAnchor.get('a')).toHaveLength(2);
    expect(grouping.suppressed.has('f1')).toBe(true);
    expect(grouping.suppressed.has('f2')).toBe(true);
    expect(grouping.unattached.size).toBe(0);
  });

  it('不同执行身份即便紧挨着也不合并：B 的反馈不许挂到 A 的结果卡下', () => {
    const grouping = groupOperationFeedback(rows(
      { id: 'a', executedExecutionId: 'exec-a' },
      { id: 'b', feedback: feedback('exec-b', 'undone', '已撤销本次执行') },
      { id: 'c', officeExecutionId: 'exec-b' },
      { id: 'd', feedback: feedback('exec-b', 'confirmed', '已确认变更') }
    ));
    expect([...grouping.byAnchor.keys()]).toEqual(['c']);
    expect(grouping.suppressed.has('b')).toBe(false);
    expect(grouping.unattached.has('b')).toBe(true);
    expect(grouping.suppressed.has('d')).toBe(true);
  });

  it('找不到执行记录时保留原消息行，不按文案或位置猜', () => {
    const grouping = groupOperationFeedback(rows(
      { id: 'a', verification: { executionId: 'exec-a' } },
      { id: 'f', feedback: feedback('exec-missing', 'discarded', '已放弃并恢复') }
    ));
    expect(grouping.byAnchor.size).toBe(0);
    expect(grouping.unattached.has('f')).toBe(true);
    expect(grouping.suppressed.size).toBe(0);
  });

  it('反馈行自己不能充当 anchor', () => {
    const grouping = groupOperationFeedback(rows(
      { id: 'f1', executedExecutionId: 'exec-a', feedback: feedback('exec-a', 'confirmed', '已确认变更') },
      { id: 'f2', feedback: feedback('exec-a', 'undone', '已撤销本次执行') }
    ));
    expect(grouping.byAnchor.size).toBe(0);
    expect(grouping.unattached.has('f2')).toBe(true);
  });

  it('执行身份只从授权过的字段取，标题/位置/顺序都不算', () => {
    expect(operationFeedbackAnchorIds({ id: 'x' })).toEqual([]);
    expect(operationFeedbackAnchorIds({
      id: 'x', executedExecutionId: 'e1', officeExecutionId: 'e1',
      changeReviewExecutionId: 'e2', verification: { executionId: 'e3' }
    })).toEqual(['e1', 'e2', 'e3']);
  });
});

describe('副本清理失败按真实副本身份去重', () => {
  const issues: ReviewCopyCleanupIssue[] = [
    { identity: 's1:EB_aaaaaaaaaaaaaaaaaaaa', label: '副本1', reason: '窗口仍在' },
    { identity: 's1:EB_bbbbbbbbbbbbbbbbbbbb', label: '副本2', reason: '窗口仍在' },
    { identity: 's1:EB_aaaaaaaaaaaaaaaaaaaa', label: '副本1', reason: '重试仍未关闭' }
  ];

  it('相同错误文本、不同副本 ⇒ 仍是两份', () => {
    const deduped = dedupeReviewCopyCleanupIssues(issues);
    expect(deduped).toHaveLength(2);
    expect(deduped.map((item) => item.identity)).toEqual([
      's1:EB_aaaaaaaaaaaaaaaaaaaa', 's1:EB_bbbbbbbbbbbbbbbbbbbb'
    ]);
    expect(deduped[0].reason).toBe('重试仍未关闭');
  });

  it('拿不到真实身份时不合并，宁可多报', () => {
    const deduped = dedupeReviewCopyCleanupIssues([
      { identity: '', reason: '同一句错误' },
      { identity: '', reason: '同一句错误' }
    ]);
    expect(deduped).toHaveLength(2);
  });

  it('摘要给数量、保留逐份原因，且不把工作簿恢复说成失败', () => {
    const summary = copyCleanupSummary({ executionId: 'exec-a', issues });
    expect(summary.count).toBe(2);
    expect(summary.headline).toBe('展示副本未清理 · 2 份');
    expect(summary.text).toContain('工作簿的恢复结果另见本次执行记录');
    expect(summary.text).not.toContain('已放弃并恢复');
    expect(summary.issues).toHaveLength(2);
  });

  it('同一次执行的多条反馈收敛成一行，副本原因按身份合并', () => {
    const digest = compactOperationFeedback([
      feedback('exec-a', 'discarded', '已放弃并恢复：主表已按本次恢复点还原。'),
      { ...feedback('exec-a', 'discarded', '重复的一条'), },
      { ...feedback('exec-a', 'copy_cleanup_failed', '展示副本未清理'), issues: [issues[0], issues[2]] },
      { ...feedback('exec-a', 'copy_cleanup_failed', '展示副本未清理'), issues: [issues[1]] }
    ]);
    expect(digest.text).toContain('已放弃并恢复');
    expect(digest.text).toContain('展示副本未清理 · 2 份');
    expect(digest.records).toHaveLength(2);
    expect(digest.records[1].issues).toHaveLength(2);
    expect(digest.text.match(/已放弃并恢复/g)).toHaveLength(1);
  });
});
