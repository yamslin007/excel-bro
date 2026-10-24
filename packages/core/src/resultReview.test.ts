import { describe, expect, it, vi } from 'vitest';
import { ResultReview, type ResultReviewOperations, type ResultReviewReceipt } from './resultReview';

function setup() {
  const receipt = { executionId: 'execution-1', verified: true, undoable: true };
  /*
   * `clearMarker` 单列一个 mock：接口标注会把字段收窄成普通函数，`mockRejectedValueOnce`
   * 就没了（测试要模拟"宿主拒绝删除"）。所以这里让 mock 活在接口之外，`operations`
   * 保持接口形状（既有测试要往里加 `disposePrepared` / `validateResult`）。
   */
  const clearMarker = vi.fn(async (_marker: string) => {});
  const operations: ResultReviewOperations<string, ResultReviewReceipt, string, string> = {
    isCurrent: vi.fn(() => true), prepare: vi.fn(async () => 'before'),
    execute: vi.fn(async () => receipt), mark: vi.fn(async () => 'marker'),
    clearMarker, showBefore: vi.fn(async () => 'copy'),
    closeBefore: vi.fn(async () => null), restore: vi.fn(async () => {}), accepted: vi.fn()
  };
  return { review: new ResultReview<string, ResultReviewReceipt, string, string>(), operations, clearMarker, receipt };
}
describe('标记清理失败是可重试状态，不是业务失败', () => {
  it('确认时清理失败：保持审阅中、标记待清；重试只清标记，不第二次写入', async () => {
    const { review, operations, clearMarker } = setup();
    await review.start('request-1', operations);
    clearMarker.mockRejectedValueOnce(new Error('宿主拒绝删除规则'));

    expect(await review.confirm('request-1')).toBe(false);
    expect(review.getSnapshot().phase).toBe('reviewing');
    expect(review.getSnapshot().markerPending).toBe(true);
    // 确认没有继续跑到"接受结果"。
    expect(operations.accepted).not.toHaveBeenCalled();

    expect(await review.retryMarkerCleanup('request-1')).toBe(true);
    expect(review.getSnapshot().markerPending).toBe(false);
    // 全程只有一次执行：重试只清标记。
    expect(operations.execute).toHaveBeenCalledOnce();
    expect(operations.restore).not.toHaveBeenCalled();
  });

  it('放弃时标记清理失败：恢复仍是成功结论，两件事分开说', async () => {
    const { review, operations, clearMarker } = setup();
    await review.start('request-1', operations);
    clearMarker.mockRejectedValueOnce(new Error('宿主拒绝删除规则'));

    expect(await review.discard('request-1')).toBe(true);
    expect(review.getSnapshot().phase).toBe('discarded');
    expect(review.getSnapshot().markerPending).toBe(true);
    // 恢复确实发生了（不是被标记失败拦下）。
    expect(operations.restore).toHaveBeenCalledOnce();
    // 提示里同时包含"已恢复"与"标记未移除"，不把成功改成失败。
    expect(review.getSnapshot().notice).toContain('主表已恢复到执行前');
    expect(review.getSnapshot().notice).toContain('审阅标记未能全部移除');
  });

  it('外部撤销同步：清标记并复位待清状态', async () => {
    const { review, operations, clearMarker } = setup();
    await review.start('request-1', operations);
    expect(await review.settleExternally('execution-1')).toBe(true);
    expect(operations.clearMarker).toHaveBeenCalledOnce();
    expect(review.getSnapshot().markerPending).toBe(false);
    expect(operations.restore).not.toHaveBeenCalled();
  });
});

/*
 * 当前性不成立 ≠ 标记清理失败。
 *
 * 真机症状：双视图把副本激活后（或用户切到别的簿）确认被拦下，界面却说
 * "可重试清理标记" —— 那是把"目标不对"说成了"收尾没做完"，用户会一直重试
 * 一件根本不会成功的事。这里把两类提示钉成互斥的。
 */
describe('当前性不成立与标记清理失败是两回事', () => {
  it('确认时当前性不成立：给出原因，不显示可重试清理标记', async () => {
    const { review, operations, clearMarker } = setup();
    await review.start('request-1', operations);
    // 标记确实建起来了，所以"标记还在"不能被当成"可以重试清理"。
    expect(review.getSnapshot().marked).toBe(true);
    vi.mocked(operations.isCurrent).mockReturnValue(false);

    expect(await review.confirm('request-1')).toBe(false);
    const notice = review.getSnapshot().notice ?? '';
    expect(notice).toContain('确认未完成');
    expect(notice).toContain('工作簿、授权或计划已变化');
    expect(notice).not.toContain('可重试清理标记');
    // 不是可重试状态 ⇒ 操作条不该给「重试清理标记」。
    expect(review.getSnapshot().markerPending).toBe(false);
    // 也确实没有动手：不清理、不接受、不恢复、不再执行。
    expect(operations.clearMarker).not.toHaveBeenCalled();
    expect(operations.accepted).not.toHaveBeenCalled();
    expect(operations.restore).not.toHaveBeenCalled();
    expect(operations.execute).toHaveBeenCalledOnce();
    expect(review.getSnapshot().phase).toBe('reviewing');
  });

  it('回到原表后再确认：同一份审阅可以正常收尾，不重新执行', async () => {
    const { review, operations, clearMarker } = setup();
    await review.start('request-1', operations);
    vi.mocked(operations.isCurrent).mockReturnValueOnce(false);
    expect(await review.confirm('request-1')).toBe(false);
    // 用户回到本次执行的原表：当前性恢复，再点一次即可收尾。
    expect(await review.confirm('request-1')).toBe(true);
    expect(review.getSnapshot().phase).toBe('confirmed');
    expect(operations.clearMarker).toHaveBeenCalledOnce();
    expect(operations.execute).toHaveBeenCalledOnce();
    expect(operations.accepted).toHaveBeenCalledOnce();
  });
});

describe('从其他入口撤销后的同步', () => {
  it('撤销已确认的结果：操作条转入终态，且不再次调用恢复', async () => {
    const { review, operations, clearMarker } = setup();
    await review.start('request-1', operations);
    await review.confirm('request-1');
    expect(review.getSnapshot().phase).toBe('confirmed');

    expect(await review.settleExternally('execution-1')).toBe(true);
    expect(review.getSnapshot().phase).toBe('discarded');
    // 撤销已经由宿主完成：这里再调一次 restore 就是二次恢复。
    expect(operations.restore).not.toHaveBeenCalled();
  });

  it('审阅中从结果卡撤销：标记被清理、按钮撤下', async () => {
    const { review, operations, clearMarker } = setup();
    await review.start('request-1', operations);
    expect(await review.settleExternally('execution-1')).toBe(true);
    expect(review.getSnapshot().phase).toBe('discarded');
    expect(operations.clearMarker).toHaveBeenCalledOnce();
    expect(operations.restore).not.toHaveBeenCalled();
  });

  it('executionId 对不上时不动：别的执行被撤销不能收掉这一份审阅', async () => {
    const { review, operations, clearMarker } = setup();
    await review.start('request-1', operations);
    expect(await review.settleExternally('execution-other')).toBe(false);
    expect(review.getSnapshot().phase).toBe('reviewing');
  });
});

describe('already executed result review', () => {
  it('disposes a native before-copy even when dual view was never opened', async () => {
    const { review, operations, clearMarker } = setup();
    operations.disposePrepared = vi.fn(async () => null);
    await review.start('request-1', operations);
    await review.confirm('request-1');
    expect(operations.disposePrepared).toHaveBeenCalledExactlyOnceWith('before');
    expect(operations.closeBefore).not.toHaveBeenCalled();
  });
  it('refuses confirmation if the real result changed after execution', async () => {
    const { review, operations, clearMarker } = setup();
    operations.validateResult = vi.fn(async () => { throw new Error('changed'); });
    await review.start('request-1', operations);
    expect(await review.confirm('request-1')).toBe(false);
    expect(operations.accepted).not.toHaveBeenCalled();
    expect(review.getSnapshot().phase).toBe('reviewing');
    expect(operations.execute).toHaveBeenCalledOnce();
  });
  it('executes once, displays the main-sheet result, and confirms without another write or copy', async () => {
    const { review, operations, clearMarker } = setup();
    expect(await review.start('request-1', operations)).toBe(true);
    expect(review.getSnapshot().phase).toBe('reviewing');
    expect(operations.showBefore).not.toHaveBeenCalled();
    expect(operations.accepted).not.toHaveBeenCalled();
    expect(await review.confirm('request-1')).toBe(true);
    expect(operations.execute).toHaveBeenCalledOnce();
    expect(operations.clearMarker).toHaveBeenCalledExactlyOnceWith('marker');
    expect(operations.accepted).toHaveBeenCalledOnce();
    expect(review.getSnapshot().phase).toBe('confirmed');
    expect(await review.confirm('request-1')).toBe(false);
  });
  it('dual view uses the pre-write snapshot and never executes again', async () => {
    const { review, operations, receipt } = setup();
    await review.start('request-1', operations);
    await review.compare('request-1');
    await review.compare('request-1');
    expect(operations.showBefore).toHaveBeenCalledExactlyOnceWith('before', receipt);
    expect(operations.execute).toHaveBeenCalledOnce();
    await review.confirm('request-1');
    expect(operations.closeBefore).toHaveBeenCalledExactlyOnceWith('copy');
  });
  it('discard invokes real conflict-aware undo before closing the before-copy', async () => {
    const { review, operations, receipt } = setup();
    await review.start('request-1', operations); await review.compare('request-1');
    await review.discard('request-1');
    expect(operations.restore).toHaveBeenCalledExactlyOnceWith(receipt);
    expect(vi.mocked(operations.restore).mock.invocationCallOrder[0]).toBeLessThan(vi.mocked(operations.closeBefore).mock.invocationCallOrder[0]);
    expect(review.getSnapshot().phase).toBe('discarded');
    expect(operations.accepted).not.toHaveBeenCalled();
  });
  it('an undo conflict keeps the result pending and preserves the comparison copy', async () => {
    const { review, operations, clearMarker } = setup();
    await review.start('request-1', operations); await review.compare('request-1');
    vi.mocked(operations.restore).mockRejectedValue(new Error('undo_conflict'));
    expect(await review.discard('request-1')).toBe(false);
    expect(review.getSnapshot()).toMatchObject({ phase: 'reviewing', notice: expect.stringContaining('undo_conflict') });
    expect(operations.closeBefore).not.toHaveBeenCalled();
  });
  it('never retries an interrupted execution on a render or another start call', async () => {
    const { review, operations, clearMarker } = setup();
    vi.mocked(operations.execute).mockRejectedValue(new Error('partial write'));
    await review.start('request-1', operations);
    await review.start('request-1', operations);
    expect(operations.execute).toHaveBeenCalledOnce();
    expect(review.getSnapshot().phase).toBe('failed');
    expect(await review.confirm('request-1')).toBe(false);
  });
  it('fails closed if the target changes during pre-write capture', async () => {
    const { review, operations, clearMarker } = setup();
    vi.mocked(operations.prepare).mockImplementation(async () => {
      vi.mocked(operations.isCurrent).mockReturnValue(false); return 'before';
    });
    await review.start('request-1', operations);
    expect(operations.execute).not.toHaveBeenCalled();
  });
  it('rejects stale card actions after a different execution', async () => {
    const { review, operations, clearMarker } = setup();
    await review.start('request-1', operations); await review.confirm('request-1');
    await review.start('request-2', operations);
    expect(await review.discard('request-1')).toBe(false);
    expect(operations.restore).not.toHaveBeenCalled();
  });
  it('a marker failure preserves the successful execution and its undo entrance', async () => {
    const { review, operations, clearMarker } = setup();
    vi.mocked(operations.mark).mockRejectedValue(new Error('unavailable'));
    await review.start('request-1', operations);
    expect(review.getSnapshot()).toMatchObject({ phase: 'reviewing', executionId: 'execution-1', notice: expect.stringContaining('颜色标记') });
    await review.discard('request-1');
    expect(operations.restore).toHaveBeenCalledOnce();
  });
  it('does not confirm if removing its marks fails; never clears unrelated formatting', async () => {
    const { review, operations, clearMarker } = setup();
    await review.start('request-1', operations);
    vi.mocked(operations.clearMarker).mockRejectedValue(new Error('mark changed'));
    expect(await review.confirm('request-1')).toBe(false);
    expect(review.getSnapshot().phase).toBe('reviewing');
    expect(operations.accepted).not.toHaveBeenCalled();
  });
  it('copy cleanup failure cannot undo the confirmed business outcome', async () => {
    const { review, operations, clearMarker } = setup();
    await review.start('request-1', operations); await review.compare('request-1');
    vi.mocked(operations.closeBefore).mockRejectedValue(new Error('copy kept'));
    expect(await review.confirm('request-1')).toBe(true);
    expect(review.getSnapshot()).toMatchObject({ phase: 'confirmed', notice: expect.stringContaining('copy kept') });
  });
  it('a failed readback does not offer confirmation or save-as-tool', async () => {
    const { review, operations, receipt } = setup(); receipt.verified = false;
    await review.start('request-1', operations);
    expect(review.getSnapshot().phase).toBe('failed');
    expect(operations.mark).not.toHaveBeenCalled();
    expect(operations.accepted).not.toHaveBeenCalled();
  });
  it('reserves the write synchronously against same-tick repeated calls', async () => {
    const { review, operations, clearMarker } = setup();
    await Promise.all([review.start('request-1', operations), review.start('request-1', operations)]);
    expect(operations.execute).toHaveBeenCalledOnce();
  });
});
