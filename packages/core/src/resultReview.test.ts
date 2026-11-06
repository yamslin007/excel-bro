import { describe, expect, it, vi } from 'vitest';
import { ResultReview, ReviewNotStartedError, type ResultReviewOperations, type ResultReviewReceipt } from './resultReview';

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
it('accepts a positive host no-start proof, while an ordinary execution error remains uncertain',async()=>{
  for(const proved of [false,true]) {
    const {review,operations}=setup();
    vi.mocked(operations.execute).mockRejectedValueOnce(proved ? new ReviewNotStartedError('package unavailable') : new Error('receipt unavailable'));
    expect(await review.start('request-1',operations)).toBe(false);
    expect(review.getSnapshot().phase).toBe(proved ? 'not_started' : 'failed');
    expect(await review.start('request-1',operations)).toBe(false);
    expect(operations.execute).toHaveBeenCalledTimes(1);
  }
});

it('退出只收起，切表再回来复用原副本，确认成功才关闭全部副本', async () => {
  const {review, operations} = setup();
  operations.comparisonSheets = () => ['A','B'];
  operations.showBefore = vi.fn(async (_before, _receipt, sheet) => `copy-${sheet}`);
  operations.activateComparison = vi.fn(async () => {});
  operations.suspendComparison = vi.fn(async () => null);
  await review.start('r', operations);
  await review.compare('r','A');
  expect(await review.exitComparison('r')).toBe(true);
  expect(review.getComparisonCopy()).toBeUndefined();
  expect(operations.closeBefore).not.toHaveBeenCalled();
  await review.compare('r','A');
  await review.compare('r','B');
  await review.compare('r','A');
  expect(operations.showBefore).toHaveBeenCalledTimes(2);
  expect(review.getComparisonCopy()).toBe('copy-A');
  expect(await review.confirm('r')).toBe(true);
  expect(operations.closeBefore).toHaveBeenCalledWith('copy-A');
  expect(operations.closeBefore).toHaveBeenCalledWith('copy-B');
  expect(operations.execute).toHaveBeenCalledTimes(1);
});

it('多子表共用一个副本，切表不退出或重新排列；持久恢复后仍只清理一次', async () => {
  const {review, operations, receipt} = setup();
  operations.comparisonSheets = () => ['A', 'B', 'C'];
  operations.switchComparison = vi.fn(async () => true);
  operations.suspendComparison = vi.fn(async () => null);
  operations.activateComparison = vi.fn(async () => {});
  operations.saveComparisons = vi.fn(async () => {});
  await review.start('r', operations);
  await review.compare('r', 'A');
  const frames: boolean[] = [];
  const unsubscribe = review.subscribe(() => frames.push(review.getSnapshot().comparing));
  await review.compare('r', 'C'); await review.compare('r', 'A');
  expect(frames.every(Boolean)).toBe(true);
  unsubscribe();
  expect(operations.showBefore).toHaveBeenCalledTimes(1);
  expect(operations.activateComparison).toHaveBeenCalledTimes(1);
  expect(operations.suspendComparison).not.toHaveBeenCalled();
  expect(operations.saveComparisons).toHaveBeenCalledTimes(1);
  await review.exitComparison('r'); await review.compare('r', 'B');
  expect(operations.closeBefore).not.toHaveBeenCalled();
  expect(operations.showBefore).toHaveBeenCalledTimes(1);
  const restored = setup().review;
  operations.loadComparisons = async () => [['A', 'copy']];
  await restored.adoptCompleted('restored', 'before', receipt, operations);
  expect(await restored.compare('restored', 'C')).toBe(true);
  expect(operations.showBefore).toHaveBeenCalledTimes(1);
  expect(await restored.confirm('restored')).toBe(true);
  expect(operations.closeBefore).toHaveBeenCalledTimes(1);
  expect(operations.execute).toHaveBeenCalledTimes(1);
});

it('多子表切换失败保留副本与审阅，不新建副本或重复执行', async () => {
  const {review, operations} = setup();
  operations.comparisonSheets = () => ['A', 'C'];
  operations.suspendComparison = vi.fn(async () => null);
  operations.switchComparison = vi.fn(async () => {throw new Error('副本目标缺失');});
  await review.start('r', operations); await review.compare('r', 'A');
  expect(await review.compare('r', 'C')).toBe(false);
  expect(review.getSnapshot()).toMatchObject({phase: 'reviewing', comparisonSheet: 'A', notice: '副本目标缺失'});
  expect(review.getComparisonCopy()).toBe('copy');
  expect(operations.showBefore).toHaveBeenCalledTimes(1);
  expect(operations.closeBefore).not.toHaveBeenCalled();
  expect(operations.execute).toHaveBeenCalledTimes(1);
});

it('旧逐表副本升级为一份多子表副本，旧句柄不丢失且只在决定成功后清理', async () => {
  const {review, operations, receipt} = setup();
  operations.comparisonSheets = () => ['A','B','C'];
  operations.loadComparisons = async () => [['A','legacy-A'], ['B','legacy-B']];
  operations.suspendComparison = vi.fn(async () => null);
  operations.switchComparison = vi.fn(async copy => copy === 'copy');
  operations.saveComparisons = vi.fn(async () => {});
  await review.adoptCompleted('r', 'before', receipt, operations);
  await review.compare('r', 'A'); await review.compare('r', 'C');
  expect(operations.showBefore).toHaveBeenCalledTimes(1);
  expect(operations.closeBefore).not.toHaveBeenCalled();
  expect(operations.saveComparisons).toHaveBeenCalledWith(receipt, expect.arrayContaining([
    ['A','copy'], ['B','legacy-B'], [expect.stringMatching(/^__legacy__:/),'legacy-A']
  ]));
  await review.confirm('r');
  expect(operations.closeBefore).toHaveBeenCalledTimes(3);
  expect(operations.closeBefore).toHaveBeenCalledWith('legacy-A');
  expect(operations.closeBefore).toHaveBeenCalledWith('legacy-B');
  expect(operations.closeBefore).toHaveBeenCalledWith('copy');
  expect(operations.execute).not.toHaveBeenCalled();
});

it('恢复或审阅落账失败保留副本；重试落账不重复恢复，成功才清理', async () => {
  const {review, operations} = setup();
  operations.suspendComparison = vi.fn(async () => null);
  const restore = vi.fn(async () => {}), settle = vi.fn(async () => {});
  operations.restore = restore; operations.settle = settle;
  await review.start('r',operations); await review.compare('r');
  restore.mockRejectedValueOnce(new Error('现场变化'));
  expect(await review.discard('r')).toBe(false);
  expect(operations.closeBefore).not.toHaveBeenCalled();
  settle.mockRejectedValueOnce(new Error('离线'));
  await review.discard('r');
  expect(operations.closeBefore).not.toHaveBeenCalled();
  expect(await review.retryDecision('r')).toBe(true);
  expect(restore).toHaveBeenCalledTimes(2);
  expect(operations.closeBefore).toHaveBeenCalledTimes(1);
});

it('重开审阅接回所有已保留副本，不重建或执行；未查看的副本也在确认后清理', async () => {
  const {review, operations, receipt} = setup();
  operations.comparisonSheets = () => ['A','B'];
  operations.suspendComparison = async () => null;
  operations.activateComparison = vi.fn(async () => {});
  operations.loadComparisons = async () => [['A','retained-A'], ['B','retained-B']];
  operations.saveComparisons = vi.fn(async () => {});
  await review.adoptCompleted('restored','before',receipt,operations);
  expect(await review.compare('restored','B')).toBe(true);
  expect(operations.activateComparison).toHaveBeenCalledWith('retained-B');
  expect(operations.showBefore).not.toHaveBeenCalled();
  expect(operations.execute).not.toHaveBeenCalled();
  await review.confirm('restored');
  expect(operations.closeBefore).toHaveBeenCalledWith('retained-A');
  expect(operations.closeBefore).toHaveBeenCalledWith('retained-B');
  expect(operations.saveComparisons).toHaveBeenLastCalledWith(receipt, []);
});

it('旧副本身份读不到只阻止重新比对，不抹掉审阅和放弃入口', async () => {
  const {review, operations, receipt} = setup();
  operations.loadComparisons = async () => { throw new Error('旧副本身份无法核实'); };
  await review.adoptCompleted('restored','before',receipt,operations);
  expect(review.getSnapshot().phase).toBe('reviewing');
  expect(await review.compare('restored')).toBe(false);
  expect(operations.showBefore).not.toHaveBeenCalled();
  expect(await review.discard('restored')).toBe(true);
  expect(operations.restore).toHaveBeenCalledTimes(1);
});
describe('标记清理失败是可重试状态，不是业务失败', () => {
  it('刷新后新原型方法可处理保留状态的旧实例，安全关闭原副本', async () => {
    const {review, operations} = setup();
    await review.start('request-1', operations);
    await review.compare('request-1');
    const old = Object.create(Object.getPrototypeOf(review));
    old.exitComparison = vi.fn(async () => false);
    Object.setPrototypeOf(review, old);
    expect(await ResultReview.prototype.exitComparison.call(review, 'request-1')).toBe(true);
    expect(old.exitComparison).not.toHaveBeenCalled();
    expect(operations.closeBefore).toHaveBeenCalledTimes(1);
    expect(review.getSnapshot().phase).toBe('reviewing');
    expect(operations.execute).toHaveBeenCalledTimes(1);
  });
  it('退出对比只关闭受控副本；关闭受阻保留副本和审阅，之后可重试', async () => {
    const {review, operations} = setup();
    await review.start('request-1', operations);
    await review.compare('request-1');
    vi.mocked(operations.closeBefore).mockResolvedValueOnce('副本已被修改，请手动核对。');
    expect(await review.exitComparison('request-1')).toBe(false);
    expect(review.getSnapshot().comparing).toBe(true);
    expect(review.getSnapshot().notice).toContain('副本已被修改');
    expect(await review.exitComparison('request-1')).toBe(true);
    expect(review.getSnapshot().phase).toBe('reviewing');
    expect(review.getSnapshot().comparing).toBe(false);
    expect(operations.execute).toHaveBeenCalledTimes(1);
    expect(operations.restore).not.toHaveBeenCalled();
    expect(operations.accepted).not.toHaveBeenCalled();
    await review.compare('request-1');
    expect(operations.showBefore).toHaveBeenCalledTimes(2);
  });
  it('退出时身份不符或没有受控副本，明确提示且不触碰任何工作簿', async () => {
    const {review, operations} = setup();
    await review.start('request-1', operations);
    expect(await review.exitComparison('old')).toBe(false);
    expect(review.getSnapshot().notice).toContain('审阅任务已变化');
    expect(await review.exitComparison('request-1')).toBe(false);
    expect(review.getSnapshot().notice).toContain('安全关闭记录');
    expect(operations.closeBefore).not.toHaveBeenCalled();
    expect(operations.accepted).not.toHaveBeenCalled();
    expect(operations.restore).not.toHaveBeenCalled();
  });
  it('结果核对失败保留具体区域原因，不误导重试清理标记，也不接受或重写', async () => {
    const { review, operations } = setup();
    operations.validateResult = vi.fn(async () => { throw new Error('区域「结果4!B2」的公式与写前、写后都不一致。'); });
    await review.start('request-1', operations);
    expect(await review.confirm('request-1')).toBe(false);
    expect(review.getSnapshot().notice).toContain('结果4!B2');
    expect(review.getSnapshot().notice).not.toContain('重试清理');
    expect(review.getSnapshot().markerPending).toBe(false);
    expect(review.getSnapshot().phase).toBe('reviewing');
    expect(operations.accepted).not.toHaveBeenCalled();
    expect(operations.execute).toHaveBeenCalledTimes(1);
  });
  it('旧任务按钮不能静默失败，也不能接受当前任务', async () => {
    const { review, operations, clearMarker } = setup();
    await review.start('request-1', operations);
    expect(await review.confirm('old-request')).toBe(false);
    expect(review.getSnapshot().notice).toContain('此按钮对应的任务已变化');
    expect(clearMarker).not.toHaveBeenCalled();
    expect(operations.accepted).not.toHaveBeenCalled();
    expect(operations.execute).toHaveBeenCalledTimes(1);
    expect(await review.confirm('request-1')).toBe(true);
  });
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
    expect(operations.showBefore).toHaveBeenCalledExactlyOnceWith('before', receipt, undefined);
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

describe('adopt recovered result without execution',()=>{
  it.each(['confirm','discard'] as const)('adopts a failed review and supports %s without preparing/executing again',async action=>{
    const {review,operations,receipt}=setup();
    vi.mocked(operations.execute).mockRejectedValueOnce(new Error('interrupted'));
    await review.start('old',operations);
    expect(await review.adoptCompleted('recovered','original-before',receipt,operations)).toBe(true);
    await review.compare('recovered');
    expect(operations.showBefore).toHaveBeenCalledWith('original-before',receipt, undefined);
    await review[action]('recovered');
    expect(operations.prepare).toHaveBeenCalledTimes(1);expect(operations.execute).toHaveBeenCalledTimes(1);
    expect(review.getSnapshot().phase).toBe(action==='confirm'?'confirmed':'discarded');
    expect(await review.adoptCompleted('recovered','before',receipt,operations)).toBe(false);
  });
  it('unverified result cannot expose review actions',async()=>{
    const {review,operations,receipt}=setup();
    expect(await review.adoptCompleted('bad','before',{...receipt,verified:false},operations)).toBe(false);
    expect(await review.confirm('bad')).toBe(false);expect(operations.mark).not.toHaveBeenCalled();
    expect(operations.execute).not.toHaveBeenCalled();expect(operations.prepare).not.toHaveBeenCalled();
  });
});

describe('preparation refusal versus uncertain execution', () => {
  it('a preflight refusal reports not started and never enters execute or review', async () => {
    const {review,operations}=setup();
    vi.mocked(operations.prepare).mockRejectedValue(new Error('source anchor missing'));
    await review.start('preflight',operations);
    expect(review.getSnapshot().phase).toBe('not_started');
    expect(operations.execute).not.toHaveBeenCalled();
    expect(operations.mark).not.toHaveBeenCalled();
  });
  it('an exception after entering execute remains failed, never claimed zero write', async () => {
    const {review,operations}=setup();
    vi.mocked(operations.execute).mockRejectedValue(new Error('host response lost'));
    await review.start('in-flight',operations);
    expect(review.getSnapshot().phase).toBe('failed');
    expect(operations.execute).toHaveBeenCalledOnce();
  });
});


describe('逐表结果对比', () => {
  it('切换关闭前一个副本，只取所选表快照，确认仍只接受整次结果', async () => {
    const {review,operations}=setup();
    operations.comparisonSheets=()=>['结果1','结果2','结果3'];
    const opened:string[]=[];
    operations.showBefore=vi.fn(async (_before,_receipt,sheet)=>{opened.push(sheet!);return sheet!;});
    await review.start('multi',operations);
    expect(await review.compare('multi','结果1')).toBe(true);
    expect(await review.compare('multi','结果2')).toBe(true);
    expect(operations.closeBefore).toHaveBeenCalledWith('结果1');
    expect(await review.compare('multi','未授权')).toBe(false);
    expect(opened).toEqual(['结果1','结果2']);
    expect(await review.confirm('multi')).toBe(true);
    expect(operations.closeBefore).toHaveBeenLastCalledWith('结果2');
    expect(operations.execute).toHaveBeenCalledOnce();
    expect(operations.restore).not.toHaveBeenCalled();
  });
  it('副本关闭不明时不切换，不遗失句柄；排列失败也保留可关闭副本', async () => {
    const {review,operations}=setup();
    operations.comparisonSheets=()=>['A','B'];
    operations.activateComparison=vi.fn().mockRejectedValueOnce(new Error('排列失败')).mockResolvedValue(undefined);
    await review.start('multi',operations);
    expect(await review.compare('multi','A')).toBe(false);
    expect(await review.compare('multi','A')).toBe(true);
    expect(operations.showBefore).toHaveBeenCalledOnce();
    vi.mocked(operations.closeBefore).mockResolvedValueOnce('副本有人工修改');
    expect(await review.compare('multi','B')).toBe(false);
    expect(review.getSnapshot().comparisonSheet).toBe('A');
    expect(operations.showBefore).toHaveBeenCalledOnce();
    expect(await review.discard('multi')).toBe(true);
    expect(operations.restore).toHaveBeenCalledOnce();
    expect(operations.execute).toHaveBeenCalledOnce();
  });
});

describe('审阅决定必须落进持久记录（2026-11-03 刷新后待审阅消失）', () => {
  /*
   * 缺陷：审阅状态只活在 `ResultReview` 的内存里，刷新窗格后整段消失，而原表停在
   * 一个没人收尾的写入现场上。所以"确认/放弃"必须把事实写进执行账本；写不进去时
   * **不许**宣告处理完成，也不许重放任何写入 —— 只留一条可重试的持久化待办。
   */
  function setupWithSettle(settle: (decision: string) => Promise<void>) {
    const base = setup();
    const settleSpy = vi.fn(async (_receipt: unknown, decision: string) => settle(decision));
    base.operations.settle = settleSpy;
    return { ...base, settleSpy };
  }

  it('确认：先落账，再把阶段推进到已确认；执行回调一次都没有再跑', async () => {
    const { review, operations, settleSpy } = setupWithSettle(async () => {});
    await review.start('request-1', operations);
    expect(await review.confirm('request-1')).toBe(true);
    expect(review.getSnapshot().phase).toBe('confirmed');
    expect(review.getSnapshot().decisionPending).toBeNull();
    expect(settleSpy).toHaveBeenCalledWith(expect.anything(), 'confirmed');
    expect(operations.execute).toHaveBeenCalledTimes(1);
    expect(operations.accepted).toHaveBeenCalledTimes(1);
  });

  it('确认后落账失败：不谎报成功，保留可重试状态，重试只重发这条决定', async () => {
    let fail = true;
    const { review, operations, settleSpy } = setupWithSettle(async () => {
      if (fail) throw new Error('服务端 500');
    });
    await review.start('request-1', operations);
    expect(await review.confirm('request-1')).toBe(false);
    const snapshot = review.getSnapshot();
    // 宿主侧结论仍然成立（结果已被接受），但持久事实没写成 —— 两件事分开说。
    expect(snapshot.phase).toBe('confirmed');
    expect(snapshot.decisionPending).toBe('confirmed');
    expect(snapshot.notice).toContain('没有写进执行记录');

    fail = false;
    expect(await review.retryDecision('request-1')).toBe(true);
    expect(review.getSnapshot().decisionPending).toBeNull();
    expect(review.getSnapshot().notice).toBeNull();
    expect(settleSpy).toHaveBeenCalledTimes(2);
    // 重试只重发决定：没有再执行、没有再接受、没有再清标记。
    expect(operations.execute).toHaveBeenCalledTimes(1);
    expect(operations.accepted).toHaveBeenCalledTimes(1);
  });

  it('放弃：恢复已经发生，落账失败只留可重试项，重试不会二次恢复', async () => {
    let fail = true;
    const { review, operations } = setupWithSettle(async () => {
      if (fail) throw new Error('账本忙');
    });
    await review.start('request-1', operations);
    expect(await review.discard('request-1')).toBe(true);
    expect(review.getSnapshot().phase).toBe('discarded');
    expect(review.getSnapshot().decisionPending).toBe('discarded');
    expect(review.getSnapshot().notice).toContain('不会重复恢复');

    fail = false;
    expect(await review.retryDecision('request-1')).toBe(true);
    expect(review.getSnapshot().decisionPending).toBeNull();
    // 关键：重试**没有**再走恢复（那会把用户此后的编辑一起抹掉）。
    expect(operations.restore).toHaveBeenCalledTimes(1);
    expect(operations.execute).toHaveBeenCalledTimes(1);
  });

  it('没有持久化适配器时不假装写过：阶段照常推进，不产生待办', async () => {
    const { review, operations } = setup();
    await review.start('request-1', operations);
    expect(await review.confirm('request-1')).toBe(true);
    expect(review.getSnapshot().decisionPending).toBeNull();
  });

  it('外部撤销（settleExternally）也要把"已撤销"写进记录；失败可重试', async () => {
    let fail = true;
    const { review, operations } = setupWithSettle(async () => {
      if (fail) throw new Error('服务端 503');
    });
    await review.start('request-1', operations);
    expect(await review.settleExternally('execution-1')).toBe(true);
    expect(review.getSnapshot().phase).toBe('discarded');
    expect(review.getSnapshot().decisionPending).toBe('undone');
    fail = false;
    expect(await review.retryDecision('request-1')).toBe(true);
    expect(operations.restore).not.toHaveBeenCalled();
  });
});

describe('adoptCompleted 的恢复路径（刷新后重建审阅）', () => {
  function adoptOps() {
    const { operations } = setup();
    const { prepare: _prepare, execute: _execute, ...rest } = operations;
    return rest;
  }

  it('恢复路径上没有 execute：重建审阅**物理上**不能重跑计划', async () => {
    const review = new ResultReview<string, ResultReviewReceipt, string, string>();
    const ops = adoptOps();
    const ok = await review.adoptCompleted(
      'restored-review:execution-1',
      'before',
      { executionId: 'execution-1', verified: true, undoable: true },
      ops as never
    );
    expect(ok).toBe(true);
    expect(review.getSnapshot().phase).toBe('reviewing');
    expect((ops as { execute?: unknown }).execute).toBeUndefined();
  });

  it('没有执行身份的审阅不能摆出来', async () => {
    const review = new ResultReview<string, ResultReviewReceipt, string, string>();
    const ok = await review.adoptCompleted(
      'restored-review:x',
      'before',
      { executionId: '', verified: true, undoable: true },
      adoptOps() as never,
      { restored: true }
    );
    expect(ok).toBe(false);
    expect(review.getSnapshot().phase).toBe('failed');
  });

  it('`restored` 允许"恢复材料缺失"的现场进入审阅，但默认仍要自检', async () => {
    const receipt = { executionId: 'execution-1', verified: true, undoable: false };
    const strict = new ResultReview<string, ResultReviewReceipt, string, string>();
    expect(
      await strict.adoptCompleted('k1', 'before', receipt, adoptOps() as never)
    ).toBe(false);

    const restored = new ResultReview<string, ResultReviewReceipt, string, string>();
    expect(
      await restored.adoptCompleted('k2', 'before', receipt, adoptOps() as never, {
        restored: true
      })
    ).toBe(true);
  });
});

it('exposes only the active comparison to pane sizing, releasing it on exit and settlement', async () => {
  const {review, operations} = setup();
  expect(review.getComparisonCopy()).toBeUndefined();
  await review.start('r', operations);
  expect(review.getComparisonCopy()).toBeUndefined();
  await review.compare('r');
  expect(review.getComparisonCopy()).toBe('copy');
  await review.exitComparison('r');
  expect(review.getComparisonCopy()).toBeUndefined();
  await review.compare('r');
  expect(review.getComparisonCopy()).toBe('copy');
  await review.confirm('r');
  expect(review.getComparisonCopy()).toBeUndefined();
  expect(operations.execute).toHaveBeenCalledTimes(1);
});

describe('展示副本清理失败：按真实副本身份去重，不冒充恢复失败', () => {
  function multiCopySetup() {
    const built = setup();
    built.operations.comparisonSheets = () => ['A', 'B'];
    built.operations.showBefore = vi.fn(async (_before, _receipt, sheet) => `copy-${sheet}`);
    built.operations.suspendComparison = vi.fn(async () => null);
    built.operations.activateComparison = vi.fn(async () => {});
    built.operations.saveComparisons = vi.fn(async () => {});
    // 两份不同副本、同一句错误文本。副本身份由宿主给出（真实 nonce），不是错误文本。
    built.operations.copyIdentity = (copy) => `session-1:${copy}`;
    built.operations.closeBefore = vi.fn(async () => '窗口仍在');
    return built;
  }

  it('相同错误文本来自两份副本 ⇒ 仍算两份，提示只出现一次', async () => {
    const { review, operations } = multiCopySetup();
    await review.start('r', operations);
    await review.compare('r', 'A');
    await review.compare('r', 'B');
    expect(await review.confirm('r')).toBe(true);
    const snapshot = review.getSnapshot();
    expect(snapshot.phase).toBe('confirmed');
    expect(snapshot.copyCleanup?.count).toBe(2);
    expect(snapshot.copyCleanup?.issues.map((issue) => issue.identity)).toEqual([
      'session-1:copy-A', 'session-1:copy-B'
    ]);
    expect(snapshot.copyCleanup?.headline).toContain('展示副本未清理');
    expect(snapshot.notice).toContain('窗口仍在（2 份）');
    expect(snapshot.notice?.match(/窗口仍在/g)).toHaveLength(1);
    expect(snapshot.notice).not.toContain('恢复失败');
  });

  it('多个收尾入口重复触发也只更新同一份清单，不会越拼越长', async () => {
    const { review, operations } = multiCopySetup();
    await review.start('r', operations);
    await review.compare('r', 'A');
    await review.compare('r', 'B');
    await review.confirm('r');
    const first = review.getSnapshot().notice;
    expect(await review.settleExternally('execution-1')).toBe(true);
    expect(review.getSnapshot().copyCleanup?.count).toBe(2);
    expect(review.getSnapshot().notice).toBe(first);
    expect(review.getSnapshot().notice?.match(/窗口仍在/g)).toHaveLength(1);
  });

  it('副本真的关掉之后清掉清单，且不动其他提示', async () => {
    const { review, operations } = multiCopySetup();
    await review.start('r', operations);
    await review.compare('r', 'A');
    await review.compare('r', 'B');
    await review.confirm('r');
    vi.mocked(operations.closeBefore).mockResolvedValue(null);
    expect(await review.settleExternally('execution-1')).toBe(true);
    expect(review.getSnapshot().copyCleanup).toBeNull();
    expect(review.getSnapshot().notice).toBeNull();
  });

  it('业务结论与副本清理分开：确认成功仍为 confirmed', async () => {
    const { review, operations } = multiCopySetup();
    await review.start('r', operations);
    await review.compare('r', 'A');
    expect(await review.confirm('r')).toBe(true);
    expect(review.getSnapshot().phase).toBe('confirmed');
    expect(operations.execute).toHaveBeenCalledTimes(1);
  });
});

it('常规范围和布局帮助独立于操作错误，重试和切表不重放写入', async () => {
  const {review,operations}=setup();
  operations.comparisonSheets=()=>['A','B'];
  operations.switchComparison=vi.fn(async()=>true);
  operations.comparisonHelp=()=>({scope:'静态快照',layout:'手动排列'});
  operations.comparisonNotice=()=>null;
  await review.start('help',operations);
  await review.compare('help','A');
  expect(review.getSnapshot()).toMatchObject({notice:null,comparisonHelp:{scope:'静态快照',layout:'手动排列'}});
  vi.mocked(operations.switchComparison!).mockRejectedValueOnce(new Error('切表身份异常'));
  expect(await review.compare('help','B')).toBe(false);
  expect(review.getSnapshot().notice).toBe('切表身份异常');
  expect(operations.execute).toHaveBeenCalledTimes(1);
});

it('同一工作簿登记多个子表只关闭一次，失败重试绝不再次恢复', async () => {
 const {review,operations,receipt}=setup();
 operations.suspendComparison=async()=>null;
 operations.copyIdentity=()=> 'owned-workbook';
 operations.loadComparisons=async()=>[['结果1','copy'],['结果2','copy'],['结果3','copy']];
 operations.saveComparisons=vi.fn(async()=>{});
 operations.closeBefore=vi.fn(async()=> 'identity_lost' as string|null);
 await review.adoptCompleted('discard-copy','before',receipt,operations);
 const first=review.discard('discard-copy');
 expect(await review.discard('discard-copy')).toBe(false);
 expect(await first).toBe(true);
 expect(review.getSnapshot()).toMatchObject({phase:'discarded',copyCleanup:{count:1}});
 expect(operations.restore).toHaveBeenCalledTimes(1);
 expect(operations.closeBefore).toHaveBeenCalledTimes(1);
 vi.mocked(operations.closeBefore).mockResolvedValue(null);
 expect(await review.retryCopyCleanup('discard-copy')).toBe(true);
 expect(operations.closeBefore).toHaveBeenCalledTimes(2);
 expect(operations.restore).toHaveBeenCalledTimes(1);
 expect(review.getSnapshot()).toMatchObject({phase:'discarded',comparing:false,copyCleanup:null});
});
it('副本登记清理失败不把恢复成功改为放弃失败',async()=>{
 const {review,operations}=setup();
 operations.suspendComparison=async()=>null;
 const save=vi.fn(async()=>{}); operations.saveComparisons=save;
 await review.start('registry',operations); await review.compare('registry');
 save.mockRejectedValueOnce(new Error('登记存储失败'));
 expect(await review.discard('registry')).toBe(true);
 expect(review.getSnapshot()).toMatchObject({phase:'discarded',copyCleanup:{count:1}});
 expect(await review.retryCopyCleanup('registry')).toBe(true);
 expect(operations.restore).toHaveBeenCalledTimes(1);
});
