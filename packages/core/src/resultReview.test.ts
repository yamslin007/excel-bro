import { describe, expect, it, vi } from 'vitest';
import { ResultReview, type ResultReviewOperations, type ResultReviewReceipt } from './resultReview';

function setup() {
  const receipt = { executionId: 'execution-1', verified: true, undoable: true };
  const operations: ResultReviewOperations<string, ResultReviewReceipt, string, string> = {
    isCurrent: vi.fn(() => true), prepare: vi.fn(async () => 'before'),
    execute: vi.fn(async () => receipt), mark: vi.fn(async () => 'marker'),
    clearMarker: vi.fn(async () => {}), showBefore: vi.fn(async () => 'copy'),
    closeBefore: vi.fn(async () => null), restore: vi.fn(async () => {}), accepted: vi.fn()
  };
  return { review: new ResultReview<string, ResultReviewReceipt, string, string>(), operations, receipt };
}
describe('already executed result review', () => {
  it('disposes a native before-copy even when dual view was never opened', async () => {
    const { review, operations } = setup();
    operations.disposePrepared = vi.fn(async () => null);
    await review.start('request-1', operations);
    await review.confirm('request-1');
    expect(operations.disposePrepared).toHaveBeenCalledExactlyOnceWith('before');
    expect(operations.closeBefore).not.toHaveBeenCalled();
  });
  it('refuses confirmation if the real result changed after execution', async () => {
    const { review, operations } = setup();
    operations.validateResult = vi.fn(async () => { throw new Error('changed'); });
    await review.start('request-1', operations);
    expect(await review.confirm('request-1')).toBe(false);
    expect(operations.accepted).not.toHaveBeenCalled();
    expect(review.getSnapshot().phase).toBe('reviewing');
    expect(operations.execute).toHaveBeenCalledOnce();
  });
  it('executes once, displays the main-sheet result, and confirms without another write or copy', async () => {
    const { review, operations } = setup();
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
    const { review, operations } = setup();
    await review.start('request-1', operations); await review.compare('request-1');
    vi.mocked(operations.restore).mockRejectedValue(new Error('undo_conflict'));
    expect(await review.discard('request-1')).toBe(false);
    expect(review.getSnapshot()).toMatchObject({ phase: 'reviewing', notice: expect.stringContaining('undo_conflict') });
    expect(operations.closeBefore).not.toHaveBeenCalled();
  });
  it('never retries an interrupted execution on a render or another start call', async () => {
    const { review, operations } = setup();
    vi.mocked(operations.execute).mockRejectedValue(new Error('partial write'));
    await review.start('request-1', operations);
    await review.start('request-1', operations);
    expect(operations.execute).toHaveBeenCalledOnce();
    expect(review.getSnapshot().phase).toBe('failed');
    expect(await review.confirm('request-1')).toBe(false);
  });
  it('fails closed if the target changes during pre-write capture', async () => {
    const { review, operations } = setup();
    vi.mocked(operations.prepare).mockImplementation(async () => {
      vi.mocked(operations.isCurrent).mockReturnValue(false); return 'before';
    });
    await review.start('request-1', operations);
    expect(operations.execute).not.toHaveBeenCalled();
  });
  it('rejects stale card actions after a different execution', async () => {
    const { review, operations } = setup();
    await review.start('request-1', operations); await review.confirm('request-1');
    await review.start('request-2', operations);
    expect(await review.discard('request-1')).toBe(false);
    expect(operations.restore).not.toHaveBeenCalled();
  });
  it('a marker failure preserves the successful execution and its undo entrance', async () => {
    const { review, operations } = setup();
    vi.mocked(operations.mark).mockRejectedValue(new Error('unavailable'));
    await review.start('request-1', operations);
    expect(review.getSnapshot()).toMatchObject({ phase: 'reviewing', executionId: 'execution-1', notice: expect.stringContaining('颜色标记') });
    await review.discard('request-1');
    expect(operations.restore).toHaveBeenCalledOnce();
  });
  it('does not confirm if removing its marks fails; never clears unrelated formatting', async () => {
    const { review, operations } = setup();
    await review.start('request-1', operations);
    vi.mocked(operations.clearMarker).mockRejectedValue(new Error('mark changed'));
    expect(await review.confirm('request-1')).toBe(false);
    expect(review.getSnapshot().phase).toBe('reviewing');
    expect(operations.accepted).not.toHaveBeenCalled();
  });
  it('copy cleanup failure cannot undo the confirmed business outcome', async () => {
    const { review, operations } = setup();
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
    const { review, operations } = setup();
    await Promise.all([review.start('request-1', operations), review.start('request-1', operations)]);
    expect(operations.execute).toHaveBeenCalledOnce();
  });
});
