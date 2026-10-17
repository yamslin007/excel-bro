import { describe, expect, it } from 'vitest';
import capabilities from '../../../config/capabilities.json';
import {
  describeCopyDiscardSafety, planTempFileCleanup, reviewCopyCleanupState, reviewCopyCreationSummary, validateTempFilePath,
  type ReviewCopyBusinessOutcome, type ReviewCopyCloseOutcome
} from './reviewCopyLifecycle';

const ROOT = 'C:/Users/tester/AppData/Local/Temp/excel-bro-review';
const EVERY_CLOSE: ReviewCopyCloseOutcome[] = ['closed', 'already_closed', 'close_unconfirmed',
  'user_saved', 'identity_lost', 'ambiguous', 'unsupported', 'failed', 'not_attempted',
  'extra_edits', 'edits_unverified'];

describe('审阅副本生命周期 · 业务结果与副本清理结果互不冒充', () => {
  it('已写回并验证 ⇒ 任何关闭结果下都不会被改写成"执行失败"', () => {
    for (const copy of EVERY_CLOSE) {
      const state = reviewCopyCleanupState({ business: 'verified', copy, tempFilePath: null });
      // 业务结论始终在场，且顺序在副本清理之前。
      expect(state.copy.headline.startsWith('已写回并验证')).toBe(true);
      expect(state.copy.headline).not.toContain('执行失败');
      // 关不掉时明确说"请手动关闭"，而不是含糊其辞。
      if (copy === 'failed') expect(state.copy.headline).toContain('未能自动关闭，请手动关闭');
    }
  });
  it('已放弃变更 ⇒ 任何关闭结果下都保留"原表未修改"这一业务事实', () => {
    for (const copy of EVERY_CLOSE) {
      const state = reviewCopyCleanupState({ business: 'discarded', copy, tempFilePath: null });
      expect(state.copy.headline.startsWith('已放弃变更，原表未修改')).toBe(true);
    }
  });
  it('写回失败或结果不确定时保留副本不自动丢弃排查材料', () => {
    for (const business of ['write_failed', 'partial_or_unknown'] as const) {
      // 就算适配器声称"已经关掉了"，也不许把排查材料说成已清理。
      for (const copy of EVERY_CLOSE) {
        const state = reviewCopyCleanupState({ business, copy, tempFilePath: null });
        expect(state.copy.outcome).toBe('not_attempted');
        expect(state.copy.headline).toContain('已保留审阅副本供你核对');
        expect(state.copy.headline).not.toContain('已关闭');
        expect(state.copy.retryable).toBe(false);
        expect(state.copy.needsUserChoice).toBe(true);
      }
      // 文件清理也不许在"后果未确认"时报成功。
      expect(reviewCopyCleanupState({ business, copy: 'closed', file: 'deleted', tempFilePath: `${ROOT}/x.tmp` }).file.outcome)
        .toBe('out_of_scope');
    }
  });
  it('未知的关闭结果按失败处理绝不报已关闭', () => {
    for (const bogus of [undefined, null, 'weird', ''] as unknown[]) {
      const state = reviewCopyCleanupState({ business: 'verified', copy: bogus as ReviewCopyCloseOutcome, tempFilePath: null });
      expect(state.copy.outcome).toBe('failed');
      expect(state.copy.headline).toContain('未能自动关闭，请手动关闭');
    }
  });
  it('只有确认收尾的关闭结果才允许接着声称磁盘文件已清理', () => {
    const allowed: ReviewCopyCloseOutcome[] = ['closed', 'already_closed', 'close_unconfirmed'];
    for (const copy of EVERY_CLOSE) {
      const file = reviewCopyCleanupState({ business: 'verified', copy, file: 'deleted', tempFilePath: `${ROOT}/x.tmp` }).file;
      expect(file.outcome).toBe(allowed.includes(copy) ? 'deleted' : 'out_of_scope');
    }
  });
  it('可重试 / 需用户决定的结果不会被包装成"已完成"', () => {
    const retryable = reviewCopyCleanupState({ business: 'verified', copy: 'failed', tempFilePath: null });
    expect(retryable.copy.retryable).toBe(true);
    expect(retryable.copy.needsUserChoice).toBe(true);
    // 关闭待确认与"副本里有额外修改"都必须保留重试/确认入口：副本还在，用户还要再处理。
    for (const copy of ['close_unconfirmed', 'extra_edits', 'edits_unverified'] as const) {
      const state = reviewCopyCleanupState({ business: 'discarded', copy, tempFilePath: null });
      expect(state.copy.retryable).toBe(true);
      expect(state.copy.needsUserChoice).toBe(true);
    }
    // 用户已自行关闭 / 用户另存 都是终态：不应再兜售"重试关闭"。
    for (const copy of ['already_closed', 'user_saved', 'ambiguous', 'unsupported'] as const) {
      expect(reviewCopyCleanupState({ business: 'discarded', copy, tempFilePath: null }).copy.retryable).toBe(false);
    }
    // 用户另存的文件永远不自动丢弃：必须明确交还给用户决定。
    expect(reviewCopyCleanupState({ business: 'verified', copy: 'user_saved', tempFilePath: null }).copy.needsUserChoice).toBe(true);
    expect(reviewCopyCleanupState({ business: 'verified', copy: 'user_saved', tempFilePath: null }).copy.headline).toContain('未自动关闭');
  });
});

describe('第三项判断：副本里没有额外人工修改（身份可信 + 未另存 都推不出它）', () => {
  it('逐格读回一致才放行；确认有改动 ⇒ extra_edits；读不到 ⇒ edits_unverified', () => {
    expect(describeCopyDiscardSafety('unchanged')).toBeNull();
    expect(describeCopyDiscardSafety('changed')).toBe('extra_edits');
    expect(describeCopyDiscardSafety('unknown')).toBe('edits_unverified');
  });
  it('"确认被改过"与"读不到读数"必须分成两档，不许互相冒充', () => {
    const changed = reviewCopyCleanupState({ business: 'discarded', copy: describeCopyDiscardSafety('changed')!, tempFilePath: null });
    const unknown = reviewCopyCleanupState({ business: 'discarded', copy: describeCopyDiscardSafety('unknown')!, tempFilePath: null });
    expect(changed.copy.outcome).toBe('extra_edits');
    expect(unknown.copy.outcome).toBe('edits_unverified');
    expect(changed.copy.headline).toContain('未经确认的人工修改');
    expect(unknown.copy.headline).toContain('无法确认');
    for (const state of [changed, unknown]) {
      // 副本还在 ⇒ 绝不许说"已关闭"，且必须把决定权交回用户。
      expect(state.copy.headline).not.toContain('已关闭');
      expect(state.copy.outcome).not.toBe('closed');
      expect(state.copy.needsUserChoice).toBe(true);
      // 业务结论照旧在任何关闭结果下都成立。
      expect(state.copy.headline.startsWith('已放弃变更，原表未修改')).toBe(true);
    }
  });
  it('第三项没通过时不许接着声称磁盘文件已清理', () => {
    for (const copy of ['extra_edits', 'edits_unverified'] as const) {
      expect(reviewCopyCleanupState({ business: 'verified', copy, file: 'deleted', tempFilePath: `${ROOT}/x.tmp` }).file.outcome)
        .toBe('out_of_scope');
    }
  });
});

describe('关闭尚未确认：窗口还在但拿不到对话框证据时不推断"保存提示"', () => {
  it('close_unconfirmed 只说"未确认"，不声称宿主在问是否保存', () => {
    for (const business of ['verified', 'discarded'] as const) {
      const state = reviewCopyCleanupState({ business, copy: 'close_unconfirmed', tempFilePath: null });
      expect(state.copy.outcome).toBe('close_unconfirmed');
      expect(state.copy.headline).toContain('关闭尚未确认');
      // 反向锚点：不许把"窗口仍然存在"解释成宿主正在询问是否保存。
      expect(state.copy.headline).not.toContain('询问是否保存');
      expect(state.copy.detail).not.toContain('正在询问是否保存');
      expect(state.copy.detail).toContain('对话框证据');
      // 也绝不许说成已关闭，或把副本问题说成业务执行失败。
      expect(state.copy.headline).not.toContain('副本已关闭');
      expect(state.copy.headline).not.toContain('执行失败');
    }
  });
  it('四种关闭结论各自可辨：成功 / 待确认 / 失败 / 此前已关闭', () => {
    const headlineOf = (copy: ReviewCopyCloseOutcome) =>
      reviewCopyCleanupState({ business: 'verified', copy, tempFilePath: null }).copy.headline;
    expect(headlineOf('closed')).toContain('审阅副本已关闭');
    expect(headlineOf('close_unconfirmed')).toContain('关闭尚未确认');
    expect(headlineOf('failed')).toContain('未能自动关闭');
    expect(headlineOf('already_closed')).toContain('此前已由你关闭');
    // 四种结论的文案互不重合：不许用同一句话糊弄四种完全不同的处境。
    const headlines = (['closed', 'close_unconfirmed', 'failed', 'already_closed'] as const).map(headlineOf);
    expect(new Set(headlines).size).toBe(4);
  });
  it('关闭待确认不算业务失败：原表结论一个字都不改', () => {
    const state = reviewCopyCleanupState({ business: 'verified', copy: 'close_unconfirmed', tempFilePath: null });
    expect(state.copy.headline.startsWith('已写回并验证')).toBe(true);
    expect(state.copy.outcome).not.toBe('failed');
  });
});

describe('审阅副本生命周期 · 磁盘临时文件', () => {
  it('没有登记的临时文件时不得声称已删除文件', () => {
    const state = reviewCopyCleanupState({ business: 'verified', copy: 'closed', tempFilePath: null });
    expect(state.file.outcome).toBe('none');
    expect(state.file.headline).toBe('本次没有需要清理的磁盘文件。');
    expect(state.file.detail).toContain('未保存工作簿');
    // 反向锚点：绝不许出现"已删除/已清理临时文件"这种编造的成功。
    expect(state.file.headline).not.toContain('已删除');
    expect(state.file.headline).not.toContain('已清理本次登记');
    // 真正做清理决策的那个函数也不许把"没有文件"当成"删过文件"。
    expect(planTempFileCleanup({ tempFilePath: null, tempRoot: ROOT })).toBe('none');
    // 正向锚点：确实登记过、路径合法的文件才会被判为可删（否则上面只是"什么都返回 none"）。
    expect(planTempFileCleanup({ tempFilePath: `${ROOT}/mine.tmp`, tempRoot: ROOT, registered: true })).toBe('deleted');
  });
  it('被占用 ⇒ 保留待清理，不强杀宿主；未确认归属 ⇒ 拒绝清理', () => {
    const inUse = reviewCopyCleanupState({ business: 'verified', copy: 'closed', file: 'in_use', tempFilePath: `${ROOT}/review-9f.tmp` });
    expect(inUse.file.outcome).toBe('in_use');
    expect(inUse.file.headline).toContain('保留待清理');
    expect(inUse.file.detail).toContain('不会强杀宿主进程');
    expect(inUse.file.label).toBe('review-9f.tmp');
    expect(inUse.file.retryable).toBe(true);
  });
  it('专用临时目录之外的路径一律拒绝清理（越界不是"清理失败"，是"不该动"）', () => {
    const outside = reviewCopyCleanupState({ business: 'verified', copy: 'closed', file: 'out_of_scope', tempFilePath: 'C:/Users/tester/Documents/important.xlsx', tempRoot: ROOT });
    expect(outside.file.outcome).toBe('out_of_scope');
    expect(outside.file.headline).toContain('已拒绝清理');
    expect(outside.file.retryable).toBe(false);
  });
});

describe('专用临时目录的路径判据', () => {
  it('磁盘清理拒绝越出专用临时目录', () => {
    // 三个方向都试一遍：同前缀的兄弟目录、用户目录、以及用 `..` 爬出去。
    for (const candidate of [
      'C:/Users/tester/AppData/Local/Temp/excel-bro-review-evil/b.tmp',
      'C:/Users/tester/Documents/important.xlsx',
      `${ROOT}/../escape.tmp`,
      `${ROOT}/a/../../escape.tmp`
    ]) {
      expect(validateTempFilePath(ROOT, candidate).ok).toBe(false);
    }
    // 正向锚点：目录内真的文件必须通过，否则上一条断言只是"什么都拒绝"。
    expect(validateTempFilePath(ROOT, `${ROOT}/mine.tmp`).ok).toBe(true);
  });

  it('接受目录内的本机绝对路径（含反斜杠写法）', () => {
    for (const candidate of [`${ROOT}/a/b.tmp`, `${ROOT}\\a\\b.tmp`, `${ROOT}/b.tmp`]) {
      const result = validateTempFilePath(ROOT, candidate);
      expect(result.ok).toBe(true);
      expect(result.normalized?.startsWith(ROOT)).toBe(true);
    }
  });
  it.each([
    ['', 'empty'],
    ['   ', 'empty'],
    ['b.tmp', 'not-absolute-local-path'],
    ['//server/share/b.tmp', 'not-absolute-local-path'],
    ['/tmp/b.tmp', 'not-absolute-local-path'],
    [`${ROOT}/a\u0000.tmp`, 'control-characters'],
    ['C:/Users/tester/AppData/Local/Temp/excel-bro-review-evil/b.tmp', 'outside-root'],
    ['C:/Users/tester/Documents/b.tmp', 'outside-root'],
    [`${ROOT}`, 'outside-root'],
    [`${ROOT}/../escape.tmp`, 'traversal'],
    [`${ROOT}/a/../../escape.tmp`, 'traversal']
  ])('拒绝 %s（%s）', (candidate, reason) => {
    const result = validateTempFilePath(ROOT, candidate);
    expect(result.ok).toBe(false);
    expect(result.reason).toBe(reason);
    expect(result.normalized).toBeNull();
  });
  it('根目录本身非法时同样拒绝（不能把用户目录当成专用临时目录）', () => {
    expect(validateTempFilePath('C:/', `${ROOT}/b.tmp`).reason).toBe('invalid-root');
    expect(validateTempFilePath('relative/root', `${ROOT}/b.tmp`).reason).toBe('invalid-root');
    expect(validateTempFilePath('//server/share', `${ROOT}/b.tmp`).reason).toBe('invalid-root');
  });
  it('清理决策：未登记不删、越界不删、占用保留、否则才删', () => {
    expect(planTempFileCleanup({ tempFilePath: null, tempRoot: ROOT })).toBe('none');
    expect(planTempFileCleanup({ tempFilePath: `${ROOT}/b.tmp`, tempRoot: ROOT, registered: false })).toBe('out_of_scope');
    expect(planTempFileCleanup({ tempFilePath: 'C:/Users/tester/Documents/b.tmp', tempRoot: ROOT, registered: true })).toBe('out_of_scope');
    expect(planTempFileCleanup({ tempFilePath: `${ROOT}/b.tmp`, tempRoot: ROOT, registered: true, locked: true })).toBe('in_use');
    expect(planTempFileCleanup({ tempFilePath: `${ROOT}/b.tmp`, tempRoot: ROOT, registered: true })).toBe('deleted');
  });
});

describe('两端共用：实际创建了什么 / 能不能自动关闭', () => {
  it('三种副本都不是磁盘文件，且都说得出真正的关闭途径', () => {
    for (const kind of ['excel-native', 'excel-compat', 'wps-native'] as const) {
      const summary = reviewCopyCreationSummary(kind);
      expect(summary.hasDiskFile).toBe(false);
      expect(summary.created).toContain('未保存');
      expect(summary.canClose.trim().length).toBeGreaterThan(0);
    }
  });
  it('兼容副本不再宣称"不能自动关闭"，而是本机桥按身份关闭', () => {
    const compat = reviewCopyCreationSummary('excel-compat');
    expect(compat.canClose).toContain('本地桥');
    expect(compat.canClose).toContain('WM_CLOSE');
    expect(compat.canClose).not.toContain('不能自动关闭');
    // 兼容副本读不回来 ⇒ 不许宣称"已确认没有额外修改"。
    expect(compat.canClose).toContain('无法回读');
    const native = reviewCopyCreationSummary('excel-native');
    expect(native.canClose).toContain('Window.close');
    // 原生关闭也必须先说清第三项判断。
    expect(native.canClose).toContain('逐格回读');
    const wps = reviewCopyCreationSummary('wps-native');
    expect(wps.canClose).toContain('Close(false)');
    expect(wps.canClose).toContain('extra_edits');
    expect(wps.canClose).toContain('edits_unverified');
  });
  it('限额来自能力配置，且只列真正被消费的项（未消费的不当已实现能力）', () => {
    expect(capabilities.reviewCopyLifecycle.closeWaitSeconds).toBeGreaterThan(0);
    expect(capabilities.reviewCopyLifecycle.closePollIntervalMs).toBeGreaterThan(0);
    // 只保留服务端真正消费的两个键：不再挂着"自动重试上限""临时清理等待"这类空头承诺。
    expect(Object.keys(capabilities.reviewCopyLifecycle).sort()).toEqual(['closePollIntervalMs', 'closeWaitSeconds']);
  });
});
