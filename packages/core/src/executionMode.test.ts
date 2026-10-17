import { describe, expect, it } from 'vitest';
import {
  COPY_COMPARE_EXECUTION_DETAIL,
  DIRECT_EXECUTION_DETAIL,
  DIRECT_EXECUTION_SKIPS,
  EXECUTION_MODE_PREFACE,
  PLAN_EXECUTION_MODE_LABELS,
  directExecutionStillRuns,
  executionModeOptions,
  executionModeSwitchDecision,
  executionModeTargetChanged,
  executionSubmitBlocked,
  modeSwitchBlockReason,
  type ExecutionModeAvailability,
  type ModeSwitchState
} from './executionMode';

const available: ExecutionModeAvailability = { 'copy-compare': { available: true, reason: null } };
const unavailable: ExecutionModeAvailability = {
  'copy-compare': { available: false, reason: '当前宿主缺少创建副本所需的基础 API。' }
};

function state(over: Partial<ModeSwitchState> = {}): ModeSwitchState {
  return { current: 'copy-compare', next: 'direct', copyExists: false, inFlight: false, editingPlan: false, ...over };
}

describe('执行方式选项', () => {
  it('两项都在，且文案区分「跳过什么」与「不跳过什么」', () => {
    const options = executionModeOptions('copy-compare', available);
    expect(options.map(option => option.mode)).toEqual(['direct', 'copy-compare']);
    expect(options[0].label).toBe(PLAN_EXECUTION_MODE_LABELS.direct);
    expect(options[1].label).toBe(PLAN_EXECUTION_MODE_LABELS['copy-compare']);
    // 直接执行必须点明「当前工作簿」而不是磁盘文件，且必须说清守卫仍在。
    expect(DIRECT_EXECUTION_DETAIL).toContain('当前工作簿');
    // 不许把「写回当前工作簿」说成「保存到磁盘文件」：用户说的"文件"就是这本工作簿。
    expect(DIRECT_EXECUTION_DETAIL).not.toMatch(/另存为|保存到文件|保存磁盘/);
    expect(DIRECT_EXECUTION_DETAIL).toContain('授权范围');
    expect(DIRECT_EXECUTION_DETAIL).toContain('回读');
    expect(DIRECT_EXECUTION_DETAIL).toContain('撤销');
    expect(COPY_COMPARE_EXECUTION_DETAIL).toContain('原表始终保持打开');
    expect(EXECUTION_MODE_PREFACE).toContain('白名单计划');
  });

  it('副本路径不可用时如实说原因，不静默消失', () => {
    const options = executionModeOptions('direct', unavailable);
    const copy = options.find(option => option.mode === 'copy-compare')!;
    expect(copy.available).toBe(false);
    expect(copy.blockedReason).toBe('当前宿主缺少创建副本所需的基础 API。');
    const direct = options.find(option => option.mode === 'direct')!;
    expect(direct.available).toBe(true);
    expect(direct.blockedReason).toBeNull();
  });
});

describe('直接执行只跳过副本比对', () => {
  it('除副本比对外的每一项都必须仍执行', () => {
    expect(directExecutionStillRuns()).toEqual({
      authorizationCheck: true,
      planValidation: true,
      preExecutionRecheck: true,
      recoverySnapshot: true,
      readbackVerification: true,
      undoAvailable: true
    });
    expect(DIRECT_EXECUTION_SKIPS).toEqual(['copy-comparison']);
  });
});

describe('切换判定', () => {
  it('副本生成前可以自由切换', () => {
    expect(modeSwitchBlockReason(state({ copyExists: false }), available)).toBeNull();
    expect(modeSwitchBlockReason(state({ current: 'direct', next: 'copy-compare' }), available)).toBeNull();
  });

  it('执行在途时禁止切换与重复提交', () => {
    const decision = executionModeSwitchDecision(state({ inFlight: true }), available);
    expect(decision.allowed).toBe(false);
    expect(decision.reason).toBe('in_flight');
    expect(decision.invalidatesReview).toBe(false);
    expect(executionSubmitBlocked({ inFlight: true })).toBe(true);
    expect(executionSubmitBlocked({ inFlight: false })).toBe(false);
  });

  it('计划编辑中不提供切换', () => {
    expect(modeSwitchBlockReason(state({ editingPlan: true }), available)).toBe('editing_plan');
  });

  it('切到当前方式不算切换', () => {
    expect(modeSwitchBlockReason(state({ next: 'copy-compare' }), available)).toBe('no_change');
  });

  it('副本路径不可用时切不进去', () => {
    const decision = executionModeSwitchDecision(state({ current: 'direct', next: 'copy-compare' }), unavailable);
    expect(decision.allowed).toBe(false);
    expect(decision.reason).toBe('copy_compare_unavailable');
  });

  it('副本已存在时切到直接执行 ⇒ 旧审阅失效，且交代旧副本怎么办', () => {
    const decision = executionModeSwitchDecision(state({ copyExists: true }), available);
    expect(decision.allowed).toBe(true);
    // 不失效就会同时留下两个有效写回入口。
    expect(decision.invalidatesReview).toBe(true);
    expect(decision.reviewNotice).toContain('失效');
    // 失效 ≠ 自动关闭：那份副本可能是用户自己要留的。
    expect(decision.reviewNotice).toContain('不会被自动关闭');
  });

  it('副本不在场时切到直接执行不需要失效任何东西', () => {
    expect(executionModeSwitchDecision(state({ copyExists: false }), available).invalidatesReview).toBe(false);
  });

  it('从直接执行切回副本不算失效（本来就没有审阅可失效）', () => {
    const decision = executionModeSwitchDecision(state({ current: 'direct', next: 'copy-compare', copyExists: true }), available);
    expect(decision.allowed).toBe(true);
    expect(decision.invalidatesReview).toBe(false);
  });
});

describe('写回请求的身份必须含执行方式', () => {
  const base = { review: {}, plan: {}, workbookIdentity: 'w1', mode: 'copy-compare' as const };

  it('同一现场同一方式 ⇒ 仍有效', () => {
    expect(executionModeTargetChanged(base, { ...base })).toBe(false);
  });

  it('方式变了就算变了（三元组全同也不放行）', () => {
    expect(executionModeTargetChanged(base, { ...base, mode: 'direct' })).toBe(true);
  });

  it('换审阅 / 换计划 / 换原工作簿都算变了', () => {
    expect(executionModeTargetChanged(base, { ...base, review: {} })).toBe(true);
    expect(executionModeTargetChanged(base, { ...base, plan: {} })).toBe(true);
    expect(executionModeTargetChanged(base, { ...base, workbookIdentity: 'w2' })).toBe(true);
  });

  it('任一侧缺席一律 fail-closed', () => {
    expect(executionModeTargetChanged(null, base)).toBe(true);
    expect(executionModeTargetChanged(base, null)).toBe(true);
  });
});