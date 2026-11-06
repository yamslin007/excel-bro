// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { cleanup, fireEvent, render, screen } from '@testing-library/react';
import type { ResultReviewState } from '@excel-bro/core/resultReview';
import { ResultReviewBar, resultReviewComposerActive } from './ResultReviewBar';

afterEach(cleanup);
const state: ResultReviewState = { phase: 'reviewing', key: 'task', executionId: 'execution', pending: false, comparing: false, notice: null };
const callbacks = () => ({ onCompare: vi.fn(), onConfirm: vi.fn(), onDiscard: vi.fn(), onRetryMarker: vi.fn() });

it('对比成功后才显示选表，收起和展开不触发对比或审阅', () => {
  const actions = callbacks(); const onComparisonSheet = vi.fn(); const onExitComparison = vi.fn();
  const initial = {...state, comparisonSheets:['结果1','结果2']};
  const view = render(<ResultReviewBar inComposer state={initial} title="多表" {...actions} onComparisonSheet={onComparisonSheet} onExitComparison={onExitComparison}/>);
  expect(screen.queryByRole('button',{name:'结果1'})).toBeNull();
  fireEvent.click(screen.getByRole('button',{name:'双视图对比'}));
  expect(actions.onCompare).toHaveBeenCalledTimes(1);
  view.rerender(<ResultReviewBar inComposer state={{...initial,notice:'副本无法打开'}} title="多表" {...actions} onComparisonSheet={onComparisonSheet}/>);
  expect(screen.getByRole('status').textContent).toContain('副本无法打开');
  expect(screen.queryByRole('button',{name:'结果1'})).toBeNull();
  view.rerender(<ResultReviewBar inComposer state={{...initial,comparing:true,comparisonSheet:'结果1'}} title="多表" {...actions} onComparisonSheet={onComparisonSheet} onExitComparison={onExitComparison}/>);
  expect(screen.getByRole('button',{name:'结果1'})).toBeTruthy();
  fireEvent.click(screen.getByRole('button',{name:'退出双视图'}));
  expect(onExitComparison).toHaveBeenCalledTimes(1);
  fireEvent.click(screen.getByRole('button',{name:'收起选表'}));
  expect(screen.queryByRole('button',{name:'结果1'})).toBeNull();
  fireEvent.click(screen.getByRole('button',{name:'展开选表'}));
  expect(screen.getByRole('button',{name:'结果1'})).toBeTruthy();
  expect(actions.onCompare).toHaveBeenCalledTimes(1);
  expect(actions.onConfirm).not.toHaveBeenCalled();
  expect(actions.onDiscard).not.toHaveBeenCalled();
});

it('长表名可搜索并完整选择，摘要不重复任务，确认仍针对全部修改', () => {
  const longName = '华东区域销售数据2026年第四季度汇总表';
  const actions = callbacks(); const onComparisonSheet = vi.fn();
  const { container } = render(<ResultReviewBar inComposer title="原始任务说明" state={{...state, comparing:true, comparisonSheets: ['结果1', '结果2', '结果3', longName], comparisonSheet:'结果1'}} {...actions} onComparisonSheet={onComparisonSheet}/>);
  expect(screen.getByText('待审阅 · 4 张表')).toBeTruthy();
  expect(screen.getByText('原始任务说明').closest('details')?.open).toBe(false);
  expect(screen.queryByText('请查看表格，选择保留或撤销本次修改。')).toBeNull();
  fireEvent.change(screen.getByRole('textbox', {name:'搜索对比工作表'}), {target:{value:'第四季度'}});
  fireEvent.click(screen.getByRole('button', {name:longName}));
  expect(onComparisonSheet).toHaveBeenCalledExactlyOnceWith(longName);
  expect((screen.getByRole('textbox') as HTMLInputElement).value).toBe('第四季度');
  expect(actions.onConfirm).not.toHaveBeenCalled();
  expect(container.querySelector('.result-review-actions')?.closest('.result-review-composer-scroll')).toBeNull();
});

describe('输入框结果审阅', () => {
  it('说明默认折叠，错误保持可见，动作不提交表单', () => {
    const actions = callbacks();
    const submit = vi.fn(e => e.preventDefault());
    const { container } = render(<form onSubmit={submit}><ResultReviewBar inComposer state={{ ...state, notice: '目标已变化，请返回原工作簿。' }} title="三步公式" {...actions} /></form>);
    expect(container.querySelector('details')?.open).toBe(false);
    expect(screen.getByRole('status').textContent).toContain('目标已变化');
    expect(screen.getByRole('status').closest('.result-review-composer-scroll')).toBeNull();
    fireEvent.click(screen.getByRole('button', { name: '确认全部变更' }));
    expect(actions.onConfirm).toHaveBeenCalledTimes(1);
    expect(actions.onDiscard).not.toHaveBeenCalled();
    expect(submit).not.toHaveBeenCalled();
    expect(container.querySelector('.result-review-actions')?.parentElement).toBe(container.querySelector('.result-review-composer'));
  });
  it('不支持对比时不画假入口，忙碌时禁止重复处理', () => {
    const actions = callbacks();
    render(<ResultReviewBar inComposer state={{ ...state, pending: true }} title="公式" compareUnavailable="没有执行前副本" {...actions} />);
    expect(screen.queryByRole('button', { name: '双视图对比' })).toBeNull();
    expect(screen.getByText('没有执行前副本').closest('details')).toBeNull();
    expect(screen.getByText('处理中，请稍候…').closest('.result-review-composer-scroll')).toBeNull();
    fireEvent.click(screen.getByRole('button', { name: '放弃全部变更' }));
    fireEvent.click(screen.getByRole('button', { name: '确认全部变更' }));
    expect(actions.onDiscard).not.toHaveBeenCalled();
    expect(actions.onConfirm).not.toHaveBeenCalled();
  });
  it('收尾失败仅重试标记清理，成功后退出操作区', () => {
    const actions = callbacks();
    const failed = { ...state, phase: 'confirmed' as const, markerPending: true, notice: '标记未清理' };
    const view = render(<ResultReviewBar inComposer state={failed} title="公式" {...actions} />);
    expect(resultReviewComposerActive(failed, false, false)).toBe(true);
    expect(screen.queryByRole('button', { name: '确认全部变更' })).toBeNull();
    fireEvent.click(screen.getByRole('button', { name: '重试清理标记' }));
    expect(actions.onRetryMarker).toHaveBeenCalledTimes(1);
    const done = { ...failed, markerPending: false, notice: null };
    view.rerender(<ResultReviewBar inComposer state={done} title="公式" {...actions} />);
    expect(view.container.textContent).toBe('');
    expect(resultReviewComposerActive(done, false, false)).toBe(false);
  });
  it('澄清和输入法组合期间不卸载输入框，恢复后接管', () => {
    expect(resultReviewComposerActive(state, true, false)).toBe(false);
    expect(resultReviewComposerActive(state, false, true)).toBe(false);
    expect(resultReviewComposerActive(state, false, false)).toBe(true);
    expect(resultReviewComposerActive({ ...state, phase: 'failed' }, false, false)).toBe(false);
  });
});

it('逐表选择只派发对比请求，忙碌时不允许切换', () => {
  const onComparisonSheet=vi.fn(); const actions=callbacks();
  const view=render(<ResultReviewBar inComposer title="多表" state={{...state,comparing:true,comparisonSheets:['结果1','结果2'],comparisonSheet:'结果1'}} {...actions} onComparisonSheet={onComparisonSheet}/>);
  fireEvent.click(screen.getByRole('button',{name:'结果2'}));
  expect(onComparisonSheet).toHaveBeenCalledExactlyOnceWith('结果2');
  expect(actions.onConfirm).not.toHaveBeenCalled();
  view.rerender(<ResultReviewBar inComposer title="多表" state={{...state,comparing:true,pending:true,comparisonSheets:['结果1','结果2']}} {...actions} onComparisonSheet={onComparisonSheet}/>);
  expect((screen.getByRole('button',{name:'结果2'}) as HTMLButtonElement).disabled).toBe(true);
});

it('shows manual pane sizing guidance only for an active comparison that needs it', () => {
  const actions = callbacks();
  const view = render(<ResultReviewBar inComposer state={state} title="结果" {...actions} paneNeedsResize />);
  expect(screen.queryByText(/拖动窗格边界/)).toBeNull();
  view.rerender(<ResultReviewBar inComposer state={{...state,comparing:true}} title="结果" {...actions} paneNeedsResize />);
  expect(screen.getByText(/拖动窗格边界/)).toBeTruthy();
  view.rerender(<ResultReviewBar inComposer state={{...state,comparing:true}} title="结果" {...actions} paneNeedsResize={false} />);
  expect(screen.queryByText(/拖动窗格边界/)).toBeNull();
});

for (const [host, scope] of [
  ['Excel', '执行前单元格值、基础格式与合并区域；公式冻结，不包含图表与条件格式。'],
  ['WPS', '基础格式与列宽已恢复；图表、条件格式来自打开副本时的原表。'],
]) it(`${host} 常规范围和布局帮助只在默认折叠详情中，错误保持在主操作区`, () => {
  const actions = callbacks();
  const { container, rerender } = render(<ResultReviewBar inComposer paneNeedsResize title="多表任务"
    state={{...state, comparing:true, comparisonHelp:{scope, layout:'可手动排列窗口。'}}} {...actions} />);
  const scopeNode = screen.getByText(scope);
  expect(scopeNode.closest('details')?.open).toBe(false);
  expect(scopeNode.closest('details')?.querySelector('summary')?.textContent).toBe('比对范围');
  expect(scopeNode.closest('details')?.parentElement?.closest('details')?.open).toBe(false);
  expect(screen.getByText(/宿主未能自动调窄/).closest('details')?.open).toBe(false);
  expect(screen.queryByRole('status')).toBeNull();
  fireEvent.click(screen.getByRole('button',{name:'确认全部变更'}));
  fireEvent.click(screen.getByRole('button',{name:'放弃全部变更'}));
  expect(actions.onConfirm).toHaveBeenCalledTimes(1);
  expect(actions.onDiscard).toHaveBeenCalledTimes(1);
  for (const notice of ['工作簿身份不匹配', '缺少执行前快照', '切换工作表失败', '恢复失败']) {
    rerender(<ResultReviewBar inComposer paneNeedsResize title="任务" state={{...state, comparing:true, notice, comparisonHelp:{scope}}} {...actions}/>);
    expect(screen.getByRole('status').textContent).toBe(notice);
    expect(screen.getByRole('status').closest('details')).toBeNull();
  }
  expect(container.querySelector('.result-review-actions')?.closest('.result-review-composer-scroll')).toBeNull();
});

it('放弃恢复成功后的独立清理和记账重试不再显示放弃按钮',()=>{
 const retryCopy=vi.fn(),retryDecision=vi.fn(),actions=callbacks();
 const view=render(<ResultReviewBar title="任务" state={{...state,phase:'discarded',decisionPending:'discarded',copyCleanup:{count:1,headline:'展示副本未关闭',issues:[{identity:'copy',reason:'identity_lost'}]},notice:'记账失败'}} {...actions} onRetryCopy={retryCopy} onRetryDecision={retryDecision}/>);
 expect(view.queryByRole('button',{name:'放弃全部变更'})).toBeNull();
 expect((view.getByRole('button',{name:'重试清理副本'}) as HTMLButtonElement).disabled).toBe(true);
 fireEvent.click(view.getByRole('button',{name:'重试记录结果'}));
 expect(retryDecision).toHaveBeenCalledTimes(1);
 expect(actions.onDiscard).not.toHaveBeenCalled();
});

it('独立错误清掉 notice 后，副本清理清单仍可见且可重试',()=>{
 const retryCopy=vi.fn();
 const view=render(<ResultReviewBar title="已恢复" state={{...state,phase:'discarded',notice:null,copyCleanup:{count:1,headline:'展示副本未关闭',issues:[{identity:'copy',reason:'user_saved'}]}}} {...callbacks()} onRetryCopy={retryCopy}/>);
 expect(view.getByRole('region',{name:'展示副本清理结果'})).toBeTruthy();
 fireEvent.click(view.getByRole('button',{name:'重试清理副本'}));
 expect(retryCopy).toHaveBeenCalledTimes(1);
});
