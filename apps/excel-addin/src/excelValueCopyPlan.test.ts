import { beforeEach, describe, expect, it, vi } from 'vitest';
import { prepareExcelValueCopies, valueCopyReviewGeometry } from './excelValueCopyPlan';
import { resultReviewPlanIssue, nativeComparisonPlanIssue } from './excelNativeComparison';
import type { AnalysisPlan } from './contracts';
import { captureWorkbookSourceFingerprint, currentOfficeWorkbookIdentity } from './excel';
vi.mock('./excel', async original => ({...await original<typeof import('./excel')>(),
  captureWorkbookSourceFingerprint:vi.fn(),currentOfficeWorkbookIdentity:vi.fn()}));
const plan:AnalysisPlan={id:'copy',title:'copy',summary:'copy',assumptions:[],warnings:[],actions:[{
  type:'copyRange',sheet:'Target',sourceSheet:'Source',sourceRange:'B2:C3',targetRange:'D4',copyType:'values',skipBlanks:false,transpose:false}]};
beforeEach(()=>{vi.clearAllMocks();vi.mocked(currentOfficeWorkbookIdentity).mockReturnValue('book');
  vi.mocked(captureWorkbookSourceFingerprint).mockResolvedValue('baseline');});
describe('value copy automatic review',()=>{
  it('admits only the result workflow, leaves legacy copy execution disabled and original plan intact',()=>{
    const original=JSON.stringify(plan);
    expect(resultReviewPlanIssue(plan,['Source','Target'])).toBeNull();
    expect(nativeComparisonPlanIssue(plan,['Target'])).not.toBeNull();
    expect(valueCopyReviewGeometry(plan,['Source','Target']).actions[0]).toMatchObject({range:'D4:E5'});
    expect(JSON.stringify(plan)).toBe(original);
  });
  it.each(['Source','Target'])('rejects missing authorization for %s', missing=>{
    expect(resultReviewPlanIssue(plan,['Source','Target'].filter(s=>s!==missing))).not.toBeNull();
  });
  it.each([
    {copyType:'all'}, {skipBlanks:true}, {sourceSheet:'Target'}, {targetRange:'D4:D5'},
    {targetRange:'XFD1048576'}, {sourceRange:'B:B'}
  ])('does not silently change copy semantics: %o',change=>{
    expect(resultReviewPlanIssue({...plan,actions:[{...plan.actions[0],...change}]} as AnalysisPlan,['Source','Target'])).not.toBeNull();
  });
  it('rejects overlapping targets',()=>{
    expect(resultReviewPlanIssue({...plan,actions:[plan.actions[0],plan.actions[0]]},['Source','Target'])).not.toBeNull();
  });
  it('does not refresh away an obsolete planning baseline',async()=>{
    await expect(prepareExcelValueCopies({...plan,sourceFingerprint:'old'},['Source','Target'],'book',()=>true)).rejects.toThrow('规划后发生变化');
  });
  it('binds source values, geometry and explicit readback criteria without modifying the source plan',async()=>{
    const sync=vi.fn(async()=>{}); const reads:string[]=[];
    const context = {sync,workbook:{worksheets:{getItem:(sheet:string)=>{
      reads.push(sheet);
      return {getRange:(address:string)=>{
        reads.push(address);
        return {load:vi.fn(),values:[[1,2],[3,4]],formulas:[[1,2],[3,4]]};
      }};
    }}}};
    vi.stubGlobal('Excel',{run:async(fn:(context:unknown)=>unknown)=>fn(context)});
    try {
      const output=await prepareExcelValueCopies(plan,['Source','Target'],'book',()=>true);
      expect(output.actions).toEqual([
        {type:'writeValues',sheet:'Target',range:'D4:E4',values:[[1,2]]},
        {type:'writeValues',sheet:'Target',range:'D5:E5',values:[[3,4]]}]);
      expect(output.acceptanceCriteria).toHaveLength(2);
      expect(output.sourceDependencySheets).toEqual(['Source']);
      expect(output.sourceFingerprintSheets).toEqual(['Source']);
      expect(reads).toEqual(['Source','B2:C3']);
      expect(plan.actions[0].type).toBe('copyRange');
      vi.mocked(captureWorkbookSourceFingerprint).mockResolvedValueOnce('baseline').mockResolvedValueOnce('changed');
      await expect(prepareExcelValueCopies(plan,['Source','Target'],'book',()=>true)).rejects.toThrow('读取期间');
    } finally {vi.unstubAllGlobals();}
  });
});
