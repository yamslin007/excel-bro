import { expect, it, vi } from 'vitest';
import { prepareFormulaSourceEvidence } from './excelFormulaSourceEvidence';
import type { AnalysisPlan } from './contracts';

function plan(): AnalysisPlan {
  return { id:'formula', title:'formula', summary:'formula', assumptions:[], warnings:[],
    sourceDependencySheets:['forged'],
    actions:[{type:'writeFormulas',sheet:'Target',range:'B2',formulas:[['=Source!B5*3']]}] };
}
it('attests actual formula sources only after the local read completes', async () => {
  const p = plan();
  const read = vi.fn(async () => 'fingerprint');
  const result = await prepareFormulaSourceEvidence(p, ['Source','Target'], read, () => true);
  expect(read).toHaveBeenCalledExactlyOnceWith(['Source']);
  expect(result.sourceDependencySheets).toEqual(['Source']);
  expect(result.sourceDependencyFingerprint).toBe('fingerprint');
  expect(p.sourceDependencySheets).toEqual(['forged']);
});
it('unauthorized sources are rejected before reading', async () => {
  const read = vi.fn();
  await expect(prepareFormulaSourceEvidence(plan(), ['Target'], read, () => true)).rejects.toThrow('不在当前授权');
  expect(read).not.toHaveBeenCalled();
});
it('a failed read cannot produce attestation', async () => {
  const p = plan();
  await expect(prepareFormulaSourceEvidence(p, ['Source','Target'], async () => { throw Error('read failed'); }, () => true)).rejects.toThrow('read failed');
  expect(p.sourceDependencyFingerprint).toBeUndefined();
});
it('scope changes during reading cannot produce a baseline', async () => {
  const current = vi.fn().mockReturnValueOnce(true).mockReturnValueOnce(false);
  await expect(prepareFormulaSourceEvidence(plan(), ['Source','Target'], async () => 'fp', current)).rejects.toThrow('读取期间');
});
it('never replaces planning evidence for mixed fixed-value plans', async () => {
  const p = plan();
  p.actions.push({type:'writeValues',sheet:'Target',range:'C2',values:[[8]]});
  const read = vi.fn();
  expect(await prepareFormulaSourceEvidence(p, ['Source','Target'], read, () => true)).toBe(p);
  expect(read).not.toHaveBeenCalled();
});
