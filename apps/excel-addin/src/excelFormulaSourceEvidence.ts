import { formulaSourceEvidenceRequirement } from '@excel-bro/core/formulaWriteback';
import { attestSourceDependency } from '@excel-bro/core/planSourceDependency';
import type { AnalysisPlan } from './contracts';

/** First execution only. Never use a fresh read to replace a recovery baseline. */
export async function prepareFormulaSourceEvidence(
  plan: AnalysisPlan,
  authorization: readonly string[],
  readFingerprint: (sheets: string[]) => Promise<string>,
  isCurrent: () => boolean
): Promise<AnalysisPlan> {
  // Mixed/baked plans retain their original planning-time evidence.
  if (!plan.actions.length || !plan.actions.every(action => action.type === 'writeFormulas')) return plan;
  const requirement = formulaSourceEvidenceRequirement(plan);
  if (!requirement.ok || !requirement.sheets.length) return plan;
  const sheets = requirement.sheets.map(name => authorization.find(item => item.toLocaleLowerCase() === name.toLocaleLowerCase()));
  if (sheets.some(name => name === undefined)) throw new Error('公式依赖的来源表不在当前授权范围内；未读取来源、未执行写入。');
  const authorizedSources = sheets.filter((name): name is string => name !== undefined);
  if (!isCurrent()) throw new Error('工作簿或授权已变化；未读取公式来源、未执行写入。');
  const fingerprint = await readFingerprint(authorizedSources);
  if (!isCurrent()) throw new Error('读取期间工作簿或授权已变化；未登记来源、未执行写入。');
  if (!fingerprint.trim()) throw new Error('读不到公式依赖的来源值，无法登记来源证据。');
  // Attest only after the deterministic host read succeeds. No values go to the model.
  return { ...attestSourceDependency(plan, { origin: 'deterministic-local-read', sheets: authorizedSources }),
    sourceDependencyFingerprint: fingerprint };
}
