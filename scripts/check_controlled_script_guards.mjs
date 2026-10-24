import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const output = resolve(root, 'outputs/controlled-script-guards');
mkdirSync(output, { recursive: true });
const hash = value => createHash('sha256').update(value).digest('hex');
const run = (id, cwd, file, label) => {
  const result = spawnSync(process.execPath, [resolve(root, 'node_modules/vitest/vitest.mjs'), 'run', file], {
    cwd: resolve(root, cwd), encoding: 'utf8', timeout: 120000, env: { ...process.env, NO_COLOR: '1' }
  });
  const log = `${result.stdout ?? ''}\n${result.stderr ?? ''}`.replace(/\u001b\[[0-9;]*m/g, '');
  writeFileSync(resolve(output, `${id}-${label}.log`), log);
  if (result.error || result.signal || /No test files found|Tests\s+no tests|Failed Suites/.test(log)) throw new Error(`${id}: test infrastructure failed`);
  return { code: result.status, log };
};
const cases = [
  {
    id: 'scope', path: 'packages/core/src/controlledScript.ts', cwd: 'packages/core', file: 'src/controlledScript.test.ts',
    needle: 'if (!authorizedSheets.includes(action.sheet))', replacement: 'if (false)',
    evidence: 'rejects the entire compilation if any emitted target is unauthorized'
  },
  {
    id: 'multi-sheet', path: 'apps/wps-addin/src/wpsNativeComparison.ts', cwd: 'apps/wps-addin', file: 'src/wpsPlanExecutor.test.ts',
    needle: 'export function captureWpsBeforeResult(application: WpsApplication, prepared: WpsPreparedPlan): WpsBeforeResult {',
    replacement: 'export function captureWpsBeforeResult(application: WpsApplication, prepared: WpsPreparedPlan): WpsBeforeResult {\n  writeSheetOf(prepared);',
    evidence: 'controlled JS executes and restores two selected sheets'
  }
];
for (const item of cases) {
  const path = resolve(root, item.path);
  const original = readFileSync(path);
  const text = original.toString('utf8');
  if (text.split(item.needle).length !== 2) throw new Error(`${item.id}: unique seed missing`);
  const baseline = run(item.id, item.cwd, item.file, 'baseline');
  if (baseline.code !== 0) throw new Error(`${item.id}: baseline failed`);
  const mutated = Buffer.from(text.replace(item.needle, item.replacement));
  try {
    writeFileSync(path, mutated);
    const red = run(item.id, item.cwd, item.file, 'red');
    if (red.code === 0 || !/Tests\s+\d+ failed/.test(red.log) || !red.log.includes(item.evidence) || !red.log.includes('AssertionError')) throw new Error(`${item.id}: missing business failure evidence`);
  } finally {
    if (!readFileSync(path).equals(mutated)) throw new Error(`${item.id}: concurrent edit; refusing to overwrite it`);
    writeFileSync(path, original);
    if (hash(readFileSync(path)) !== hash(original)) throw new Error(`${item.id}: restoration mismatch`);
  }
  if (run(item.id, item.cwd, item.file, 'restored').code !== 0) throw new Error(`${item.id}: restored suite failed`);
  console.log(`${item.id}: RED; restored SHA256 ${hash(original)}; GREEN`);
}
