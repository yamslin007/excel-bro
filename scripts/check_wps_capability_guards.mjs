import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createHash, randomUUID } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const output = resolve(root, 'outputs/wps-capability-guards');
mkdirSync(output, { recursive: true });
const hash = data => createHash('sha256').update(data).digest('hex');
const cases = [
  {
    "id": "ole-color",
    "path": "apps/wps-addin/src/wpsFormatCapability.ts",
    "frontend": true,
    "needle": "  return ((rgb & 0xff) << 16) | (rgb & 0xff00) | ((rgb >>> 16) & 0xff);",
    "replacement": "  return rgb;",
    "test": "src/wpsExtendedFormats.test.ts",
    "evidence": "RGB channels are not swapped"
  },
  {
    "id": "number-format-restore",
    "path": "apps/wps-addin/src/wpsFormatSnapshot.ts",
    "frontend": true,
    "needle": "if (fields.has(\"numberFormat\") && saved.numberFormat != null) cell.NumberFormat = saved.numberFormat;",
    "replacement": "if (fields.has(\"numberFormat\") && saved.numberFormat != null) void saved.numberFormat;",
    "test": "src/wpsExtendedFormats.test.ts",
    "evidence": "restores mixed formats"
  },
  {
    "id": "copy-source-change",
    "path": "apps/wps-addin/src/wpsPlanExecutor.ts",
    "frontend": true,
    "needle": "    if (current.fingerprint !== source.previewFingerprint) {",
    "replacement": "    if (false && current.fingerprint !== source.previewFingerprint) {",
    "test": "src/wpsCopyRange.test.ts",
    "evidence": "source changes after preparation"
  },
  {
    "id": "copy-skip-blanks",
    "path": "packages/core/src/copyRangePlan.ts",
    "frontend": true,
    "needle": "if(action.skipBlanks&&(value===null||value===\"\"))",
    "replacement": "if(false && action.skipBlanks&&(value===null||value===\"\"))",
    "test": "src/wpsCopyRange.test.ts",
    "evidence": "skips empty source cells"
  },
  {
    "id": "overlapping-data",
    "path": "apps/wps-addin/src/wpsPlanExecutor.ts",
    "frontend": true,
    "needle": "  if (isFormatActionType(left.type) || isFormatActionType(right.type)) return false;",
    "replacement": "  return false;",
    "test": "src/wpsWriteTable.test.ts",
    "evidence": "continues to reject overlapping data writes"
  },
  {
    "id": "copy-tool-template",
    "path": "apps/wps-addin/src/WpsAgentPanel.tsx",
    "frontend": true,
    "needle": "plan: prepared.originalPlan ?? prepared.plan,",
    "replacement": "plan: prepared.plan,",
    "test": "src/WpsAgentPanel.test.tsx",
    "evidence": "保存动态复制工具保留原动作"
  },
  {
    "id": "copy-plan-capability",
    "path": "server/app/excel_agent.py",
    "needle": "if action.copyType not in capabilities().get(\"wpsWrite\", {}).get(\"copyTypes\", []):",
    "replacement": "if False:",
    "test": "server/tests/test_planning_capability_narrowing.py",
    "evidence": "test_wps_copy_variants_are_validated_before_plan_submission"
  },
  {
  "id": "compute-presentation-change",
  "path": "apps/wps-addin/src/wpsPlanExecutor.ts",
  "frontend": true,
  "needle": "if (source.presentationFingerprint !== undefined && JSON.stringify(computationPresentation(",
  "replacement": "if (false && source.presentationFingerprint !== undefined && JSON.stringify(computationPresentation(",
  "test": "src/wpsComputedPlan.test.ts",
  "evidence": "changed source numberFormat"
},
  {
  "id": "checkpoint-source-change",
  "path": "apps/wps-addin/src/wpsPlanExecutor.ts",
  "frontend": true,
  "needle": "    assertSourcesCurrent();\n    firstWriteChecked = true;",
  "replacement": "    firstWriteChecked = true;",
  "test": "src/wpsComputedPlan.test.ts",
  "evidence": "source changes while the durable checkpoint"
},
  {
  "id": "checkpoint-target-change",
  "path": "apps/wps-addin/src/wpsPlanExecutor.ts",
  "frontend": true,
  "needle": "|| readMatrices(checkedRange, action.rowCount, action.columnCount).fingerprint !== action.beforeFingerprint)",
  "replacement": "|| false)",
  "test": "src/wpsComputedPlan.test.ts",
  "evidence": "rechecks target contents"
},
  {
  "id": "filtered-link-row",
  "path": "packages/core/src/copyRangePlan.ts",
  "frontend": true,
  "needle": "const sr=sourceRows[action.transpose?c:r]",
  "replacement": "const sr=action.transpose?c:r",
  "test": "src/wpsFilteredPlan.test.ts",
  "evidence": "preserves link addresses through transposition"
},
  {
  "id": "vertical-restore",
  "path": "apps/wps-addin/src/wpsFormatSnapshot.ts",
  "frontend": true,
  "needle": "if (fields.has(\"verticalAlignment\") && saved.verticalAlignment != null) cell.VerticalAlignment = saved.verticalAlignment;",
  "replacement": "if (fields.has(\"verticalAlignment\") && saved.verticalAlignment != null) void saved.verticalAlignment;",
  "test": "src/wpsVerticalAlignment.test.ts",
  "evidence": "restores mixed vertical positions"
},
{
  "id": "recovery-attempt-ownership",
  "path": "apps/wps-addin/src/wpsPlanExecutor.ts",
  "frontend": true,
  "needle": "if (!attemptedActions.has(action.actionIndex)) continue;",
  "replacement": "if (false && !attemptedActions.has(action.actionIndex)) continue;",
  "test": "src/wpsComputedPlan.test.ts",
  "evidence": "rechecks target contents"
},
{
  "id": "partial-overlap-evidence",
  "path": "apps/wps-addin/src/wpsPlanExecutor.ts",
  "frontend": true,
  "needle": "if (!interruptedHostWrite && action.afterFormatSnapshot) continue;",
  "replacement": "if (action.afterFormatSnapshot) continue;",
  "test": "src/wpsVerticalAlignment.test.ts",
  "evidence": "restores overlapping formats even when the second action stops halfway"
},
{
  id: "border-restore", path: "apps/wps-addin/src/wpsFormatSnapshot.ts", frontend: true,
  needle: "restoreWpsBorder(cell,field,saved[field]);",
  replacement: "void saved[field];",
  test: "src/wpsBorders.test.ts", evidence: "restores mixed automatic/theme/RGB borders"
},
{
  id: "border-alias", path: "apps/wps-addin/src/wpsBorders.ts", frontend: true,
  needle: "if(previous && previous!==owner)", replacement: "if(false && previous && previous!==owner)",
  test: "src/wpsBorders.test.ts", evidence: "rejects conflicting adjacent ownership"
},
{
  id: "border-mark-visibility", path: "packages/core/src/reviewMarkStyle.ts", frontend: true,
  needle: 'const borderOwned = overlapsType("setBorders");', replacement: 'const borderOwned = false;',
  test: "src/wpsResultReviewMarkers.test.ts", evidence: "does not overlay business borders"
}
];

function run(item, stage) {
  const command = item.frontend ? process.execPath : resolve(root, '.venv/Scripts/python.exe');
  const args = item.frontend
    ? [resolve(root, 'node_modules/vitest/vitest.mjs'), 'run', item.test]
    : ['-m', 'pytest', item.test, '-q', '--tb=short', '-p', 'no:cacheprovider',
       `--basetemp=${resolve(output, `tmp-${randomUUID()}`)}`];
  const result = spawnSync(command, args, {
    cwd: item.frontend ? resolve(root, 'apps/wps-addin') : root,
    encoding: 'utf8', timeout: 120000, env: { ...process.env, NO_COLOR: '1' },
  });
  const log = `${result.stdout ?? ''}\n${result.stderr ?? ''}`.replace(/\u001b\[[0-9;]*m/g, '');
  writeFileSync(resolve(output, `${item.id}-${stage}.log`), log);
  if (result.error || result.signal || /no tests ran|Tests\s+no tests|No test files found|Failed Suites|ERROR collecting|ERROR at setup|SyntaxError|ImportError/.test(log)) {
    throw new Error(`${item.id}: test infrastructure failure (${stage})`);
  }
  return { code: result.status, log };
}

const selected = process.argv.slice(2);
for (const item of cases.filter(item => !selected.length || selected.includes(item.id))) {
  const path = resolve(root, item.path);
  const original = readFileSync(path);
  const source = original.toString('utf8');
  if (source.split(item.needle).length !== 2) throw new Error(`${item.id}: unique seed missing`);
  if (run(item, 'baseline').code !== 0) throw new Error(`${item.id}: baseline failed`);
  const mutated = Buffer.from(source.replace(item.needle, item.replacement));
  let installed = false;
  try {
    if (!readFileSync(path).equals(original)) throw new Error(`${item.id}: changed during baseline; no mutation applied`);
    writeFileSync(path, mutated);
    installed = true;
    const red = run(item, 'red');
    if (red.code !== 1 || !/\d+ failed/.test(red.log) || !red.log.includes(item.evidence)
        || !/AssertionError|DID NOT RAISE|E\s+assert /.test(red.log)) {
      throw new Error(`${item.id}: no discriminating business assertion`);
    }
  } finally {
    if (installed) {
      if (!readFileSync(path).equals(mutated)) throw new Error(`${item.id}: concurrent edit, refusing overwrite`);
      writeFileSync(path, original);
      if (hash(readFileSync(path)) !== hash(original)) throw new Error(`${item.id}: SHA256 differs`);
    }
  }
  if (run(item, 'restored').code !== 0) throw new Error(`${item.id}: restored suite failed`);
  console.log(`${item.id}: RED; SHA256 restored ${hash(original)}; GREEN`);
}
