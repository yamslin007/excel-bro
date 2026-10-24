/**
 * 排序（sortRange）守卫的**反向判别力**验证。
 *
 * 只做一件事：把某一处守卫拆掉，跑它对应的业务用例，**必须**出现
 * `AssertionError` 且失败原文里带那条用例自己的证据；随后逐字节还原并核对 SHA256。
 *
 * 纪律（与 `check_wps_capability_guards.mjs` 相同）：
 *  - 一次只装一个变异；
 *  - 装之前先跑 baseline，必须绿；
 *  - 变异期间别人改过文件就拒绝覆盖（并发编辑保护）；
 *  - finally 还原并核对 SHA256；
 *  - 语法错误 / 采集失败（Failed Suites / SyntaxError / no tests）**不算**通过。
 */
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createHash, randomUUID } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const output = resolve(root, 'outputs/wps-sort-guards');
mkdirSync(output, { recursive: true });
const hash = (data) => createHash('sha256').update(data).digest('hex');

const cases = [
  {
    id: 'sort-rejects-formulas',
    path: 'apps/wps-addin/src/wpsPlanExecutor.ts',
    cwd: 'apps/wps-addin',
    needle: '  if (formulaCells.length) {',
    replacement: '  if (false && formulaCells.length) {',
    test: 'src/wpsSortRange.test.ts',
    evidence: 'toThrow(/包含',
    label: '区域含公式必须写入前拒绝'
  },
  {
    id: 'sort-rejects-mixed-format',
    path: 'apps/wps-addin/src/wpsPlanExecutor.ts',
    cwd: 'apps/wps-addin',
    needle: '  if (!style.uniform) {',
    replacement: '  if (false && !style.uniform) {',
    test: 'src/wpsSortRange.test.ts',
    evidence: '数据行存在不同格式',
    label: '数据行格式不一致必须拒绝'
  },
  {
    id: 'sort-rejects-table-object',
    path: 'apps/wps-addin/src/wpsPlanExecutor.ts',
    cwd: 'apps/wps-addin',
    needle: '    if (geometric) {',
    replacement: '    if (false && geometric) {',
    test: 'src/wpsSortRange.test.ts',
    evidence: '表格对象「明细表」重叠',
    label: '与表格对象重叠必须拒绝'
  },
  {
    id: 'sort-blank-always-last',
    path: 'packages/core/src/tableTransforms.ts',
    cwd: 'packages/core',
    needle: '      if (blankA) return 1;\n      if (blankB) return -1;',
    replacement: '      if (blankA) return key.ascending ? 1 : -1;\n      if (blankB) return key.ascending ? -1 : 1;',
    test: 'src/tableTransforms.test.ts',
    evidence: '空白永远排在最后',
    label: '空白恒排末位（不随方向翻转）'
  }
];

function run(item, stage) {
  const result = spawnSync(
    process.execPath,
    [resolve(root, 'node_modules/vitest/vitest.mjs'), 'run', item.test],
    { cwd: resolve(root, item.cwd), encoding: 'utf8', timeout: 180000, env: { ...process.env, NO_COLOR: '1' } }
  );
  const log = `${result.stdout ?? ''}\n${result.stderr ?? ''}`.replace(/\u001b\[[0-9;]*m/g, '');
  writeFileSync(resolve(output, `${item.id}-${stage}.log`), log);
  if (
    result.error || result.signal ||
    /no tests ran|Tests\s+no tests|No test files found|Failed Suites|SyntaxError|ImportError|Transform failed/.test(log)
  ) {
    throw new Error(`${item.id}: 测试基础设施失败（${stage}），不能算判别力证据`);
  }
  return { code: result.status, log };
}

const selected = process.argv.slice(2);
for (const item of cases.filter((entry) => !selected.length || selected.includes(entry.id))) {
  const path = resolve(root, item.path);
  const original = readFileSync(path);
  const source = original.toString('utf8');
  if (source.split(item.needle).length !== 2) throw new Error(`${item.id}: 种子不唯一或缺失`);

  if (run(item, 'baseline').code !== 0) throw new Error(`${item.id}: baseline 未通过`);

  const mutated = Buffer.from(source.replace(item.needle, item.replacement));
  let installed = false;
  try {
    if (!readFileSync(path).equals(original)) throw new Error(`${item.id}: baseline 期间文件被改动，未安装变异`);
    writeFileSync(path, mutated);
    installed = true;
    const red = run(item, 'red');
    if (red.code !== 1 || !/\d+ failed/.test(red.log) || !red.log.includes(item.evidence)
      || !/AssertionError|DID NOT RAISE/.test(red.log)) {
      throw new Error(`${item.id}: 未出现预期的业务断言失败（${item.label}）`);
    }
  } finally {
    if (installed) {
      if (!readFileSync(path).equals(mutated)) throw new Error(`${item.id}: 并发编辑，拒绝覆盖`);
      writeFileSync(path, original);
      if (hash(readFileSync(path)) !== hash(original)) throw new Error(`${item.id}: SHA256 不一致`);
    }
  }
  if (run(item, 'restored').code !== 0) throw new Error(`${item.id}: 还原后用例未通过`);
  console.log(`${item.id}: RED（${item.label}）；SHA256 还原 ${hash(original).slice(0, 12)}；GREEN`);
}
