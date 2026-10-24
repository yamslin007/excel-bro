import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createHash, randomUUID } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const output = resolve(root, 'outputs/sdk-route-guards');
mkdirSync(output, { recursive: true });
const hash = data => createHash('sha256').update(data).digest('hex');
const cases = [
  {
    id: 'natural-language', path: 'server/app/planner.py',
    needle: '    sdk_request = request.model_copy',
    replacement: '    if request.prompt.startswith("清空"):\n        from .models import AnswerResponse\n        return AnswerResponse(message="旧本地拦截", provider="model")\n    sdk_request = request.model_copy',
    test: 'server/tests/test_sdk_agent_route.py::test_every_natural_request_reaches_sdk_unchanged',
    evidence: 'test_every_natural_request_reaches_sdk_unchanged',
  },
  {
    id: 'authorized-sheets', path: 'server/app/plan_tools.py',
    needle: '    if not candidates:\n        raise ValueError(',
    replacement: '    if not candidates:\n        return request.workbook.worksheets[0]\n        raise ValueError(',
    test: 'server/tests/test_sdk_agent_route.py', evidence: 'test_structured_delete_never_guesses',
  },
  {
    id: 'client-route', path: 'apps/excel-addin/src/modelSelection.ts', frontend: true,
    needle: '  return isRealModelId(modelId);',
    replacement: '  return (modelId ?? "").toLowerCase().startsWith("claude-agent-sdk");',
    test: 'src/modelSelection.test.ts', evidence: 'uses SDK policy for existing connection IDs too',
  },
  {
    id: 'followup-context', path: 'server/app/claude_sdk_agent.py',
    needle: '"title": request.priorResult.title,', replacement: '"title": "",',
    test: 'server/tests/test_main.py::test_model_intent_retries_invalid_structure_and_keeps_prior_intent',
    evidence: 'test_model_intent_retries_invalid_structure_and_keeps_prior_intent',
  },
];

function run(item, stage) {
  const command = item.frontend ? process.execPath : resolve(root, '.venv/Scripts/python.exe');
  const args = item.frontend
    ? [resolve(root, 'node_modules/vitest/vitest.mjs'), 'run', item.test]
    : ['-m', 'pytest', item.test, '-q', '--tb=short', '-p', 'no:cacheprovider',
       `--basetemp=${resolve(output, `tmp-${randomUUID()}`)}`];
  const result = spawnSync(command, args, {
    cwd: item.frontend ? resolve(root, 'apps/excel-addin') : root,
    encoding: 'utf8', timeout: 120000, env: { ...process.env, NO_COLOR: '1' },
  });
  const log = `${result.stdout ?? ''}\n${result.stderr ?? ''}`.replace(/\u001b\[[0-9;]*m/g, '');
  writeFileSync(resolve(output, `${item.id}-${stage}.log`), log);
  if (result.error || result.signal || /no tests ran|Tests\s+no tests|No test files found|Failed Suites|ERROR collecting|ERROR at setup|SyntaxError|ImportError/.test(log)) {
    throw new Error(`${item.id}: test infrastructure failure (${stage})`);
  }
  return { code: result.status, log };
}

for (const item of cases) {
  const path = resolve(root, item.path);
  const original = readFileSync(path);
  const source = original.toString('utf8');
  if (source.split(item.needle).length !== 2) throw new Error(`${item.id}: unique seed missing`);
  if (run(item, 'baseline').code !== 0) throw new Error(`${item.id}: baseline failed`);
  const mutated = Buffer.from(source.replace(item.needle, item.replacement));
  try {
    writeFileSync(path, mutated);
    const red = run(item, 'red');
    if (red.code !== 1 || !/\d+ failed/.test(red.log) || !red.log.includes(item.evidence)
        || !/AssertionError|DID NOT RAISE|E\s+assert /.test(red.log)) {
      throw new Error(`${item.id}: no discriminating business assertion`);
    }
  } finally {
    if (!readFileSync(path).equals(mutated)) throw new Error(`${item.id}: concurrent edit, refusing overwrite`);
    writeFileSync(path, original);
    if (hash(readFileSync(path)) !== hash(original)) throw new Error(`${item.id}: SHA256 differs`);
  }
  if (run(item, 'restored').code !== 0) throw new Error(`${item.id}: restored suite failed`);
  console.log(`${item.id}: RED; SHA256 restored ${hash(original)}; GREEN`);
}
