import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createHash, randomUUID } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const output = resolve(root, 'outputs/sdk-connection-guards');
mkdirSync(output, { recursive: true });
const hash = data => createHash('sha256').update(data).digest('hex');
const cases = [
  {
    "id": "sdk-tool-receipt",
    "path": "server/app/claude_sdk_agent.py",
    "needle": "    if not received:\n",
    "replacement": "    if False:\n",
    "test": "server/tests/test_model_connection_purposes.py",
    "evidence": "test_sdk_probe_requires_real_tool_receipt"
  },
  {
    "id": "protocol-persistence",
    "path": "server/app/llm/config.py",
    "needle": "                \"supportsVision\": connection.supports_vision,\n                \"protocol\": connection.protocol,",
    "replacement": "                \"supportsVision\": connection.supports_vision,\n                \"protocol\": \"openai_compatible\",",
    "test": "server/tests/test_model_connection_purposes.py",
    "evidence": "test_explicit_protocol_roundtrip"
  },
  {
    "id": "no-provider-fallback",
    "path": "apps/excel-addin/src/modelSelection.ts",
    "frontend": true,
    "needle": "  return Boolean(model?.available && model.agentAvailable !== false);",
    "replacement": "  return Boolean(model?.available);",
    "test": "src/modelSelection.test.ts",
    "evidence": "does not silently replace"
  },
  {
    "id": "purpose-receipt",
    "path": "apps/excel-addin/src/hooks/useModelManagement.ts",
    "frontend": true,
    "needle": "      const verified = result.ok && result.testedFor === purpose;",
    "replacement": "      const verified = result.ok;",
    "test": "src/hooks/useModelManagement.test.ts",
    "evidence": "does not accept an old or formula-only success"
  }
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
