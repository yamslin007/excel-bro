"""Run sequentially after normal tests; always restore exact bytes, including on failure."""
from pathlib import Path
import hashlib
import json
import subprocess

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'outputs/formula-probe-fix-20261030'
CASES = [
    ('F1', 'apps/excel-addin/src/excelFormulaSourceEvidence.ts',
     "  // Mixed/baked plans retain their original planning-time evidence.",
     "  if (plan.actions.length) return plan; // mutation: omit deterministic read",
     '模型读取=false', '三张结果表的公式计划自动执行', 'expected "spy" to be called 1 times, but got 0 times'),
    ('F2', 'apps/excel-addin/src/App.tsx',
     'if (onResultReview) throw new ReviewNotStartedError(reason.message);',
     'if (onResultReview) return false; // mutation: lose not-started classification',
     '探针取消显示具体同表依赖原因', '探针取消显示具体同表依赖原因', 'Received: "本次执行需要核对'),
]

def run(pattern):
    return subprocess.run(['npx.cmd', 'vitest', 'run', 'src/excelWriteEndToEnd.test.tsx', '-t', pattern],
        cwd=ROOT / 'apps/excel-addin', capture_output=True, timeout=120)

def sha(data):
    return hashlib.sha256(data).hexdigest()

results = []
for name, relative, needle, replacement, pattern, test_name, evidence in CASES:
    path = ROOT / relative
    baseline = path.read_bytes()
    source = baseline.decode('utf-8')
    assert source.count(needle) == 1, (name, 'seed must match exactly once')
    try:
        path.write_bytes(source.replace(needle, replacement).encode('utf-8'))
        red = run(pattern)
        output = (red.stdout + red.stderr).decode('utf-8', errors='replace')
        (OUT / f'{name}-red.log').write_text(output, encoding='utf-8')
    finally:
        path.write_bytes(baseline)
    restored = path.read_bytes() == baseline
    green = run(pattern)
    (OUT / f'{name}-restored.log').write_bytes(green.stdout + green.stderr)
    result = dict(case=name, nonzero=red.returncode != 0, expected_test=test_name in output,
                  evidence=evidence in output, restored=restored, sha256=sha(baseline),
                  restored_test_exit=green.returncode)
    results.append(result)
    (OUT / 'mutations.json').write_text(json.dumps(results, indent=2), encoding='utf-8')
    print(result, flush=True)
    assert all(result[key] for key in ('nonzero', 'expected_test', 'evidence', 'restored'))
    assert green.returncode == 0
