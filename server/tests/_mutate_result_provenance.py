"""反向变异：证明「结果来源指纹 / 实际来源依赖」的测试有判别力。

用法：python server/tests/_mutate_result_provenance.py
每个变异必须 RED（测试变红），还原后必须 GREEN。
"""
from __future__ import annotations

import os
import pathlib
import re
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parents[2]
TESTS = ["server/tests/test_python_coordinator.py"]
TARGETS = {
    "coordinator": ROOT / "server" / "app" / "python_coordinator.py",
    "task_bridge": ROOT / "server" / "app" / "task_bridge.py",
}
ORIGINALS = {key: path.read_bytes() for key, path in TARGETS.items()}

MUTATIONS = [
    ("M1-来源依赖虚报成登记的全部来源", "coordinator",
     "    source_sheets = binding.sheets_of(handles)",
     "    source_sheets = [s.sheetName for s in binding.scope.sources]"),
    ("M2-查不到来源下标仍放行", "coordinator",
     "    if source_sheets is None:",
     "    if False:"),
    ("M3-结果不带源指纹", "coordinator",
     "            task_id, data_result, binding.source_fingerprint",
     "            task_id, data_result, None"),
    ("M4-来源下标按回执迭代顺序而非块声明", "task_bridge",
     "                source_indexes[receipt.datasetHandle] = block.sourceIndex",
     "                source_indexes[receipt.datasetHandle] = len(receipts) - 1"),
]


def _norm(text: str) -> bytes:
    return text.replace("\n", "\r\n").encode("utf-8")


def run_tests() -> tuple[bool, str]:
    env = {"TMP": "C:/Users/jia/AppData/Local/Temp",
           "TEMP": "C:/Users/jia/AppData/Local/Temp"}
    proc = subprocess.run(
        [str(ROOT / ".venv" / "Scripts" / "python.exe"), "-m", "pytest",
         *TESTS, "-p", "no:cacheprovider", "-q",
         "--basetemp=C:/Users/jia/AppData/Local/Temp/pytest-excelbro"],
        cwd=str(ROOT), capture_output=True, text=True, encoding="utf-8",
        errors="replace", env={**os.environ, **env},
    )
    out = re.sub(r"\x1b\[[0-9;]*m", "", proc.stdout or "")
    tail = [line for line in out.splitlines() if line.strip()][-1:]
    return "failed" not in out and "error" not in out, tail[0] if tail else out[-200:]


def main() -> int:
    ok = True
    for name, key, seed, replacement in MUTATIONS:
        target = TARGETS[key]
        seed_b = _norm(seed)
        if seed_b not in ORIGINALS[key]:
            print(f"{name}: SEED 未命中 -> SKIP")
            ok = False
            continue
        target.write_bytes(ORIGINALS[key].replace(seed_b, _norm(replacement), 1))
        passed, summary = run_tests()
        verdict = "GREEN(NG!)" if passed else "RED(OK)"
        print(f"{name}: {verdict}  {summary}")
        if passed:
            ok = False
    for key, target in TARGETS.items():
        target.write_bytes(ORIGINALS[key])
    passed, summary = run_tests()
    print(f"还原: {'OK' if passed else 'FAILED'}  {summary}")
    return 0 if (ok and passed) else 1


if __name__ == "__main__":
    sys.exit(main())
