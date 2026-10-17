"""反向变异：证明「datasets 完整性闸 / 部分登记清理」的测试真有判别力。

用法：python server/tests/_mutate_prepare_gate.py
每个变异必须 RED（测试变红），还原后必须 GREEN。
"""
from __future__ import annotations

import pathlib
import re
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parents[2]
TARGET = ROOT / "server" / "app" / "task_bridge.py"
TEST = "server/tests/test_prepare_python_data_bridge.py"

ORIGINAL = TARGET.read_bytes()

GATE = """        if (
            len(datasets) != len(scope.sources)
            or len(set(indexes)) != len(indexes)
            or set(indexes) != set(range(len(scope.sources)))
        ):
            return _error_receipt("invalid_arguments")"""

MUTATIONS = [
    ("M1-完整性闸整体失效", GATE, "        if False:\n            return _error_receipt(\"invalid_arguments\")"),
    # 只拆"块数比对"那一层是拆不红的（下面 `set(indexes) != set(range(...))`
    # 会兜住）—— 冗余层要一次拆全：同时去掉块数比对与声明数比对，只留重复检查。
    ("M2-完整性闸只剩重复检查", GATE,
     "        if len(set(indexes)) != len(indexes):\n"
     "            return _error_receipt(\"invalid_arguments\")"),
    ("M3-部分登记不清理",
     "            _discard_python_grant(registry, token, scope, origin)\n"
     "            return _error_receipt(\"result_too_large\")",
     "            return _error_receipt(\"result_too_large\")"),
]


def _norm(text: str) -> bytes:
    return text.replace("\n", "\r\n").encode("utf-8")


def run_tests() -> tuple[bool, str]:
    env = {"TMP": "C:/Users/jia/AppData/Local/Temp",
           "TEMP": "C:/Users/jia/AppData/Local/Temp"}
    proc = subprocess.run(
        [str(ROOT / ".venv" / "Scripts" / "python.exe"), "-m", "pytest",
         TEST, "-p", "no:cacheprovider", "-q",
         "--basetemp=C:/Users/jia/AppData/Local/Temp/pytest-excelbro"],
        cwd=str(ROOT), capture_output=True, text=True, encoding="utf-8",
        errors="replace", env={**__import__("os").environ, **env},
    )
    out = re.sub(r"\x1b\[[0-9;]*m", "", proc.stdout or "")
    tail = [line for line in out.splitlines() if line.strip()][-1:]
    return "failed" not in out and "error" not in out, tail[0] if tail else out[-200:]


def main() -> int:
    ok = True
    for name, seed, replacement in MUTATIONS:
        seed_b = _norm(seed)
        if seed_b not in ORIGINAL:
            print(f"{name}: SEED 未命中 -> SKIP")
            ok = False
            continue
        TARGET.write_bytes(ORIGINAL.replace(seed_b, _norm(replacement), 1))
        passed, summary = run_tests()
        verdict = "GREEN(NG!)" if passed else "RED(OK)"
        print(f"{name}: {verdict}  {summary}")
        if passed:
            ok = False
    TARGET.write_bytes(ORIGINAL)
    passed, summary = run_tests()
    print(f"还原: {'OK' if passed else 'FAILED'}  {summary}")
    return 0 if (ok and passed) else 1


if __name__ == "__main__":
    sys.exit(main())
