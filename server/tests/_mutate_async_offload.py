"""反向变异：证明「隔离执行不阻塞事件循环 / 取消真实传导」的测试有判别力。

用法：python server/tests/_mutate_async_offload.py
每个变异必须 RED（测试变红），还原后必须 GREEN。
"""
from __future__ import annotations

import os
import pathlib
import re
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parents[2]
TEST = "server/tests/test_python_async_offload.py"
TARGETS = {
    "coordinator": ROOT / "server" / "app" / "python_coordinator.py",
    "excel_agent": ROOT / "server" / "app" / "excel_agent.py",
}
ORIGINALS = {key: path.read_bytes() for key, path in TARGETS.items()}

OFFLOAD = """    future = loop.run_in_executor(
        None,
        lambda: execute_python_analysis(request, arguments, cancel_check=guard),
    )"""

WAIT = """        cancelled.set()
        while not future.done():
            try:
                await asyncio.shield(future)
            except asyncio.CancelledError:
                if future.done():
                    break
                await asyncio.sleep(0)
        raise"""

MUTATIONS = [
    ("M1-同步调用阻塞事件循环", "coordinator", OFFLOAD,
     "    future = asyncio.ensure_future(asyncio.sleep("
     "0, result=execute_python_analysis(request, arguments, cancel_check=guard)))"),
    ("M2-取消不传导给执行器", "coordinator",
     "        cancelled.set()\n", "        pass\n"),
    ("M3-取消后不等真实执行终止", "coordinator", WAIT, "        raise"),
    ("M4-普通后端仍同步调用", "excel_agent",
     """    if name == "run_python_analysis":
        from .python_coordinator import execute_python_analysis_async

        return await execute_python_analysis_async(request, arguments)
    return _execute_tool(request, name, arguments)""",
     "    return _execute_tool(request, name, arguments)"),
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
