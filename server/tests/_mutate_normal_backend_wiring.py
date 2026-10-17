"""反向变异：证明「普通后端 prepare_python_data 注册/分派/可信绑定」测试有判别力。

用法：python server/tests/_mutate_normal_backend_wiring.py
每个变异必须 RED（测试变红），还原后必须 GREEN。
"""
from __future__ import annotations

import os
import pathlib
import re
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parents[2]
TESTS = ["server/tests/test_python_normal_backend_wiring.py"]
TARGET = ROOT / "server" / "app" / "excel_agent.py"
ORIGINAL = TARGET.read_bytes()

MUTATIONS = [
    ("M1-只注册分析工具不注册登记通道",
     "    return [prepare_python_data_tool_definition(), python_analysis_tool_definition()]",
     "    return [python_analysis_tool_definition()]"),
    ("M2-无桥会话仍放行登记",
     "        if bridge_session is None:",
     "        if False:"),
    ("M3-登记不走桥（服务端凭自报区域放行）",
     '        return await bridge_session.request(name, {"sources": sources})',
     '        return {"status": "ok", "datasetHandles": []}'),
    ("M4-桥请求不发指针事件",
     "        on_request=on_bridge_request,",
     "        on_request=None,"),
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
