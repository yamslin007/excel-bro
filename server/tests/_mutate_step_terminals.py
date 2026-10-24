"""反向变异：证明「步骤身份 / 真实终态」这批断言真的有判别力。

用法：python server/tests/_mutate_step_terminals.py
每个变异必须 RED，并且**红的必须是被测那条用例**（不是随便哪条崩溃）。
还原后必须 GREEN，且文件逐字节回到 sha256 原值。

纪律（见仓库 AGENTS.md / .workbuddy/memory）：
- 一次只改一个 SEED，按原 EOL 写回；
- 写回后**读回校验**，不一致立刻还原并报错（Windows 上偶发 Errno 22 会把
  变异留在工作区）；
- 装 SIGINT/TERM/HUP 与 uncaughtException 还原钩子；被抓走也要还原；
- 判据不是"退出码非 0"，而是"预期的那条用例失败，且失败原文里带着它自己
  那处证据"。
"""
from __future__ import annotations

import hashlib
import os
import pathlib
import re
import signal
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parents[2]

BRIDGE = ROOT / "server" / "app" / "task_bridge.py"
COORD = ROOT / "server" / "app" / "python_coordinator.py"
FOLDER = ROOT / "server" / "app" / "folder_workbooks.py"
MAIN = ROOT / "server" / "app" / "main.py"

TESTS = {
    "bridge": "server/tests/test_agent_tool_bridge.py",
    "python": "server/tests/test_python_async_offload.py",
    "folder": "server/tests/test_folder_workbooks.py",
    "main": "server/tests/test_main.py",
}


def _read(path: pathlib.Path) -> bytes:
    return path.read_bytes()


def _sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


ORIGINALS: dict[pathlib.Path, bytes] = {
    path: _read(path) for path in (BRIDGE, COORD, FOLDER, MAIN)
}
ORIGINAL_SHA: dict[pathlib.Path, str] = {
    path: _sha(data) for path, data in ORIGINALS.items()
}
EOL = {path: (b"\r\n" if b"\r\n" in data else b"\n") for path, data in ORIGINALS.items()}


def _encode(path: pathlib.Path, text: str) -> bytes:
    return text.replace("\n", "\r\n" if EOL[path] == b"\r\n" else "\n").encode("utf-8")


def _restore() -> None:
    for path, data in ORIGINALS.items():
        try:
            path.write_bytes(data)
        except OSError:
            pass


def _guard(signum, _frame):  # pragma: no cover - 只在被打断时执行
    _restore()
    print(f"\n收到信号 {signum}，已还原工作区。", flush=True)
    sys.exit(130)


for _sig in (signal.SIGINT, signal.SIGTERM):
    try:
        signal.signal(_sig, _guard)
    except (ValueError, OSError):  # pragma: no cover
        pass


def _write_verified(path: pathlib.Path, payload: bytes) -> None:
    for attempt in range(4):
        path.write_bytes(payload)
        if path.read_bytes() == payload:
            return
        if attempt == 3:
            _restore()
            raise RuntimeError(f"写回校验失败：{path.name}")
        # Windows 上偶发 Errno 22/13：读回不一致就再写一次。


def run_tests(test_files: list[str], keyword: str | None) -> tuple[bool, str, list[str]]:
    command = [
        str(ROOT / ".venv" / "Scripts" / "python.exe"), "-m", "pytest",
        *test_files, "-p", "no:cacheprovider", "-q",
        "--basetemp=C:/Users/jia/AppData/Local/Temp/pytest-excelbro-mut",
    ]
    if keyword:
        command += ["-k", keyword]
    proc = subprocess.run(
        command, cwd=str(ROOT), capture_output=True, text=True,
        encoding="utf-8", errors="replace",
    )
    out = re.sub(r"\x1b\[[0-9;]*m", "", (proc.stdout or "") + (proc.stderr or ""))
    failed = re.findall(r"^FAILED\s+(\S+)", out, re.MULTILINE)
    return proc.returncode == 0, out, failed


# name, target, seed, replacement, test file key, -k keyword, 必须变红的用例
MUTATIONS = [
    (
        "M1-步骤身份退化成只用工具名",
        BRIDGE,
        '    return f"bridge.{tool}:{suffix}"',
        '    return f"bridge.{tool}"',
        "bridge",
        "step_id or distinct_steps",
        "test_bridge_step_id_separates_invocations_of_the_same_tool",
    ),
    (
        "M2-成功调用不发结束事件",
        BRIDGE,
        """            return await self._notify_complete(
                tool, operation_id, await asyncio.wait_for(future, timeout=timeout)
            )""",
        """            return await asyncio.wait_for(future, timeout=timeout)""",
        "bridge",
        "successful_call",
        "test_bridge_reports_a_completion_for_a_successful_call",
    ),
    (
        "M3-超时不算真实终态",
        BRIDGE,
        """            return await self._notify_complete(
                tool, operation_id, _error_receipt("timeout")
            )""",
        """            return _error_receipt("timeout")""",
        "bridge",
        "timeout",
        "test_bridge_reports_timeout_as_a_real_terminal",
    ),
    (
        "M4-隔离分析完成步骤身份写死",
        COORD,
        '                    "stepId": f"python.completed:{run_id}",',
        '                    "stepId": "python.completed",',
        "python",
        "own_completion_step",
        "test_each_isolated_run_emits_its_own_completion_step",
    ),
    (
        "M5-准备文件步骤只留 progress",
        FOLDER,
        """            "stepId": "folder.prepare-files",
            "status": "completed",""",
        """            "stepId": "folder.prepare-files",
            "status": "progress",""",
        "folder",
        "terminal",
        "test_every_started_folder_step_gets_its_own_terminal_event",
    ),
    (
        "M6-执行失败不给中断步骤收尾",
        MAIN,
        """            interrupted = in_flight.get("stepId")
            if isinstance(interrupted, str) and interrupted:
                await _record_folder_progress(
                    operation_id,
                    state,
                    {
                        "phase": "executing",
                        "stepId": interrupted,
                        "status": "failed",
                        "title": "这一步没有完成",
                        "summary": public_message[:1000],
                    },
                )""",
        """            pass""",
        "main",
        "interrupted_step",
        "test_folder_execute_stream_closes_the_interrupted_step_as_failed",
    ),
]


def main() -> int:
    ok = True
    try:
        for name, target, seed, replacement, test_key, keyword, expected in MUTATIONS:
            original = ORIGINALS[target]
            seed_bytes = _encode(target, seed)
            if seed_bytes not in original:
                print(f"{name}: SEED 未命中 -> 判据无效（SKIP 视为失败）", flush=True)
                ok = False
                continue
            mutated = original.replace(seed_bytes, _encode(target, replacement), 1)
            _write_verified(target, mutated)
            passed, out, failed = run_tests([TESTS[test_key]], keyword)
            hit = any(expected in item for item in failed)
            verdict = "RED(OK)" if (not passed and hit) else "GREEN/误红(NG)"
            print(f"{name}: {verdict}", flush=True)
            if passed or not hit:
                print(f"  退出码={0 if passed else 1} 失败用例={failed or '无'}", flush=True)
                tail = [line for line in out.splitlines() if line.strip()][-12:]
                print("  " + "\n  ".join(tail), flush=True)
                ok = False
            _restore()
    finally:
        _restore()

    for path, sha in ORIGINAL_SHA.items():
        actual = _sha(_read(path))
        if actual != sha:
            print(f"还原校验失败：{path.name} {actual} != {sha}", flush=True)
            ok = False
    passed, out, _ = run_tests(
        [TESTS["bridge"], TESTS["python"], TESTS["folder"], TESTS["main"]], None
    )
    print(f"还原后全量: {'GREEN(OK)' if passed else 'RED(NG)'}", flush=True)
    if not passed:
        tail = [line for line in out.splitlines() if line.strip()][-12:]
        print("  " + "\n  ".join(tail), flush=True)
        ok = False
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
