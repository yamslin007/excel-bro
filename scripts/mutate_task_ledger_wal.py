"""反向变异：任务账本 SQLite 并发读写缺陷（Request L · 焦点四）修复的判别力。

用法：python scripts/mutate_task_ledger_wal.py M13|M14|M27|M28

口径与其余变异脚本一致（铁律：① 基线先 GREEN；② 信号钩子还原；③ touched 登记；
④ SEED 唯一，从函数声明行/唯一片段开始，避免留残体；⑤ 换行符归一化；
⑥ 基线先跑一遍确认 GREEN 再下刀）。

每个变异拆掉**对应行为本身**，且各变异有唯一判别点：
- M13  拆掉可写连接的重开重试（首建后并发连接被持成只读即崩溃）。
- M14  拆掉 bootstrap 连接的重开重试（bootstrap 自身被持成只读即崩溃）。
- M27  拆掉「读走只读连接」——读路径重新走可写连接并取写锁（纯读被串行到写锁）。
- M28  拆掉 `_schema_is_current` 短路——新实例在已有库上重跑 bootstrap（与开启写者冲突）。
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
PY = REPO / ".venv" / "Scripts" / "python.exe"
LEDGER = REPO / "server" / "app" / "task_ledger.py"
TESTS = "server/tests/test_task_ledger.py"

NODE = "\n"


def _j(*lines: str) -> str:
    return NODE.join(lines)


# (文件, 期望变红的测试名)
TARGETS = {
    # 焦点四：并发首建只读——拆掉普通可写连接的重开重试
    "M13": (LEDGER, ["test_first_initialization_under_concurrency_never_reports_readonly"]),
    # 焦点四：并发首建只读——拆掉 bootstrap 连接的重开重试
    "M14": (LEDGER, ["test_first_initialization_under_concurrency_never_reports_readonly"]),
    # 焦点四：读路径不得走可写连接（否则纯读被写锁串行）
    "M27": (
        LEDGER,
        [
            "test_read_is_not_serialized_behind_concurrent_write_lock",
            "test_read_after_write_commit_reflects_new_state",
        ],
    ),
    # 焦点四：新实例不得在已有库上重跑 bootstrap
    "M28": (
        LEDGER,
        ["test_new_instance_on_existing_db_does_not_collide_with_open_writer"],
    ),
}

SEEDS = {
    # _open_writable_connection 的可写探测 + 重试循环
    "M13": _j(
        "        last_error: sqlite3.OperationalError | None = None",
        "        for attempt in range(_READONLY_RETRY_ATTEMPTS):",
        "            connection = sqlite3.connect(path, timeout=10, isolation_level=None)",
        "            connection.row_factory = sqlite3.Row",
        "            connection.execute(\"PRAGMA foreign_keys = ON\")",
        "            connection.execute(\"PRAGMA busy_timeout = 10000\")",
        "            if _connection_is_writable(connection):",
        "                return connection",
        "            connection.close()",
        "            last_error = sqlite3.OperationalError(\"attempt to write a readonly database\")",
        "            time.sleep(0.005 * (attempt + 1))",
    ),
    # _run_schema_bootstrap 的可写重试循环
    "M14": _j(
        "        last_error: sqlite3.OperationalError | None = None",
        "        for attempt in range(_READONLY_RETRY_ATTEMPTS):",
        "            bootstrap = sqlite3.connect(path, timeout=10, isolation_level=None)",
        "            try:",
        "                bootstrap.row_factory = sqlite3.Row",
        "                self._initialize_schema(bootstrap)",
        "            except sqlite3.OperationalError as error:",
        "                bootstrap.close()",
        "                if not _is_readonly_error(error):",
        "                    raise",
        "                last_error = error",
        "                time.sleep(0.005 * (attempt + 1))",
        "                continue",
        "            bootstrap.close()",
        "            return",
    ),
    # _connect_readonly 的只读连接构造
    "M27": _j(
        "        path = self._path()",
        "        path.parent.mkdir(parents=True, exist_ok=True)",
        "        self._ensure_schema(path)",
        "        return self._open_readonly_connection(path)",
    ),
    # _run_schema_bootstrap 的当前版本短路
    "M28": _j(
        "        if _schema_is_current(path):",
        "            return",
    ),
}

MUTANTS = {
    # 只连接一次，不做可写判别（并发只读句柄会被当成成功返回 → 后续 BEGIN IMMEDIATE 崩）
    "M13": _j(
        "        # [变异M13] 拆掉可写重试：只连一次",
        "        connection = sqlite3.connect(path, timeout=10, isolation_level=None)",
        "        connection.row_factory = sqlite3.Row",
        "        connection.execute(\"PRAGMA foreign_keys = ON\")",
        "        connection.execute(\"PRAGMA busy_timeout = 10000\")",
        "        return connection",
    ),
    # 只 bootstrap 一次，不做可写重试（bootstrap 自身被持成只读即崩）
    "M14": _j(
        "        # [变异M14] 拆掉 bootstrap 可写重试：只试一次",
        "        bootstrap = sqlite3.connect(path, timeout=10, isolation_level=None)",
        "        try:",
        "            bootstrap.row_factory = sqlite3.Row",
        "            self._initialize_schema(bootstrap)",
        "        finally:",
        "            bootstrap.close()",
    ),
    # 读也走可写连接（重新引入写锁串行）
    "M27": _j(
        "        # [变异M27] 读走可写连接：纯读也取写锁",
        "        path = self._path()",
        "        path.parent.mkdir(parents=True, exist_ok=True)",
        "        self._ensure_schema(path)",
        "        return self._open_writable_connection(path)",
    ),
    # 不判定当前版本，已有库也重跑 bootstrap
    "M28": _j(
        "        # [变异M28] 拆掉当前版本短路：已有库也重跑 bootstrap",
        "        # 无短路",
    ),
}

_touched: dict[Path, str] = {}


def restore_all() -> None:
    for path, content in list(_touched.items()):
        path.write_text(content, encoding="utf-8", newline="")
    _touched.clear()


def _install_signal_hooks() -> None:
    import signal

    def handler(signum, frame):
        restore_all()
        sys.exit(130)

    for sig in (signal.SIGINT, signal.SIGTERM):
        try:
            signal.signal(sig, handler)
        except (ValueError, OSError):
            pass


def run_pytest(test_path: str, name: str) -> str:
    base = tempfile.mkdtemp(prefix=f"mutwal_{name}_")
    env = dict(os.environ)
    env["TMP"] = base
    env["TEMP"] = base
    env["PYTEST_ADDOPTS"] = ""
    proc = subprocess.run(
        [str(PY), "-m", "pytest", test_path, "-q",
         "-p", "no:cacheprovider", "--basetemp", str(Path(base) / "bt")],
        cwd=str(REPO), capture_output=True, text=True, timeout=300, env=env,
    )
    return (proc.stdout or "") + "\n" + (proc.stderr or "")


def _counts(text: str) -> tuple[int, int]:
    p = re.search(r"(\d+) passed", text)
    f = re.search(r"(\d+) failed", text)
    return (int(p.group(1)) if p else 0, int(f.group(1)) if f else 0)


def main() -> int:
    which = sys.argv[1] if len(sys.argv) > 1 else ""
    if which not in TARGETS:
        print("用法: python scripts/mutate_task_ledger_wal.py M13|M14|M27|M28")
        return 2
    _install_signal_hooks()
    file, expect_fail = TARGETS[which]

    baseline = run_pytest(TESTS, "baseline")
    bp, bf = _counts(baseline)
    print(f"[基线] {bp} passed, {bf} failed")
    if bp == 0 or bf:
        restore_all()
        print("BASELINE-NOT-GREEN：基线未全绿，终止（源码未改动）。")
        print(baseline[-1500:])
        return 4

    raw = file.read_bytes().decode("utf-8")
    crlf = "\r\n" in raw
    original = raw
    norm = original.replace("\r\n", "\n")
    seed = SEEDS[which]
    occurrences = norm.count(seed)
    if occurrences != 1:
        print(f"SKIP: {which} SEED 出现 {occurrences} 次（期望 1）。")
        print("SEED 片段：\n" + seed)
        return 3

    _touched[file] = original
    try:
        mutated_norm = norm.replace(seed, MUTANTS[which])
        mutated = mutated_norm.replace("\n", "\r\n") if crlf else mutated_norm
        file.write_text(mutated, encoding="utf-8", newline="")
        out = run_pytest(TESTS, "mutated")
        failed = re.findall(r"FAILED (server/tests/\w+\.py::[\w\[\]\-]+)", out)
        failed_names = sorted({n.split("::")[1].split("[")[0] for n in failed})
        hits = [n for n in expect_fail if n in failed_names]
        print(f"[变异后] 失败 {len(failed_names)} 条：{', '.join(failed_names) or '（无）'}")
        print(f"[命中判别点] {', '.join(hits) if hits else '无'}")
        red = bool(failed_names) and bool(hits)
        print("RED-OK：变异使判别测试独立变红。" if red else "RED-MISS：变异未按预期变红。")
        if not red:
            print(out[-2500:])
    finally:
        restore_all()

    restored = run_pytest(TESTS, "restored")
    rp, rf = _counts(restored)
    print(f"[还原后] {rp} passed, {rf} failed")
    green = rp > 0 and rf == 0
    print("RESTORED-GREEN：还原后复跑全绿。" if green else "RESTORED-NOT-GREEN：还原后仍有失败！")
    return 0 if (red and green) else 5


if __name__ == "__main__":
    sys.exit(main())
