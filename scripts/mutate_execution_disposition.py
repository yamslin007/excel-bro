"""反向变异：用户处置留痕（放弃 / 撤销）的判别力。

背景（2026-10-29 真机）：`execution-c52ad791…` 三步写入、第二步后中断、约 2.5 小时后只续
第三步、确认变更 —— 账本里 `status/revision/next/host_writes` 全对，但"用户做过什么"一个字
都没有。补上 `abandoned_at` / `undone_at` 之后，这三条性质必须有判别力：

- D1  拆掉「首次写赢」（`AND <col> IS NULL`）⇒ 重复上报会把第一次的时刻往后改，
      留痕就不再是"第一次发生"，而变成"最后一次上报"。
- D2  让处置顺手推进 `revision` ⇒ 留痕混进进度，"用户放弃了"会被后续续跑判据读成执行进展。
- D3  把两种取值映射到同一列 ⇒ 放弃与撤销互相冒充。

四判据（不许只卡"退出码非 0"）：
  ① 变异后 pytest 退出码非 0；
  ② 预期用例名出现在 FAILED 列表里；
  ③ 失败原文带它**自己那处**证据（断言行/消息）；
  ④ 还原后逐字节 sha256 与基线一致 + 复跑全绿。

用法：
  python scripts/mutate_execution_disposition.py            # 打印基线 sha256（不改文件）
  python scripts/mutate_execution_disposition.py D1        # 应用变异
  python scripts/mutate_execution_disposition.py D1 restore
  python scripts/mutate_execution_disposition.py D1 check  # 应用→跑→还原→四判据
"""
from __future__ import annotations

import hashlib
import os
import re
import signal
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
PY = REPO / ".venv" / "Scripts" / "python.exe"
LEDGER = REPO / "server" / "app" / "task_ledger.py"
TESTS = "server/tests/test_execution_disposition.py"
NODE = "\n"


def _j(*lines: str) -> str:
    return NODE.join(lines)


UPDATE_SEED = _j(
    "            connection.execute(",
    '                f"UPDATE office_executions SET {column} = ? "',
    '                f"WHERE operation_id = ? AND {column} IS NULL",',
)

TARGETS: dict[str, tuple[list[str], str, str]] = {
    # D1：拆掉"首次写赢"
    "D1": (
        ["test_first_write_wins_and_the_two_kinds_are_independent"],
        _j(
            "            connection.execute(",
            '                f"UPDATE office_executions SET {column} = ? "',
            '                f"WHERE operation_id = ?",  # D1：重复上报会改写时刻（变异体，勿提交）',
        ),
        "首次写赢：重复上报不许推后时刻",
    ),
    # D2：处置顺手推进 revision
    "D2": (
        ["test_office_abandoned_is_recorded_and_progress_is_untouched"],
        _j(
            "            connection.execute(",
            '                f"UPDATE office_executions SET {column} = ?, revision = revision + 1 "  # D2（变异体，勿提交）',
            '                f"WHERE operation_id = ? AND {column} IS NULL",',
        ),
        "assert after.revision == before.revision",
    ),
    # D3：两种取值映射到同一列
    "D3": (
        ["test_wps_undone_is_recorded_on_its_own_namespace"],
        '    DISPOSITION_COLUMNS = {"abandoned": "abandoned_at", "undone": "abandoned_at"}  # D3（变异体，勿提交）',
        'assert disposition["undoneAt"] is not None',
    ),
}

_touched: dict[Path, str] = {}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _write_verified(path: Path, text: str, crlf: bool) -> bool:
    """写回必须读回校验并重试：变异体留在工作区比变异失败更糟。"""
    expected = text.replace("\r\n", "\n")
    for _ in range(5):
        path.write_bytes(text.encode("utf-8"))
        if path.read_bytes().decode("utf-8").replace("\r\n", "\n") == expected:
            return True
    return False


def _restore_all() -> None:
    for path, original in list(_touched.items()):
        crlf = "\r\n" in original
        _write_verified(path, original, crlf)
        _touched.pop(path, None)


def _install_hooks() -> None:
    def handler(signum, _frame):
        _restore_all()
        sys.exit(130)

    for sig in (signal.SIGINT, signal.SIGTERM, getattr(signal, "SIGHUP", None)):
        if sig is None:
            # Windows 上没有 SIGHUP；缺它不算缺陷，别为此跳过其余钩子。
            continue
        try:
            signal.signal(sig, handler)
        except (ValueError, OSError):
            pass


def run_pytest(label: str) -> str:
    basetemp = Path(tempfile.gettempdir()) / f"eb-mut-dispo-{label}"
    completed = subprocess.run(
        [str(PY), "-m", "pytest", TESTS, "-q", f"--basetemp={basetemp}"],
        cwd=REPO,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return f"{completed.stdout}\n{completed.stderr}"


def _counts(output: str) -> tuple[int, int]:
    passed = re.search(r"(\d+) passed", output)
    failed = re.search(r"(\d+) failed", output)
    return (int(passed.group(1)) if passed else 0, int(failed.group(1)) if failed else 0)


def main() -> int:
    _install_hooks()
    which = sys.argv[1] if len(sys.argv) > 1 else ""
    baseline_hash = sha256(LEDGER)
    if not which:
        print(f"BASELINE {LEDGER} sha256={baseline_hash}")
        return 0
    if which not in TARGETS:
        print("用法: python scripts/mutate_execution_disposition.py [D1|D2|D3] [restore|check]")
        return 2

    expect_fail, mutant, evidence = TARGETS[which]
    raw = LEDGER.read_bytes().decode("utf-8")
    crlf = "\r\n" in raw
    norm = raw.replace("\r\n", "\n")
    seed = UPDATE_SEED if which in ("D1", "D2") else '    DISPOSITION_COLUMNS = {"abandoned": "abandoned_at", "undone": "undone_at"}'

    if sys.argv[2:] == ["restore"]:
        hits = norm.count(mutant)
        if hits == 0:
            print("ALREADY-CLEAN")
            return 0
        _touched[LEDGER] = raw
        _restore_all()
        print(f"restore: OK sha256={sha256(LEDGER)}")
        return 0

    occurrences = norm.count(seed)
    if occurrences != 1:
        # SEED 未命中 = 假装有证据。
        print(f"SKIP: {which} SEED 出现 {occurrences} 次（期望 1）。")
        print("SEED 片段：\n" + seed)
        return 3

    if sys.argv[2:] != ["check"]:
        _touched[LEDGER] = raw
        mutated = norm.replace(seed, mutant).replace("\n", "\r\n") if crlf else norm.replace(seed, mutant)
        if not _write_verified(LEDGER, mutated, crlf):
            _restore_all()
            return 4
        print(f"APPLIED {which} → {LEDGER} sha256={sha256(LEDGER)}")
        return 0

    # ---- check：基线必须先是绿的 ----
    baseline = run_pytest("baseline")
    bp, bf = _counts(baseline)
    print(f"[基线] {bp} passed, {bf} failed")
    if bp == 0 or bf:
        print("BASELINE-NOT-GREEN：基线未全绿，终止（源码未改动）。")
        print(baseline[-1500:])
        return 4

    red = False
    try:
        _touched[LEDGER] = raw
        mutated = norm.replace(seed, mutant).replace("\n", "\r\n") if crlf else norm.replace(seed, mutant)
        if not _write_verified(LEDGER, mutated, crlf):
            return 4
        out = run_pytest("mutated")
        failed = re.findall(r"FAILED (server/tests/[\w]+\.py::[\w\-\[\]]+)", out)
        names = sorted({item.split("::")[1].split("[")[0] for item in failed})
        hits = [name for name in expect_fail if name in names]
        mp, mf = _counts(out)
        print(f"[变异后] {mp} passed, {mf} failed；红了 {len(names)} 条：{', '.join(names) or '（无）'}")
        print(f"[命中判别点] {', '.join(hits) if hits else '无'}")
        criteria = [
            ("①退出码非 0", mf > 0),
            ("②预期用例红了", bool(hits)),
            ("③失败原文带自家证据", evidence in out),
        ]
        for label, ok in criteria:
            print(f"{'OK  ' if ok else 'FAIL'} {label}")
        red = all(ok for _, ok in criteria)
        if not red:
            print(out[-2500:])
    finally:
        _restore_all()

    restored = sha256(LEDGER) == baseline_hash
    print(f"{'OK  ' if restored else 'FAIL'} ④还原后逐字节一致 :: {baseline_hash} → {sha256(LEDGER)}")
    after = run_pytest("restored")
    ap, af = _counts(after)
    print(f"[还原后] {ap} passed, {af} failed")
    green = ap > 0 and af == 0
    print("RESTORED-GREEN：还原后复跑全绿。" if green else "RESTORED-NOT-GREEN：还原后仍有失败！")
    ok = red and restored and green
    print(f"\n{which}: 四判据{'全部成立' if ok else '不成立，别当成\"已经验证过\"'}")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
