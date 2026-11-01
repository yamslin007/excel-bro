"""反向变异：删除前的恢复材料留痕（U5）的判别力。

背景（`docs/A1_GATE_REMAINING_2026-10-28.md` §1 U5）：材料一删就没得算了。补上的顺序是
**先取指纹 → 落账 → 才删**。三件性质各有一个判别点，任何一个被拆掉，对应用例必须红：

- R1  不再取删除前的指纹（`material = None`）⇒ 账本里根本没有留痕。
- R2  落账失败也照删（去掉那条 `raise`）⇒ "材料没了、证据也没了"正是要防的状态。
- R3  把三列从 `_REQUIRED_OFFICE_COLUMNS` 里拿掉 ⇒ 旧库（`user_version` 已是当前版本）
      永远不会被补列，`recovery_package_retired_at` 直接 `no such column`。
      **"版本号只是意图，列才是事实"** —— 这一条钉的就是那份名单本身。

四判据（不许只卡"退出码非 0"）：①pytest 退出码非 0；②预期用例名出现在 FAILED 里；
③失败原文带它**自己那处**证据；④还原后逐字节 sha256 与基线一致 + 复跑全绿。

用法（**一次一个 SEED**，不要并发）：
  python scripts/mutate_recovery_package_retirement.py            # 打印基线 sha256
  python scripts/mutate_recovery_package_retirement.py R1         # 应用
  python scripts/mutate_recovery_package_retirement.py R1 restore # 还原
  python scripts/mutate_recovery_package_retirement.py R1 check   # 应用→跑→还原→四判据
"""
from __future__ import annotations

import hashlib
import re
import signal
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
PY = REPO / ".venv" / "Scripts" / "python.exe"
MAIN = REPO / "server" / "app" / "main.py"
LEDGER = REPO / "server" / "app" / "task_ledger.py"
TESTS = "server/tests/test_recovery_package_retirement.py"
NODE = "\n"


def _j(*lines: str) -> str:
    return NODE.join(lines)


DIGEST_SEED = '    material = await run_in_threadpool(execution_recovery.raw_digest, operation_id)'
DIGEST_MUTANT = '    material = None  # R1（变异体，勿提交）：删前不再取指纹'

RAISE_SEED = _j(
    '            logger.exception("恢复材料回收留痕不可用")',
    "            raise service_error(",
    "                503,",
    '                "RECOVERY_PACKAGE_RETIREMENT_UNRECORDED",',
    '                "无法留下这次材料回收的留痕，已停止删除；材料仍在，会随保留期到期清理。",',
    "                retryable=True,",
    "            ) from error",
)
RAISE_MUTANT = '            logger.exception("恢复材料回收留痕不可用")  # R2（变异体，勿提交）：失败也照删'

COLUMNS_SEED = _j(
    '    "recovery_package_digest",',
    '    "recovery_package_bytes",',
    '    "recovery_package_retired_at",',
    '    "recovery_package_removed_at",',
    ")",
)
COLUMNS_MUTANT = ")"

TARGETS: dict[str, dict[str, object]] = {
    # R1：不再取删除前的指纹 ⇒ 账本里没有留痕
    "R1": {
        "file": MAIN,
        "seed": DIGEST_SEED,
        "mutant": DIGEST_MUTANT,
        "expect": ["test_delete_records_the_digest_of_the_bytes_it_deleted"],
        "evidence": "assert retired is not None",
    },
    # R2：落账失败也照删
    "R2": {
        "file": MAIN,
        "seed": RAISE_SEED,
        "mutant": RAISE_MUTANT,
        "expect": ["test_ledger_failure_refuses_to_delete_and_keeps_the_material"],
        "evidence": "assert 204 == 503",
    },
    # R3：四列从"必需列"名单里拿掉 ⇒ 旧库永不补列
    "R3": {
        "file": LEDGER,
        "seed": COLUMNS_SEED,
        "mutant": COLUMNS_MUTANT,
        "expect": ["test_legacy_ledger_is_migrated_with_the_three_retirement_columns"],
        "evidence": "no such column",
    },
    # R4：删除**失败**也写"删除确认" ⇒ 账本会出现假的删除结论
    "R4": {
        "file": MAIN,
        "seed": _j(
            "    except Exception as error:",
            "        # 落账已经发生、删除失败：账本留下的是\"记了要删的是哪一份\"，`removed_at` 仍为空。",
            "        # 这正是\"两个事实\"的意义 —— 不能因为 `retired_at` 有值就说删掉了。",
            "        raise _recovery_package_error(error) from error",
        ),
        "mutant": _j(
            "    except Exception as error:",
            "        await run_in_threadpool(  # R4（变异体，勿提交）：删除失败也宣称删掉了",
            "            _folder_operation_ledger.mark_recovery_package_removed, operation_id",
            "        )",
            "        raise _recovery_package_error(error) from error",
        ),
        "expect": ["test_removal_failure_records_the_digest_but_no_deletion_confirmation"],
        "evidence": 'assert trail["removedAt"] is None',
    },
    # R5：删除确认不"首次写赢" ⇒ 第二次回收会改写第一次的删除时刻
    "R5": {
        "file": LEDGER,
        "seed": _j(
            '                "UPDATE office_executions SET recovery_package_removed_at = ? "',
            '                "WHERE operation_id = ? AND recovery_package_removed_at IS NULL",',
        ),
        "mutant": _j(
            '                "UPDATE office_executions SET recovery_package_removed_at = ? "',
            '                "WHERE operation_id = ?",  # R5（变异体，勿提交）',
        ),
        "expect": ["test_first_digest_and_first_removal_survive_a_second_retirement"],
        "evidence": "assert instance.recovery_package_retirement(OPERATION) == first",
    },
}

_touched: dict[Path, str] = {}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _write_verified(path: Path, text: str) -> bool:
    """写回必须读回校验并重试：变异体留在工作区比变异失败更糟。"""
    expected = text.replace("\r\n", "\n")
    for _ in range(5):
        path.write_bytes(text.encode("utf-8"))
        if path.read_bytes().decode("utf-8").replace("\r\n", "\n") == expected:
            return True
    return False


def _restore_all() -> None:
    for path, original in list(_touched.items()):
        _write_verified(path, original)
        _touched.pop(path, None)


def _install_hooks() -> None:
    def handler(_signum, _frame):
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
    basetemp = Path(tempfile.gettempdir()) / f"eb-mut-retire-{label}"
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


def _apply(path: Path, seed: str, mutant: str) -> bool:
    raw = path.read_bytes().decode("utf-8")
    crlf = "\r\n" in raw
    norm = raw.replace("\r\n", "\n")
    if norm.count(seed) != 1:
        print(f"SKIP: {path.name} SEED 出现 {norm.count(seed)} 次（期望 1）")
        print("SEED 片段：\n" + seed)
        return False
    _touched[path] = raw
    mutated = norm.replace(seed, mutant)
    return _write_verified(path, mutated.replace("\n", "\r\n") if crlf else mutated)


def main() -> int:
    _install_hooks()
    which = sys.argv[1] if len(sys.argv) > 1 else ""
    if not which:
        for name, target in TARGETS.items():
            path = target["file"]
            print(f"BASELINE {name} {path} sha256={sha256(path)}")
        return 0
    if which not in TARGETS:
        print("用法: python scripts/mutate_recovery_package_retirement.py [R1|R2|R3] [restore|check]")
        return 2

    target = TARGETS[which]
    path: Path = target["file"]  # type: ignore[assignment]
    seed: str = target["seed"]  # type: ignore[assignment]
    mutant: str = target["mutant"]  # type: ignore[assignment]
    baseline_hash = sha256(path)

    if sys.argv[2:] == ["restore"]:
        raw = path.read_bytes().decode("utf-8")
        norm = raw.replace("\r\n", "\n")
        if norm.count(mutant) == 0:
            print("ALREADY-CLEAN")
            return 0
        _touched[path] = raw
        _restore_all()
        print(f"restore: OK sha256={sha256(path)}")
        return 0

    if sys.argv[2:] != ["check"]:
        if not _apply(path, seed, mutant):
            _restore_all()
            return 3
        print(f"APPLIED {which} → {path} sha256={sha256(path)}")
        return 0

    baseline = run_pytest("baseline")
    bp, bf = _counts(baseline)
    print(f"[基线] {bp} passed, {bf} failed")
    if bp == 0 or bf:
        print("BASELINE-NOT-GREEN：基线未全绿，终止（源码未改动）。")
        print(baseline[-1500:])
        return 4

    red = False
    try:
        if not _apply(path, seed, mutant):
            return 3
        out = run_pytest("mutated")
        failed = re.findall(r"FAILED (server/tests/[\w]+\.py::[\w\-\[\]]+)", out)
        names = sorted({item.split("::")[1].split("[")[0] for item in failed})
        hits = [name for name in target["expect"] if name in names]  # type: ignore[union-attr]
        mp, mf = _counts(out)
        print(f"[变异后] {mp} passed, {mf} failed；红了 {len(names)} 条：{', '.join(names) or '（无）'}")
        print(f"[命中判别点] {', '.join(hits) if hits else '无'}")
        criteria = [
            ("①退出码非 0", mf > 0),
            ("②预期用例红了", bool(hits)),
            ("③失败原文带自家证据", str(target["evidence"]) in out),
        ]
        for label, ok in criteria:
            print(f"{'OK  ' if ok else 'FAIL'} {label}")
        red = all(ok for _, ok in criteria)
        if not red:
            print(out[-2500:])
    finally:
        _restore_all()

    restored = sha256(path) == baseline_hash
    print(f"{'OK  ' if restored else 'FAIL'} ④还原后逐字节一致 :: {baseline_hash} → {sha256(path)}")
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
