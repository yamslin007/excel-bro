"""反向变异：Docker 隔离执行器七类缺陷修复的判别力（2026-10-07, Request H）

用法：python scripts/mutate_python_docker_defects.py M1|M5|M6

与 ``mutate_python_docker_defects.mjs`` 等价（node 在本沙箱无法 spawn 子进程，
故用 Python 实现；口径与其余 .mjs 变异脚本一致）。

铁律：① 基线先 GREEN；② 信号钩子还原；③ touched 登记；④ SEED 唯一。
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
RUNNER = REPO / "server" / "app" / "python_docker_runner.py"
DOCKERFILE = REPO / "server" / "docker" / "Dockerfile"
REALTESTS = REPO / "server" / "tests" / "test_python_docker_real.py"
DEFECT_TESTS = "server/tests/test_python_docker_defects.py"

NODE = "\n"


def _j(*lines: str) -> str:
    return NODE.join(lines)


# (文件, 期望变红的测试名)
# 注意：第八批（Request I）重写执行路径后，第七批部分判别测试已并入
# test_python_docker_defects2.py。此处保留仍有效的第七批判别点。
TARGETS = {
    "M1": (DOCKERFILE, ["test_dockerfile_copy_source_exists_in_build_context"]),
    "M5": (RUNNER, ["test_isolation_config_validation_rejects_incomplete_config"]),
    "M6": (RUNNER, ["test_isolation_config_validation_rejects_root_uid_variants"]),
}

SEEDS = {
    "M1": "COPY app/docker_worker.py /app/worker.py",
    "M5": _j(
        "    problems: list[str] = []",
        "    if config.network != \"none\":",
        "        problems.append(\"network 必须为 none\")",
    ),
    "M6": _j(
        "    uid_part = text.split(\":\", 1)[0].strip()",
        "    if uid_part.lower() in (\"root\", \"\"):",
        "        return True",
        "    # 纯数字（允许前导零/加号）→ 数值 0 即 root",
        "    if re.fullmatch(r\"[0-9]+\", uid_part):",
        "        try:",
        "            return int(uid_part) == 0",
        "        except ValueError:",
        "            return True",
        "    return False",
    ),
}

MUTANTS = {
    "M1": "COPY docker_worker.py /app/worker.py",
    "M5": _j(
        "    problems: list[str] = []",
        "    return True, None  # [变异M5] 拆隔离配置校验",
        "    if config.network != \"none\":",
        "        problems.append(\"network 必须为 none\")",
    ),
    "M6": _j(
        "    # [变异M6] 只拒裸 0：漏过 root / 0:1000 / root:1000 等变体",
        "    return uid_part == \"0\"",
    ),
}

_touched: dict[Path, str] = {}


def restore_all() -> None:
    for path, content in list(_touched.items()):
        path.write_text(content, encoding="utf-8", newline="")
    _touched.clear()


def run_pytest(test_path: str, name: str) -> str:
    base = tempfile.mkdtemp(prefix=f"mutdef_{name}_")
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
        print("用法: python scripts/mutate_python_docker_defects.py M1|M5|M6")
        return 2
    file, expect_fail = TARGETS[which]

    baseline = run_pytest(DEFECT_TESTS, "baseline")
    bp, bf = _counts(baseline)
    print(f"[基线] {bp} passed, {bf} failed")
    if bp == 0 or bf:
        restore_all()  # 防御：若上一次异常退出残留了 touched，先还原
        print("BASELINE-NOT-GREEN：基线未全绿，终止（源码未改动）。")
        print(baseline[-1500:])
        return 4

    original = file.read_text(encoding="utf-8")
    seed = SEEDS[which]
    occurrences = original.count(seed)
    if occurrences != 1:
        print(f"SKIP: {which} SEED 出现 {occurrences} 次（期望 1）。")
        print("SEED 片段：\n" + seed)
        return 3

    _touched[file] = original
    try:
        file.write_text(original.replace(seed, MUTANTS[which]), encoding="utf-8", newline="")
        mutated = run_pytest(DEFECT_TESTS, "mutated")
        failed = re.findall(r"FAILED (server/tests/\w+\.py::[\w\[\]\-]+)", mutated)
        failed_names = sorted({n.split("::")[1].split("[")[0] for n in failed})
        hits = [n for n in expect_fail if n in failed_names]
        print(f"[变异后] 失败 {len(failed_names)} 条：{', '.join(failed_names) or '（无）'}")
        print(f"[命中判别点] {', '.join(hits) if hits else '无'}")
        red = bool(failed_names) and bool(hits)
        print("RED-OK：变异使判别测试独立变红。" if red else "RED-MISS：变异未按预期变红。")
        if not red:
            print(mutated[-2500:])
    finally:
        restore_all()

    restored = run_pytest(DEFECT_TESTS, "restored")
    rp, rf = _counts(restored)
    print(f"[还原后] {rp} passed, {rf} failed")
    green = rp > 0 and rf == 0
    print("RESTORED-GREEN：还原后复跑全绿。" if green else "RESTORED-NOT-GREEN：还原后仍有失败！")
    return 0 if (red and green) else 5


if __name__ == "__main__":
    sys.exit(main())
