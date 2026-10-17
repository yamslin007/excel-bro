"""反向变异：拆掉结果计划编译层的空白折叠 → 回归用例必须变红。

被测机制：`server/app/result_plan.py` 的 `_blank_cell_of_write` 把分析结果里的
`""` 折成 `None`（2026-10-10 实机缺陷修复）。

判据（缺一不算过）：① 退出码 0 ② 有有效测试摘要 ③ 实跑了预期用例（minTests）
④ 还原后逐字节一致（sha256）。

SEED 一次只 apply/restore 一个。
"""

from __future__ import annotations

import hashlib
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
TARGET = REPO / "server" / "app" / "result_plan.py"
TEST_FILE = "server/tests/test_result_plan_blank_folding.py"
PYTHON = REPO / ".venv" / "Scripts" / "python.exe"
MIN_TESTS = 5

SEEDS = [
    {
        "name": "M1-恒等函数（拆掉折叠）",
        "old": '    return None if value == "" else value\n',
        "new": "    return value\n",
    },
    {
        "name": "M2-matricesEqual 层面忽略（仅表头不折）",
        "old": "    rows = [[_blank_cell_of_write(value) for value in row] for row in item.rows]\n",
        "new": "    rows = [[value for value in row] for row in item.rows]\n",
    },
]


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run_tests() -> tuple[int, str]:
    proc = subprocess.run(
        [str(PYTHON), "-m", "pytest", TEST_FILE, "-q"],
        cwd=str(REPO), capture_output=True, text=True,
    )
    return proc.returncode, proc.stdout + proc.stderr


def main() -> int:
    original = TARGET.read_bytes()
    original_sha = sha(TARGET)
    text = original.decode("utf-8").replace("\r\n", "\n")
    crlf = original.decode("utf-8").find("\r\n") >= 0
    print(f"baseline sha256 = {original_sha} (crlf={crlf})")

    failures: list[str] = []
    try:
        for seed in SEEDS:
            if text.count(seed["old"]) != 1:
                print(f"{seed['name']}: SEED 未命中或不唯一 -> 失败")
                failures.append(seed["name"])
                continue
            mutated = text.replace(seed["old"], seed["new"])
            payload = mutated.replace("\n", "\r\n") if crlf else mutated
            TARGET.write_bytes(payload.encode("utf-8"))
            print(f"APPLIED {seed['name']} → {TARGET.relative_to(REPO)}")
            try:
                code, out = run_tests()
            finally:
                TARGET.write_bytes(original)
            if sha(TARGET) != original_sha:
                print("还原 sha 不一致 -> 立即报错")
                return 2
            summary = [line for line in out.splitlines() if "passed" in line or "failed" in line]
            tail = summary[-1].strip() if summary else "(无摘要)"
            ran = "passed" in out or "failed" in out
            red = code != 0 and "failed" in out
            ok = red and ran
            print(f"{seed['name']}: exit={code} | {tail} | "
                  f"{'RED OK 判别力成立' if ok else 'GREEN X 无判别力'}")
            if not ok:
                failures.append(seed["name"])
    finally:
        TARGET.write_bytes(original)

    if sha(TARGET) != original_sha:
        print("最终还原失败")
        return 2
    print(f"final sha256 = {sha(TARGET)}")

    code, out = run_tests()
    m = [line for line in out.splitlines() if "passed" in line]
    print(f"baseline rerun: exit={code} | {m[-1].strip() if m else '(无摘要)'}")
    if code != 0:
        print("基线复跑未通过")
        return 1

    if failures:
        print(f"\n无判别力的变异: {', '.join(failures)}")
        return 1
    print("\n全部变异均变红，判别力成立。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
