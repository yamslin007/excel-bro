"""反向变异：拆掉"收窄为空集时不得产出空 anyOf"这道机制 → 用例必须变红。

被测机制（2026-10-11 修复，`server/app/agent_tools.py::host_pruned_terminal_definitions`）：

- M1：**省略空能力工具** —— 拆掉 `if not candidates: continue`，恢复原样
  `{"anyOf": []}` ⇒「上界为空集时不得产出空 anyOf」用例变红
  （空 anyOf 是非法 JSON Schema，实测 DeepSeek anthropic 端点会 400 整轮请求）。
- M2：**收窄仍然生效** —— 把候选集合改成**无条件**全量
  ⇒「收窄非空时仍是原语义」用例变红
  （这条防止用"永远不裁剪"来假装修好 M1）。
- M3：**未声明宿主逐字不变** —— 把 `if not request.hostId` 的早退改成总是裁剪
  ⇒「未声明宿主定义逐字不变」用例变红。

判据（缺一不算过）：① 退出码 0 ② 有有效测试摘要 ③ 实跑了预期用例（minTests）
④ 还原后逐字节一致（sha256）。

SEED 一次只 apply/restore 一个；写回前三处先归一化 LF。
"""

from __future__ import annotations

import hashlib
import re
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
TOOLS = REPO / "server" / "app" / "agent_tools.py"
TEST_FILE = "server/tests/test_planning_capability_narrowing.py"
PYTHON = REPO / ".venv" / "Scripts" / "python.exe"
MIN_TESTS = 15

SEEDS = [
    {
        "name": "M1-省略空能力工具被拆（恢复空 anyOf）",
        "old": "        if not candidates:\n            continue\n",
        "new": "        if False:\n            continue\n",
    },
    {
        "name": "M2-收窄被架空（无条件全量变体）",
        "old": "    candidates = [variants[name] for name in selected]\n",
        "new": "    candidates = list(variants.values())\n",
    },
    {
        "name": "M3-未声明宿主的早退被拆（未声明也裁剪）",
        "old": "    if not request.hostId:\n        return [tool for tool in AGENT_TOOLS_WITH_DATA]\n",
        "new": "    if False:\n        return [tool for tool in AGENT_TOOLS_WITH_DATA]\n",
    },
]


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run_tests() -> tuple[int, str]:
    proc = subprocess.run(
        [str(PYTHON), "-m", "pytest", TEST_FILE, "-q", "-p", "no:cacheprovider"],
        cwd=str(REPO), capture_output=True, text=True,
    )
    return proc.returncode, proc.stdout + proc.stderr


def count_tests(out: str) -> int:
    total = 0
    for match in re.finditer(r"(\d+) (passed|failed)", out):
        total += int(match.group(1))
    return total


def main() -> int:
    original = TOOLS.read_bytes()
    original_sha = sha(TOOLS)
    print(f"baseline sha256 {TOOLS.name} = {original_sha}")

    crlf = b"\r\n" in original
    text = original.decode("utf-8").replace("\r\n", "\n")

    failures: list[str] = []
    try:
        for seed in SEEDS:
            if text.count(seed["old"]) != 1:
                print(f"{seed['name']}: SEED 未命中或不唯一 -> 失败")
                failures.append(seed["name"])
                continue
            mutated = text.replace(seed["old"], seed["new"])
            payload = mutated.replace("\n", "\r\n") if crlf else mutated
            TOOLS.write_bytes(payload.encode("utf-8"))
            print(f"APPLIED {seed['name']} → {TOOLS.relative_to(REPO)}")
            try:
                code, out = run_tests()
            finally:
                TOOLS.write_bytes(original)
            if sha(TOOLS) != original_sha:
                print("还原 sha 不一致 -> 立即报错")
                return 2
            total = count_tests(out)
            summary = [
                line for line in out.splitlines() if "passed" in line or "failed" in line
            ]
            tail = summary[-1].strip() if summary else "(无摘要)"
            red = code != 0 and "failed" in out
            ran = total >= MIN_TESTS
            ok = red and ran
            print(
                f"{seed['name']}: exit={code} | {tail} | tests={total} | "
                f"{'RED OK 判别力成立' if ok else 'GREEN X 无判别力'}"
            )
            if not ok:
                failures.append(seed["name"])
    finally:
        TOOLS.write_bytes(original)

    if sha(TOOLS) != original_sha:
        print(f"最终还原失败: {TOOLS}")
        return 2
    print(f"final sha256 {TOOLS.name} = {sha(TOOLS)}")

    code, out = run_tests()
    summary = [line for line in out.splitlines() if "passed" in line]
    print(f"baseline rerun: exit={code} | {summary[-1].strip() if summary else '(无摘要)'}")
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
