"""反向变异：拆掉规划时间策略的每一道机制 → 时间策略用例必须变红。

被测机制（2026-10-10 改造，`server/app/excel_agent.py` +
`server/app/agent_loop_protection.py`）：

- M1：``note_progress`` 的**非阻塞提醒** —— 拆成不发提醒 ⇒ 提醒用例变红。
- M2：``note_progress(enforce_total_wall_time=True)`` 的**总时限收口** —— 拆成不抛
  暂停 ⇒ 收口用例变红。
- M3：`should_stop` 的 ``timeout`` **回归** —— 把墙上时间判定加回去（旧行为）
  ⇒ 判"收口不在 should_stop"的用例变红。
- M4：`note_progress` 的**检查点保存** —— 拆成不调 ``save_boundary``
  ⇒ 检查点用例变红（证明"取消暂停 ≠ 取消检查点"这条被判据钉住）。

判据（缺一不算过）：① 退出码 0 ② 有有效测试摘要 ③ 实跑了预期用例（minTests）
④ 还原后逐字节一致（sha256）。

SEED 一次只 apply/restore 一个。同一文件多个 SEED 不叠加。
"""

from __future__ import annotations

import hashlib
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
AGENT = REPO / "server" / "app" / "excel_agent.py"
PROTECTION = REPO / "server" / "app" / "agent_loop_protection.py"
TEST_FILE = "server/tests/test_planning_time_policy.py"
PYTHON = REPO / ".venv" / "Scripts" / "python.exe"
MIN_TESTS = 7

# 每个 SEED 说明改哪个文件、替换什么、期望哪条用例变红。
SEEDS = [
    {
        "name": "M1-提醒不发送（拆掉非阻塞提醒）",
        "target": AGENT,
        "old": '        await emit(\n'
               '            f"planning.reminder.{int(state.elapsed_time // reminder_seconds)}",\n',
        "new": '        await emit(\n'
               '            f"planning.noreminder.{int(state.elapsed_time // reminder_seconds)}",\n',
    },
    {
        "name": "M2-总时限不收口（拆掉 enforce 的抛错）",
        "target": AGENT,
        "old": "        raise AgentBudgetPause(checkpoint, loop_config.max_wall_time_seconds)\n",
        "new": "        return\n",
    },
    {
        "name": "M3-should_stop 复活墙上时间熔断（旧行为回归）",
        "target": PROTECTION,
        "old": "    return False, None\n",
        "new": "    if state.elapsed_time >= config.max_wall_time_seconds:\n"
               '        return True, "timeout"\n'
               "    return False, None\n",
    },
    {
        "name": "M4-边界不落检查点（拆掉 save_boundary）",
        "target": AGENT,
        "old": "        nonlocal last_reminder_at\n"
               "        checkpoint = await save_boundary(\n"
               "            stage,\n"
               "            upcoming_turn=upcoming_turn,\n"
               "            pending_message=pending_message,\n"
               "            upcoming_tool=upcoming_tool,\n"
               "            tool_batch_state=tool_batch_state,\n"
               "        )\n",
        "new": "        nonlocal last_reminder_at\n"
               "        checkpoint = None\n",
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


def count_tests(out: str) -> int:
    """从摘要里数实际跑了几条（passed/failed 之和）。"""
    import re

    total = 0
    for match in re.finditer(r"(\d+) (passed|failed)", out):
        total += int(match.group(1))
    return total


def main() -> int:
    originals = {AGENT: AGENT.read_bytes(), PROTECTION: PROTECTION.read_bytes()}
    original_shas = {p: sha(p) for p in originals}
    for path, digest in original_shas.items():
        print(f"baseline sha256 {path.name} = {digest}")

    crlf = {p: b"\r\n" in data for p, data in originals.items()}
    texts = {
        p: data.decode("utf-8").replace("\r\n", "\n") for p, data in originals.items()
    }

    failures: list[str] = []
    try:
        for seed in SEEDS:
            target: Path = seed["target"]
            text = texts[target]
            if text.count(seed["old"]) != 1:
                print(f"{seed['name']}: SEED 未命中或不唯一 -> 失败")
                failures.append(seed["name"])
                continue
            mutated = text.replace(seed["old"], seed["new"])
            payload = mutated.replace("\n", "\r\n") if crlf[target] else mutated
            target.write_bytes(payload.encode("utf-8"))
            print(f"APPLIED {seed['name']} → {target.relative_to(REPO)}")
            try:
                code, out = run_tests()
            finally:
                target.write_bytes(originals[target])
            if sha(target) != original_shas[target]:
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
        for path, data in originals.items():
            path.write_bytes(data)

    for path, digest in original_shas.items():
        if sha(path) != digest:
            print(f"最终还原失败: {path}")
            return 2
        print(f"final sha256 {path.name} = {sha(path)}")

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
