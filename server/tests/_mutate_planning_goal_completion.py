"""反向变异：拆掉规划目标完成判据的每一道机制 → 判据用例必须变红。

被测机制（2026-10-11 修复，`server/app/result_submission.py` +
`server/app/excel_agent.py` + `server/app/claude_sdk_agent.py`）：

- M1：**共享判据本体** —— 把 ``evaluate_local_result_submission`` 拆成永不抛错
  ⇒「声明 write_preview 却只交结果卡」的用例变红。
- M2：**普通后端接线** —— 拆掉 ``_execute_tool`` 里那一次判据调用
  ⇒「普通后端拒绝提前回答」的用例变红。
- M3：**SDK 侧不置终局** —— 让未完成分支照样 ``terminal = ...``
  ⇒「SDK 不锁死后续工具」的用例变红（这是缺陷 3 的直接成因）。
- M4：**默认值方向** —— 把 ``normalize_planning_goal`` 的缺省改成
  ``write_preview`` ⇒「缺省按只读放行」的用例变红。

判据（缺一不算过）：① 退出码 0 ② 有有效测试摘要 ③ 实跑了预期用例（minTests）
④ 还原后逐字节一致（sha256）。

SEED 一次只 apply/restore 一个。同一文件多个 SEED 不叠加。
"""

from __future__ import annotations

import hashlib
import re
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SUBMISSION = REPO / "server" / "app" / "result_submission.py"
AGENT = REPO / "server" / "app" / "excel_agent.py"
SDK = REPO / "server" / "app" / "claude_sdk_agent.py"
TEST_FILE = "server/tests/test_planning_goal_completion.py"
PYTHON = REPO / ".venv" / "Scripts" / "python.exe"
MIN_TESTS = 12

SEEDS = [
    {
        "name": "M1-判据本体失效（拆掉 evaluate 的抛错）",
        "target": SUBMISSION,
        "old": '    if goal == "write_preview":\n        raise LocalResultIncomplete(\n',
        "new": '    if goal == "write_preview" and False:\n        raise LocalResultIncomplete(\n',
    },
    {
        "name": "M2-普通后端不接线（拆掉 _execute_tool 的判据调用）",
        "target": AGENT,
        "old": "        evaluate_local_result_submission(\n"
               '            normalize_planning_goal(arguments.get("planningGoal"))\n'
               "        )\n",
        "new": "        pass  # mutation: judgement removed\n",
    },
    {
        "name": "M3-SDK 未完成分支照样置终局（锁死后续工具）",
        "target": SDK,
        "old": "                    except LocalResultIncomplete as error:\n"
               "                        return {\n",
        "new": "                    except LocalResultIncomplete as error:\n"
               "                        terminal = build_local_result_answer(item)\n"
               "                        return {\n",
    },
    {
        "name": "M4-缺省值方向反了（改成默认 write_preview）",
        "target": SUBMISSION,
        "old": '    return "write_preview" if str(value or "").strip() == "write_preview" else "read_only"\n',
        "new": '    return "read_only" if str(value or "").strip() == "read_only" else "write_preview"\n',
    },
    {
        "name": "M5-编译被拒回执不标错误（拆掉 is_error）",
        "target": SDK,
        "old": "                                ensure_ascii=False,\n"
               "                            )}],\n"
               "                            \"is_error\": True,\n"
               "                        }\n"
               "                    terminal = prepared\n",
        "new": "                                ensure_ascii=False,\n"
               "                            )}],\n"
               "                        }\n"
               "                    terminal = prepared\n",
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
    originals = {
        SUBMISSION: SUBMISSION.read_bytes(),
        AGENT: AGENT.read_bytes(),
        SDK: SDK.read_bytes(),
    }
    original_shas = {p: sha(p) for p in originals}
    for path, digest in original_shas.items():
        print(f"baseline sha256 {path.name} = {digest}")

    crlf = {p: b"\r\n" in data for p, data in originals.items()}
    texts = {p: data.decode("utf-8").replace("\r\n", "\n") for p, data in originals.items()}

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
