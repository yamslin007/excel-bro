#!/usr/bin/env python
"""反向变异：拆掉 `run_python_analysis` 描述里的代码契约 -> 用例必须变红。

纪律（见项目 MEMORY）：
- 同文件多 SEED **一次只 apply/restore 一个**；
- 读/写都先归一化 LF；写回按原换行转回；
- 还原后**逐字节一致**（sha256）；任何一处不符立刻中止；
- 判定用各变异器自己的词形。
"""
from __future__ import annotations

import hashlib
import re
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
TARGET = ROOT / "server" / "app" / "python_analysis.py"
TEST_FILE = "server/tests/test_python_analysis_tool_contract.py"

# 契约段原文（真实文件词形：源码里是字面反斜杠 + n 两个字符）。
CONTRACT_BLOCK = (
    '                "\\n\\n代码契约（必须遵守，否则本工具返回不可用回执、拿不到任何结果）："\n'
    '                "\\n1. 输入数据只从代码顶层变量 input 读取：input[\\"datasets\\"] 是数组，"\n'
    '                "每一项形如 {\\"source\\": {...}, \\"values\\": [[表头行], [数据行], ...]}，"\n'
    '                "其顺序与你传给 prepare_python_data 的 sources 顺序一致。"\n'
    '                "\\n2. 代码必须在**模块级（顶层）**给变量 result 赋值；不要用函数返回值"\n'
    '                "或 print 传递结果 —— 只认模块级 result。"\n'
    '                "\\n3. result 必须是对象 {\\"kind\\": \\"matrix\\", \\"values\\": [[...], ...]}："\n'
    '                "kind 固定为字符串 \\"matrix\\"，values 是二维矩阵（字符串/数字/布尔/null），"\n'
    '                "且每一行长度相同、至少一行一列。"\n'
    '                "\\n4. 把要展示的列名作为 values 的第一行；汇总、总计、明细都应作为普通"\n'
    '                "行写进同一个矩阵（结果会以固定值矩阵的形式进入结果卡与写入计划）。"\n'
)

CODE_PARAM_BLOCK = (
    '                            "要运行的 Python 分析代码。从 input[\\"datasets\\"][i][\\"values\\"] "\n'
    '                            "读取数据，并在模块级给 result 赋值："\n'
    '                            "result = {\\"kind\\": \\"matrix\\", \\"values\\": [[...], ...]}。"\n'
    '                            "只认模块级 result（不用函数返回值、不靠 print）。"\n'
)

MATRIX_SHAPE_LINE = (
    '                "\\n3. result 必须是对象 {\\"kind\\": \\"matrix\\", \\"values\\": [[...], ...]}："\n'
    '                "kind 固定为字符串 \\"matrix\\"，values 是二维矩阵（字符串/数字/布尔/null），"\n'
)

SEEDS = [
    ("M1", CONTRACT_BLOCK, '                "\\n\\n请遵守代码契约。"\n'),
    ("M2", CODE_PARAM_BLOCK, '                            "要运行的 Python 分析代码。"\n'),
    ("M3", MATRIX_SHAPE_LINE, '                "\\n3. result 必须是对象（形状未说明）。"\n'),
]


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write_retry(path: Path, data: bytes, attempts: int = 15) -> None:
    """Windows 上文件可能被瞬时占用（索引/杀软/编辑器）；重试而不是失败。"""
    last: OSError | None = None
    for _ in range(attempts):
        try:
            path.write_bytes(data)
            return
        except OSError as exc:
            last = exc
            time.sleep(0.4)
    assert last is not None
    raise last


def run_tests() -> tuple[int, str]:
    proc = subprocess.run(
        [sys.executable, "-m", "pytest", TEST_FILE, "-q"],
        cwd=ROOT, capture_output=True, text=True,
    )
    out = re.sub(r"\x1b\[[0-9;]*m", "", proc.stdout + proc.stderr)
    return proc.returncode, out


def main() -> int:
    original_bytes = TARGET.read_bytes()
    original_sha = sha(TARGET)
    original_text = original_bytes.decode("utf-8")
    lf_text = original_text.replace("\r\n", "\n")
    crlf = "\r\n" in original_text
    print(f"baseline sha256 = {original_sha} (crlf={crlf})")

    def restore() -> None:
        write_retry(TARGET, original_bytes)
        got = sha(TARGET)
        if got != original_sha:
            raise RuntimeError(f"还原失败：{got} != {original_sha}")

    failures: list[str] = []
    try:
        for name, old, new in SEEDS:
            if old not in lf_text:
                print(f"{name}: SEED 未命中（检查词形）-> 记为失败")
                failures.append(name)
                continue
            if lf_text.count(old) != 1:
                print(f"{name}: SEED 命中 {lf_text.count(old)} 次，不唯一 -> 记为失败")
                failures.append(name)
                continue
            mutated_lf = lf_text.replace(old, new)
            payload = mutated_lf.replace("\n", "\r\n") if crlf else mutated_lf
            write_retry(TARGET, payload.encode("utf-8"))
            try:
                code, out = run_tests()
            finally:
                restore()
            summary = [
                ln for ln in out.splitlines()
                if "passed" in ln or "failed" in ln or "error" in ln
            ]
            tail = summary[-1] if summary else "(无摘要)"
            ok = code != 0 and bool(summary) and "failed" in " ".join(summary)
            verdict = "RED  OK 判别力成立" if ok else "GREEN/无摘要 X 无判别力"
            print(f"{name}: exit={code} | {tail} | {verdict}")
            if not ok:
                failures.append(name)
    finally:
        restore()
        print(f"final sha256 = {sha(TARGET)}")

    if failures:
        print(f"\n无判别力的变异: {failures}")
        return 1
    print("\n全部变异均变红，判别力成立。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
