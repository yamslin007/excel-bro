"""反向变异：证明「普通后端结果终局工具（注册 / 分派 / 归属与授权校验）」有判别力。

用法：python server/tests/_mutate_normal_backend_result_tools.py

判据（缺一不算过）：
  ① 退出码 0
  ② 有有效测试摘要
  ③ 实跑了预期用例（minTests）
  ④ 每个变异 RED、还原后逐字节一致（sha256）

每个变异必须有**唯一判别点**：拆掉它，对应断言必须变红。两个后端各有一份独立
实现时变异要分拆各证各的；这里普通后端自带实现的部分（注册、分派）单独拆，
共享层的归属 / 授权校验在各自的数据源上单独拆。
"""
from __future__ import annotations

import hashlib
import os
import pathlib
import re
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parents[2]
TESTS = ["server/tests/test_normal_backend_result_tools.py"]
BASETEMP = "C:/Users/jia/AppData/Local/Temp/pytest-excelbro"
MIN_TESTS = 17

EXCEL = ROOT / "server" / "app" / "excel_agent.py"
REGISTRY = ROOT / "server" / "app" / "local_results.py"

# (名称, [(文件, seed, 替换), ...])
#
# M5 是**冗余层**：重授权既会在 ``_bind_locked`` 里直接作废该任务的旧值，``resolve``
# 里还会再比一次登记时的授权版本。单拆任一道都 GREEN —— 那是判别点选错，不是断言
# 没判别力。冗余层必须**一次拆全**，才能证明「重授权后旧值读不出」这条结论真的被
# 钉住。
MUTATIONS: list[tuple[str, list[tuple[pathlib.Path, str, str]]]] = [
    ("M1-终局工具不注册（模型拿不到引用入口）", [
        (EXCEL,
         "    extra = list(local_terminal_tool_definitions())",
         "    extra = []"),
    ]),
    ("M2-不派发 submit_local_result（旧行为：未知工具）", [
        (EXCEL, "    if name == SUBMIT_LOCAL_RESULT_TOOL:", "    if False:"),
    ]),
    ("M3-不派发 submit_result_plan（旧行为：未知工具）", [
        (EXCEL, "    if name == SUBMIT_RESULT_PLAN_TOOL:", "    if False:"),
    ]),
    ("M4-注册表不校验归属（知道句柄就能读）", [
        (REGISTRY,
         "        if item is None or item.task_id != task_id:",
         "        if item is None:"),
    ]),
    ("M5-重授权既不换版本也不作废旧值（冗余两道一起拆）", [
        (REGISTRY,
         "            self._discard_task_locked(task_id)\n"
         "            revision = existing.revision + 1",
         "            revision = existing.revision + 1"),
        (REGISTRY,
         "        if (\n"
         "            claim is not None\n"
         "            and item.claim_revision is not None\n"
         "            and item.claim_revision != claim.revision\n"
         "        ):",
         "        if False:"),
    ]),
]


def _encode_for(text: str, raw: bytes) -> bytes:
    """按**目标文件自己的换行风格**编码 seed / 替换文本。

    仓库里并非所有文件同一种换行（``excel_agent.py`` 是 CRLF，
    ``local_results.py`` 是 LF）；统一按一种转会在另一种上"SEED 未命中"，
    那是脚本的错，不是断言没判别力。
    """
    newline = "\r\n" if b"\r\n" in raw else "\n"
    return text.replace("\r\n", "\n").replace("\n", newline).encode("utf-8")


def run_tests() -> tuple[bool, str, int]:
    env = {"TMP": "C:/Users/jia/AppData/Local/Temp",
           "TEMP": "C:/Users/jia/AppData/Local/Temp"}
    proc = subprocess.run(
        [str(ROOT / ".venv" / "Scripts" / "python.exe"), "-m", "pytest",
         *TESTS, "-p", "no:cacheprovider", "-q", f"--basetemp={BASETEMP}"],
        cwd=str(ROOT), capture_output=True, text=True, encoding="utf-8",
        errors="replace", env={**os.environ, **env},
    )
    out = re.sub(r"\x1b\[[0-9;]*m", "", proc.stdout or "")
    passed_match = re.search(r"(\d+) passed", out)
    failed_match = re.search(r"(\d+) failed", out)
    ran = (int(passed_match.group(1)) if passed_match else 0) + (
        int(failed_match.group(1)) if failed_match else 0
    )
    lines = [line for line in out.splitlines() if line.strip()]
    summary = lines[-1] if lines else out[-200:]
    green = proc.returncode == 0 and "passed" in out and "failed" not in out
    return green, summary, ran


def main() -> int:
    ok = True
    originals = {
        EXCEL: EXCEL.read_bytes(),
        REGISTRY: REGISTRY.read_bytes(),
    }
    digest = {path: hashlib.sha256(raw).hexdigest() for path, raw in originals.items()}

    for name, edits in MUTATIONS:
        current = dict(originals)
        applied = True
        for path, seed, replacement in edits:
            raw = current[path]
            seed_b = _encode_for(seed, raw)
            if seed_b not in raw:
                print(f"{name}: SEED 未命中({path.name}) -> FAIL")
                ok = False
                applied = False
                break
            current[path] = raw.replace(
                seed_b, _encode_for(replacement, raw), 1
            )
        if not applied:
            continue
        for path, raw in current.items():
            path.write_bytes(raw)
        green, summary, ran = run_tests()
        ran_ok = ran >= MIN_TESTS
        verdict = "GREEN(NG!)" if green else "RED(OK)"
        if green or not ran_ok:
            ok = False
        print(f"{name}: {verdict} ran={ran}  {summary}")
        for path, raw in originals.items():
            path.write_bytes(raw)

    for path, raw in originals.items():
        if hashlib.sha256(path.read_bytes()).hexdigest() != digest[path]:
            print(f"还原失败（逐字节不一致）：{path}")
            ok = False
    green, summary, ran = run_tests()
    print(f"还原: {'OK' if green else 'FAILED'} ran={ran}  {summary}")
    return 0 if (ok and green) else 1


if __name__ == "__main__":
    sys.exit(main())
