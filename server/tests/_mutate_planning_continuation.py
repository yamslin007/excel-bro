"""反向变异：证明 A0「继续分析」的守卫真的有判别力。

纪律（与仓库既有约定一致）：

1. **一个变异只有一个判别点**，且必须落在该机制**真正管辖**的属性上。
2. **RED 的四条判据**：退出码非 0、摘要里确有 failed、**点名的那条用例**真的在
   FAILED 列表里、失败原文里带着它自己那一处证据。只看"退出码非 0"不算——一行
   写崩了语法也是非 0，那是运行期错误，不是判据在起作用。
3. 还原后必须逐字节 sha256 一致，并复跑全绿。
4. 字节级 IO、按原 EOL 写回、异常/中断也要还原。

用法：python server/tests/_mutate_planning_continuation.py
"""

from __future__ import annotations

import hashlib
import os
import pathlib
import re
import signal
import subprocess
import sys
import tempfile

ROOT = pathlib.Path(__file__).resolve().parents[2]
PYTHON = ROOT / ".venv" / "Scripts" / "python.exe"
TESTS = [
    "server/tests/test_planning_continuation.py",
    "server/tests/test_claude_sdk_agent.py",
]
TARGET = ROOT / "server" / "app" / "main.py"
ADAPTER = ROOT / "server" / "app" / "claude_sdk_agent.py"
TARGETS = {"main": TARGET, "adapter": ADAPTER}
ORIGINALS = {key: path.read_bytes() for key, path in TARGETS.items()}
DIGESTS = {
    key: hashlib.sha256(payload).hexdigest() for key, payload in ORIGINALS.items()
}

# (名字, 文件, 原文种子, 变异后, 必须变红的用例 id 片段, 失败原文里应出现的证据)
MUTATIONS: list[tuple[str, str, str, str, str, str]] = [
    (
        # 校验照跑，但把**所有**拒绝都报成同一个码：接口仍然 409，只是理由错了。
        # 这是"契约判别点没在起作用"最接近真实事故的形态（放行/错码），
        # 而且不会顺手制造运行期错误 —— 后者只证明代码崩了，不证明判据有判别力。
        "M1-不再区分拒绝码（把任何不合格现场都说成 INVALID）",
        "main",
        "        code = verdict.code or planning_context.CODE_INVALID",
        "        code = planning_context.CODE_INVALID",
        "test_planning_continuation.py::test_bad_checkpoint_is_refused_by_the_continue_endpoint",
        "assert",
    ),
    (
        "M2-不看执行账本就放行继续",
        "main",
        "    if _folder_operation_ledger.has_execution_for_task(record.turn_id):",
        "    if False:",
        "test_planning_continuation.py::test_registered_execution_blocks_continuation_in_every_state",
        "assert",
    ),
    (
        "M3-不比对工作簿身份（换簿也能继续）",
        "main",
        "    if current_identity != recorded_identity:\n"
        "        return planning_context.CODE_TARGET_CHANGED",
        "    if False:  # [变异M3] 不比对工作簿身份\n"
        "        return planning_context.CODE_TARGET_CHANGED",
        "test_planning_continuation.py::test_checkpoint_with_foreign_workbook_is_refused_before_planning",
        "assert",
    ),
    (
        "M4-不比对授权范围（收窄/扩大都放行）",
        "main",
        "    if current_sheets != recorded_sheets:\n"
        "        return planning_context.CODE_SCOPE_CHANGED",
        "    if False:  # [变异M4] 不比对授权成员\n"
        "        return planning_context.CODE_SCOPE_CHANGED",
        "test_planning_continuation.py::test_narrowed_authorization_does_not_reuse_the_old_checkpoint",
        "assert",
    ),
    (
        "M5-不比对来源指纹（数据变了也继续）",
        "main",
        "        if submitted != recorded_fingerprint.strip():\n"
        "            return planning_context.CODE_SOURCE_CHANGED",
        "        if False:  # [变异M5] 不比对来源指纹\n"
        "            return planning_context.CODE_SOURCE_CHANGED",
        "test_planning_continuation.py::test_scope_or_source_change_is_refused_by_the_server",
        "assert",
    ),
    (
        "M6-Excel 登记不写 taskId（执行闸门在 Excel 侧恒假）",
        "main",
        "            task_id=request.taskId,",
        "            task_id=None,",
        "test_planning_continuation.py::test_registered_execution_blocks_continuation_in_every_state",
        "assert",
    ),
    (
        "M7-适配器不向静调用方透传 on_checkpoint",
        "adapter",
        "        await _save_checkpoint(\"before_model\")",
        "        pass",
        "test_claude_sdk_agent.py::test_sdk_reports_before_model_checkpoint_with_only_whitelisted_facts",
        "assert",
    ),
    (
        "M8-落盘失败直接抛出（把可继续问题变成本轮失败）",
        "adapter",
        "        except Exception:\n"
        "            # 落盘失败只降低\"能不能继续\"的能力，不能让本轮规划崩掉，\n"
        "            # 更不能因此把任务判成失败。\n"
        "            logger.warning(\"无法保存规划现场 stage=%s\", stage, exc_info=True)",
        "        except ZeroDivisionError:\n"
        "            logger.warning(\"无法保存规划现场 stage=%s\", stage, exc_info=True)",
        "test_claude_sdk_agent.py::test_sdk_checkpoint_failure_does_not_break_the_turn",
        "OSError",
    ),
    (
        "M9-续跑投影不再过滤引用形状（把任意字符串带进提示）",
        "adapter",
        "        if (\n            isinstance(ref, str)\n            and ref.startswith(\"res_\")\n            and len(ref) == 36\n            and all(char in \"0123456789abcdef\" for char in ref[4:])\n        ):",
        "        if isinstance(ref, str):",
        "test_claude_sdk_agent.py::test_resume_projection_carries_only_whitelisted_strings",
        "assert",
    ),
]


def _restore_all() -> None:
    for key, path in TARGETS.items():
        path.write_bytes(ORIGINALS[key])


def _install_restore_hooks() -> None:
    def handler(signum, _frame):  # noqa: ANN001
        _restore_all()
        sys.exit(128 + signum)

    for name in ("SIGINT", "SIGTERM", "SIGHUP"):
        number = getattr(signal, name, None)
        if number is not None:
            signal.signal(number, handler)


def _run_tests() -> tuple[int, str]:
    # basetemp 必须**每次唯一**：固定的 basetemp 目录被一次异常中断留下后，
    # 后续每次运行都会在 setup 阶段踩 `PermissionError [WinError 5]`，
    # 于是"45 errors + 18 passed"会被误读成"变异没变红"。基线不绿的时候，
    # 任何 RED 都不算证据 —— 所以这里既换新目录，也在 main() 里加了基线闸。
    base = pathlib.Path(tempfile.mkdtemp(prefix="pytest-excelbro-a0-", dir=str(ROOT / "outputs")))
    proc = subprocess.run(
        [
            str(PYTHON),
            "-m",
            "pytest",
            *TESTS,
            "-p",
            "no:cacheprovider",
            "-q",
            f"--basetemp={base}",
        ],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env={
            **os.environ,
            "TMP": str(pathlib.Path(tempfile.gettempdir())),
            "TEMP": str(pathlib.Path(tempfile.gettempdir())),
        },
    )
    output = re.sub(r"\x1b\[[0-9;]*m", "", (proc.stdout or "") + (proc.stderr or ""))
    # 刻意**不**删 basetemp：一次 rmtree 会同时删掉几十个文件，触发环境的批量删除闸门，
    # 把驱动自己打断（实测 [SAFE_DELETE_BULK_CONFIRM_REQUIRED]）。留着的代价只是几个
    # `outputs/pytest-excelbro-a0-*` 目录，而"驱动被自己的清理杀掉"会让整轮结论作废。
    return proc.returncode, output


def _summary_line(output: str) -> str:
    """pytest -q 的摘要行形如 `8 failed, 13 passed in 6.17s`（没有 `=` 分隔）。"""
    lines = [
        line.strip()
        for line in output.splitlines()
        if re.search(r"\b\d+ (?:passed|failed)\b", line)
    ]
    return lines[-1] if lines else "(无测试摘要)"


def _verify_red(
    code: int, output: str, expected_test: str, evidence: str
) -> tuple[bool, str]:
    """RED 的四条判据（缺一条都算"没证明判别力"）。"""
    summary = _summary_line(output)
    if code == 0:
        return False, f"退出码 0（未变红） {summary}"
    # ② 确实有失败的用例，而不是踩到运行期错误/采集失败。
    failed_match = re.search(r"(\d+) failed", output)
    if failed_match is None or int(failed_match.group(1)) < 1:
        return False, f"退出码非 0 但没有 failed 计数（更像运行期错误） {summary}"
    if int(failed_match.group(1)) > 10:
        return False, f"大面积崩溃（{failed_match.group(1)} 条），不是这一处守卫的证据 {summary}"
    # ③ 红的必须是**点名的那条**用例。
    failed_lines = [line for line in output.splitlines() if line.startswith("FAILED ")]
    if not any(expected_test in line for line in failed_lines):
        return False, f"红的不是点名的那条 {summary}"
    # ④ 失败原文里要有**它自己那一处**：断言发生在该测试文件内（行号可定位），
    #    而且带着本变异特有的证据词 —— 只卡退出码会把"写崩了 JSX/Python"也算红。
    module = expected_test.split("::")[0]
    location = re.search(
        rf"{re.escape(module)}:(\d+):", output.replace("/", "\\")
    )
    if location is None:
        return False, f"失败位置不在被测用例里（可能是运行期错误） {summary}"
    if evidence not in output:
        return False, f"失败原文里没有证据词「{evidence}」 {summary}"
    ran = re.search(r"(\d+) passed", output)
    if ran is None:
        return False, f"没有有效测试计数 {summary}"
    return True, f"{summary}（首处 {module}:{location.group(1)}）"


def main() -> int:
    _install_restore_hooks()

    # 基线闸：未变异的代码必须**先绿**。基线红/报错时，后面每一条"变红"都可能是
    # 同一个运行期故障的重复，无法归因到被测守卫 —— 这时直接拒绝出结论。
    baseline_code, baseline_output = _run_tests()
    baseline_summary = _summary_line(baseline_output)
    if baseline_code != 0:
        print(f"基线不绿，拒绝继续（变异结论此时无效）：{baseline_summary}")
        for line in baseline_output.splitlines():
            if line.startswith(("ERROR ", "FAILED ")):
                print(f"  {line}")
        return 1
    print(f"基线 GREEN：{baseline_summary}")

    ok = True
    try:
        for name, key, seed, replacement, expected_test, evidence in MUTATIONS:
            target = TARGETS[key]
            original = ORIGINALS[key]
            seed_bytes = seed.replace("\n", "\r\n").encode("utf-8")
            if seed_bytes not in original:
                seed_bytes = seed.encode("utf-8")
            if seed_bytes not in original:
                print(f"{name}: SEED 未命中 -> 无法作证")
                ok = False
                continue
            mutated = original.replace(
                seed_bytes,
                replacement.replace("\n", "\r\n").encode("utf-8")
                if seed_bytes.endswith(b"\r\n")
                else replacement.encode("utf-8"),
                1,
            )
            target.write_bytes(mutated)
            code, output = _run_tests()
            good, detail = _verify_red(code, output, expected_test, evidence)
            print(f"{name}: {'RED(OK)' if good else 'GREEN(NG!)'}  {detail}")
            if not good:
                ok = False
            target.write_bytes(original)
    finally:
        _restore_all()

    changed = [
        key
        for key, path in TARGETS.items()
        if hashlib.sha256(path.read_bytes()).hexdigest() != DIGESTS[key]
    ]
    if changed:
        print(f"还原失败（字节不一致）: {', '.join(changed)}")
        return 1
    code, output = _run_tests()
    summary = _summary_line(output)
    print(f"还原后逐字节一致；复跑 {'GREEN' if code == 0 else 'RED'}：{summary}")
    return 0 if ok and code == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
