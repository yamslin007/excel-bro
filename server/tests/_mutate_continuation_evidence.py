"""反向变异：证明 A0 补充修复的「继续必须带本轮核实过的现场证据」真有判别力。

被测机制（缺一即失真）：

1. **协议必填**：`PlanningContinueRequest` 的三项证据不许省略（省略拿不到续跑权）。
2. **服务端内容判定**：空身份/空成员 ⇒ `EVIDENCE_MISSING`；来源三态各自有明确落点。
3. **不为旧客户端保留"缺证据就放行"的路径**。

E1 就是把这些一起退回旧缺陷 —— 字段重新可选、缺了就跳过比对、缺证据直接放行。
它**必须**让「缺失证据被拒」那条用例变红；变不绿说明这条闸只是装饰。

纪律与仓库既有驱动一致：
- 一个变异只有一个判别点（E1 = "缺证据能不能拿到续跑权"），多 hunk 同属这一点；
- RED 四判据：退出码非 0、摘要确有 failed、点名用例真在 FAILED 里、失败原文带自己那处证据；
- 基线必须先绿（基线红时任何 RED 都无法归因）；
- 还原后逐字节 sha256 一致并复跑全绿；
- basetemp 每次唯一（固定目录被中断留下后会 PermissionError，把"没变红"伪装成"环境错误"）。

用法：
    python server/tests/_mutate_continuation_evidence.py            # 跑全部
    python server/tests/_mutate_continuation_evidence.py --probe    # 只打印变异后的失败原文
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
import time

ROOT = pathlib.Path(__file__).resolve().parents[2]
PYTHON = ROOT / ".venv" / "Scripts" / "python.exe"
TESTS = ["server/tests/test_planning_continuation.py"]
FILES = {
    "models": ROOT / "server" / "app" / "models.py",
    "main": ROOT / "server" / "app" / "main.py",
}
ORIGINALS = {key: path.read_bytes() for key, path in FILES.items()}
DIGESTS = {key: hashlib.sha256(payload).hexdigest() for key, payload in ORIGINALS.items()}

# (名字, [(文件, 原文种子, 变异后)], 必须变红的用例 id 片段, 失败原文里应出现的证据)
MUTATIONS: list[tuple[str, list[tuple[str, str, str]], str, str]] = [
    (
        "E1-恢复旧缺陷：证据字段可选且缺失即跳过校验",
        [
            (
                "models",
                # 1/3：三项证据里两项重新变回可选。
                "    workbookIdentity: Annotated[\n"
                "        str,\n"
                "        StringConstraints(strip_whitespace=True, min_length=5, max_length=4000),\n"
                "    ]\n"
                "    authorizedSheetNames: list[str] = Field(min_length=1, max_length=200)",
                "    # [变异E1] 证据重新变回可选\n"
                "    workbookIdentity: Annotated[\n"
                "        str,\n"
                "        StringConstraints(strip_whitespace=True, min_length=5, max_length=4000),\n"
                "    ] | None = None\n"
                "    authorizedSheetNames: list[str] = Field(default_factory=list, max_length=200)",
            ),
            (
                "models",
                '    sourceVerification: Literal["verified", "not_collected", "failed"]\n',
                '    sourceVerification: Literal["verified", "not_collected", "failed"] | None = None\n',
            ),
            (
                "models",
                '        if self.sourceVerification == "verified" and not (\n'
                "            self.sourceFingerprint and self.sourceFingerprint.strip()\n"
                "        ):\n"
                '            raise ValueError("声明已核实来源时必须提供来源指纹")\n',
                "        if self.sourceVerification is None:\n"
                "            # [变异E1] 缺字段就跳过一致性校验\n"
                "            return self\n"
                '        if self.sourceVerification == "verified" and not (\n'
                "            self.sourceFingerprint and self.sourceFingerprint.strip()\n"
                "        ):\n"
                '            raise ValueError("声明已核实来源时必须提供来源指纹")\n',
            ),
            (
                "main",
                # 2/3：工作簿身份缺失不再拒绝。
                "    if not current_identity:\n"
                "        return planning_context.CODE_EVIDENCE_MISSING\n",
                "    if not current_identity:\n"
                "        return None  # [变异E1] 缺身份就放行\n",
            ),
            (
                "main",
                # 3/3：授权成员缺失不再拒绝。
                "    if not current_sheets:\n"
                "        return planning_context.CODE_EVIDENCE_MISSING\n",
                "    if not current_sheets:\n"
                "        return None  # [变异E1] 缺授权成员就放行\n",
            ),
        ],
        "test_planning_continuation.py::test_missing_evidence_field_is_rejected_before_anything_runs",
        # 实测原文：缺证据的请求拿到了 200，桩里的规划器被真的调起来了。
        "assert 200 == 422",
    ),
]


def _write_verified(path: pathlib.Path, payload: bytes, *, attempts: int = 8) -> None:
    """写回并**读回校验**。

    实测踩过：连续快速写同一个 `.py` 会被索引/杀软短暂占住，`write_bytes` 抛
    `OSError: [Errno 22] Invalid argument`。变异驱动一旦在还原那一步抛错，就可能
    把变异留在工作区里 —— 所以这里重试并且每次都读回比对，不允许"写了就 assume 成功"。
    """
    last: Exception | None = None
    for index in range(attempts):
        try:
            path.write_bytes(payload)
            if path.read_bytes() == payload:
                return
            last = OSError(f"{path} 读回不一致")
        except OSError as error:  # pragma: no cover - 环境相关
            last = error
        time.sleep(0.2 * (index + 1))
    raise SystemExit(f"写回失败且无法恢复：{path}（{last}）")


def _restore_all() -> None:
    for key, path in FILES.items():
        _write_verified(path, ORIGINALS[key])


def _install_restore_hooks() -> None:
    def handler(signum, _frame):  # noqa: ANN001
        _restore_all()
        sys.exit(128 + signum)

    for name in ("SIGINT", "SIGTERM", "SIGHUP"):
        number = getattr(signal, name, None)
        if number is not None:
            signal.signal(number, handler)


def _run_tests() -> tuple[int, str]:
    base = pathlib.Path(
        tempfile.mkdtemp(prefix="pytest-excelbro-ev-", dir=str(ROOT / "outputs"))
    )
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
    )
    output = re.sub(r"\x1b\[[0-9;]*m", "", (proc.stdout or "") + (proc.stderr or ""))
    return proc.returncode, output


def _summary_line(output: str) -> str:
    lines = [
        line.strip()
        for line in output.splitlines()
        if re.search(r"\b\d+ (?:passed|failed)\b", line)
    ]
    return lines[-1] if lines else "(无测试摘要)"


def _apply(mutation: list[tuple[str, str, str]]) -> str | None:
    """按 hunk 依次替换字节；任何一个种子没命中就整体还原并报错。"""
    payloads = {key: ORIGINALS[key] for key in FILES}
    for key, seed, mutant in mutation:
        original = payloads[key]
        seed_bytes = seed.replace("\n", "\r\n").encode("utf-8")
        newline = b"\r\n"
        if seed_bytes not in original:
            seed_bytes = seed.encode("utf-8")
            newline = b"\n"
        if seed_bytes not in original:
            return f"{key} 的 SEED 未命中：{seed.splitlines()[0][:60]}"
        mutant_bytes = mutant.replace("\n", newline.decode()).encode("utf-8")
        payloads[key] = original.replace(seed_bytes, mutant_bytes, 1)
    for key, payload in payloads.items():
        _write_verified(FILES[key], payload)
    return None


def _verify_red(code: int, output: str, expected_test: str, evidence: str) -> tuple[bool, str]:
    summary = _summary_line(output)
    if code == 0:
        return False, f"退出码 0（未变红） {summary}"
    failed_match = re.search(r"(\d+) failed", output)
    if failed_match is None or int(failed_match.group(1)) < 1:
        return False, f"退出码非 0 但没有 failed 计数（更像运行期错误） {summary}"
    if int(failed_match.group(1)) > 10:
        return False, f"大面积崩溃（{failed_match.group(1)} 条），不是这处守卫的证据 {summary}"
    failed_lines = [line for line in output.splitlines() if line.startswith("FAILED ")]
    if not any(expected_test in line for line in failed_lines):
        return False, f"红的不是点名的那条 {summary}"
    module = expected_test.split("::")[0]
    location = re.search(rf"{re.escape(module)}:(\d+):", output.replace("/", "\\"))
    if location is None:
        return False, f"失败位置不在被测用例里（可能是运行期错误） {summary}"
    if evidence not in output:
        return False, f"失败原文里没有证据词「{evidence}」 {summary}"
    if re.search(r"(\d+) passed", output) is None:
        return False, f"没有有效测试计数 {summary}"
    return True, f"{summary}（首处 {module}:{location.group(1)}）"


def main() -> int:
    _install_restore_hooks()
    probe = "--probe" in sys.argv

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
        for name, mutation, expected_test, evidence in MUTATIONS:
            applied = _apply(mutation)
            if applied is not None:
                print(f"{name}: {applied} -> 无法作证")
                _restore_all()
                ok = False
                continue
            code, output = _run_tests()
            if probe:
                print(output)
            good, detail = _verify_red(code, output, expected_test, evidence)
            print(f"{name}: {'RED(OK)' if good else 'GREEN(NG!)'}  {detail}")
            if not good:
                ok = False
                for line in output.splitlines():
                    if line.startswith("FAILED "):
                        print(f"  {line}")
            _restore_all()
    finally:
        _restore_all()

    changed = [
        key
        for key, path in FILES.items()
        if hashlib.sha256(path.read_bytes()).hexdigest() != DIGESTS[key]
    ]
    if changed:
        print(f"还原失败（字节不一致）: {', '.join(changed)}")
        return 1
    code, output = _run_tests()
    print(f"还原后逐字节一致；复跑 {'GREEN' if code == 0 else 'RED'}：{_summary_line(output)}")
    return 0 if ok and code == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
