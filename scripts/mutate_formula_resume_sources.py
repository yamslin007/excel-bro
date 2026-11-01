"""定点变异：证明「公式恢复要求来源证据一致」这一轮用例真有判别力。

用法：python scripts/mutate_formula_resume_sources.py N2（一次只开一个 SEED）。

纪律（沿用仓库约定，逐条都在这份脚本里落实）：
- 一次只开一个 SEED；归一化 LF 读写、按原 EOL 写回；写回后**读回校验**
  （`Errno 22` 会把变异留在工作区，所以写失败要重试并在恢复时再验 sha256）；
- SIGINT/TERM/HUP + 恢复钩子；杀驱动前先收尾；
- 四判据：退出码、实跑到的用例（红名单里有它）、失败原文含它自己那处证据、
  还原后逐字节 sha256 与变异前一致；
- **ANSI 转义会让 `"× " in line` 恒假**，所以先 strip 再判（`strip_ansi`）。

与上一版（`mutate_resume_acceptance.py`）的关系：那一版里的 `M4B` 指向
`apps/excel-addin/src/excel.ts` 的本地烘焙清单，而本轮的改动把这个清单**删掉并
统一到 core**，于是 `M4B` 变成 `SEED 未命中`（= 假装有证据）。本文件是新定点，
`M4B` 与 `M6` 不再存在；`M1/M2/M3/M5` 原样搬过来继续有效。
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import signal
import subprocess
import sys
import tempfile
from pathlib import Path

ANSI = re.compile(r"\x1b\[[0-9;]*m")


def strip_ansi(text: str) -> str:
    return ANSI.sub("", text)


ROOT = Path(r"E:\codex\excel-bro")
PYTHON = ROOT / ".venv" / "Scripts" / "python.exe"
NPX = "npx.cmd" if sys.platform == "win32" else "npx"

SERVER_TEST = "server/tests/test_office_resume_acceptance.py"
EXCEL_TEST = "src/excelFormulaResume.test.ts"
CORE_TEST = "src/formulaWriteback.test.ts"
WPS_TEST = "src/wpsExecutionResume.test.ts"
# vitest 的环境与配置按 **cwd** 解析（WPS 用 `environment: "jsdom"`，在仓库根跑会
# 报 `window is not defined`），所以每个 SEED 都要带自己的 cwd。
EXCEL_DIR = "apps/excel-addin"
CORE_DIR = "packages/core"
WPS_DIR = "apps/wps-addin"

SEEDS = {
    # ── 旧定点（仍在）：协议字段 / 验收条件推导 / live 分支结构体检 ──────────
    "M1-DROP-R1C1-FIELD": {
        "file": "server/app/models.py",
        "old": "    formulaR1C1: str | None = Field(default=None, max_length=512)\n",
        "new": "",
        "runner": "pytest",
        "test": SERVER_TEST,
        "expect_case": "test_r1c1_formula_plan_is_rejected_at_registration",
        "expect_evidence": "formulaR1C1",
    },
    "M2-DROP-CRITERIA-ON-STORE": {
        "file": "server/app/main.py",
        "old": "            plan=request.plan.model_dump(mode=\"json\"),",
        "new": "            plan=request.plan.model_dump(mode=\"json\", exclude={\"acceptanceCriteria\"}),",
        "runner": "pytest",
        "test": SERVER_TEST,
        "expect_case": "test_resumed_state_plan_carries_formulas_equal_for_formula_actions",
        "expect_evidence": "formulasEqual",
        "known_survivor": "读路径会重新推导，持久化里有没有它对响应没有区别",
    },
    "M3-NO-FORMULAS-EQUAL-DERIVATION": {
        "file": "server/app/models.py",
        "old": '            elif action.type == "writeFormulas":\n',
        "new": '            elif False and action.type == "writeFormulas":\n',
        "runner": "pytest",
        "test": SERVER_TEST,
        "expect_case": "test_resumed_state_plan_carries_formulas_equal_for_formula_actions",
        "expect_evidence": "formulasEqual",
    },
    "M5-NO-SCHEMA-CHECK-FOR-LIVE": {
        "file": "apps/excel-addin/src/excel.ts",
        "old": "      if (missing.length > 0) {",
        "new": "      if (false && missing.length > 0) {",
        "runner": "vitest",
        "test": EXCEL_TEST,
        "cwd": EXCEL_DIR,
        "expect_case": "续跑时点名的来源表已不在工作簿里",
        "expect_evidence": "计划所需的来源已发生变化",
    },
    # ── 本轮新增定点：来源证据这条链，逐环拆 ────────────────────────────────
    # 1) 服务端少声明字段 ⇒ 客户端带上来的指纹被静默丢掉 ⇒ 读响应里没有证据。
    "N1-DROP-DEPENDENCY-FINGERPRINT-FIELD": {
        "file": "server/app/models.py",
        "old": "    sourceDependencyFingerprint: str | None = Field(default=None, min_length=1, max_length=128)\n",
        "new": "",
        "runner": "pytest",
        "test": SERVER_TEST,
        "expect_case": "test_source_dependency_fingerprint_survives_registration_and_read",
        "expect_evidence": "sourceDependencyFingerprint",
    },
    # 2) 服务端不校验"有指纹没口径" ⇒ 一串没有表集合的哈希也能登记。
    "N2-NO-SERVER-EVIDENCE-PAIRING": {
        "file": "server/app/main.py",
        "old": "    if request.plan.sourceDependencyFingerprint and not dependency:",
        "new": "    if False and request.plan.sourceDependencyFingerprint and not dependency:",
        "runner": "pytest",
        "test": SERVER_TEST,
        "expect_case": "test_fingerprint_without_dependency_sheets_is_rejected",
        # 变异后这条请求被放行：失败原文里是那份**登记成功**的响应，
        # 里面同时有指纹和空口径 —— 那就是"缺配对校验"的证据本身。
        "expect_evidence": "\"sourceDependencySheets\":[]",
    },
    # 3) 端上撤掉续跑前的来源校验 ⇒ 来源变了照写（这正是本轮要修的行为）。
    "N3-NO-RESUME-SOURCE-GATE": {
        "file": "apps/excel-addin/src/excel.ts",
        "old": "  if (startActionIndex > 0) {\n    const resumeSourceVerdict",
        "new": "  if (false && startActionIndex > 0) {\n    const resumeSourceVerdict",
        "runner": "vitest",
        "test": EXCEL_TEST,
        "cwd": EXCEL_DIR,
        "expect_case": "续跑时来源值变了",
        "expect_evidence": "来源已变化",
    },
    # 4) 判据本体改成"只看有没有登记过"（等于拿公式文本/登记存在性替代来源校验）：
    #    来源值变了也放行。判别点唯一 —— 阈值就在最后那一行的比较上。
    "N4-TEXT-INSTEAD-OF-VALUES": {
        "file": "packages/core/src/formulaWriteback.ts",
        "old": "  return current === recorded\n    ? { verdict: \"ok\", sheets: requirement.sheets }\n"
        "    : { verdict: \"refused\", reason: FORMULA_SOURCE_CHANGED_MESSAGE };",
        "new": "  return { verdict: \"ok\", sheets: requirement.sheets };",
        "runner": "vitest",
        "test": CORE_TEST,
        "cwd": CORE_DIR,
        "expect_case": "公式文本没变、来源值变了",
        "expect_evidence": "来源已变化",
    },
    # 5) 登记闸拆掉：依赖面不可分离（同表/动态/外部/名称）仍登记精确恢复。
    "N5-NO-SAME-SHEET-REFUSAL": {
        # 判别点放在**共享判据**里那个具体的拒绝理由上（只拆"同表引用"这一条），
        # 不动 excel.ts 的转发：这样其他几条理由各自仍绿，判别点唯一。
        "file": "packages/core/src/formulaWriteback.ts",
        "old": "    if (deps.local) {",
        "new": "    if (false && deps.local) {",
        "runner": "vitest",
        "test": EXCEL_TEST,
        "cwd": EXCEL_DIR,
        "expect_case": "同表公式（=A1*2）",
        "expect_evidence": "写入目标表自身",
    },
    # 6) 登记闸只判"依赖面可分离"，不判"依赖表有没有本机读取记录" ⇒
    #    读过 A 却写 `=B!A1` 的计划也能登记精确恢复（没有 B 的值基线）。
    "N6-NO-LEDGER-COVERAGE": {
        "file": "apps/excel-addin/src/excel.ts",
        "old": "    if (uncovered.length > 0) {",
        "new": "    if (false && uncovered.length > 0) {",
        "runner": "vitest",
        "test": EXCEL_TEST,
        "cwd": EXCEL_DIR,
        "expect_case": "公式读了记账里没有的表",
        "expect_evidence": "本机读取记录",
    },
    # 7) WPS 侧的同一道闸拆掉 ⇒ "读了没有读取记录的表"仍会登记精确续跑材料。
    # 拆这道闸时故意选一个**看似能跑**的替代：把"不登记"改成"登记一个没有来源的
    # 空上下文"。这比 `if (false)` 更接近真实误改（会直接 TypeError 崩在下一行，
    # 那样的红证明不了判据的语义）。
    "N7-WPS-NO-FORMULA-GATE": {
        "file": "apps/wps-addin/src/wpsPlanExecutor.ts",
        "old": "  if (!formulaRequirement.ok) return null;",
        "new": "  if (!formulaRequirement.ok) return {version: 1, hostVersion: prepared.hostVersion, sources: []};",
        "runner": "vitest",
        "test": WPS_TEST,
        "cwd": WPS_DIR,
        "expect_case": "动态引用的公式",
        "expect_evidence": "to be null",
    },
    # 8) 烘焙清单合并后只剩一份：把它扩成含公式 ⇒ 公式计划走"烘焙"分支，
    #    live 语义被改写（这条证明那个清单确实在主导 liveness，而不是死代码）。
    "N8-BAKED-LIST-INCLUDES-FORMULAS": {
        "file": "packages/core/src/planSourceAnchor.ts",
        "old": 'export const BAKED_WRITE_ACTION_TYPES = ["writeValues", "writeTable"] as const;',
        "new": 'export const BAKED_WRITE_ACTION_TYPES = ["writeValues", "writeTable", "writeFormulas"] as const;',
        "runner": "vitest",
        "test": EXCEL_TEST,
        "cwd": EXCEL_DIR,
        "expect_case": "续跑被拒之后",
        "expect_evidence": "算好的固定值",
    },
    # ── 本轮新增定点：WPS 的来源**覆盖**（以前"表名出现在读取记录里"就算覆盖） ──
    # 9) 登记阶段不再要求"有证据真的盖住这张依赖表" ⇒ 证据盖不住整表也照样登记
    #    精确恢复（超限那张表就只能靠裸证据"猜它没变"）。
    "N9-WPS-REGISTRATION-NO-COVERAGE": {
        "file": "apps/wps-addin/src/wpsPlanExecutor.ts",
        "old": "    return !range || !prepared.sources.some(source => source.sheetName === name && source.range === range);",
        "new": "    return false && !range && !prepared.sources.some(source => source.sheetName === name && source.range === range);",
        "runner": "vitest",
        "test": WPS_TEST,
        "cwd": WPS_DIR,
        "expect_case": "依赖表有效区域大到无法登记证据",
        "expect_evidence": "to be null",
    },
    # 10) 采集退回旧口径：**表名**出现在已有证据里就直接复用，不再看有没有盖住
    #     整张有效区域 ⇒ "公式读 B10、读取记录只到 A1:A3"重新变回可登记，
    #     清空 B10 也发现不了（有效区域缩回登记范围，只有整块指纹看得出来）。
    "N10-WPS-COVERAGE-BY-SHEET-NAME": {
        "file": "apps/wps-addin/src/wpsPlanExecutor.ts",
        "old": "        .find((source) => source.sheetName === name && rectContains(parseA1Range(source.range), used));",
        "new": "        .find((source) => source.sheetName === name);",
        "runner": "vitest",
        "test": WPS_TEST,
        "cwd": WPS_DIR,
        "expect_case": "公式引用的那一格被清空（有效区域缩回登记范围）",
        "expect_evidence": "来源已变化",
    },
    # 11) 覆盖核对整体失效 ⇒ 只比登记矩形里的值，矩形之外长大的一格看不见
    #     （"有效区域变大"那条会走完续跑，而不是停在来源已变化）。
    "N11-WPS-NO-USED-RANGE-CONTAINMENT": {
        "file": "apps/wps-addin/src/wpsPlanExecutor.ts",
        "old": "  if (requirement.sheets.length === 0) return;\n  const sources = state.descriptor.resumeContext?.sources ?? [];",
        "new": "  if (true) return;\n  const sources = state.descriptor.resumeContext?.sources ?? [];",
        "runner": "vitest",
        "test": WPS_TEST,
        "cwd": WPS_DIR,
        "expect_case": "依赖表原来为空的一格被填上",
        "expect_evidence": "来源已变化",
    },
    # 12) 把"只为恢复而采的基线"也塞进执行期来源复核 ⇒ 预览之后改一个只出现在基线里的格，
    #     正常公式写入被拦（"请重新预览"），与"不改变正常公式自动重算"这条规则冲突。
    "N12-WPS-COVERAGE-BLOCKS-EXECUTION": {
        "file": "apps/wps-addin/src/wpsPlanExecutor.ts",
        "old": "    if (source.coverageOnly) continue;\n",
        "new": "    if (false && source.coverageOnly) continue;\n",
        "runner": "vitest",
        "test": WPS_TEST,
        "cwd": WPS_DIR,
        "expect_case": "预览后依赖表变了",
        "expect_evidence": "在预览后发生变化",
    },
    # ── 持久化往返（第二十三轮核实）：判据的输入必须真的跨进程活着 ──
    # 13) 恢复包只投影"模型读过的那块"（把补采的覆盖矩形丢掉）⇒ 重启后材料里没有 A1:B10：
    #     正向用例直接红；"B10 被清空"那条会**通过**（区域缩回、指纹只覆盖 A1:A3），
    #     正是"覆盖证据没落盘"的样子。
    "P1-PACKAGE-DROPS-COVERAGE-RECT": {
        "file": "apps/wps-addin/src/wpsRecoveryPackage.ts",
        "old": "      sources: record.resume.context.sources.map(source => ({ ...source }))",
        "new": "      sources: record.resume.context.sources.filter(source => source.range !== \"A1:B10\").map(source => ({ ...source }))",
        "runner": "vitest",
        "test": WPS_TEST,
        "cwd": WPS_DIR,
        "expect_case": "重启后只用恢复包",
        "expect_evidence": "A1:B10",
    },
}

_state: dict[str, object] = {}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def restore() -> None:
    path = _state.get("path")
    original = _state.get("original")
    if isinstance(path, Path) and isinstance(original, bytes):
        for _ in range(5):
            try:
                path.write_bytes(original)
            except OSError:
                continue
            if sha256(path) == _state["digest"]:
                break
        print(f"[restore] {path.name} sha256={sha256(path)}")


def _hook(signum, frame):  # noqa: ANN001
    restore()
    sys.exit(130)


def run(seed_name: str) -> None:
    seed = SEEDS[seed_name]
    path = ROOT / seed["file"]
    original_bytes = path.read_bytes()
    digest = hashlib.sha256(original_bytes).hexdigest()
    text = original_bytes.decode("utf-8").replace("\r\n", "\n")
    count = text.count(seed["old"])
    if count != 1:
        raise SystemExit(f"SEED 未命中或命中多处：{seed_name} count={count}")
    mutated = text.replace(seed["old"], seed["new"], 1)
    newline = "\r\n" if b"\r\n" in original_bytes else "\n"
    _state.update({"path": path, "original": original_bytes, "digest": digest})

    for _ in range(5):
        path.write_bytes(mutated.replace("\n", newline).encode("utf-8"))
        if path.read_bytes().decode("utf-8").replace("\r\n", "\n") == mutated:
            break
    else:
        restore()
        raise SystemExit(f"写回校验失败：{seed_name}")

    workdir = ROOT / seed.get("cwd", ".")
    if seed["runner"] == "pytest":
        basetemp = tempfile.mkdtemp(prefix="eb-mut-")
        command = [str(PYTHON), "-m", "pytest", seed["test"], "-q", f"--basetemp={basetemp}"]
    else:
        basetemp = None
        command = [NPX, "vitest", "run", seed["test"], "--reporter=verbose"]

    artifacts = ROOT / "outputs" / "mutation-2026-10-30-formula-sources"
    artifacts.mkdir(parents=True, exist_ok=True)
    log = artifacts / f"{seed_name}.log"
    try:
        completed = subprocess.run(
            command,
            cwd=str(workdir),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env={**os.environ, "NO_COLOR": "1", "FORCE_COLOR": "0", "CI": "1"},
        )
        combined = strip_ansi(completed.stdout + completed.stderr)
        log.write_text(combined, encoding="utf-8")
        red_cases = [
            line.strip()
            for line in combined.splitlines()
            if "× " in line or line.strip().startswith(("FAILED", "ERROR", "FAIL "))
        ]
        red_cases = [line for line in red_cases if "test" in line or ">" in line]
        matched = [line for line in red_cases if seed["expect_case"] in line]
        if not matched:
            matched = [
                line.strip()
                for line in combined.splitlines()
                if "FAIL " in line and seed["expect_case"] in line
            ]
        evidence = seed["expect_evidence"] in combined
        summary_line = [
            line.strip()
            for line in combined.splitlines()
            if " passed" in line or " failed" in line
        ]
        print("=" * 70)
        print(f"SEED {seed_name}  [{seed['runner']}] {seed.get('cwd', '.')}/{seed['test']}")
        print(f"  退出码        : {completed.returncode}")
        print(f"  红/错误条目   : {len(red_cases)}")
        for line in red_cases[:6]:
            print(f"    {line}")
        print(f"  命中预期用例  : {bool(matched)}")
        print(f"  失败原文含证据: {evidence} ({seed['expect_evidence']!r})")
        print(f"  摘要行        : {summary_line[-1] if summary_line else '(无)'}")
        if seed.get("known_survivor") and completed.returncode == 0:
            print(f"  已知存活理由  : {seed['known_survivor']}")
        json.dump(
            {
                "seed": seed_name,
                "file": seed["file"],
                "runner": seed["runner"],
                "exit": completed.returncode,
                "red": red_cases,
                "expected_case_red": bool(matched),
                "evidence_in_output": evidence,
                "summary": summary_line[-1] if summary_line else None,
                "known_survivor": seed.get("known_survivor"),
            },
            open(artifacts / f"{seed_name}.json", "w", encoding="utf-8"),
            ensure_ascii=False,
            indent=2,
        )
    finally:
        restore()
        if basetemp:
            shutil.rmtree(basetemp, ignore_errors=True)
        after = sha256(path)
        print(f"  还原后 sha256 : {after} {'== 原始' if after == digest else '!! 与原始不同 !!'}")


if __name__ == "__main__":
    for sig in (signal.SIGINT, signal.SIGTERM, getattr(signal, "SIGHUP", None)):
        if sig is None:
            continue
        try:
            signal.signal(sig, _hook)
        except (ValueError, AttributeError, OSError):
            pass
    run(sys.argv[1])
