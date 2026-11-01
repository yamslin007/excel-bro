"""定点变异：证明落盘本轮新增/改写的用例判别力（用法：python scripts/mutate_resume_acceptance.py M1-DROP-FIELD）。

纪律（沿用仓库约定）：
- 一次只开一个 SEED；
- 归一化 LF 读写、按原 EOL 写回；
- 写回后读回校验；
- SIGINT/TERM + uncaughtException 都装还原钩子；
- 还原后逐字节 sha256 必须与变异前一致；
- 四判据：退出码、实跑到的用例、失败原文里有它自己那处证据、还原后哈希。
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
EXCEL_TEST = "apps/excel-addin/src/excelFormulaResume.test.ts"

SEEDS = {
    "M1-DROP-FIELD": {
        "file": "server/app/models.py",
        "old": '    formulaR1C1: str | None = Field(default=None, max_length=512)\n',
        "new": "",
        "runner": "pytest",
        "test": SERVER_TEST,
        "expect_case": "test_resumed_state_plan_carries_formulas_equal_for_formula_actions",
        "expect_evidence": "formulaR1C1",
    },
    "M2-DROP-CRITERIA-ON-STORE": {
        "file": "server/app/main.py",
        "old": '            plan=request.plan.model_dump(mode="json"),',
        "new": '            plan=request.plan.model_dump(mode="json", exclude={"acceptanceCriteria"}),',
        "runner": "pytest",
        "test": SERVER_TEST,
        "expect_case": "test_resumed_state_plan_carries_formulas_equal_for_formula_actions",
        "expect_evidence": "formulasEqual",
        "known_survivor": "读路径会重新推导，持久化里有没有它对响应没有区别",
    },
    "M3-NO-FORMULAS-EQUAL-DERIVATION": {
        "file": "server/app/models.py",
        "old": '            elif action.type == "writeFormulas":\n'
        "                add_criterion(\n"
        "                    FormulasEqualCriterion(\n"
        '                        type="formulasEqual",\n'
        "                        sheet=action.sheet,\n"
        "                        range=action.range,\n"
        "                        expected=action.formulas,\n"
        "                    )\n"
        "                )\n",
        "new": '            elif False and action.type == "writeFormulas":\n'
        "                add_criterion(\n"
        "                    FormulasEqualCriterion(\n"
        '                        type="formulasEqual",\n'
        "                        sheet=action.sheet,\n"
        "                        range=action.range,\n"
        "                        expected=action.formulas,\n"
        "                    )\n"
        "                )\n",
        "runner": "pytest",
        "test": SERVER_TEST,
        "expect_case": "test_resumed_state_plan_carries_formulas_equal_for_formula_actions",
        "expect_evidence": "formulasEqual",
    },
    # 让公式被当成"烘焙矩阵"（改**共享判据**，不是 excel.ts 那份本地清单）⇒ 走来源值指纹比对。
    #
    # 注：先试过只改 `apps/excel-addin/src/excel.ts` 的 `BAKED_ACTION_TYPES`，结果
    # 「来源值变了」那条**仍然绿**：`planLiveness` 变 mixed 后落进 else 分支，而那里
    # 的 `judgePlanSourceGate` 读的是 core 的 `BAKED_WRITE_ACTION_TYPES`，仍判
    # runtime-evaluated ⇒ ok。这反过来说明 excel.ts 与 core 各存一份"烘焙动作清单"，
    # 两边漂移时结论不随任一份单独改动而变 —— 证据见 M4B。
    "M4-FORMULAS-COUNT-AS-BAKED": {
        "file": "packages/core/src/planSourceAnchor.ts",
        "old": 'export const BAKED_WRITE_ACTION_TYPES = ["writeValues", "writeTable"] as const;',
        "new": 'export const BAKED_WRITE_ACTION_TYPES = ["writeValues", "writeTable", "writeFormulas"] as const;',
        "runner": "vitest",
        "test": EXCEL_TEST,
        "expect_case": "续跑时来源值变了",
        "expect_evidence": "数据来源在预览后发生了变化",
    },
    # 只改 excel.ts 本地清单：预期「来源值变了」存活、只有"来源表不在"那条红。
    "M4B-LOCAL-BAKED-LIST-ONLY": {
        "file": "apps/excel-addin/src/excel.ts",
        "old": 'const BAKED_ACTION_TYPES = new Set<ExcelAction["type"]>([\n  "writeValues",\n  "writeTable"\n]);',
        "new": 'const BAKED_ACTION_TYPES = new Set<ExcelAction["type"]>([\n  "writeValues",\n  "writeTable",\n  "writeFormulas"\n]);',
        "runner": "vitest",
        "test": EXCEL_TEST,
        "expect_case": "续跑时点名的来源表已不在工作簿里",
        "expect_evidence": "instead of rejecting",
        "known_survivor": "「来源值变了」那条仍绿：共享判据仍判 runtime-evaluated",
    },
    # 关掉 live 分支的结构体检 ⇒ 来源表不在现场也照写。
    "M5-NO-SCHEMA-CHECK-FOR-LIVE": {
        "file": "apps/excel-addin/src/excel.ts",
        "old": "      if (missing.length > 0) {",
        "new": "      if (false && missing.length > 0) {",
        "runner": "vitest",
        "test": EXCEL_TEST,
        "expect_case": "续跑时点名的来源表已不在工作簿里",
        "expect_evidence": "instead of rejecting",
    },
    # 反过来：往 live 分支**加**一条来源值比对（模拟"有人把公式也当烘焙"）⇒
    # 「来源值变了照常续跑」那条必须红，而结构体检那条仍绿（判别点唯一）。
    "M6-LIVE-BRANCH-ALSO-COMPARES-VALUES": {
        "file": "apps/excel-addin/src/excel.ts",
        "old": "      const missing = missingSchemaRequirements(\n"
        "        plan,\n"
        "        await captureWorkbookStructure(liveSourceSheets)\n"
        "      );\n",
        "new": "      const missing = missingSchemaRequirements(\n"
        "        plan,\n"
        "        await captureWorkbookStructure(liveSourceSheets)\n"
        "      );\n"
        "      /* 变异：让 live 分支也比对来源值指纹（公式计划本该不比对）。 */\n"
        "      if ((await captureWorkbookSourceFingerprint(liveSourceSheets)) !== anchor!.fingerprint)\n"
        "        missing.push(\"来源值已变化\");\n",
        "runner": "vitest",
        "test": EXCEL_TEST,
        "expect_case": "续跑时来源值变了",
        "expect_evidence": "计划所需的来源已发生变化",
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
    if text.count(seed["old"]) != 1:
        raise SystemExit(f"SEED 未命中或命中多处：{seed_name} count={text.count(seed['old'])}")
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

    if seed["runner"] == "pytest":
        basetemp = tempfile.mkdtemp(prefix="eb-mut-")
        command = [str(PYTHON), "-m", "pytest", seed["test"], "-q", f"--basetemp={basetemp}"]
    else:
        basetemp = None
        command = [NPX, "vitest", "run", seed["test"], "--reporter=verbose"]

    artifacts = ROOT / "outputs" / "mutation-2026-10-30"
    artifacts.mkdir(parents=True, exist_ok=True)
    log = artifacts / f"{seed_name}.log"
    try:
        completed = subprocess.run(
            command,
            cwd=str(ROOT),
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
        red_cases = [line for line in red_cases if "test" in line]
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
        print(f"SEED {seed_name}  [{seed['runner']}] {seed['test']}")
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
