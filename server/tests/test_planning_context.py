"""planning_context 契约的两端一致性判据。

跑的是**共享夹具** `packages/core/src/planningContext.fixtures.json` —— 与
`packages/core/src/planningContext.test.ts` 同一批用例。任何一端放宽判据，
这里的期望码就对不上。
"""

from __future__ import annotations

import copy
import json
from pathlib import Path

import pytest

from server.app import planning_context as pc

REPOSITORY = Path(__file__).resolve().parents[2]
FIXTURES = REPOSITORY / "packages" / "core" / "src" / "planningContext.fixtures.json"


@pytest.fixture(scope="module")
def fixtures() -> dict:
    return json.loads(FIXTURES.read_text(encoding="utf-8"))


def _materialize(entry: dict, base: dict) -> object:
    if "record" in entry:
        return entry["record"]
    record = copy.deepcopy(base)
    for key in entry.get("remove", []):
        record.pop(key, None)
    for key, value in entry.get("patch", {}).items():
        if value == "__LONG_PROMPT__":
            value = "长" * (pc.MAX_PROMPT_CHARS + 1)
        record[key] = value
    return record


def test_fixtures_cover_every_rejection_code(fixtures: dict) -> None:
    codes = {case["expected"].get("code") for case in fixtures["cases"]}
    for code in pc.ALL_CODES:
        assert code in codes, f"共享夹具缺少拒绝码 {code}"


def test_allowed_fields_match_fixture_base(fixtures: dict) -> None:
    """base 记录的键集合必须正好等于白名单 —— 两端字段漂移会在这里暴露。"""
    assert tuple(sorted(fixtures["base"].keys())) == tuple(sorted(pc.FIELDS))


@pytest.mark.parametrize("case", json.loads(FIXTURES.read_text(encoding="utf-8"))["cases"],
                         ids=lambda case: case["name"])
def test_shared_fixture_cases(case: dict, fixtures: dict) -> None:
    record = _materialize(case, fixtures["base"])
    verdict = pc.validate(
        record,
        turn_id=fixtures["expect"]["turnId"],
        revision=fixtures["expect"]["revision"],
    )
    assert verdict.ok is case["expected"]["ok"]
    if not case["expected"]["ok"]:
        assert verdict.code == case["expected"]["code"]


def test_valid_context_projects_only_whitelisted_fields(fixtures: dict) -> None:
    verdict = pc.validate(
        fixtures["base"],
        turn_id=fixtures["expect"]["turnId"],
        revision=fixtures["expect"]["revision"],
    )
    assert verdict.ok and verdict.record is not None
    assert tuple(sorted(verdict.record.keys())) == tuple(sorted(pc.FIELDS))
    assert verdict.record["kind"] == pc.KIND


def test_rejections_have_distinct_messages() -> None:
    messages = {code: pc.Verdict(ok=False, code=code).message() for code in pc.ALL_CODES}
    assert len(set(messages.values())) == len(pc.ALL_CODES)
    for code, text in messages.items():
        assert len(text) > 8, code


def test_builder_output_passes_its_own_validator() -> None:
    record = pc.build_context(
        turn_id="turn-builder",
        executor_id="exec-1",
        task_revision=2,
        stage="before_model",
        prompt="统计订单金额",
        planning_goal=None,
        confirmations=[{"questionId": "q1", "answerLabel": "按金额"}],
        target_identity="url:https://example.invalid/book.xlsx",
        target_evidence_key="k",
        authorized_sheet_names=["订单表"],
        source_fingerprint="fp",
        source_fingerprint_sheets=["订单表"],
        result_refs=[],
        completed_tools=[],
        failed_tool=None,
        saved_at="2026-10-19T22:00:00+08:00",
    )
    verdict = pc.validate(record, turn_id="turn-builder", revision=2)
    assert verdict.ok, verdict.detail


def test_builder_drops_invalid_and_sensitive_inputs() -> None:
    """构造器是唯一入口：非法引用与值矩阵都进不来。"""
    record = pc.build_context(
        turn_id="turn-drop",
        executor_id="exec-1",
        task_revision=1,
        stage="after_tool",
        prompt="p",
        planning_goal=None,
        confirmations=[
            {"questionId": "q1", "answerLabel": "a"},
            {"questionId": "  ", "answerLabel": "should drop"},
        ],
        target_identity="url:x",
        target_evidence_key="k",
        authorized_sheet_names=["S1"],
        source_fingerprint=None,
        source_fingerprint_sheets=[],
        result_refs=[
            {"resultRef": "res_not-a-ref"},
            {"resultRef": "res_" + "a" * 32, "sourceFingerprint": None},
        ],
        completed_tools=["query_table", ""],
        failed_tool=None,
        saved_at="2026-10-19T22:00:00+08:00",
    )
    assert len(record["confirmations"]) == 1
    assert [item["resultRef"] for item in record["resultRefs"]] == ["res_" + "a" * 32]
    assert [item["tool"] for item in record["completedTools"]] == ["query_table"]
    assert pc.find_forbidden_key(record) is None
    assert pc.validate(record, turn_id="turn-drop", revision=1).ok
