"""WPS **公式**计划的写入前恢复登记（四表、三条 ``writeFormulas``）。

存在理由（2026-11-01 真机）：用户真实场景是

- 来源表 ``来源!A1:B5``（只读）；
- 三个独立动作：``结果1!B2 =来源!B2*1``、``结果2!B2 =来源!B4*2``、``结果3!B2 =来源!B5*3``；
- 目标原先为空。

计划生成正常，写入前登记被服务端拒掉，界面给的是**一句合并过的**话：

    无法持久登记本次写入现场，已取消写入；原表未被修改。续跑来源越界、形状不符或不支持该动作组合。

那句话来自 ``_validate_wps_recovery_descriptor`` 里一个 ``valid`` 布尔：来源越界、
矩形形状、动作组合三件事被折叠成一个 ``all(...)``，任一条不成立都报同一句。
真正命中的是**动作组合**那条 —— 端上 ``wpsResumeContext`` 早已把
``writeFormulas`` 算作可带来源证据的动作（``isResumableValueAction``），
服务端却还写着 ``all(action.type == "writeValues")``。于是「公式计划一律登记不了」，
而且从错误文案上看不出是哪一条。

本文件守三件事：

1. 四表三公式的同形状载荷**必须能登记**（回归的正面锚点）；
2. 每一条拒绝原因**各自有独立的错误码与文案**（不许再合并成一句）；
3. 两端"哪些动作可以带来源证据"是**同一份判据**（``isResumableValueAction``）。
"""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import server.app.main as main_module
import server.app.llm.config as llm_config
from server.app.main import app
from server.app.turn_state import turn_registry

REPO_ROOT = Path(__file__).resolve().parents[2]

SOURCE_SHEETS = ["来源", "结果1", "结果2", "结果3"]


@pytest.fixture(autouse=True)
def isolate(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    turn_registry.reset()
    main_module._folder_operations.clear()
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    monkeypatch.setattr(llm_config, "_config_path", lambda: tmp_path / "config.env")


def _formula_action(index: int, sheet: str) -> dict:
    return {
        "actionIndex": index,
        "type": "writeFormulas",
        "sheet": sheet,
        "range": "B2:B2",
        "rowCount": 1,
        "columnCount": 1,
        "cellCount": 1,
        # 公式动作的写前/写后指纹都按**公式文本**算（不是计算值），形状与值动作相同。
        "previewFingerprint": f"v2:0:0000000{index}",
        "expectedFingerprint": f"v2:7:aaaaaaaa",
    }


def _four_sheet_formula_payload(operation_id: str = "wps-formula-resume-0001") -> dict:
    """与真机同形状：来源 A1:B5 只读，三个独立公式动作写三个空表。"""
    return {
        "operationId": operation_id,
        "taskId": "task-formula",
        "confirmed": True,
        "workbookIdentity": "wps-path:c:/work/formula.xlsx",
        "authorizedSheetNames": list(SOURCE_SHEETS),
        "descriptor": {
            "resumeContext": {
                "version": 1,
                "hostVersion": "12.0",
                "sources": [
                    {
                        "sheetName": "来源",
                        "range": "A1:B5",
                        "rowCount": 5,
                        "columnCount": 2,
                        "previewFingerprint": "v2:42:12345678",
                        "valueFingerprint": "v2:40:abcdef12",
                    }
                ],
            },
            "actions": [_formula_action(0, "结果1"), _formula_action(1, "结果2"), _formula_action(2, "结果3")],
        },
    }


@pytest.fixture(scope="module")
def _shared_client() -> TestClient:
    return TestClient(app)


def test_four_sheet_formula_plan_registers() -> None:
    """真机同形状：来源证据 + 三条 writeFormulas 必须登记成功。

    这条是**正面锚点**：把它拆掉之后，"公式计划登记不了"就能一路静默回归。
    """
    payload = _four_sheet_formula_payload()
    response = TestClient(app).post("/api/wps/executions/register", json=payload)
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["ownerToken"], "登记成功必须下发写入权"
    assert [action["type"] for action in body["descriptor"]["actions"]] == ["writeFormulas"] * 3
    assert body["descriptor"]["resumeContext"] == payload["descriptor"]["resumeContext"]


def test_four_sheet_formula_plan_persists_no_cell_content() -> None:
    """登记成功也不能把公式文本或来源值带进账本（来源证据只是定位 + 指纹）。"""
    payload = _four_sheet_formula_payload("wps-formula-resume-0002")
    client = TestClient(app)
    assert client.post("/api/wps/executions/register", json=payload).status_code == 200
    fetched = client.get("/api/wps/executions/wps-formula-resume-0002").json()
    serialized = json.dumps(fetched, ensure_ascii=False)
    assert "=来源!" not in serialized
    assert "valueFingerprint" in serialized
    # 字段名整体比对：`"preview"` 是 `"previewFingerprint"` 的子串，不能用朴素包含。
    for forbidden in ("values", "formulas", "preview", "cells"):
        assert f'"{forbidden}"' not in serialized


def test_action_type_is_the_only_input_that_decides_this_branch() -> None:
    """同一份载荷**只把动作类型换成 `writeValues`** 就能登记 —— 唯一判别点。

    这是"命中哪一条分支"的证据，不是猜：来源表名、区域矩形、行列数、授权范围、
    单元格总数、两个指纹、`resumeContext.version` 全部保持不动，只有 `type` 变了。
    换掉就能过 ⇒ 命中的是**动作组合**那条，不是"来源越界"也不是"形状不符"。
    """
    payload = _four_sheet_formula_payload("wps-formula-control-values")
    for action in payload["descriptor"]["actions"]:
        action["type"] = "writeValues"
    response = TestClient(app).post("/api/wps/executions/register", json=payload)
    assert response.status_code == 200, response.text


@pytest.mark.parametrize(
    ("issue", "status", "code"),
    [
        # 来源表不在本次授权里（授权类沿用 403，与目标表越界同一口径）
        ("unauthorized", 403, "WPS_RESUME_SOURCE_OUT_OF_SCOPE"),
        # 来源表同时是写入目标（读写重叠）
        ("overlap", 422, "WPS_RESUME_SOURCE_TARGET_OVERLAP"),
        # 声明的行列数与区域不一致
        ("shape", 422, "WPS_RESUME_SOURCE_SHAPE_MISMATCH"),
        # 来源单元格总数超限
        ("limit", 422, "WPS_RESUME_SOURCE_LIMIT"),
        # 格式动作带来源证据：语义上不存在这种组合（端上也不会产出）
        ("format", 422, "WPS_RESUME_ACTION_UNSUPPORTED"),
    ],
)
def test_resume_refusals_are_reported_one_branch_at_a_time(issue: str, status: int, code: str) -> None:
    """五种拒绝**各自**有错误码。

    以前它们共用一个 ``valid`` 布尔和服务端同一句话，用户与排查都只能看到
    「续跑来源越界、形状不符或不支持该动作组合」。
    """
    payload = _four_sheet_formula_payload(f"wps-formula-refuse-{issue}")
    context = payload["descriptor"]["resumeContext"]
    source = context["sources"][0]
    if issue == "unauthorized":
        source["sheetName"] = "没授权"
    elif issue == "overlap":
        source["sheetName"] = "结果1"
    elif issue == "shape":
        source["rowCount"] = 4
    elif issue == "limit":
        source["range"] = "A1:B1048576"
        source["rowCount"] = 1048576
    elif issue == "format":
        payload["descriptor"]["actions"][1]["type"] = "setFont"
        payload["descriptor"]["actions"][1]["formatFields"] = ["bold"]
    response = TestClient(app).post("/api/wps/executions/register", json=payload)
    assert response.status_code == status, response.text
    assert response.json()["detail"]["code"] == code


def test_unparsable_source_range_is_refused_by_the_guard_itself() -> None:
    """区域解析不出来时，函数里的兜底闸仍然拒绝（且不被合并成别的原因）。

    HTTP 层走不到这里：`WpsResumeSource.range` 的 pattern 在模型层就把
    ``A1:B`` 这类形状挡下了。所以这条守卫要**直接**测 —— 留着它是因为
    pattern 一旦放宽，`range_boundaries` 就是唯一还能 fail-closed 的地方。
    """
    from fastapi import HTTPException

    from server.app.models import WpsExecutionRegisterRequest, WpsResumeSource

    payload = _four_sheet_formula_payload("wps-formula-bad-range")
    request = WpsExecutionRegisterRequest.model_validate(payload)
    assert request.descriptor.resumeContext is not None
    request.descriptor.resumeContext.sources[0] = WpsResumeSource.model_construct(
        sheetName="来源",
        range="A1:B",
        rowCount=5,
        columnCount=2,
        previewFingerprint="v2:42:12345678",
        valueFingerprint="v2:40:abcdef12",
    )
    with pytest.raises(HTTPException) as raised:
        main_module._validate_wps_recovery_descriptor(request)
    detail = raised.value.detail
    assert isinstance(detail, dict)
    assert detail["code"] == "WPS_RESUME_SOURCE_RANGE_INVALID"


def test_literal_plan_still_refuses_source_evidence_and_non_value_actions() -> None:
    """version 2（无来源依赖的字面值计划）的两条边界一个字都不放宽。

    「公式可以带来源证据」不等于「什么都能带」：v2 是"这份计划不读别的表"的声明，
    它既不接受来源列表（模型层），也只允许值写入（模型层 + 这里不重复）。
    """
    payload = _four_sheet_formula_payload("wps-formula-literal")
    payload["descriptor"]["resumeContext"] = {"version": 2, "hostVersion": "12.0", "sources": []}
    response = TestClient(app).post("/api/wps/executions/register", json=payload)
    assert response.status_code == 422, response.text

    payload = _four_sheet_formula_payload("wps-formula-literal-2")
    payload["descriptor"]["resumeContext"] = {"version": 2, "hostVersion": "12.0", "sources": []}
    payload["descriptor"]["actions"] = [_formula_action(0, "结果1")]
    payload["descriptor"]["actions"][0]["type"] = "writeValues"
    assert TestClient(app).post("/api/wps/executions/register", json=payload).status_code == 200


def test_resume_source_action_types_match_the_ts_criterion() -> None:
    """两端"哪些动作可以带来源证据"必须是**同一份判据**。

    TS 侧唯一出处是 `apps/wps-addin/src/wpsPlanExecutor.ts` 的 `isResumableValueAction`；
    它一旦扩容而服务端没跟上，症状就是本轮真机那个：端上产出 `resumeContext`，
    服务端一律 422 —— 而用户看到的只是一句合并过的"不支持该动作组合"。
    """
    source = (REPO_ROOT / "apps" / "wps-addin" / "src" / "wpsPlanExecutor.ts").read_text(encoding="utf-8")
    match = re.search(
        r"function isResumableValueAction\(type: string\): type is (?P<types>[^\{]+)\{",
        source,
    )
    assert match is not None, "wpsPlanExecutor.ts 里找不到 isResumableValueAction"
    from_ts = set(re.findall(r'"([A-Za-z]+)"', match.group("types")))
    assert from_ts == set(main_module._WPS_RESUME_SOURCE_ACTION_TYPES)
