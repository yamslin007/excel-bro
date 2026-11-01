"""恢复状态里的**验收条件**：公式计划续跑后还认得出"哪几格该是什么公式"。

为什么单独钉这一环（2026-10-30）：

Excel 端续跑完成后要用 `verifyPlan` 判定 `verified / executed_unverified / failed`，
而 `verificationGaps` **只认计划里声明的验收条件**（`plan.acceptanceCriteria`）。
服务端 `AnalysisPlan.model_validate` 的 `add_deterministic_acceptance_criteria`
会为每个 `writeFormulas` 自动补 `formulasEqual`。

这不是"补过就够了"的问题，而是**这条判据要真的走到客户端**。真实路径是
`App.tsx` 的恢复续跑：`getOfficeExecutionStatus()` 取回 `state.plan`
（服务端 `office_executions.plan_json` → `AnalysisPlan.model_validate` →
`OfficeExecutionStateResponse.plan`）→ `executePlan(state.plan, {startActionIndex})`
→ `verifyPlan`。也就是说客户端手里那份计划**来自服务端的读响应**，不是它自己的
副本；读响应少了 `formulasEqual`，`verifyPlan` 只能给出 `executed_unverified`，
而 `ResultReview.adoptCompleted` 在 `verified !== true` 时直接抛错 —— 恢复收尾就
被挡在审阅卡之前，用户点了"继续执行"却拿不到确认/放弃。

所以这里钉的是**恢复读路径**：一份登记时**没写** `acceptanceCriteria` 的公式计划，
从 `/api/office/executions/{id}` 读回来时，必须已经带上对应区域的 `formulasEqual`，
且 `expected` 是**公式文本**（不是算出来的值）。

反向变异（2026-10-30，证据在 `outputs/mutation-2026-10-30/`）：
- 删掉 `WriteFormulasAction.formulaR1C1` → 两条用例全红（`formulaR1C1` 出现在
  失败原文里），说明"客户端声明的字段服务端没声明"是会被抓到的；
- 拆掉 `writeFormulas` 的 `formulasEqual` 推导 → 只有验收条件那条红，R1C1 那条
  仍绿（判别点唯一）；
- **把登记时的 `acceptanceCriteria` 从持久化里剔掉 → 仍然全绿**。这条存活是结论
  不是缺口：读路径会重新 `model_validate` 再推导一遍，推导幂等，客户端看到的
  响应与"存下来"完全一致。真正有判别力的是"推导在不在"，不是"存没存"。
"""

from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

import server.app.main as main_module
from server.app.main import app

OPERATION_ID = "office-resume-acceptance-0001"
IDENTITY = "url:c:/work/resume-acceptance.xlsx"


@pytest.fixture(autouse=True)
def isolate(monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    main_module._folder_operations.clear()


def _formula_plan() -> dict[str, object]:
    """刻意**不**带 `acceptanceCriteria`：验收条件必须由模型层补齐。

    `formulaR1C1` 显式写 `None`（而不是省略）：两处声明都必须接受 **null**，
    这条 200 就是「可空」的证据本身 —— 少了它，"A1 公式能登记"只在"字段缺省"
    这一种形状下被验证过，`null` 形状靠推断。
    """
    return {
        "id": "office-resume-acceptance-plan",
        "title": "写公式",
        "summary": "两格公式",
        "assumptions": [],
        "warnings": [],
        "actions": [
            {
                "type": "writeFormulas",
                "sheet": "结果",
                "range": "B2:C2",
                "formulas": [["=数据!A1*2", "=数据!A1*3"]],
                "formulaR1C1": None,
            }
        ],
    }


def test_resumed_state_plan_carries_formulas_equal_for_formula_actions() -> None:
    client = TestClient(app)
    registered = client.post(
        "/api/office/executions/register",
        json={
            "operationId": OPERATION_ID,
            "plan": _formula_plan(),
            "authorizedSheetNames": ["结果", "数据"],
            "allowedExternalFiles": [],
            "workbookIdentity": IDENTITY,
            "confirmed": True,
        },
    )
    assert registered.status_code == 200, registered.text

    response = client.get(f"/api/office/executions/{OPERATION_ID}")
    assert response.status_code == 200, response.text
    criteria = response.json()["plan"]["acceptanceCriteria"]

    formula_checks = [
        criterion for criterion in criteria if criterion["type"] == "formulasEqual"
    ]
    assert formula_checks == [
        {
            "type": "formulasEqual",
            "sheet": "结果",
            "range": "B2:C2",
            # 公式身份 = 公式文本：这里存的就是文本，不是 `=数据!A1*2` 算出来的数。
            "expected": [["=数据!A1*2", "=数据!A1*3"]],
        }
    ]
    # 区域存在性判据同时保留（`worksheetExists`），顺序上先声明表、再声明内容。
    assert [criterion["type"] for criterion in criteria] == [
        "worksheetExists",
        "formulasEqual",
    ]


def test_r1c1_formula_plan_is_rejected_at_registration() -> None:
    """R1C1 铺满公式的**拒绝闸门**必须在登记时真的生效。

    这条用例的由来是一次真实缺陷：`WriteFormulasAction` 原本**没声明**
    `formulaR1C1`（客户端 `contracts.ts` 有），pydantic 默认 `extra="ignore"`
    把字段静默丢掉了，于是 `_validate_office_recovery_plan` 读属性时
    `AttributeError` —— 任何公式计划的登记直接 500，"拒绝 R1C1"这道闸门
    从来没真正判过。用例钉的是闸门的**判别点**：同样的计划只差一个
    `formulaR1C1`，一个必须 200、一个必须 422。
    """
    client = TestClient(app)
    plan = _formula_plan()
    plan["actions"] = [
        {
            "type": "writeFormulas",
            "sheet": "结果",
            "range": "B2:C2",
            "formulas": [["=数据!A1*2", "=数据!A1*3"]],
            "formulaR1C1": "=RC[-1]*2",
        }
    ]
    response = client.post(
        "/api/office/executions/register",
        json={
            "operationId": "office-resume-acceptance-r1c1",
            "plan": plan,
            "authorizedSheetNames": ["结果", "数据"],
            "allowedExternalFiles": [],
            "workbookIdentity": IDENTITY,
            "confirmed": True,
        },
    )
    assert response.status_code == 422, response.text
    detail = response.json()["detail"]
    assert detail["code"] == "OFFICE_RECOVERY_UNSUPPORTED_ACTION"
    assert "R1C1" in detail["message"]


def test_empty_string_r1c1_is_treated_as_absent_on_every_end() -> None:
    """`formulaR1C1: ""` 必须与 `null` / 省略**同义**：不是 R1C1 计划，照常登记。

    三处消费必须同一口径，缺一处就会让同一个空字段在两端得到不同命运：
    - 服务端 `if action.formulaR1C1`（真值）⇒ 空串不拒；
    - Excel `if (action.formulaR1C1)` ⇒ 走 A1 公式分支；
    - WPS 以前写的是 `!== undefined` ⇒ 空串被当成 R1C1 整份拒绝（2026-10-30 改正）。

    这条钉的是**运行时约束**这一面：字段名一致、类型一致之外，"什么算有值"
    也必须一致。判别点唯一 —— 只差 `""` 与非空 R1C1 一个字符串。
    """
    client = TestClient(app)
    plan = _formula_plan()
    plan["actions"] = [
        {
            "type": "writeFormulas",
            "sheet": "结果",
            "range": "B2:C2",
            "formulas": [["=数据!A1*2", "=数据!A1*3"]],
            "formulaR1C1": "",
        }
    ]
    response = client.post(
        "/api/office/executions/register",
        json={
            "operationId": "office-resume-acceptance-empty-r1c1",
            "plan": plan,
            "authorizedSheetNames": ["结果", "数据"],
            "allowedExternalFiles": [],
            "workbookIdentity": IDENTITY,
            "confirmed": True,
        },
    )
    assert response.status_code == 200, response.text


def test_source_dependency_fingerprint_survives_registration_and_read() -> None:
    """公式来源证据必须**真的落进读响应**，而不是被 pydantic 静默丢掉。

    为什么单独钉（2026-10-30）：`sourceDependencyFingerprint` 由宿主在登记前算好、
    随计划一起上来。pydantic 默认 `extra="ignore"` —— 服务端 `AnalysisPlan` 少声明一个
    字段，客户端带上来的值就会被丢掉，续跑读回来的计划里永远没有证据，于是每一次
    公式续跑都会被判成"没有登记来源证据"。这与 `formulaR1C1` 那次 500 是同一个陷阱
    的两个面：那边是 AttributeError，这边是**静默为空**。

    判别点唯一：同一份计划，只差这个字段 —— 一个读回来有值，一个没有。
    """
    client = TestClient(app)
    plan = _formula_plan()
    plan["sourceDependencySheets"] = ["数据"]
    plan["sourceDependencyFingerprint"] = "fnv1a32x2:deadbeef"
    registered = client.post(
        "/api/office/executions/register",
        json={
            "operationId": OPERATION_ID,
            "plan": plan,
            "authorizedSheetNames": ["结果", "数据"],
            "allowedExternalFiles": [],
            "workbookIdentity": IDENTITY,
            "confirmed": True,
        },
    )
    assert registered.status_code == 200, registered.text

    stored = client.get(f"/api/office/executions/{OPERATION_ID}").json()["plan"]
    assert stored["sourceDependencyFingerprint"] == "fnv1a32x2:deadbeef"
    assert stored["sourceDependencySheets"] == ["数据"]


def test_fingerprint_without_dependency_sheets_is_rejected() -> None:
    """有指纹没口径 = 一串哈希，不能登记（也没有表可读回来比对）。"""
    client = TestClient(app)
    plan = _formula_plan()
    plan["sourceDependencyFingerprint"] = "fnv1a32x2:deadbeef"
    response = client.post(
        "/api/office/executions/register",
        json={
            "operationId": "office-resume-acceptance-orphan-fingerprint",
            "plan": plan,
            "authorizedSheetNames": ["结果", "数据"],
            "allowedExternalFiles": [],
            "workbookIdentity": IDENTITY,
            "confirmed": True,
        },
    )
    assert response.status_code == 422, response.text
    detail = response.json()["detail"]
    assert detail["code"] == "OFFICE_RECOVERY_SOURCE_EVIDENCE_INCOMPLETE"
    assert "来源表" in detail["message"]
