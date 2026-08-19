"""rule_generator 测试：/function 公式生成请求、提示词注入与安全放行。"""

from __future__ import annotations

from server.app.rule_generator import (
    DictionarySheet,
    ExtraSheet,
    GenerateFormulaRequest,
    _build_formula_generation_prompt,
)
from server.app.safety import dangerous_formula


def test_generate_formula_request_extra_sheets_defaults_empty() -> None:
    """extraSheets 为可选字段，缺省时保持向后兼容。"""
    request = GenerateFormulaRequest.model_validate(
        {"description": "求和", "activeCell": "A1"}
    )
    assert request.extraSheets == []


def test_formula_prompt_injects_external_sheets() -> None:
    """勾选的外部工作簿工作表应注入提示词，含文件引用语法与列头映射。"""
    request = GenerateFormulaRequest(
        description="把B表的单价按编号匹配过来",
        activeCell="E2",
        headers=["编号", "名称"],
        columns=["A", "B"],
        sampleRows=[["1", "苹果"]],
        extraSheets=[
            ExtraSheet(
                sourceFile="B.xlsx",
                sourcePath="子目录/B.xlsx",
                sheetName="Sheet2",
                headers=["编号", "单价"],
                columns=["A", "B"],
                sampleRows=[["1", "3.5"]],
            )
        ],
    )
    messages = _build_formula_generation_prompt(request)
    user_content = messages[1]["content"]
    assert "外部工作簿「B.xlsx」工作表「Sheet2」" in user_content
    assert "[B.xlsx]Sheet2!" in user_content
    assert "列头→列字母映射：A列=编号、B列=单价" in user_content
    assert '["1", "3.5"]' in user_content
    system_content = messages[0]["content"]
    assert "[文件名]表名!区域" in system_content
    assert "禁止编造未提供的文件或工作表" in system_content


def test_formula_prompt_keeps_literal_markers() -> None:
    """提示词必须要求把字典里的 / 等符号当字面值原样输出，不得当空值。"""
    request = GenerateFormulaRequest(
        description="无问题填入/",
        activeCell="E2",
        dictionary=DictionarySheet(
            name="数据字典", rows=[["问题", "结果"], ["无问题", "/"]]
        ),
    )
    messages = _build_formula_generation_prompt(request)
    system_content = messages[0]["content"]
    assert "「/」「－」「—」等符号是**字面值**" in system_content
    assert "绝不能把这些符号当作空值" in system_content
    assert "要填入 / 就写 \"/\"" in system_content
    user_content = messages[1]["content"]
    assert '["无问题", "/"]' in user_content


def test_generate_formula_whitelist_passes_selected_external_refs() -> None:
    """勾选文件名集合作为 allowed_external 时，对应外部引用放行。"""
    allowed = {"B.xlsx"}
    assert (
        dangerous_formula(
            "='[B.xlsx]Sheet2'!A1", allowed_external=allowed
        )
        is None
    )
    assert (
        dangerous_formula(
            "='[C.xlsx]Sheet1'!A1", allowed_external=allowed
        )
        == "EXTERNAL_REF"
    )
