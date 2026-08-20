"""formula_tools 测试：数据模型、编译器、函数注册表、Prompt 与集成回退。"""

from __future__ import annotations

import asyncio
import json
import logging

import pytest
from pydantic import ValidationError

from server.app.formula_tools.compiler import compile_formula
from server.app.formula_tools.function_registry import (
    EXCEL_FUNCTIONS,
    get_functions_prompt,
)
from server.app.formula_tools.model_prompt import SYSTEM_PROMPT, build_user_message
from server.app.formula_tools.schema import (
    BooleanLiteral,
    CellRef,
    FunctionCall,
    NumberLiteral,
    RangeRef,
    TextLiteral,
)


# ============================================================================
# 编译器
# ============================================================================


def test_compile_simple_sum() -> None:
    call = FunctionCall(
        function="SUM",
        args=[
            RangeRef(type="range", workbook=None, sheet="数据", col="A", row1=1, row2=10)
        ],
    )
    assert compile_formula(call) == "=SUM(数据!$A$1:$A$10)"


def test_compile_xlookup() -> None:
    call = FunctionCall(
        function="XLOOKUP",
        args=[
            CellRef(type="cell", address="D2"),
            RangeRef(type="range", workbook=None, sheet="字典", col="A", row1=2, row2=100),
            RangeRef(type="range", workbook=None, sheet="字典", col="B", row1=2, row2=100),
            TextLiteral(type="text", value=""),
        ],
    )
    result = compile_formula(call)
    assert result == '=XLOOKUP(D2,字典!$A$2:$A$100,字典!$B$2:$B$100,"")'


def test_compile_external_workbook() -> None:
    call = FunctionCall(
        function="XLOOKUP",
        args=[
            CellRef(type="cell", address="D2"),
            RangeRef(type="range", workbook="tt.xlsx", sheet="tt", col="A", row1=2, row2=999),
            RangeRef(type="range", workbook="tt.xlsx", sheet="tt", col="B", row1=2, row2=999),
            TextLiteral(type="text", value=""),
        ],
    )
    result = compile_formula(call)
    assert result == '=XLOOKUP(D2,[tt.xlsx]tt!$A$2:$A$999,[tt.xlsx]tt!$B$2:$B$999,"")'


def test_compile_whole_column_range() -> None:
    call = FunctionCall(
        function="SUM",
        args=[RangeRef(type="range", workbook=None, sheet="数据", col="a", row1=1, row2=None)],
    )
    # 列字母统一大写；row2=None 编译成整列引用
    assert compile_formula(call) == "=SUM(数据!$A:$A)"


def test_compile_comparison() -> None:
    call = FunctionCall(
        function="EQ",
        args=[CellRef(type="cell", address="A2"), TextLiteral(type="text", value="")],
    )
    assert compile_formula(call) == '=(A2="")'


def test_compile_arithmetic() -> None:
    call = FunctionCall(
        function="MULTIPLY",
        args=[CellRef(type="cell", address="A2"), NumberLiteral(type="number", value=1.1)],
    )
    assert compile_formula(call) == "=(A2*1.1)"


def test_compile_nested_if() -> None:
    call = FunctionCall(
        function="IF",
        args=[
            FunctionCall(
                function="GT",
                args=[
                    CellRef(type="cell", address="A2"),
                    NumberLiteral(type="number", value=0),
                ],
            ),
            TextLiteral(type="text", value="正数"),
            TextLiteral(type="text", value="非正数"),
        ],
    )
    # Pydantic 会把 0 存成 0.0，但编译输出必须保持 >0（整数不带 .0）
    assert compile_formula(call) == '=IF((A2>0),"正数","非正数")'


def test_compile_number_integral_without_decimal() -> None:
    call = FunctionCall(
        function="GT",
        args=[NumberLiteral(type="number", value=0), NumberLiteral(type="number", value=0.0)],
    )
    assert compile_formula(call) == "=(0>0)"


def test_compile_boolean() -> None:
    call = FunctionCall(
        function="VLOOKUP",
        args=[
            CellRef(type="cell", address="A2"),
            RangeRef(type="range", workbook=None, sheet="字典", col="A", row1=2, row2=100),
            NumberLiteral(type="number", value=2),
            BooleanLiteral(type="boolean", value=False),
        ],
    )
    assert compile_formula(call) == '=VLOOKUP(A2,字典!$A$2:$A$100,2,FALSE)'


def test_compile_special_characters() -> None:
    call = FunctionCall(
        function="IF",
        args=[
            FunctionCall(
                function="EQ",
                args=[CellRef(type="cell", address="A2"), TextLiteral(type="text", value="/")],
            ),
            TextLiteral(type="text", value="斜杠"),
            TextLiteral(type="text", value="其他"),
        ],
    )
    assert compile_formula(call) == '=IF((A2="/"),"斜杠","其他")'


def test_compile_text_with_quotes() -> None:
    call = FunctionCall(
        function="IF",
        args=[
            FunctionCall(
                function="EQ",
                args=[
                    CellRef(type="cell", address="A2"),
                    TextLiteral(type="text", value='他说"你好"'),
                ],
            ),
            TextLiteral(type="text", value="匹配"),
            TextLiteral(type="text", value="不匹配"),
        ],
    )
    assert compile_formula(call) == '=IF((A2="他说""你好"""),"匹配","不匹配")'


def test_compile_unknown_function_passthrough() -> None:
    """任意 Excel 函数名都能编译（注册表只约束 Prompt，不限制编译器）。"""
    call = FunctionCall(
        function="TEXTJOIN",
        args=[
            TextLiteral(type="text", value="-"),
            BooleanLiteral(type="boolean", value=True),
            CellRef(type="cell", address="A2"),
        ],
    )
    assert compile_formula(call) == '=TEXTJOIN("-",TRUE,A2)'


def test_compile_sheet_name_quoting() -> None:
    """含空格/特殊字符/单引号的工作表名必须加单引号，内部单引号翻倍。"""
    call = FunctionCall(
        function="SUM",
        args=[RangeRef(type="range", workbook=None, sheet="数据 明细", col="A", row1=1, row2=10)],
    )
    assert compile_formula(call) == "=SUM('数据 明细'!$A$1:$A$10)"

    call = FunctionCall(
        function="SUM",
        args=[RangeRef(type="range", workbook=None, sheet="It's", col="A", row1=1, row2=10)],
    )
    assert compile_formula(call) == "=SUM('It''s'!$A$1:$A$10)"

    # 看起来像单元格地址的表名也要加引号
    call = FunctionCall(
        function="SUM",
        args=[RangeRef(type="range", workbook=None, sheet="A1", col="A", row1=1, row2=10)],
    )
    assert compile_formula(call) == "=SUM('A1'!$A$1:$A$10)"


def test_compile_operator_arity_error() -> None:
    call = FunctionCall(
        function="EQ",
        args=[CellRef(type="cell", address="A2")],
    )
    with pytest.raises(ValueError, match="EQ 需要 2 个参数"):
        compile_formula(call)


def test_compile_sumproduct_priority_match() -> None:
    """优先级关键词匹配公式（完整示例）必须与预期逐字符一致。"""
    call = FunctionCall(
        function="IF",
        args=[
            FunctionCall(
                function="EQ",
                args=[
                    CellRef(type="cell", address="D2"),
                    TextLiteral(type="text", value=""),
                ],
            ),
            TextLiteral(type="text", value=""),
            FunctionCall(
                function="IF",
                args=[
                    FunctionCall(
                        function="GT",
                        args=[
                            FunctionCall(
                                function="SUMPRODUCT",
                                args=[
                                    FunctionCall(
                                        function="MULTIPLY",
                                        args=[
                                            FunctionCall(
                                                function="ISNUMBER",
                                                args=[
                                                    FunctionCall(
                                                        function="SEARCH",
                                                        args=[
                                                            RangeRef(
                                                                type="range",
                                                                workbook="tt.xlsx",
                                                                sheet="tt",
                                                                col="A",
                                                                row1=2,
                                                                row2=999,
                                                            ),
                                                            CellRef(type="cell", address="D2"),
                                                        ],
                                                    )
                                                ],
                                            ),
                                            FunctionCall(
                                                function="EQ",
                                                args=[
                                                    RangeRef(
                                                        type="range",
                                                        workbook="tt.xlsx",
                                                        sheet="tt",
                                                        col="B",
                                                        row1=2,
                                                        row2=999,
                                                    ),
                                                    TextLiteral(type="text", value="p0"),
                                                ],
                                            ),
                                        ],
                                    )
                                ],
                            ),
                            NumberLiteral(type="number", value=0),
                        ],
                    ),
                    TextLiteral(type="text", value="p0"),
                    TextLiteral(type="text", value=""),
                ],
            ),
        ],
    )

    result = compile_formula(call)
    expected = (
        '=IF((D2=""),"",IF((SUMPRODUCT((ISNUMBER(SEARCH('
        "[tt.xlsx]tt!$A$2:$A$999,D2))*([tt.xlsx]tt!$B$2:$B$999=\"p0\")))>0),\"p0\",\"\"))"
    )
    assert result == expected


# ============================================================================
# 数据模型
# ============================================================================


def test_schema_parses_nested_json() -> None:
    raw = {
        "formula": {
            "function": "IF",
            "args": [
                {
                    "function": "EQ",
                    "args": [
                        {"type": "cell", "address": "D2"},
                        {"type": "text", "value": ""},
                    ],
                },
                {"type": "text", "value": ""},
                {
                    "function": "XLOOKUP",
                    "args": [
                        {"type": "cell", "address": "D2"},
                        {
                            "type": "range",
                            "workbook": "tt.xlsx",
                            "sheet": "tt",
                            "col": "A",
                            "row1": 2,
                            "row2": 999,
                        },
                        {
                            "type": "range",
                            "workbook": "tt.xlsx",
                            "sheet": "tt",
                            "col": "B",
                            "row1": 2,
                            "row2": 999,
                        },
                        {"type": "text", "value": ""},
                    ],
                },
            ],
        },
        "explanation": "测试",
    }
    from server.app.formula_tools.schema import FormulaToolCallResponse

    parsed = FormulaToolCallResponse.model_validate(raw)
    assert parsed.formula.function == "IF"
    assert parsed.formula.args[2].function == "XLOOKUP"
    assert parsed.formula.args[2].args[1].workbook == "tt.xlsx"


def test_schema_rejects_unknown_node_type() -> None:
    raw = {
        "formula": {
            "function": "IF",
            "args": [{"type": "datetime", "value": "2026-01-01"}],
        },
        "explanation": "测试",
    }
    from server.app.formula_tools.schema import FormulaToolCallResponse

    with pytest.raises(ValidationError):
        FormulaToolCallResponse.model_validate(raw)


# ============================================================================
# 函数注册表
# ============================================================================


def test_registry_covers_plan_functions() -> None:
    names = {
        func["name"]
        for functions in EXCEL_FUNCTIONS.values()
        for func in functions
    }
    required = {
        "XLOOKUP",
        "VLOOKUP",
        "IF",
        "AND",
        "OR",
        "SEARCH",
        "ISNUMBER",
        "SUM",
        "SUMPRODUCT",
        "EQ",
        "GT",
        "MULTIPLY",
    }
    assert required <= names


def test_registry_examples_are_valid_json() -> None:
    """注册表示例必须是可以直接解析的完整 JSON，不能有 [...] 占位符。"""
    for functions in EXCEL_FUNCTIONS.values():
        for func in functions:
            data = json.loads(func["example"])
            assert data["function"] == func["name"]


def test_get_functions_prompt_contains_examples() -> None:
    prompt = get_functions_prompt()
    assert "XLOOKUP" in prompt
    assert "SUMPRODUCT" in prompt
    assert '"function": "XLOOKUP"' in prompt


# ============================================================================
# 模型 Prompt
# ============================================================================


def test_system_prompt_contains_core_rules() -> None:
    assert "你不直接生成公式文本" in SYSTEM_PROMPT
    assert "引用完整区域" in SYSTEM_PROMPT
    assert "字面量原样保留" in SYSTEM_PROMPT
    assert "循环引用" in SYSTEM_PROMPT
    assert "禁止近似/模糊匹配" in SYSTEM_PROMPT


def test_build_user_message_with_dictionary_and_extra_sheets() -> None:
    message = build_user_message(
        active_cell="E2",
        description="根据 D 列编号返回对应级别",
        headers=["编号", "级别"],
        columns=["D", "E"],
        dictionary={
            "name": "数据字典",
            "rows": [["编号", "级别"], ["1", "p0"]],
        },
        extra_sheets=[
            {
                "sourceFile": "tt.xlsx",
                "sheetName": "tt",
                "headers": ["关键词", "级别"],
                "columns": ["A", "B"],
                "rowCount": 20,
            }
        ],
    )
    assert "目标单元格：E2" in message
    assert "列头映射：D列=编号、E列=级别" in message
    assert "字典表「数据字典」" in message
    assert "共 2 行数据" in message
    assert "外部工作簿「tt.xlsx」工作表「tt」" in message
    assert "共 20 行数据" in message
    # 原始数据行不得进入模型消息（意图/生成阶段都只发结构与摘要）
    assert '"1"' not in message
    assert "p0" not in message


def test_build_user_message_without_context() -> None:
    message = build_user_message(
        active_cell="B10",
        description="统计 B 列大于 100 的个数",
        headers=None,
        columns=None,
        dictionary=None,
        extra_sheets=None,
    )
    assert message == "目标单元格：B10\n需求：统计 B 列大于 100 的个数\n\n请生成结构化的公式调用 JSON。"


# ============================================================================
# 集成：工具调用路径
# ============================================================================


class _FakeClient:
    def __init__(self, config, timeout):
        self._config = config
        self._timeout = timeout
        self.captured_kwargs: dict = {}

    async def __aenter__(self):
        return self

    async def __aexit__(self, exc_type, exc, tb):
        return False

    async def chat_completions(self, *, messages, max_tokens=None, temperature=None, **kwargs):
        self.captured_kwargs = {
            "messages": messages,
            "max_tokens": max_tokens,
            "temperature": temperature,
        }
        return {
            "choices": [
                {"message": {"content": self.content}, "finish_reason": "stop"}
            ]
        }


def _run(coro):
    return asyncio.run(coro)


def test_generate_formula_with_tool_call_success(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from server.app import formula_tools_integration

    client = _FakeClient(None, 0)
    client.content = json.dumps(
        {
            "formula": {
                "function": "IF",
                "args": [
                    {
                        "function": "EQ",
                        "args": [
                            {"type": "cell", "address": "D2"},
                            {"type": "text", "value": ""},
                        ],
                    },
                    {"type": "text", "value": ""},
                    {
                        "function": "XLOOKUP",
                        "args": [
                            {"type": "cell", "address": "D2"},
                            {
                                "type": "range",
                                "workbook": "tt.xlsx",
                                "sheet": "tt",
                                "col": "A",
                                "row1": 2,
                                "row2": 999,
                            },
                            {
                                "type": "range",
                                "workbook": "tt.xlsx",
                                "sheet": "tt",
                                "col": "B",
                                "row1": 2,
                                "row2": 999,
                            },
                            {"type": "text", "value": ""},
                        ],
                    },
                ],
            },
            "explanation": "根据 D 列编号查表返回级别",
        }
    )
    monkeypatch.setattr(
        formula_tools_integration, "OpenAICompatibleClient", lambda config, timeout: client
    )
    monkeypatch.setattr(formula_tools_integration, "formula_model_config", lambda: object())

    result = _run(
        formula_tools_integration.generate_formula_with_tool_call(
            active_cell="E2",
            description="根据 D 列编号返回对应级别",
            headers=["编号", "级别"],
            columns=["D", "E"],
            dictionary={"name": "数据字典", "rows": [["编号", "级别"]]},
            extra_sheets=[
                {
                    "sourceFile": "tt.xlsx",
                    "sheetName": "tt",
                    "headers": ["关键词", "级别"],
                    "columns": ["A", "B"],
                    "rowCount": 20,
                }
            ],
            model_id=None,
        )
    )

    assert result["modernFormula"] == (
        '=IF((D2=""),"",XLOOKUP(D2,[tt.xlsx]tt!$A$2:$A$999,'
        '[tt.xlsx]tt!$B$2:$B$999,""))'
    )
    assert result["compatFormula"] == result["modernFormula"]
    assert result["modernExplanation"] == "根据 D 列编号查表返回级别"
    # 兼容版与 modern 相同，但说明必须带上不兼容旧版本的警示
    assert result["compatExplanation"] == (
        "⚠️ 此公式使用了 Excel 365+ 函数，可能不兼容旧版本。"
        "根据 D 列编号查表返回级别"
    )
    # 低温度与输出限额必须随请求发送
    assert client.captured_kwargs["temperature"] == 0.1
    assert client.captured_kwargs["max_tokens"] > 0
    assert client.captured_kwargs["messages"][0]["role"] == "system"


def test_generate_formula_with_tool_call_temperature_unsupported_logs(
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """模型拒绝 temperature 参数时，记录 info 日志后回退（异常继续上抛）。"""
    from server.app import formula_tools_integration

    class _RejectingClient:
        def __init__(self, config, timeout):
            self.config = config

        async def __aenter__(self):
            return self

        async def __aexit__(self, exc_type, exc, tb):
            return False

        async def chat_completions(self, **kwargs):
            raise ValueError("Unsupported parameter: 'temperature'")

    monkeypatch.setattr(
        formula_tools_integration,
        "OpenAICompatibleClient",
        lambda config, timeout: _RejectingClient(config, timeout),
    )
    monkeypatch.setattr(
        formula_tools_integration,
        "formula_model_config",
        lambda: type("Fake", (), {"model": "reasoning-model" })(),
    )

    with caplog.at_level(logging.INFO, logger="server.app.formula_tools_integration"):
        with pytest.raises(ValueError, match="temperature"):
            _run(
                formula_tools_integration.generate_formula_with_tool_call(
                    active_cell="E2",
                    description="求和",
                    headers=None,
                    columns=None,
                    dictionary=None,
                    extra_sheets=None,
                    model_id=None,
                )
            )

    assert "reasoning-model 不支持 temperature 参数" in caplog.text


def test_generate_formula_with_tool_call_generic_failure_logs_warning(
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """非 temperature 的调用失败记录 warning 后回退（异常继续上抛）。"""
    from server.app import formula_tools_integration

    class _FailingClient:
        def __init__(self, config, timeout):
            self.config = config

        async def __aenter__(self):
            return self

        async def __aexit__(self, exc_type, exc, tb):
            return False

        async def chat_completions(self, **kwargs):
            raise TimeoutError("模型响应超时")

    monkeypatch.setattr(
        formula_tools_integration,
        "OpenAICompatibleClient",
        lambda config, timeout: _FailingClient(config, timeout),
    )
    monkeypatch.setattr(
        formula_tools_integration,
        "formula_model_config",
        lambda: type("Fake", (), {"model": "fast-model"})(),
    )

    with caplog.at_level(logging.WARNING, logger="server.app.formula_tools_integration"):
        with pytest.raises(TimeoutError):
            _run(
                formula_tools_integration.generate_formula_with_tool_call(
                    active_cell="E2",
                    description="求和",
                    headers=None,
                    columns=None,
                    dictionary=None,
                    extra_sheets=None,
                    model_id=None,
                )
            )

    assert "fast-model" in caplog.text or "工具调用失败" in caplog.text


def test_generate_formula_with_tool_call_invalid_json_raises(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from server.app import formula_tools_integration

    client = _FakeClient(None, 0)
    client.content = "抱歉，我无法生成 JSON。"
    monkeypatch.setattr(
        formula_tools_integration, "OpenAICompatibleClient", lambda config, timeout: client
    )
    monkeypatch.setattr(formula_tools_integration, "formula_model_config", lambda: object())

    with pytest.raises(ValueError, match="模型返回格式错误"):
        _run(
            formula_tools_integration.generate_formula_with_tool_call(
                active_cell="E2",
                description="求和",
                headers=None,
                columns=None,
                dictionary=None,
                extra_sheets=None,
                model_id=None,
            )
        )


def test_generate_formula_with_tool_call_rejects_dangerous_formula(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from server.app import formula_tools_integration

    client = _FakeClient(None, 0)
    client.content = json.dumps(
        {
            "formula": {
                "function": "WEBSERVICE",
                "args": [{"type": "text", "value": "https://example.com"}],
            },
            "explanation": "取网页内容",
        }
    )
    monkeypatch.setattr(
        formula_tools_integration, "OpenAICompatibleClient", lambda config, timeout: client
    )
    monkeypatch.setattr(formula_tools_integration, "formula_model_config", lambda: object())

    with pytest.raises(ValueError, match="被禁用的函数"):
        _run(
            formula_tools_integration.generate_formula_with_tool_call(
                active_cell="E2",
                description="取网页内容",
                headers=None,
                columns=None,
                dictionary=None,
                extra_sheets=None,
                model_id=None,
            )
        )


def test_generate_formula_with_tool_call_rejects_unknown_external_ref(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """未勾选的外部工作簿引用必须被拦截，与主流程安全护栏一致。"""
    from server.app import formula_tools_integration

    client = _FakeClient(None, 0)
    client.content = json.dumps(
        {
            "formula": {
                "function": "XLOOKUP",
                "args": [
                    {"type": "cell", "address": "D2"},
                    {
                        "type": "range",
                        "workbook": "secret.xlsx",
                        "sheet": "tt",
                        "col": "A",
                        "row1": 2,
                        "row2": 999,
                    },
                    {
                        "type": "range",
                        "workbook": "secret.xlsx",
                        "sheet": "tt",
                        "col": "B",
                        "row1": 2,
                        "row2": 999,
                    },
                    {"type": "text", "value": ""},
                ],
            },
            "explanation": "跨文件查表",
        }
    )
    monkeypatch.setattr(
        formula_tools_integration, "OpenAICompatibleClient", lambda config, timeout: client
    )
    monkeypatch.setattr(formula_tools_integration, "formula_model_config", lambda: object())

    with pytest.raises(ValueError, match="被禁用的函数"):
        _run(
            formula_tools_integration.generate_formula_with_tool_call(
                active_cell="E2",
                description="跨文件查表",
                headers=None,
                columns=None,
                dictionary=None,
                extra_sheets=[{"sourceFile": "tt.xlsx", "sheetName": "tt"}],
                model_id=None,
            )
        )


# ============================================================================
# 集成：generate_formula 回退
# ============================================================================


def test_generate_formula_falls_back_when_tool_call_fails(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """工具调用失败时，主流程自动回退到原有端到端生成，不破坏现有功能。"""
    from server.app import formula_tools_integration, rule_generator

    async def failing_tool_call(**kwargs):
        raise ValueError("模型返回格式错误: 测试")

    monkeypatch.setattr(
        formula_tools_integration,
        "generate_formula_with_tool_call",
        failing_tool_call,
    )
    monkeypatch.setattr(rule_generator, "formula_model_config", lambda: object())

    class _FallbackClient(_FakeClient):
        content = (
            '{"modernFormula": "=SUM(A1:A2)", "modernExplanation": "求和", '
            '"compatFormula": "=SUM(A1:A2)", "compatExplanation": "求和"}'
        )

    monkeypatch.setattr(
        rule_generator, "OpenAICompatibleClient", lambda config, timeout: _FallbackClient(None, 0)
    )

    request = rule_generator.GenerateFormulaRequest(
        description="求和",
        activeCell="A3",
    )
    response = _run(rule_generator.generate_formula(request))

    assert response.modernFormula == "=SUM(A1:A2)"
    assert response.compatFormula == "=SUM(A1:A2)"
    assert response.modernExplanation == "求和"


def test_generate_formula_uses_tool_call_result_when_successful(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """工具调用成功时，主流程直接采用编译结果，不再调用原有模型路径。"""
    from server.app import formula_tools_integration, rule_generator

    async def succeeding_tool_call(**kwargs):
        return {
            "modernFormula": "=SUM(A1:A2)",
            "modernExplanation": "求和",
            "compatFormula": "=SUM(A1:A2)",
            "compatExplanation": "求和",
        }

    monkeypatch.setattr(
        formula_tools_integration,
        "generate_formula_with_tool_call",
        succeeding_tool_call,
    )

    called: list[bool] = []

    class _ShouldNotBeUsed:
        def __init__(self, *args, **kwargs):
            called.append(True)

    monkeypatch.setattr(
        rule_generator, "OpenAICompatibleClient", _ShouldNotBeUsed
    )

    request = rule_generator.GenerateFormulaRequest(
        description="求和",
        activeCell="A3",
    )
    response = _run(rule_generator.generate_formula(request))

    assert response.modernFormula == "=SUM(A1:A2)"
    assert called == []
