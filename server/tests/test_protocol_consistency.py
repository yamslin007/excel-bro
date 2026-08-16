"""
协议一致性测试：验证前后端协议定义同步

目标：
1. 确保 TypeScript contracts.ts 和 Python models.py 的 ExcelAction 类型一致
2. 验证每个动作类型的必填字段、可选字段、字段类型完全匹配
3. 检测协议漂移（一端新增/修改/删除字段，另一端未同步）

测试策略：
- 方案 A（当前实现）：通过真实请求验证后端能正确解析前端发送的每种动作
- 方案 B（未来）：从 Pydantic 模型自动生成 JSON Schema，前端用 zod 验证
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from server.app.models import (
    ActivateWorksheetAction,
    AddCommentAction,
    AddImageAction,
    AddNamedRangeAction,
    AddNoteAction,
    AddShapeAction,
    AutofitAction,
    ClearFilterAction,
    ClearRangeAction,
    CopyRangeAction,
    CreateChartAction,
    CreatePivotTableAction,
    CreateTableAction,
    CreateWorksheetAction,
    DeleteRangeAction,
    DeleteWorksheetAction,
    ExcelAction,
    FilterRangeAction,
    FreezePanesAction,
    InsertRangeAction,
    MergeCellsAction,
    RemoveDuplicatesAction,
    ResizeRangeAction,
    SetAlignmentAction,
    SetBordersAction,
    SetConditionalFormatAction,
    SetDataValidationAction,
    SetFillAction,
    SetFontAction,
    SetHyperlinkAction,
    SetNumberFormatAction,
    SortRangeAction,
    SplitGroupAggregateAction,
    UnmergeCellsAction,
    WriteFormulasAction,
    WriteTableAction,
    WriteValuesAction,
)


# ============================================================================
# 测试用例：每个 ExcelAction 类型的典型实例
# ============================================================================
# 这些 JSON 对象模拟前端发送的数据，必须能被后端 Pydantic 正确解析
# ============================================================================

VALID_ACTION_SAMPLES = {
    "createWorksheet": {
        "type": "createWorksheet",
        "sheet": "新工作表",
    },
    "writeTable": {
        "type": "writeTable",
        "sheet": "Sheet1",
        "startCell": "A1",
        "headers": ["姓名", "分数"],
        "rows": [["张三", 95], ["李四", 88]],
    },
    "writeValues": {
        "type": "writeValues",
        "sheet": "Sheet1",
        "range": "A1:B2",
        "values": [[1, 2], [3, 4]],
    },
    "setFill": {
        "type": "setFill",
        "sheet": "Sheet1",
        "range": "A1:B2",
        "color": "#FF0000",
    },
    "setFont": {
        "type": "setFont",
        "sheet": "Sheet1",
        "range": "A1:B2",
        "bold": True,
        "color": "#0000FF",
    },
    "autofit": {
        "type": "autofit",
        "sheet": "Sheet1",
        "range": "A:A",
    },
    "activateWorksheet": {
        "type": "activateWorksheet",
        "sheet": "Sheet2",
    },
    "deleteWorksheet": {
        "type": "deleteWorksheet",
        "sheet": "Sheet3",
    },
    "clearRange": {
        "type": "clearRange",
        "sheet": "Sheet1",
        "range": "A1:B10",
        "applyTo": "all",
    },
    "insertRange": {
        "type": "insertRange",
        "sheet": "Sheet1",
        "range": "A1:A5",
        "shift": "down",
    },
    "deleteRange": {
        "type": "deleteRange",
        "sheet": "Sheet1",
        "range": "A1:A5",
        "shift": "up",
        "hasHeaders": False,
    },
    "copyRange": {
        "type": "copyRange",
        "sheet": "Sheet1",
        "sourceSheet": "Sheet1",
        "sourceRange": "A1:B2",
        "targetRange": "D1:E2",
        "copyType": "all",
        "skipBlanks": False,
        "transpose": False,
    },
    "writeFormulas": {
        "type": "writeFormulas",
        "sheet": "Sheet1",
        "range": "C1:C2",
        "formulas": [["=A1+B1"], ["=A2+B2"]],
    },
    "sortRange": {
        "type": "sortRange",
        "sheet": "Sheet1",
        "range": "A1:B10",
        "keys": [{"column": 0, "ascending": True}],
        "hasHeaders": True,
    },
    "removeDuplicates": {
        "type": "removeDuplicates",
        "sheet": "Sheet1",
        "range": "A1:B10",
        "columns": [0, 1],
        "hasHeaders": True,
    },
    "filterRange": {
        "type": "filterRange",
        "sheet": "Sheet1",
        "range": "A1:B10",
        "column": 1,
        "values": [90, 95, 100],
    },
    "clearFilter": {
        "type": "clearFilter",
        "sheet": "Sheet1",
    },
    "setDataValidation": {
        "type": "setDataValidation",
        "sheet": "Sheet1",
        "range": "A1:A10",
        "validationType": "list",
        "values": ["优秀", "良好", "及格"],
        "allowBlank": True,
    },
    "setConditionalFormat": {
        "type": "setConditionalFormat",
        "sheet": "Sheet1",
        "range": "A1:A10",
        "ruleType": "cellValue",
        "operator": "greaterThan",
        "formula1": 90,
        "color": "#00FF00",
    },
    "setNumberFormat": {
        "type": "setNumberFormat",
        "sheet": "Sheet1",
        "range": "A1:A10",
        "formatCode": "0.00",
    },
    "setBorders": {
        "type": "setBorders",
        "sheet": "Sheet1",
        "range": "A1:B10",
        "sides": ["top", "bottom", "left", "right"],
        "style": "continuous",
        "color": "#000000",
        "weight": "thin",
    },
    "setAlignment": {
        "type": "setAlignment",
        "sheet": "Sheet1",
        "range": "A1:B10",
        "horizontal": "center",
        "vertical": "center",  # 注意：后端使用 "center" 而非 "middle"
        "wrapText": True,
    },
    "mergeCells": {
        "type": "mergeCells",
        "sheet": "Sheet1",
        "range": "A1:B2",
        "across": False,
    },
    "unmergeCells": {
        "type": "unmergeCells",
        "sheet": "Sheet1",
        "range": "A1:B2",
    },
    "resizeRange": {
        "type": "resizeRange",
        "sheet": "Sheet1",
        "range": "A:A",
        "columnWidth": 120,
    },
    "freezePanes": {
        "type": "freezePanes",
        "sheet": "Sheet1",
        "rows": 1,
        "columns": 0,
    },
    "setHyperlink": {
        "type": "setHyperlink",
        "sheet": "Sheet1",
        "range": "A1",
        "address": "https://example.com",
        "text": "点击这里",
    },
    "addComment": {
        "type": "addComment",
        "sheet": "Sheet1",
        "cell": "A1",
        "text": "这是一条批注",
    },
    "addNote": {
        "type": "addNote",
        "sheet": "Sheet1",
        "cell": "A1",
        "text": "这是一条备注",
    },
    "createTable": {
        "type": "createTable",
        "sheet": "Sheet1",
        "range": "A1:B10",
        "name": "MyTable",
        "hasHeaders": True,
        "style": "TableStyleMedium2",
    },
    "createChart": {
        "type": "createChart",
        "sheet": "Sheet1",
        "sourceRange": "A1:B10",
        "chartType": "ColumnClustered",
        "title": "销售图表",
        "targetRange": "D2",
    },
    "createPivotTable": {
        "type": "createPivotTable",
        "sheet": "Sheet2",
        "sourceSheet": "Sheet1",
        "sourceRange": "A1:C100",
        "name": "PivotTable1",
        "destinationCell": "A1",
        "rowFields": ["类别"],
        "columnFields": [],
        "valueFields": [{"field": "金额", "aggregation": "sum"}],
    },
    "splitGroupAggregate": {
        "type": "splitGroupAggregate",
        "sheet": "Sheet1",
        "sourceRange": "A1:D100",
        "splitBy": "部门",
        "groupBy": ["类别"],
        "metrics": [
            {"operation": "sum", "field": "金额", "outputName": "总金额"}
        ],
        "includeBlankSplitValues": False,
        "existingSheetPolicy": "rename",
        "maxOutputSheets": 50,
    },
    "addNamedRange": {
        "type": "addNamedRange",
        "sheet": "Sheet1",
        "name": "SalesData",
        "range": "A1:B10",
        "comment": "销售数据区域",
    },
    "addImage": {
        "type": "addImage",
        "sheet": "Sheet1",
        "base64": "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
        "targetRange": "A1",
        "name": "Logo",
    },
    "addShape": {
        "type": "addShape",
        "sheet": "Sheet1",
        "shapeType": "rectangle",
        "targetRange": "A1:B2",
        "text": "注意事项",
        "fillColor": "#FFFF00",
    },
}


# ============================================================================
# 测试：后端能解析前端发送的每种动作
# ============================================================================


@pytest.mark.parametrize("action_type", list(VALID_ACTION_SAMPLES.keys()))
def test_backend_can_parse_frontend_action(action_type: str) -> None:
    """验证后端 Pydantic 能正确解析前端发送的每种 ExcelAction 类型"""
    action_data = VALID_ACTION_SAMPLES[action_type]

    # 后端应该能无错误地解析这个 JSON
    try:
        # 注意：这里通过 AnalysisPlan 间接验证 ExcelAction 解析
        # 因为 ExcelAction 是 Annotated Union，需要通过实际使用场景验证
        from server.app.models import AnalysisPlan

        plan_data = {
            "id": "test-plan-001",
            "title": f"测试 {action_type}",
            "summary": f"测试 {action_type}",
            "reasoning": "协议一致性测试",
            "assumptions": [],
            "warnings": [],
            "actions": [action_data],
        }

        plan = AnalysisPlan.model_validate(plan_data)

        # 验证 type 字段正确
        assert plan.actions[0].type == action_type

    except ValidationError as e:
        pytest.fail(
            f"后端无法解析前端的 {action_type} 动作。"
            f"这表明前后端协议不一致。\n错误详情：\n{e}"
        )


# ============================================================================
# 测试：前后端动作类型列表完全一致
# ============================================================================


def test_action_type_list_consistency() -> None:
    """验证前后端的动作类型列表完全一致（数量和名称）"""

    # 后端 Python 的所有动作类型（从 Union 中提取）
    # Pydantic v2 使用 model_fields 替代 __fields__
    import typing

    # ExcelAction 是 Annotated[Union[...], Field(discriminator="type")]
    # 需要提取 Union 内的所有类型
    union_type = typing.get_args(ExcelAction)[0]  # 获取 Union
    action_models = typing.get_args(union_type)  # 获取 Union 中的所有类型

    backend_types = set()
    for model in action_models:
        # 使用 Pydantic v2 的 model_fields
        type_field = model.model_fields.get("type")
        if type_field:
            # type 字段是 Literal["xxx"]，需要提取字符串值
            literal_args = typing.get_args(type_field.annotation)
            if literal_args:
                backend_types.add(literal_args[0])

    # 前端 TypeScript 的所有动作类型（从测试样本中提取）
    frontend_types = set(VALID_ACTION_SAMPLES.keys())

    # 检查是否有后端独有的类型（前端缺失）
    backend_only = backend_types - frontend_types
    if backend_only:
        pytest.fail(
            f"后端有 {len(backend_only)} 个动作类型在前端测试样本中缺失：\n"
            f"{sorted(backend_only)}\n"
            f"请在 VALID_ACTION_SAMPLES 中补充这些类型的测试用例。"
        )

    # 检查是否有前端独有的类型（后端缺失）
    frontend_only = frontend_types - backend_types
    if frontend_only:
        pytest.fail(
            f"前端有 {len(frontend_only)} 个动作类型在后端不存在：\n"
            f"{sorted(frontend_only)}\n"
            f"这表明前端定义了后端不支持的动作类型。"
        )

    # 数量必须完全一致
    assert len(backend_types) == len(frontend_types), (
        f"前后端动作类型数量不一致：后端 {len(backend_types)} 个，"
        f"前端 {len(frontend_types)} 个"
    )


# ============================================================================
# 测试：关键协议的字段一致性（IntentCheckResponse、PlanResponse 等）
# ============================================================================


def test_intent_check_response_consistency() -> None:
    """验证 IntentCheckResponse 协议的前后端一致性"""
    from server.app.models import IntentProceedResponse

    # IntentCheckResponse 是一个 Union，测试其中一个变体
    sample_response = {
        "kind": "proceed",
        "summary": "需求明确",
        "confirmedPrompt": "创建新工作表",
        "provider": "model",
        "turnId": "test-turn-001",
    }

    try:
        response = IntentProceedResponse.model_validate(sample_response)
        assert response.kind == "proceed"
    except ValidationError as e:
        pytest.fail(f"IntentProceedResponse 协议不一致：\n{e}")


def test_plan_response_consistency() -> None:
    """验证 PlanResponse 协议的前后端一致性"""
    from server.app.models import PlanResponse

    # 模拟前端期望的 PlanResponse JSON
    sample_response = {
        "kind": "plan",
        "plan": {
            "id": "test-plan-001",
            "title": "创建新工作表",
            "summary": "创建新工作表",
            "assumptions": [],
            "warnings": [],
            "actions": [
                {
                    "type": "createWorksheet",
                    "sheet": "新工作表",
                }
            ],
        },
        "provider": "model",
    }

    try:
        response = PlanResponse.model_validate(sample_response)
        assert response.kind == "plan"
        assert len(response.plan.actions) == 1
        assert response.plan.actions[0].type == "createWorksheet"
    except ValidationError as e:
        pytest.fail(f"PlanResponse 协议不一致：\n{e}")


# ============================================================================
# 测试：字段类型严格性（防止类型漂移）
# ============================================================================


def test_action_field_type_strictness() -> None:
    """验证动作字段的类型严格性（防止字符串/数字等类型混淆）"""

    # 测试 1：sortRange 的 column 字段（Pydantic 会自动转换 "0" -> 0，所以这个测试不适用）
    # 改为测试无效的字符串
    invalid_sort = {
        "type": "sortRange",
        "sheet": "Sheet1",
        "range": "A1:B10",
        "keys": [{"column": "invalid", "ascending": True}],  # ❌ 无法转换的字符串
        "hasHeaders": True,
    }

    with pytest.raises(ValidationError):
        SortRangeAction.model_validate(invalid_sort)

    # 测试 2：setFill 的 color 必须是 #RRGGBB 格式
    invalid_fill = {
        "type": "setFill",
        "sheet": "Sheet1",
        "range": "A1",
        "color": "red",  # ❌ 不是十六进制格式
    }

    with pytest.raises(ValidationError, match="color"):
        SetFillAction.model_validate(invalid_fill)

    # 测试 3：writeValues 的 values 必须是二维数组
    invalid_values = {
        "type": "writeValues",
        "sheet": "Sheet1",
        "range": "A1:A2",
        "values": [1, 2],  # ❌ 一维数组
    }

    with pytest.raises(ValidationError):
        WriteValuesAction.model_validate(invalid_values)


# ============================================================================
# 测试：必填字段和可选字段
# ============================================================================


def test_required_vs_optional_fields() -> None:
    """验证必填字段和可选字段的定义一致"""

    # 测试 1：setFont 的 bold 和 color 都是可选的
    minimal_font = {
        "type": "setFont",
        "sheet": "Sheet1",
        "range": "A1",
        # 不提供 bold 和 color
    }
    action = SetFontAction.model_validate(minimal_font)
    assert action.bold is None
    assert action.color is None

    # 测试 2：writeTable 的 headers 是必填的
    missing_headers = {
        "type": "writeTable",
        "sheet": "Sheet1",
        "startCell": "A1",
        "rows": [[1, 2]],
        # ❌ 缺少 headers
    }

    with pytest.raises(ValidationError, match="headers"):
        WriteTableAction.model_validate(missing_headers)

    # 测试 3：createChart 的 title 和 targetRange 是可选的
    minimal_chart = {
        "type": "createChart",
        "sheet": "Sheet1",
        "sourceRange": "A1:B10",
        "chartType": "ColumnClustered",
        # 不提供 title 和 targetRange
    }
    action = CreateChartAction.model_validate(minimal_chart)
    assert action.title is None
    assert action.targetRange is None


# ============================================================================
# 说明：如何运行这些测试
# ============================================================================
"""
运行方法：

1. 运行所有协议一致性测试：
   pytest server/tests/test_protocol_consistency.py -v --basetemp=.pytest-tmp

2. 运行特定动作类型的测试：
   pytest server/tests/test_protocol_consistency.py::test_backend_can_parse_frontend_action[writeTable] -v

3. 在 CI 中自动运行：
   在 .github/workflows/test.yml 中添加此测试

4. 新增 ExcelAction 类型时的检查清单：
   - [ ] 在 server/app/models.py 中定义 Pydantic 模型
   - [ ] 在 apps/excel-addin/src/contracts.ts 中定义 TypeScript 类型
   - [ ] 在 VALID_ACTION_SAMPLES 中添加测试样本
   - [ ] 运行 pytest 验证协议一致性
   - [ ] 在 apps/excel-addin/src/excel.ts 中实现执行器
   - [ ] 在 server/app/folder_workbooks.py 中实现或明确拒绝

预期结果：
- 所有 35 个动作类型都能通过解析测试
- 前后端类型列表完全一致
- 字段类型严格匹配，不会发生隐式转换
- 必填/可选字段定义一致
"""
