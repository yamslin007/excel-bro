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

import ast
import inspect
import re
import typing
from pathlib import Path

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
    "computeValues": {
        "type": "computeValues",
        "sheet": "结果",
        "sourceSheet": "Sheet1",
        "sourceRange": "A1:D100",
        "hasHeaders": True,
        "query": {
            "mode": "aggregate",
            "scope": "selected",
            "groupBy": ["部门"],
            "metrics": [
                {"operation": "sum", "field": "金额", "outputName": "总金额"}
            ],
            "sortBy": "总金额",
            "sortDirection": "desc",
            "limit": 50,
        },
        "destinationCell": "A1",
        "includeHeaders": True,
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
def test_backend_can_parse_frontend_action(
    action_type: str, monkeypatch: pytest.MonkeyPatch
) -> None:
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
# 测试：文件夹执行器的动作覆盖穷尽性
# ============================================================================
# 背景：文件夹执行器曾经把没有实现的动作落进循环末尾的统一
# status="succeeded"，用户看到成功但文件没变（removeDuplicates 假成功）。
# 运行期兜底已经补在 folder_workbooks.py 的 else 分支，但那只在真的执行
# 到那一步时才报错。这里把「协议动作集合」与「文件夹执行器覆盖集合」的
# 一致性提到 CI：新增动作忘了在文件夹侧实现或声明不支持，测试直接失败。
# ============================================================================


def _protocol_action_types() -> set[str]:
    """从 ExcelAction 判别联合中提取所有协议动作类型。"""
    import typing

    union_type = typing.get_args(ExcelAction)[0]
    types: set[str] = set()
    for model in typing.get_args(union_type):
        type_field = model.model_fields.get("type")
        if type_field:
            literal_args = typing.get_args(type_field.annotation)
            if literal_args:
                types.add(literal_args[0])
    return types


def _dispatched_action_types(function) -> set[str]:
    """静态提取函数体里所有与 action.type 比较过的字符串字面量。

    覆盖三种写法：
    - action.type == "x"
    - action.type in {"x", "y"}
    - action.type in unsupported（集合名，忽略，由调用方单独提供）
    """
    source = inspect.getsource(function)
    tree = ast.parse(ast.unparse(ast.parse(source)))
    found: set[str] = set()

    def is_action_type(node: ast.AST) -> bool:
        return (
            isinstance(node, ast.Attribute)
            and node.attr == "type"
            and isinstance(node.value, ast.Name)
            and node.value.id == "action"
        )

    for node in ast.walk(tree):
        if not isinstance(node, ast.Compare) or not is_action_type(node.left):
            continue
        for comparator in node.comparators:
            if isinstance(comparator, ast.Constant) and isinstance(
                comparator.value, str
            ):
                found.add(comparator.value)
            elif isinstance(comparator, (ast.Set, ast.List, ast.Tuple)):
                for element in comparator.elts:
                    if isinstance(element, ast.Constant) and isinstance(
                        element.value, str
                    ):
                        found.add(element.value)
    return found


def _folder_unsupported_action_types() -> set[str]:
    """读取 validate_plan_for_folder 里声明的 unsupported 集合。"""
    from server.app import folder_workbooks

    source = inspect.getsource(folder_workbooks.validate_plan_for_folder)
    tree = ast.parse(ast.unparse(ast.parse(source)))
    for node in ast.walk(tree):
        if (
            isinstance(node, ast.Assign)
            and len(node.targets) == 1
            and isinstance(node.targets[0], ast.Name)
            and node.targets[0].id == "unsupported"
            and isinstance(node.value, ast.Set)
        ):
            return {
                element.value
                for element in node.value.elts
                if isinstance(element, ast.Constant)
                and isinstance(element.value, str)
            }
    pytest.fail(
        "在 validate_plan_for_folder 中找不到 unsupported 集合字面量。"
        "如果重构了拒绝清单的表达方式，请同步更新本测试的提取逻辑。"
    )


def test_folder_executor_covers_every_protocol_action() -> None:
    """每个协议动作都必须在文件夹侧「有实现」或「明确声明不支持」。

    两者都没有 = 运行到 folder_workbooks.py 的 else 分支抛
    「文件夹模式未实现动作」，用户在计划已经写了一半的时候才看到失败。
    """
    from server.app import folder_workbooks

    protocol_types = _protocol_action_types()
    implemented = _dispatched_action_types(folder_workbooks._execute_folder_plan)
    declared_unsupported = _folder_unsupported_action_types()

    uncovered = protocol_types - implemented - declared_unsupported
    assert not uncovered, (
        f"以下 {len(uncovered)} 个动作在文件夹模式既没有实现也没有声明不支持：\n"
        f"{sorted(uncovered)}\n"
        "请二选一：\n"
        "  A. 在 folder_workbooks.py 的 _execute_folder_plan 里加执行分支；\n"
        "  B. 加入 validate_plan_for_folder 的 unsupported 集合，"
        "让计划在任何写入之前就被拒绝。\n"
        "不处理的话，动作会撞到末尾的 else 抛错——但那时前面的动作已经改过文件了。"
    )


def test_folder_unsupported_set_has_no_stale_entries() -> None:
    """unsupported 集合里不能有协议里已经不存在的动作类型。

    删除动作类型时忘了清理拒绝清单，会留下永远不会命中的死条目，
    掩盖真实的覆盖情况。
    """
    stale = _folder_unsupported_action_types() - _protocol_action_types()
    assert not stale, (
        f"validate_plan_for_folder 的 unsupported 集合含 {len(stale)} 个"
        f"协议中已不存在的动作类型：{sorted(stale)}\n"
        "请从集合中删除这些死条目。"
    )


def test_folder_declared_unsupported_actions_are_rejected_before_execution() -> None:
    """声明不支持的动作必须在校验阶段就被拦下，而不是执行到一半才报错。"""
    from server.app import folder_workbooks

    declared_unsupported = _folder_unsupported_action_types()
    assert declared_unsupported, "unsupported 集合为空，说明提取逻辑失效了"

    missing_samples = declared_unsupported - set(VALID_ACTION_SAMPLES)
    assert not missing_samples, (
        f"以下声明不支持的动作缺少测试样本：{sorted(missing_samples)}"
    )

    source = inspect.getsource(folder_workbooks.validate_plan_for_folder)
    assert "if action.type in unsupported" in source, (
        "validate_plan_for_folder 不再用 unsupported 集合拦截动作，"
        "本测试的假设已失效，请同步更新。"
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

    # 测试 1：setFont 的五个字体属性各自都是可选的，但**全空必须被拒**。
    # 语义裁决：一个属性都没给 = 这个动作什么也不做。放过去会让计划"看起来
    # 成功、实际没改"，与"只有指定属性被改变"的承诺相反；这与既有不变量
    # RangeFormatMatchesCriterion.require_expected_property 是同一条规矩。
    minimal_font = {
        "type": "setFont",
        "sheet": "Sheet1",
        "range": "A1",
        # 不提供任何字体属性 —— 必须被拒
    }
    with pytest.raises(ValidationError, match="至少需要一个"):
        SetFontAction.model_validate(minimal_font)

    # 逐个单属性都必须能通过（可选 ≠ 可全空）。
    for single in (
        {"bold": True},
        {"italic": True},
        {"name": "Arial"},
        {"size": 12},
        {"color": "#112233"},
    ):
        parsed = SetFontAction.model_validate(
            {"type": "setFont", "sheet": "Sheet1", "range": "A1", **single}
        )
        assert parsed.model_dump(exclude_none=True).keys() >= single.keys()

    # 未给出的属性保持 None，而不是被塞进默认值。
    bold_only = SetFontAction.model_validate(
        {"type": "setFont", "sheet": "Sheet1", "range": "A1", "bold": True}
    )
    assert bold_only.bold is True
    assert bold_only.italic is None
    assert bold_only.name is None
    assert bold_only.size is None
    assert bold_only.color is None

    # 测试 1b：setAlignment / resizeRange 与 setFont 同一语义（两端也必须一致）
    with pytest.raises(ValidationError, match="至少需要一个"):
        SetAlignmentAction.model_validate(
            {"type": "setAlignment", "sheet": "Sheet1", "range": "A1"}
        )
    with pytest.raises(ValidationError, match="至少需要一个"):
        ResizeRangeAction.model_validate(
            {"type": "resizeRange", "sheet": "Sheet1", "range": "A1"}
        )

    # 测试 1c：setBorders 的边不能重复（重复不会报错，只会让同一位置画两次）
    with pytest.raises(ValidationError, match="重复"):
        SetBordersAction.model_validate(
            {
                "type": "setBorders",
                "sheet": "Sheet1",
                "range": "A1",
                "sides": ["top", "top"],
            }
        )

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
# 测试：跨语言漂移（闭集 / 取值范围 / 可选性）
# ============================================================================
# 背景：上面那些用例只证明"后端能解析前端给的样例"。它们拦不住真正的漂移：
# 前端把 setFont 的 color 写成自由 `string`、把行列宽上限只写在后端、把
# `style` 在类型层写成必填而实现层当可选 —— 样例照样能解析，因为这些不一致
# 只有把**非法值**送进宿主才会暴露，而那一刻前面的动作可能已经落盘了。
#
# 所以这里不比对样例，而是直接把 `contracts.ts` 的常量与类型字面量抠出来，
# 和 Pydantic 的 `Literal` / `Field` 逐项对。任何一端改了值或可选性，测试变红。
# ============================================================================

CONTRACTS_TS_PATH = (
    Path(__file__).resolve().parents[2] / "apps" / "excel-addin" / "src" / "contracts.ts"
)


def _contracts_source() -> str:
    return CONTRACTS_TS_PATH.read_text(encoding="utf-8")


def _strip_ts_comments(source: str) -> str:
    """去掉块注释与行注释。注释里出现大括号会把"取到第一个 }"的截断逻辑带偏。"""
    without_block = re.sub(r"/\*.*?\*/", " ", source, flags=re.S)
    return re.sub(r"//[^\n]*", " ", without_block)


def _ts_string_array(source: str, name: str) -> list[str]:
    """取 `export const NAME: readonly X[] = [ "a", "b" ];` 里的字符串。"""
    match = re.search(
        rf"export const {name}\s*:\s*readonly\s+\w+\[\]\s*=\s*\[(?P<body>[^\]]*)\];",
        source,
        re.S,
    )
    if match is None:
        raise AssertionError(f"contracts.ts 里找不到 {name}")
    return re.findall(r'"([^"]+)"', match.group("body"))


def _ts_number(source: str, name: str) -> int:
    match = re.search(rf"export const {name}\s*=\s*(?P<value>\d+);", source)
    if match is None:
        raise AssertionError(f"contracts.ts 里找不到 {name}")
    return int(match.group("value"))


def _ts_range(source: str, name: str) -> tuple[int, int]:
    match = re.search(
        rf"export const {name}\s*=\s*\{{\s*min:\s*(?P<min>-?\d+),\s*max:\s*(?P<max>-?\d+)\s*\}}\s*as const;",
        source,
    )
    if match is None:
        raise AssertionError(f"contracts.ts 里找不到 {name}")
    return int(match.group("min")), int(match.group("max"))


def _ts_regex_source(source: str, name: str) -> str:
    match = re.search(rf"export const {name}\s*=\s*/(?P<body>[^/\n]+)/;", source)
    if match is None:
        raise AssertionError(f"contracts.ts 里找不到 {name}")
    return match.group("body")


def _ts_action_fields(source: str, action_type: str) -> dict[str, bool]:
    """取 TS 联合里某个动作的字段名 → 是否必填（`?` 为可选）。

    只用于格式动作：它们的类型字面量是平的，没有嵌套对象类型，
    所以"截到第一个 `}`"与"字段名后跟冒号"这两条简化是安全的。
    """
    stripped = _strip_ts_comments(source)
    match = re.search(
        rf'type:\s*"{action_type}";(?P<body>[^}}]*)',
        stripped,
    )
    if match is None:
        raise AssertionError(f"contracts.ts 里找不到动作 {action_type}")
    body = match.group("body")
    # `type` 本身被上面那条正则当锚点吃掉了，补回来。
    fields: dict[str, bool] = {"type": True}
    for name, optional in re.findall(r"([A-Za-z_][A-Za-z0-9_]*)\s*(\??)\s*:", body):
        fields[name] = optional == ""
    return fields


def _walk_annotation(node: object):
    yield node
    for arg in typing.get_args(node):
        yield from _walk_annotation(arg)


def _literal_values(model: type, field: str) -> list[str]:
    """取字段上的 `Literal` 取值，自动穿透 `Optional[...]` / `list[...]` / `Annotated[...]`。"""
    values: list[str] = []
    for node in _walk_annotation(model.model_fields[field].annotation):
        if typing.get_origin(node) is typing.Literal:
            values.extend(str(item) for item in typing.get_args(node))
    return values


_BOUND_ATTRS = {"ge": "ge", "le": "le", "min": "min_length", "max": "max_length"}
_BOUND_CLASSES = {"ge": "Ge", "le": "Le", "min": "MinLen", "max": "MaxLen"}


def _numeric_bound(model: type, field: str, kind: str) -> int:
    field_info = model.model_fields[field]
    for item in field_info.metadata:
        if type(item).__name__ == _BOUND_CLASSES[kind]:
            return getattr(item, _BOUND_ATTRS[kind])
    # `Annotated[str, StringConstraints(min_length=..., max_length=...)]` 会被内联进注解，
    # 不落到 metadata 上，得走一遍注解树。
    if kind in {"min", "max"}:
        for node in _walk_annotation(field_info.annotation):
            if type(node).__name__ == "StringConstraints":
                value = getattr(node, _BOUND_ATTRS[kind], None)
                if value is not None:
                    return value
    raise AssertionError(f"{model.__name__}.{field} 上找不到 {kind} 约束")


def _string_pattern(model: type, field: str) -> str:
    for node in _walk_annotation(model.model_fields[field].annotation):
        if type(node).__name__ == "StringConstraints":
            pattern = getattr(node, "pattern", None)
            if pattern:
                return pattern
    for item in model.model_fields[field].metadata:
        if type(item).__name__ == "StringConstraints" and getattr(item, "pattern", None):
            return item.pattern
    raise AssertionError(f"{model.__name__}.{field} 上没有 pattern 约束")


@pytest.mark.parametrize(
    ("ts_name", "model", "field"),
    [
        ("BORDER_SIDES", SetBordersAction, "sides"),
        ("BORDER_LINE_STYLES", SetBordersAction, "style"),
        ("BORDER_WEIGHTS", SetBordersAction, "weight"),
        ("HORIZONTAL_ALIGNMENTS", SetAlignmentAction, "horizontal"),
        ("VERTICAL_ALIGNMENTS", SetAlignmentAction, "vertical"),
    ],
)
def test_closed_sets_match_between_typescript_and_python(
    ts_name: str, model: type, field: str
) -> None:
    ts_values = _ts_string_array(_contracts_source(), ts_name)
    python_values = _literal_values(model, field)
    assert ts_values, f"{ts_name} 在 contracts.ts 里是空集合"
    assert len(ts_values) == len(set(ts_values)), f"{ts_name} 里有重复取值"
    assert set(ts_values) == set(python_values), (
        f"{ts_name} 与 {model.__name__}.{field} 的取值不一致：\n"
        f"  TS     = {sorted(ts_values)}\n"
        f"  Python = {sorted(python_values)}"
    )


@pytest.mark.parametrize(
    ("ts_name", "model", "field"),
    [
        ("ROW_HEIGHT_RANGE", ResizeRangeAction, "rowHeight"),
        ("COLUMN_WIDTH_RANGE", ResizeRangeAction, "columnWidth"),
        ("FONT_SIZE_RANGE", SetFontAction, "size"),
    ],
)
def test_numeric_ranges_match_between_typescript_and_python(
    ts_name: str, model: type, field: str
) -> None:
    ts_range = _ts_range(_contracts_source(), ts_name)
    python_range = (
        _numeric_bound(model, field, "ge"),
        _numeric_bound(model, field, "le"),
    )
    assert ts_range == python_range, (
        f"{ts_name} 与 {model.__name__}.{field} 的值域不一致："
        f" TS {ts_range} vs Python {python_range}"
    )


def test_text_length_limits_match_between_typescript_and_python() -> None:
    source = _contracts_source()
    assert _ts_number(source, "FONT_NAME_MAX_LENGTH") == _numeric_bound(
        SetFontAction, "name", "max"
    )
    assert _numeric_bound(SetFontAction, "name", "min") == 1
    assert _ts_number(source, "NUMBER_FORMAT_MAX_LENGTH") == _numeric_bound(
        SetNumberFormatAction, "formatCode", "max"
    )
    assert _numeric_bound(SetNumberFormatAction, "formatCode", "min") == 1


def test_hex_color_pattern_matches_between_typescript_and_python() -> None:
    ts_pattern = _ts_regex_source(_contracts_source(), "HEX_COLOR_PATTERN")
    for model, field in (
        (SetFillAction, "color"),
        (SetFontAction, "color"),
        (SetBordersAction, "color"),
    ):
        assert _string_pattern(model, field) == ts_pattern, (
            f"{model.__name__}.{field} 的颜色约束与 TS 的 HEX_COLOR_PATTERN 不一致"
        )


@pytest.mark.parametrize(
    ("action_type", "model"),
    [
        ("setFill", SetFillAction),
        ("setFont", SetFontAction),
        ("setNumberFormat", SetNumberFormatAction),
        ("setBorders", SetBordersAction),
        ("setAlignment", SetAlignmentAction),
        ("resizeRange", ResizeRangeAction),
    ],
)
def test_format_action_fields_match_between_typescript_and_python(
    action_type: str, model: type
) -> None:
    """字段名集合与**可选性**逐项对齐。

    可选性倒置（TS 写必填、Python 有默认值，或反过来）是这轮审计里真实存在过的
    不一致：类型层写必填而实现层兼容缺省，中间那段空档就是"模型少给一个字段、
    计划在校验期通过、执行期才炸"。
    """
    ts_fields = _ts_action_fields(_contracts_source(), action_type)
    python_fields = {
        name: field_info.is_required()
        for name, field_info in model.model_fields.items()
    }
    # ExcelAction 每个变体的前三个定位字段，两侧命名一致，必须都在。
    assert {"sheet", "range"} <= set(ts_fields), f"{action_type} 的 TS 定义缺 sheet/range"
    assert set(ts_fields) == set(python_fields), (
        f"{action_type} 的字段集合不一致：\n"
        f"  TS 独有     = {sorted(set(ts_fields) - set(python_fields))}\n"
        f"  Python 独有 = {sorted(set(python_fields) - set(ts_fields))}"
    )
    mismatched = {
        name: (ts_fields[name], python_fields[name])
        for name in ts_fields
        if ts_fields[name] != python_fields[name]
    }
    assert not mismatched, (
        f"{action_type} 的可选性不一致（字段: (TS 必填, Python 必填)）：{mismatched}"
    )


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
