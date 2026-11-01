from __future__ import annotations

import pytest
from pydantic import ValidationError

from server.app.models import (
    ActionExecutionResult,
    AnalysisPlan,
    IntentOption,
    IntentToolResponse,
    TaskProgressEvent,
    VerificationReport,
)


def test_task_progress_event_supports_safe_facts_and_cancelled_status() -> None:
    event = TaskProgressEvent.model_validate(
        {
            "operationId": "operation-progress-model",
            "sequence": 2,
            "phase": "planning",
            "stepId": "planning.tool.1.1",
            "status": "cancelled",
            "title": "任务已取消",
            "toolName": "read_range",
            "facts": [
                {"label": "范围", "value": "A1:B3", "style": "code"},
                {"label": "结果", "value": "3 行 × 2 列"},
            ],
            "timestamp": "2026-09-03T00:00:00Z",
        }
    )

    assert event.status == "cancelled"
    assert event.toolName == "read_range"
    assert event.facts is not None
    assert event.facts[0].style == "code"


def test_task_progress_event_rejects_oversized_fact_payloads() -> None:
    base = {
        "operationId": "operation-progress-invalid",
        "sequence": 1,
        "phase": "planning",
        "stepId": "planning.tool.invalid",
        "status": "completed",
        "title": "已完成",
        "timestamp": "2026-09-03T00:00:00Z",
    }
    with pytest.raises(ValidationError):
        TaskProgressEvent.model_validate(
            {
                **base,
                "facts": [
                    {"label": "结果", "value": str(index)}
                    for index in range(13)
                ],
            }
        )
    with pytest.raises(ValidationError):
        TaskProgressEvent.model_validate(
            {
                **base,
                "facts": [{"label": "结果", "value": "x" * 501}],
            }
        )


def test_intent_option_supports_explicit_replacement_mode() -> None:
    replacement = IntentOption.model_validate(
        {
            "id": "by-cinema",
            "label": "改为按影院建表",
            "description": "替换原拆分维度",
            "resolution": "按影院名称和影院编码建表，表内按影片汇总",
            "resolutionMode": "replace",
        }
    )
    legacy = IntentOption.model_validate(
        {
            "id": "ratio",
            "label": "补充分母",
            "description": "只补充占比口径",
            "resolution": "分母使用全表该影院记录数",
        }
    )

    assert replacement.resolutionMode == "replace"
    assert legacy.resolutionMode == "append"

    with pytest.raises(ValidationError):
        IntentOption.model_validate(
            {
                "id": "invalid",
                "label": "无效",
                "description": "无效模式",
                "resolution": "继续",
                "resolutionMode": "overwrite",
            }
        )


def test_unknown_action_is_rejected() -> None:
    with pytest.raises(ValidationError):
        AnalysisPlan.model_validate(
            {
                "id": "unsafe",
                "title": "危险计划",
                "summary": "不应通过",
                "assumptions": [],
                "warnings": [],
                "actions": [{"type": "runMacro", "sheet": "原始数据"}],
            }
        )


def test_action_execution_result_supports_partial_failure_statuses() -> None:
    failed = ActionExecutionResult.model_validate(
        {
            "index": 1,
            "type": "writeValues",
            "sheet": "结果",
            "status": "failed",
            "message": "写入范围无效",
        }
    )
    not_run = ActionExecutionResult.model_validate(
        {
            "index": 2,
            "type": "createWorksheet",
            "sheet": "后续",
            "status": "not_run",
        }
    )

    assert failed.status == "failed"
    assert failed.message == "写入范围无效"
    assert not_run.status == "not_run"


def test_verification_report_distinguishes_unverified_execution() -> None:
    report = VerificationReport.model_validate(
        {
            "status": "executed_unverified",
            "passed": False,
            "checks": [],
            "unverifiedActions": [
                {
                    "index": 0,
                    "type": "createChart",
                    "sheet": "结果",
                    "message": "图表已创建，但缺少独立验收",
                }
            ],
        }
    )

    assert report.status == "executed_unverified"
    assert report.unverifiedActions[0].type == "createChart"

    with pytest.raises(ValidationError):
        VerificationReport.model_validate(
            {
                "status": "verified",
                "passed": True,
                "checks": [],
                "unverifiedActions": [
                    {
                        "index": 0,
                        "type": "createChart",
                        "sheet": "结果",
                        "message": "仍有未验证动作",
                    }
                ],
            }
        )


def test_non_rectangular_table_is_rejected() -> None:
    with pytest.raises(ValidationError):
        AnalysisPlan.model_validate(
            {
                "id": "invalid-table",
                "title": "错误表格",
                "summary": "不应通过",
                "assumptions": [],
                "warnings": [],
                "actions": [
                    {
                        "type": "writeTable",
                        "sheet": "结果",
                        "startCell": "A1",
                        "headers": ["A", "B"],
                        "rows": [[1]],
                    }
                ],
            }
        )


def test_write_actions_receive_deterministic_acceptance_criteria() -> None:
    plan = AnalysisPlan.model_validate(
        {
            "id": "verified-table",
            "title": "生成表格",
            "summary": "写入并验证",
            "actions": [
                {
                    "type": "writeTable",
                    "sheet": "结果",
                    "startCell": "B3",
                    "headers": ["姓名", "得分"],
                    "rows": [["阿里", 44]],
                }
            ],
        }
    )

    assert [criterion.type for criterion in plan.acceptanceCriteria] == [
        "worksheetExists",
        "rangeEquals",
    ]
    range_check = plan.acceptanceCriteria[1]
    assert range_check.type == "rangeEquals"
    assert range_check.range == "B3:C4"
    assert range_check.expected == [["姓名", "得分"], ["阿里", 44]]


def test_write_formulas_receives_formulas_equal_acceptance() -> None:
    """U1a：writeFormulas 动作必须自动补 formulasEqual 验收条件。

    公式身份 = 公式文本（不是计算值），所以自动补的验收条件必须是
    `formulasEqual`（比对公式文本），而不是 `rangeEquals`（比对显示值）。
    少了这条，恢复执行后 `verificationGaps` 判不出已核实，整份计划会被记成
    `executed_unverified`，恢复收尾（确认/放弃）被卡在审阅之前。
    """
    plan = AnalysisPlan.model_validate(
        {
            "id": "formula-verified",
            "title": "写公式",
            "summary": "写入公式并验证",
            "actions": [
                {
                    "type": "writeFormulas",
                    "sheet": "结果",
                    "range": "B2:C2",
                    "formulas": [["=数据!A1*2", "=数据!A1*3"]],
                }
            ],
        }
    )

    assert [criterion.type for criterion in plan.acceptanceCriteria] == [
        "worksheetExists",
        "formulasEqual",
    ]
    formula_check = plan.acceptanceCriteria[1]
    assert formula_check.type == "formulasEqual"
    assert formula_check.range == "B2:C2"
    assert formula_check.expected == [["=数据!A1*2", "=数据!A1*3"]]


def test_plan_infers_sort_filter_and_table_acceptance() -> None:
    plan = AnalysisPlan.model_validate(
        {
            "id": "structured-verification",
            "title": "对象验收",
            "summary": "验证排序、筛选和表格",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": "结果",
                    "range": "A1:B3",
                    "values": [["名称", "数量"], ["乙", 2], ["甲", 1]],
                },
                {
                    "type": "sortRange",
                    "sheet": "结果",
                    "range": "A1:B3",
                    "keys": [{"column": 0, "ascending": True}],
                    "hasHeaders": True,
                },
                {
                    "type": "filterRange",
                    "sheet": "结果",
                    "range": "A1:B3",
                    "column": 1,
                    "values": [1],
                },
                {
                    "type": "createTable",
                    "sheet": "结果",
                    "range": "A1:B3",
                    "name": "ResultTable",
                    "hasHeaders": True,
                },
            ],
        }
    )

    criterion_types = [item.type for item in plan.acceptanceCriteria]
    assert "rangeEquals" not in criterion_types
    assert criterion_types == [
        "worksheetExists",
        "rangeSorted",
        "filterApplied",
        "tableExists",
    ]


def test_plan_only_verifies_the_final_filter_state() -> None:
    plan = AnalysisPlan.model_validate(
        {
            "id": "final-filter-state",
            "title": "清除筛选",
            "summary": "筛选后清除条件",
            "actions": [
                {
                    "type": "filterRange",
                    "sheet": "结果",
                    "range": "A1:B3",
                    "column": 1,
                    "values": [1],
                },
                {"type": "clearFilter", "sheet": "结果"},
            ],
        }
    )

    assert [item.type for item in plan.acceptanceCriteria] == [
        "worksheetExists",
        "filterCleared",
    ]


def test_plan_infers_format_validation_and_freeze_acceptance() -> None:
    plan = AnalysisPlan.model_validate(
        {
            "id": "format-verification",
            "title": "格式验收",
            "summary": "补齐格式、验证规则和冻结窗格条件",
            "actions": [
                {
                    "type": "setFill",
                    "sheet": "结果",
                    "range": "A1:B2",
                    "color": "#DFF3E4",
                },
                {
                    "type": "setNumberFormat",
                    "sheet": "结果",
                    "range": "B2",
                    "formatCode": "0.00",
                },
                {
                    "type": "setDataValidation",
                    "sheet": "结果",
                    "range": "B2:B10",
                    "validationType": "wholeNumber",
                    "formula1": 0,
                    "formula2": 100,
                    "operator": "between",
                    "allowBlank": False,
                },
                {
                    "type": "freezePanes",
                    "sheet": "结果",
                    "rows": 1,
                    "columns": 2,
                },
            ],
        }
    )

    assert [item.type for item in plan.acceptanceCriteria] == [
        "worksheetExists",
        "rangeFormatMatches",
        "rangeFormatMatches",
        "dataValidationMatches",
        "freezePanesMatches",
    ]
    assert plan.acceptanceCriteria[1].fillColor == "#DFF3E4"
    assert plan.acceptanceCriteria[2].numberFormat == "0.00"


def test_plan_only_keeps_final_overridden_format_properties() -> None:
    plan = AnalysisPlan.model_validate(
        {
            "id": "final-format-state",
            "title": "最终格式",
            "summary": "相同属性以后一个动作作为验收状态",
            "actions": [
                {
                    "type": "setFont",
                    "sheet": "结果",
                    "range": "A1",
                    "bold": True,
                    "color": "#111111",
                },
                {
                    "type": "setFont",
                    "sheet": "结果",
                    "range": "A1",
                    "bold": False,
                },
            ],
        }
    )

    format_checks = [
        item
        for item in plan.acceptanceCriteria
        if item.type == "rangeFormatMatches"
    ]
    assert len(format_checks) == 2
    assert format_checks[0].bold is None
    assert format_checks[0].fontColor == "#111111"
    assert format_checks[1].bold is False


def test_plan_infers_pivot_table_object_acceptance() -> None:
    plan = AnalysisPlan.model_validate(
        {
            "id": "pivot-verification",
            "title": "透视表验收",
            "summary": "补齐数据源、位置和字段配置",
            "actions": [
                {
                    "type": "createPivotTable",
                    "sheet": "透视",
                    "sourceSheet": "数据",
                    "sourceRange": "A1:C10",
                    "name": "SummaryPivot",
                    "destinationCell": "A1",
                    "rowFields": ["人员"],
                    "columnFields": ["月份"],
                    "valueFields": [
                        {"field": "得分", "aggregation": "average"}
                    ],
                }
            ],
        }
    )

    assert [item.type for item in plan.acceptanceCriteria] == [
        "worksheetExists",
        "pivotTableExists",
    ]
    criterion = plan.acceptanceCriteria[1]
    assert criterion.sourceSheet == "数据"
    assert criterion.sourceRange == "A1:C10"
    assert criterion.destinationCell == "A1"
    assert criterion.valueFields[0].aggregation == "average"


def test_acceptance_range_must_match_expected_matrix() -> None:
    with pytest.raises(ValidationError, match="尺寸一致"):
        AnalysisPlan.model_validate(
            {
                "id": "invalid-check",
                "title": "错误验收范围",
                "summary": "不应通过",
                "actions": [
                    {"type": "createWorksheet", "sheet": "结果"},
                ],
                "acceptanceCriteria": [
                    {
                        "type": "rangeEquals",
                        "sheet": "结果",
                        "range": "A1:Z100",
                        "expected": [[1]],
                    }
                ],
            }
        )


def test_partial_acceptance_criteria_are_completed_from_write_actions() -> None:
    plan = AnalysisPlan.model_validate(
        {
            "id": "partial-check",
            "title": "补齐验收",
            "summary": "模型只提供工作表检查",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": "结果",
                    "range": "B2",
                    "values": [[42]],
                }
            ],
            "acceptanceCriteria": [
                {"type": "worksheetExists", "sheet": "结果"}
            ],
        }
    )

    assert [criterion.type for criterion in plan.acceptanceCriteria] == [
        "worksheetExists",
        "rangeEquals",
    ]
    range_check = plan.acceptanceCriteria[1]
    assert range_check.type == "rangeEquals"
    assert range_check.range == "B2"
    assert range_check.expected == [[42]]


def test_split_group_aggregate_action_is_validated_and_warned() -> None:
    plan = AnalysisPlan.model_validate(
        {
            "id": "split-movies",
            "title": "按影片拆分",
            "summary": "按影院汇总记录数和占比",
            "actions": [
                {
                    "type": "splitGroupAggregate",
                    "sheet": "sheet-dc1",
                    "splitBy": "影片名称",
                    "groupBy": ["影院名称", "影院编码"],
                    "metrics": [
                        {
                            "operation": "countRows",
                            "outputName": "计数项:影厅",
                            "ratioOutputName": "占比",
                        }
                    ],
                }
            ],
        }
    )

    action = plan.actions[0]
    assert action.type == "splitGroupAggregate"
    assert action.maxOutputSheets == 200
    assert any("创建多个结果工作表" in warning for warning in plan.warnings)


def test_high_level_data_tool_request_is_validated() -> None:
    response = IntentToolResponse.model_validate(
        {
            "kind": "tool_request",
            "provider": "model",
            "summary": "按电影汇总",
            "confirmedPrompt": "在已选工作表中按电影汇总影厅数",
            "request": {
                "id": "query-1",
                "tool": "query_table",
                "arguments": {
                    "mode": "aggregate",
                    "scope": "selected",
                    "groupBy": ["影片名称"],
                    "metrics": [
                        {
                            "operation": "sum",
                            "field": "计数项:影厅",
                            "outputName": "影厅数",
                            "ratioOutputName": "占比",
                        }
                    ],
                    "sortBy": "占比",
                    "sortDirection": "desc",
                    "limit": 10,
                },
            },
        }
    )

    assert response.request.arguments.metrics[0].ratioOutputName == "占比"


def test_split_aggregate_sum_requires_a_field() -> None:
    with pytest.raises(ValidationError, match="必须指定 field"):
        AnalysisPlan.model_validate(
            {
                "id": "invalid-split",
                "title": "错误拆分",
                "summary": "缺少求和字段",
                "actions": [
                    {
                        "type": "splitGroupAggregate",
                        "sheet": "源数据",
                        "splitBy": "类别",
                        "groupBy": ["门店"],
                        "metrics": [
                            {"operation": "sum", "outputName": "金额"}
                        ],
                    }
                ],
            }
        )


def test_split_aggregate_rejects_duplicate_output_names() -> None:
    with pytest.raises(ValidationError, match="输出字段名称不能重复"):
        AnalysisPlan.model_validate(
            {
                "id": "duplicate-output",
                "title": "错误拆分",
                "summary": "输出字段重复",
                "actions": [
                    {
                        "type": "splitGroupAggregate",
                        "sheet": "源数据",
                        "splitBy": "类别",
                        "groupBy": ["门店"],
                        "metrics": [
                            {
                                "operation": "countRows",
                                "outputName": "门店",
                            }
                        ],
                    }
                ],
            }
        )


def _compute_values_plan(**overrides: object) -> dict[str, object]:
    query: dict[str, object] = {
        "mode": "aggregate",
        "scope": "selected",
        "groupBy": ["部门"],
        "metrics": [
            {"operation": "sum", "field": "金额", "outputName": "总金额"}
        ],
        "sortBy": "总金额",
        "sortDirection": "desc",
        "limit": 20,
    }
    query.update(overrides.pop("query", {}))  # type: ignore[arg-type]
    action: dict[str, object] = {
        "type": "computeValues",
        "sheet": "结果",
        "sourceSheet": "原始数据",
        "sourceRange": "A1:D500",
        "hasHeaders": True,
        "query": query,
        "destinationCell": "A1",
        "includeHeaders": True,
    }
    action.update(overrides)
    return {
        "id": "compute-1",
        "title": "按部门汇总",
        "summary": "执行时读取实时数据求和",
        "actions": [action],
    }


def test_compute_values_never_gets_a_range_equals_criterion() -> None:
    """执行时求值的动作行数未知，补 rangeEquals 会被尺寸校验直接拒掉。"""

    plan = AnalysisPlan.model_validate(_compute_values_plan())

    assert [item.type for item in plan.acceptanceCriteria] == ["worksheetExists"]
    assert plan.acceptanceCriteria[0].sheet == "结果"
    assert any("执行时" in warning for warning in plan.warnings)


def test_compute_values_carries_no_baked_result() -> None:
    """动作体里只有查询配置，没有任何算好的矩阵字段。"""

    plan = AnalysisPlan.model_validate(_compute_values_plan())
    action = plan.actions[0]

    assert action.type == "computeValues"
    assert not hasattr(action, "values")
    assert not hasattr(action, "rows")
    assert action.query.groupBy == ["部门"]


def test_compute_values_requires_a_single_destination_cell() -> None:
    with pytest.raises(ValidationError, match="必须是单个单元格"):
        AnalysisPlan.model_validate(_compute_values_plan(destinationCell="A1:C9"))


def test_compute_values_rejects_blank_field_names() -> None:
    with pytest.raises(ValidationError, match="字段名不能为空"):
        AnalysisPlan.model_validate(
            _compute_values_plan(query={"groupBy": ["部门", "  "]})
        )
