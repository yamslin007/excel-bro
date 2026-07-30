from __future__ import annotations

import pytest
from pydantic import ValidationError

from server.app.models import (
    ActionExecutionResult,
    AnalysisPlan,
    IntentToolResponse,
    VerificationReport,
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
