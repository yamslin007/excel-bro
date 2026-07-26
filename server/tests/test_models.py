from __future__ import annotations

import pytest
from pydantic import ValidationError

from server.app.models import AnalysisPlan, IntentToolResponse


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
