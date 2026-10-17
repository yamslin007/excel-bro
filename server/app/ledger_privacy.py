"""落盘侧的隐私投影：规划结果进入持久账本前的去值处理。

``planning_tasks.result_json`` 是**持久**的：写进去的单元格值会留在磁盘上。
这与 ``WpsRecoveryDescriptor`` 对执行账本的要求是同一条约定 —— 值矩阵只停在本
机内存，宁可让账本少记一点，也不让单元格内容落盘。

去值之后**不能**再让那条记录长得像一个完整的 ``AssistantResponse``：
一个"矩阵全空"的计划看着完整，回放出来却是个假计划。所以持久侧存的是带
``schemaVersion`` 的**恢复描述**（``ledger_recovery_descriptor``）——它明确写出
"值被省略了"、省略了哪些字段、以及要取值时该用哪个本地结果引用。

轮内的完整响应仍走本机回环的 SSE 通道交给本地客户端（预览与执行需要值），
那条通道本来就在回环上，与 ``/api/sdk/tool-bridge/result`` 同一口径。
"""

from __future__ import annotations

from typing import Any

# 恢复描述自身的版本。字段含义变化就**升版本**，让读到旧记录的调用方能分辨。
RECOVERY_SCHEMA_VERSION = 1

# 值矩阵字段名：按**键名**判定而不是按动作/判据类型枚举，新增类型忘记登记也
# 不会把值漏进账本。``expected`` 是 ``rangeEquals`` / ``formulasEqual`` 判据里
# 的期望矩阵，和动作里的 ``values`` 一样是单元格内容。
MATRIX_FIELD_NAMES = frozenset(
    {
        "values",
        "rows",
        "headers",
        "formulas",
        "data",
        "preview",
        "matrix",
        "expected",
    }
)


def matrix_shape(value: Any) -> dict[str, Any]:
    if not isinstance(value, list):
        return {"omitted": True, "rowCount": 0, "columnCount": 0}
    row_count = len(value)
    first = value[0] if value else None
    column_count = len(first) if isinstance(first, list) else (1 if row_count else 0)
    return {"omitted": True, "rowCount": row_count, "columnCount": column_count}


def _strip_object(entry: dict[str, Any]) -> tuple[dict[str, Any], bool]:
    """Return a copy of one object without value matrices, plus whether any was dropped."""
    cleaned: dict[str, Any] = {}
    omitted: list[dict[str, Any]] = []
    for key, value in entry.items():
        if key in MATRIX_FIELD_NAMES:
            omitted.append({"field": key, **matrix_shape(value)})
        else:
            cleaned[key] = value
    if omitted:
        cleaned["omittedMatrices"] = omitted
    return cleaned, bool(omitted)


def _strip_matrices(entries: Any) -> list[dict[str, Any]]:
    if not isinstance(entries, list):
        return []
    safe: list[dict[str, Any]] = []
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        cleaned, _omitted = _strip_object(entry)
        safe.append(cleaned)
    return safe


def ledger_safe_actions(actions: Any) -> list[dict[str, Any]]:
    return _strip_matrices(actions)


def ledger_safe_criteria(criteria: Any) -> list[dict[str, Any]]:
    return _strip_matrices(criteria)


def ledger_safe_plan(plan: Any) -> Any:
    """Return a copy of a serialized ``AnalysisPlan`` without any value matrix."""
    safe, _omitted = _safe_plan(plan)
    return safe


def _safe_plan(plan: Any) -> tuple[Any, bool]:
    if not isinstance(plan, dict):
        return plan, False
    safe = dict(plan)
    omitted = False
    # 树状字段自带矩阵（``values``），验收判据同样带矩阵（``rangeEquals.expected``）。
    for key in ("actions", "acceptanceCriteria"):
        entries, dropped = _strip_list(plan.get(key))
        safe[key] = entries
        omitted = omitted or dropped
    # 值被省略后，``summary`` / ``title`` 仍是模型给出的自由文本，可能含源自
    # 数据的内容；但它们是计划元数据而非单元格值，保持原样以便账本可读。
    return safe, omitted


def _strip_list(entries: Any) -> tuple[list[dict[str, Any]], bool]:
    if not isinstance(entries, list):
        return [], False
    dropped = False
    result: list[dict[str, Any]] = []
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        cleaned, omitted = _strip_object(entry)
        dropped = dropped or omitted
        result.append(cleaned)
    return result, dropped


def _project_response(payload: dict[str, Any]) -> tuple[dict[str, Any], bool]:
    """去值投影：返回 (载荷, 是否有值被省略)。"""
    omitted = False
    plan = payload.get("plan")
    if isinstance(plan, dict):
        safe_plan, plan_omitted = _safe_plan(plan)
        payload = {**payload, "plan": safe_plan}
        omitted = omitted or plan_omitted
    # 答案里的 ``resultContext.rows`` 同样是单元格内容，必须一起去值；
    # ``localResult`` 只带引用与结构，本来就没有值。
    for key in ("resultContext", "localResult"):
        node = payload.get(key)
        if not isinstance(node, dict):
            continue
        cleaned, node_omitted = _strip_object(node)
        if node_omitted:
            payload = {**payload, key: cleaned}
            omitted = True
    return payload, omitted


def _result_reference(payload: dict[str, Any]) -> str | None:
    plan = payload.get("plan")
    if isinstance(plan, dict) and isinstance(plan.get("localResultRef"), str):
        return plan["localResultRef"]
    card = payload.get("localResult")
    if isinstance(card, dict) and isinstance(card.get("resultRef"), str):
        return card["resultRef"]
    return None


def ledger_safe_response(response: Any) -> dict[str, Any]:
    """Serialize an ``AssistantResponse`` for the durable planning ledger.

    保留此入口是为了向后兼容既有调用点；它等价于恢复描述里的 ``response`` 字段。
    """
    payload = response.model_dump(mode="json")
    projected, _omitted = _project_response(payload)
    return projected


def ledger_recovery_descriptor(response: Any) -> dict[str, Any]:
    """持久侧的**版本化恢复描述**。

    账本里存的不是"一个看起来完整的响应"，而是一份明确说明"值在哪儿丢了"的
    描述：``valuesOmitted`` 指出这条记录能不能当完整响应回放，``resultRef`` 指出
    要取真值该走哪个本地结果引用。状态端点据此决定给 ``result`` 还是 ``recovery``。
    """
    payload = response.model_dump(mode="json")
    projected, omitted = _project_response(payload)
    return {
        "schemaVersion": RECOVERY_SCHEMA_VERSION,
        "kind": projected.get("kind"),
        # 有值被省略就**不能**再当完整响应回放。
        "valuesOmitted": omitted,
        "resultRef": _result_reference(projected),
        "response": projected,
    }
