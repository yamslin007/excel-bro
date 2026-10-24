"""「规划中断后可继续」的受控现场（服务端权威判据）。

与 ``packages/core/src/planningContext.ts`` 是**同一份契约**：字段白名单、禁用键、
上限与拒绝码一一对应。两端一致性由测试保证 —— ``server/tests/test_planning_context.py``
跑的是共享夹具 ``packages/core/src/planningContext.fixtures.json``，
TypeScript 侧 ``packages/core/src/planningContext.test.ts`` 跑的是同一批用例。

边界（不得在这里放宽）：

- 这不是 SDK 会话恢复，也不是"从模型断点原样续跑"。它只保存"按当前现场重新分析"
  所需的最小事实。
- ``checkpoint_json`` 非空不是 ``canContinue`` 的充分条件。版本、归属、阶段、继续资格、
  任务版本、值矩阵任一条不成立都拒绝。
- 旧循环检查点没有 ``kind``，因此在 ``kind`` 判别点就被拒绝，不会被当成 planning_context。
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Final

KIND: Final = "planning_context"
SCHEMA_VERSION: Final = 1

# 允许出现的顶层键。夹具的 base 记录必须正好是这些键 —— 这条在测试里钉住。
FIELDS: Final = (
    "kind",
    "schemaVersion",
    "taskId",
    "turnId",
    "executorId",
    "taskRevision",
    "stage",
    "resumability",
    "savedAt",
    "prompt",
    "planningGoal",
    "confirmations",
    "target",
    "authorization",
    "sourceFingerprint",
    "sourceFingerprintSheets",
    "resultRefs",
    "completedTools",
    "failedTool",
    "unknownTool",
)

REQUIRED_FIELDS: Final = (
    "kind",
    "schemaVersion",
    "taskId",
    "turnId",
    "executorId",
    "taskRevision",
    "stage",
    "resumability",
    "savedAt",
    "prompt",
    "confirmations",
    "target",
    "authorization",
    "sourceFingerprintSheets",
    "resultRefs",
    "completedTools",
)

# 值矩阵键名：按名字拦，不按动作类型枚举。与 ledger_privacy.MATRIX_FIELD_NAMES 同一口径，
# 另加 plan/actions 等"整份计划"字段——整份计划本身就含 values。
FORBIDDEN_KEYS: Final = frozenset(
    {
        "values",
        "rows",
        "headers",
        "formulas",
        "data",
        "preview",
        "matrix",
        "expected",
        "plan",
        "actions",
        "acceptanceCriteria",
        "snapshots",
        "before",
    }
)

MAX_PROMPT_CHARS: Final = 4000
MAX_PLANNING_GOAL_CHARS: Final = 400
MAX_CONFIRMATIONS: Final = 8
MAX_RESULT_REFS: Final = 12
MAX_COMPLETED_TOOLS: Final = 32
MAX_SHEET_NAMES: Final = 256
MAX_LABEL_CHARS: Final = 80
MAX_TOOL_NAME_CHARS: Final = 80
MAX_ERROR_CODE_CHARS: Final = 80

# 工作簿身份必须是既有命名空间。`name:` 兜底（未保存工作簿）不能启用断点续跑。
TARGET_IDENTITY_PATTERN: Final = re.compile(r"^(?:url:.+|wps-path:.+)$")

CODE_MISSING: Final = "PLANNING_CONTEXT_MISSING"
CODE_KIND_MISMATCH: Final = "PLANNING_CONTEXT_KIND_MISMATCH"
CODE_UNSUPPORTED_VERSION: Final = "PLANNING_CONTEXT_UNSUPPORTED_VERSION"
CODE_INVALID: Final = "PLANNING_CONTEXT_INVALID"
CODE_SENSITIVE: Final = "PLANNING_CONTEXT_SENSITIVE"
CODE_TASK_MISMATCH: Final = "PLANNING_CONTEXT_TASK_MISMATCH"
CODE_STAGE_UNSUPPORTED: Final = "PLANNING_CONTEXT_STAGE_UNSUPPORTED"
CODE_NOT_CONTINUABLE: Final = "PLANNING_CONTEXT_NOT_CONTINUABLE"
CODE_STALE_REVISION: Final = "PLANNING_CONTEXT_STALE_REVISION"

ALL_CODES: Final = (
    CODE_MISSING,
    CODE_KIND_MISMATCH,
    CODE_UNSUPPORTED_VERSION,
    CODE_INVALID,
    CODE_SENSITIVE,
    CODE_TASK_MISMATCH,
    CODE_STAGE_UNSUPPORTED,
    CODE_NOT_CONTINUABLE,
    CODE_STALE_REVISION,
)

# 服务端闸门码：**不属于检查点契约**（`ALL_CODES` 只列契约本身的 9 条，
# TS 侧 `PLANNING_CONTEXT_CODES` 与它一一对应）。这四条描述的是"检查点本身
# 合法，但当前不允许继续"的事实，只有服务端能判断（执行账本 + 客户端核实后
# 送上来的现场），因此不进共享枚举，客户端一律按字符串处理。
CODE_HAS_EXECUTION: Final = "PLANNING_CONTEXT_HAS_EXECUTION"
CODE_TARGET_CHANGED: Final = "PLANNING_CONTEXT_TARGET_CHANGED"
CODE_SCOPE_CHANGED: Final = "PLANNING_CONTEXT_SCOPE_CHANGED"
CODE_SOURCE_CHANGED: Final = "PLANNING_CONTEXT_SOURCE_CHANGED"
# 证据类：不是"现场变了"，而是"这一次继续**根本没核实过现场**"。
# 与上面四条必须分开 —— 前者说明现场对不上，后者说明我们手上没有现场的证据，
# 两种情况的用户动作完全不同（重新发起 vs 重新读取宿主）。
CODE_EVIDENCE_MISSING: Final = "PLANNING_CONTEXT_EVIDENCE_MISSING"
CODE_EVIDENCE_UNREADABLE: Final = "PLANNING_CONTEXT_EVIDENCE_UNREADABLE"
CODE_SOURCE_UNVERIFIED: Final = "PLANNING_CONTEXT_SOURCE_UNVERIFIED"

GATE_CODES: Final = (
    CODE_HAS_EXECUTION,
    CODE_TARGET_CHANGED,
    CODE_SCOPE_CHANGED,
    CODE_SOURCE_CHANGED,
    CODE_EVIDENCE_MISSING,
    CODE_EVIDENCE_UNREADABLE,
    CODE_SOURCE_UNVERIFIED,
)

MESSAGES: Final = {
    CODE_MISSING: "这个任务没有保存到足以继续的现场，需要重新发起需求。",
    CODE_KIND_MISMATCH: "这个任务保存的现场不是可继续的分析现场，需要重新发起需求。",
    CODE_UNSUPPORTED_VERSION: "这个任务保存的现场来自不受支持的版本，需要重新发起需求。",
    CODE_INVALID: "这个任务的现场不完整或已损坏，不能据此继续。",
    CODE_SENSITIVE: "这个任务的现场含有不应保存的内容，不能据此继续。",
    CODE_TASK_MISMATCH: "这个现场属于另一个任务，不能拿来继续当前任务。",
    CODE_STAGE_UNSUPPORTED: "这个任务停在暂不支持继续的阶段，请重新发起需求。",
    CODE_NOT_CONTINUABLE: "这个任务的结果还不确定，先核对原表状态，不能直接继续分析。",
    CODE_STALE_REVISION: "这个任务的现场版本与当前记录不一致，不能继续。",
    CODE_HAS_EXECUTION: (
        "这个任务已经登记过写入。先核对原表与执行状态，不能用重新分析绕过。"
    ),
    CODE_TARGET_CHANGED: (
        "当前工作簿与原任务的不是同一份，不能继续旧任务；请按当前工作簿重新发起需求。"
    ),
    CODE_SCOPE_CHANGED: (
        "原任务的数据范围与当前授权不一致；继续会沿用旧范围，请按当前范围重新发起需求。"
    ),
    CODE_SOURCE_CHANGED: (
        "原任务使用的数据已经变化，不能继续旧的规划现场；请按当前数据重新发起需求。"
    ),
    CODE_EVIDENCE_MISSING: (
        "继续分析需要先重新核对当前工作簿与数据范围；这次没有拿到核对结果，未继续。"
    ),
    CODE_EVIDENCE_UNREADABLE: (
        "读取当前工作簿现场时失败；没有核实过现场就不继续，请重新读取后再试。"
    ),
    CODE_SOURCE_UNVERIFIED: (
        "原任务依赖的数据还没有重新核对；先读取一次相关数据，再继续分析。"
    ),
}


@dataclass(frozen=True)
class Verdict:
    ok: bool
    record: dict[str, Any] | None = None
    code: str | None = None
    detail: str = ""

    def message(self) -> str:
        return MESSAGES.get(self.code or "", "这个任务暂时不能继续，需要重新发起需求。")


def source_verification_required(record: dict[str, Any] | None) -> bool:
    """这份检查点是否**依赖来源数据**，从而要求继续前重读并比对。

    只有三种事实说明"原任务读过数据"：记录了来源指纹、记录了被读的工作表、
    或已经产出了结果引用（结果只能来自读表或计算）。三者都为空时，
    `before_model` 的纯结构任务就是合法的 —— 此时要求重读来源属于**凭空加码**，
    会把"整理表头/规划结构"这类根本不碰数据的任务挡在门外。

    客户端按状态接口给出的 `sourceFingerprintSheets` 做同样的判断；服务端不采信
    客户端的说法，自己从检查点重新推一次。两边不一致时按"需要核实"处理。
    """
    if not isinstance(record, dict):
        return False
    fingerprint = record.get("sourceFingerprint")
    if isinstance(fingerprint, str) and fingerprint.strip():
        return True
    sheets = record.get("sourceFingerprintSheets")
    if isinstance(sheets, list) and any(
        isinstance(item, str) and item.strip() for item in sheets
    ):
        return True
    refs = record.get("resultRefs")
    if isinstance(refs, list) and refs:
        return True
    return False


def _reject(code: str, detail: str) -> Verdict:
    return Verdict(ok=False, code=code, detail=detail)


def _is_plain_object(value: Any) -> bool:
    return isinstance(value, dict)


def _is_non_empty_string(value: Any) -> bool:
    return isinstance(value, str) and bool(value.strip())


def _is_int(value: Any) -> bool:
    # bool 是 int 的子类：true/false 绝不能当版本号。
    return isinstance(value, int) and not isinstance(value, bool)


def find_forbidden_key(value: Any, depth: int = 0) -> str | None:
    """递归查找禁用键；返回命中的键名。"""
    if depth > 8:
        return None
    if isinstance(value, list):
        for item in value:
            hit = find_forbidden_key(item, depth + 1)
            if hit:
                return hit
        return None
    if not isinstance(value, dict):
        return None
    for key, item in value.items():
        if key in FORBIDDEN_KEYS:
            return str(key)
        hit = find_forbidden_key(item, depth + 1)
        if hit:
            return hit
    return None


def _string_list(value: Any, maximum: int, label: str) -> tuple[list[str] | None, str]:
    if not isinstance(value, list):
        return None, f"{label} 必须是数组"
    if len(value) > maximum:
        return None, f"{label} 超出上限"
    result: list[str] = []
    for item in value:
        if not _is_non_empty_string(item):
            return None, f"{label} 含空成员"
        result.append(item.strip())
    return result, ""


def validate(value: Any, *, turn_id: str, revision: int) -> Verdict:
    """校验一份检查点能否用作「继续分析」的现场。顺序与 TS 侧逐条对应。"""
    if value is None:
        return _reject(CODE_MISSING, "没有可用的规划现场")
    if isinstance(value, str):
        if not value.strip():
            return _reject(CODE_MISSING, "规划现场为空")
        return _reject(CODE_INVALID, "规划现场不是结构化记录")
    if not _is_plain_object(value):
        return _reject(CODE_INVALID, "规划现场不是结构化记录")
    if value.get("kind") != KIND:
        return _reject(
            CODE_KIND_MISMATCH,
            "这不是可继续的规划现场（可能来自旧版本或别的恢复通道）",
        )
    if value.get("schemaVersion") != SCHEMA_VERSION:
        return _reject(CODE_UNSUPPORTED_VERSION, "规划现场版本不受支持，不能继续")
    forbidden = find_forbidden_key(value)
    if forbidden:
        return _reject(CODE_SENSITIVE, f"规划现场含不应持久化的内容（{forbidden}）")
    for key in value:
        if key not in FIELDS:
            return _reject(CODE_INVALID, f"规划现场含未知字段（{key}）")
    for key in REQUIRED_FIELDS:
        if key not in value:
            return _reject(CODE_INVALID, f"规划现场缺少字段（{key}）")
    if not _is_non_empty_string(value.get("taskId")) or not _is_non_empty_string(
        value.get("turnId")
    ):
        return _reject(CODE_INVALID, "规划现场缺少任务身份")
    if value["taskId"] != value["turnId"]:
        return _reject(CODE_TASK_MISMATCH, "任务身份不一致")
    if value["taskId"] != turn_id:
        return _reject(CODE_TASK_MISMATCH, "规划现场属于另一个任务")
    if not _is_non_empty_string(value.get("executorId")):
        return _reject(CODE_INVALID, "规划现场缺少会话归属")
    task_revision = value.get("taskRevision")
    if not _is_int(task_revision) or task_revision < 1:
        return _reject(CODE_INVALID, "规划现场版本号无效")
    if task_revision > revision:
        return _reject(CODE_STALE_REVISION, "规划现场比当前任务更新，不能据此继续")
    if value.get("stage") not in {"before_model", "after_tool"}:
        return _reject(CODE_STAGE_UNSUPPORTED, "该中断阶段暂不支持继续")
    if value.get("resumability") != "continue":
        return _reject(CODE_NOT_CONTINUABLE, "该现场不足以继续，需要重新分析")
    if value.get("unknownTool") is not None:
        return _reject(CODE_NOT_CONTINUABLE, "存在结果未知的步骤，不能直接继续规划")
    if not _is_non_empty_string(value.get("savedAt")):
        return _reject(CODE_INVALID, "规划现场缺少保存时间")
    if not _is_non_empty_string(value.get("prompt")):
        return _reject(CODE_INVALID, "规划现场缺少原需求")
    if len(value["prompt"]) > MAX_PROMPT_CHARS:
        return _reject(CODE_INVALID, "原需求超出可继续长度")
    planning_goal = value.get("planningGoal")
    if planning_goal is not None and (
        not isinstance(planning_goal, str)
        or len(planning_goal) > MAX_PLANNING_GOAL_CHARS
    ):
        return _reject(CODE_INVALID, "规划目标无效")

    confirmations_raw = value.get("confirmations")
    if not isinstance(confirmations_raw, list):
        return _reject(CODE_INVALID, "澄清记录无效")
    if len(confirmations_raw) > MAX_CONFIRMATIONS:
        return _reject(CODE_INVALID, "澄清记录过多")
    confirmations: list[dict[str, str]] = []
    for item in confirmations_raw:
        if not _is_plain_object(item) or not _is_non_empty_string(item.get("questionId")):
            return _reject(CODE_INVALID, "澄清记录缺少问题标识")
        if not isinstance(item.get("answerLabel"), str):
            return _reject(CODE_INVALID, "澄清记录缺少回答")
        label = item.get("label")
        if label is not None and (not isinstance(label, str) or len(label) > MAX_LABEL_CHARS):
            return _reject(CODE_INVALID, "澄清记录字段名无效")
        confirmations.append(
            {
                "questionId": item["questionId"].strip(),
                "label": label,
                "answerLabel": item["answerLabel"],
            }
        )

    target = value.get("target")
    if not _is_plain_object(target):
        return _reject(CODE_INVALID, "缺少工作簿与授权范围身份")
    if not _is_non_empty_string(target.get("identity")):
        return _reject(CODE_INVALID, "缺少工作簿身份依据")
    if not TARGET_IDENTITY_PATTERN.match(str(target.get("identity"))):
        return _reject(CODE_INVALID, "工作簿身份不是可核对的命名空间身份")
    evidence_key = target.get("evidenceKey")
    if evidence_key is not None and not isinstance(evidence_key, str):
        return _reject(CODE_INVALID, "授权证据键类型无效")
    target_sheets, detail = _string_list(
        target.get("sheetNames"), MAX_SHEET_NAMES, "授权成员"
    )
    if target_sheets is None:
        return _reject(CODE_INVALID, detail)
    if not target_sheets:
        return _reject(CODE_INVALID, "授权成员为空")

    authorization = value.get("authorization")
    if not _is_plain_object(authorization):
        return _reject(CODE_INVALID, "缺少授权版本")
    auth_sheets, detail = _string_list(
        authorization.get("sheetNames"), MAX_SHEET_NAMES, "授权成员"
    )
    if auth_sheets is None:
        return _reject(CODE_INVALID, detail)

    fingerprint_sheets, detail = _string_list(
        value.get("sourceFingerprintSheets"), MAX_SHEET_NAMES, "来源表清单"
    )
    if fingerprint_sheets is None:
        return _reject(CODE_INVALID, detail)
    fingerprint = value.get("sourceFingerprint")
    if fingerprint is not None and not _is_non_empty_string(fingerprint):
        return _reject(CODE_INVALID, "来源指纹无效")
    if fingerprint is None and fingerprint_sheets:
        return _reject(CODE_INVALID, "有来源表却没有来源指纹")

    result_refs_raw = value.get("resultRefs")
    if not isinstance(result_refs_raw, list):
        return _reject(CODE_INVALID, "结果引用无效")
    if len(result_refs_raw) > MAX_RESULT_REFS:
        return _reject(CODE_INVALID, "结果引用过多")
    result_refs: list[dict[str, Any]] = []
    for item in result_refs_raw:
        if not _is_plain_object(item) or not isinstance(item.get("resultRef"), str):
            return _reject(CODE_INVALID, "结果引用缺少标识")
        ref = item["resultRef"]
        if not _valid_result_ref(ref):
            return _reject(CODE_INVALID, "结果引用格式无效")
        item_fingerprint = item.get("sourceFingerprint")
        if item_fingerprint is not None and not _is_non_empty_string(item_fingerprint):
            return _reject(CODE_INVALID, "结果引用的来源指纹无效")
        result_refs.append({"resultRef": ref, "sourceFingerprint": item_fingerprint})

    completed_raw = value.get("completedTools")
    if not isinstance(completed_raw, list):
        return _reject(CODE_INVALID, "已完成步骤无效")
    if len(completed_raw) > MAX_COMPLETED_TOOLS:
        return _reject(CODE_INVALID, "已完成步骤过多")
    completed_tools: list[dict[str, str]] = []
    for item in completed_raw:
        if not _is_plain_object(item) or not _is_non_empty_string(item.get("tool")):
            return _reject(CODE_INVALID, "已完成步骤缺少工具名")
        if len(item["tool"]) > MAX_TOOL_NAME_CHARS:
            return _reject(CODE_INVALID, "工具名过长")
        if item.get("status") != "completed":
            return _reject(CODE_INVALID, "已完成步骤状态无效")
        completed_tools.append({"tool": item["tool"], "status": "completed"})

    failed_tool: dict[str, str] | None = None
    failed_raw = value.get("failedTool")
    if failed_raw is not None:
        if not _is_plain_object(failed_raw) or not _is_non_empty_string(
            failed_raw.get("tool")
        ):
            return _reject(CODE_INVALID, "失败步骤无效")
        error_code = failed_raw.get("errorCode")
        if not isinstance(error_code, str) or len(error_code) > MAX_ERROR_CODE_CHARS:
            return _reject(CODE_INVALID, "失败原因无效")
        failed_tool = {"tool": failed_raw["tool"].strip(), "errorCode": error_code}

    record: dict[str, Any] = {
        "kind": KIND,
        "schemaVersion": SCHEMA_VERSION,
        "taskId": value["taskId"],
        "turnId": value["turnId"],
        "executorId": value["executorId"],
        "taskRevision": task_revision,
        "stage": value["stage"],
        "resumability": "continue",
        "savedAt": value["savedAt"],
        "prompt": value["prompt"],
        "planningGoal": planning_goal,
        "confirmations": confirmations,
        "target": {
            "identity": target["identity"],
            "evidenceKey": evidence_key,
            "sheetNames": target_sheets,
        },
        "authorization": {"sheetNames": auth_sheets},
        "sourceFingerprint": fingerprint,
        "sourceFingerprintSheets": fingerprint_sheets,
        "resultRefs": result_refs,
        "completedTools": completed_tools,
        "failedTool": failed_tool,
        "unknownTool": None,
    }
    return Verdict(ok=True, record=record)


def _valid_result_ref(value: str) -> bool:
    if not value.startswith("res_") or len(value) != 36:
        return False
    return all(char in "0123456789abcdef" for char in value[4:])


def build_context(
    *,
    turn_id: str,
    executor_id: str,
    task_revision: int,
    stage: str,
    prompt: str,
    planning_goal: str | None,
    confirmations: list[dict[str, str]],
    target_identity: str,
    target_evidence_key: str | None = None,
    authorized_sheet_names: list[str],
    source_fingerprint: str | None,
    source_fingerprint_sheets: list[str],
    result_refs: list[dict[str, Any]],
    completed_tools: list[str],
    failed_tool: dict[str, str] | None,
    saved_at: str,
    resumability: str = "continue",
) -> dict[str, Any]:
    """唯一的构造入口：所有写入检查点的地方都走这里，避免字段漂移。"""
    return {
        "kind": KIND,
        "schemaVersion": SCHEMA_VERSION,
        "taskId": turn_id,
        "turnId": turn_id,
        "executorId": executor_id,
        "taskRevision": max(1, int(task_revision)),
        "stage": stage,
        "resumability": resumability,
        "savedAt": saved_at,
        "prompt": prompt[:MAX_PROMPT_CHARS],
        "planningGoal": (
            planning_goal[:MAX_PLANNING_GOAL_CHARS] if isinstance(planning_goal, str) else None
        ),
        "confirmations": [
            {
                "questionId": str(item.get("questionId", ""))[:80],
                "label": (str(item["label"])[:MAX_LABEL_CHARS] if item.get("label") else None),
                "answerLabel": str(item.get("answerLabel", ""))[:400],
            }
            for item in confirmations[:MAX_CONFIRMATIONS]
            if str(item.get("questionId", "")).strip()
        ],
        "target": {
            "identity": target_identity,
            "evidenceKey": target_evidence_key or None,
            "sheetNames": list(authorized_sheet_names)[:MAX_SHEET_NAMES],
        },
        "authorization": {"sheetNames": list(authorized_sheet_names)[:MAX_SHEET_NAMES]},
        "sourceFingerprint": source_fingerprint,
        "sourceFingerprintSheets": list(source_fingerprint_sheets)[:MAX_SHEET_NAMES],
        "resultRefs": [
            {
                "resultRef": str(item.get("resultRef", "")),
                "sourceFingerprint": item.get("sourceFingerprint"),
            }
            for item in result_refs[:MAX_RESULT_REFS]
            if _valid_result_ref(str(item.get("resultRef", "")))
        ],
        "completedTools": [
            {"tool": str(name)[:MAX_TOOL_NAME_CHARS], "status": "completed"}
            for name in completed_tools[:MAX_COMPLETED_TOOLS]
            if str(name).strip()
        ],
        "failedTool": failed_tool,
        "unknownTool": None,
    }
