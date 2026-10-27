"""Shared local-result **submission semantics** for both model backends.

A model never receives data rows. It only ever holds an opaque result reference
(``res_<32 hex>``) that a previous tool receipt gave it. Turning that reference
into user-visible output happens here, once, so the two adapters do not fork the
business rules:

* the normal OpenAI-compatible agent loop (``excel_agent``), and
* the Claude Agent SDK adapter (``claude_sdk_agent``).

The adapters keep only what is genuinely backend-specific — registration, tool
dispatch and wire formatting (content blocks for the SDK, typed
``AssistantResponse`` objects for the normal loop). Everything that decides
*whether a reference may be honoured* lives in this module.

Boundaries kept here:

* **Ownership before usefulness.** ``resolve_result_ref`` resolves through the
  trusted registry scoped to the *current* ``task_id``. A reference that is
  unknown, expired, belongs to another task, or whose authorization has been
  replaced resolves to nothing and is refused — a guessing model cannot turn a
  handle string into data.
* **Display is not writing.** ``build_local_result_answer`` produces a card
  pointer plus a bounded message; the matrix stays in the registry and is fetched
  by the local renderer over the loopback result endpoint.
* **Submitting is not executing.** ``build_result_plan_response`` compiles a
  whitelist ``writeValues`` plan and returns it *pending user confirmation*.
  Nothing is written, no host setter is called, and the panel still owns the
  confirm step.
* **Displaying is not completing a write task.** When the model has recognised
  (and the user has authorised) a *write-preview* goal, showing a result card is
  only an intermediate step: only a compiled plan entering preview completes the
  planning objective. ``evaluate_local_result_submission`` is the single place
  that decides this, so both backends refuse an early ``answer`` identically.
* **Fixed error codes.** Refusals carry a small, closed set of codes; the
  message never echoes workbook values or raw exception text.
"""

from __future__ import annotations

from typing import Any, Literal

from . import dev_result_ref_fault
from .local_results import RegisteredResult, local_results
from .models import AnswerResponse, LocalResultCard, PlanRequest, PlanResponse
from .result_plan import ResultPlanError, compile_result_plan

# 引用不可解析：未知句柄、已过期、属于别的任务、或登记时的授权已被替换。
# 四种情形的**对外结论相同**，不区分具体原因 —— 区分会变成枚举注册表的探针。
RESULT_REF_INVALID_CODE = "result_ref_invalid"
RESULT_REF_INVALID_MESSAGE = "结果引用无效或已过期，请重新查询。"

# 目标区域无法编译（非 A1、越界、尺寸不符之外的结构性错误）。
INVALID_TARGET_CODE = "invalid_target"
INVALID_TARGET_MESSAGE = "目标区域无效，请给出与结果行列数一致的 A1 区域。"

DEFAULT_LOCAL_RESULT_MESSAGE = "已完成本地查询，结果见下方结果卡。"
# 模型自报的说明文字会进聊天流/账本，必须截断。
MAX_MESSAGE_CHARS = 1000


class ResultSubmissionError(ValueError):
    """A result-submission proposal that must be refused.

    ``code`` is one of the fixed codes above (or a :class:`ResultPlanError`
    code, which is itself a closed set); ``str(error)`` is a bounded message
    safe to hand back to the model.
    """

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


def resolve_result_ref(task_id: str, result_ref: Any) -> RegisteredResult:
    """Resolve a model-supplied reference for *this* task, or refuse.

    The registry checks ownership (``item.task_id != task_id``), TTL and the
    authorization revision the result was registered under; this function only
    turns "no such result" into the fixed refusal.

    开发验收入口（A1-E1）在这一层生效：被绑定的那个句柄**第一次**走到这里时立刻
    作废并从登记器里删除，然后按同一个固定码拒绝。放在这条漏斗上而不是某个后端
    适配器里，是为了让两个后端、结果卡渲染和计划编译看到同一个结论；删掉句柄则
    保证没有"缓存正文"或重试能把一个已过期的引用救回来。
    """
    handle = str(result_ref or "")
    if dev_result_ref_fault.take_expiry(task_id, handle):
        local_results.discard(handle)
        raise ResultSubmissionError(
            RESULT_REF_INVALID_CODE, RESULT_REF_INVALID_MESSAGE
        )
    item = local_results.resolve(task_id, handle)
    if item is None:
        raise ResultSubmissionError(
            RESULT_REF_INVALID_CODE, RESULT_REF_INVALID_MESSAGE
        )
    return item


def local_result_card(item: RegisteredResult) -> LocalResultCard:
    """Card pointer with a *fixed* label: never re-echoes free text.

    ``item.title`` / ``calculation`` / ``warnings`` may contain text derived from
    the source data, so they are not copied into the response DTO (which is
    persisted to the ledger); the local renderer fetches the full payload over
    the loopback result endpoint.
    """
    headers = [
        header if isinstance(header, str) else str(header) for header in item.headers
    ]
    return LocalResultCard(
        resultRef=item.handle,
        title="本地查询结果",
        headers=[header[:200] for header in headers][:50] or ["（空）"],
        rowCount=item.row_count,
        columnCount=max(1, item.column_count),
        sourceSheets=[str(sheet)[:100] for sheet in item.source_sheets][:100],
        complete=bool(item.complete),
    )


def build_local_result_answer(
    item: RegisteredResult, *, message: str = ""
) -> AnswerResponse:
    """Read-only terminal: show the registered values, write nothing."""
    text = str(message or "").strip()[:MAX_MESSAGE_CHARS]
    return AnswerResponse(
        provider="model",
        message=text or DEFAULT_LOCAL_RESULT_MESSAGE,
        localResult=local_result_card(item),
    )


# --------------------------------------------------------------------------
# Planning-goal completion semantics (shared by both backends)
# --------------------------------------------------------------------------

# 模型在提交本地结果时**必须声明**它理解的本轮规划目标。取值只有两个：
#
# * ``read_only``   —— 用户只要看到分析结果，本轮规划目标就是"展示结果"；
# * ``write_preview`` —— 用户要的是"生成写入预览"，展示结果只是中间成果。
#
# 这是**模型负责语义判断**的落点：服务端不按提示词关键词猜写入意图，而是要求
# 意图阶段模型显式声明目标，双端传递、服务端按 turnId 固定；终局声明不能降级。
# 未提供事前目标的旧调用方保留兼容语义。
PlanningGoal = Literal["read_only", "write_preview"]

# 声明了写入预览目标，却提前用「只读结果卡」收尾 —— 固定未完成码。
# 文案是**结论 + 下一步**，不含任何数据值，也不承诺必然成功。
WRITE_PREVIEW_INCOMPLETE_CODE = "write_preview_incomplete"
WRITE_PREVIEW_INCOMPLETE_MESSAGE = (
    "本轮规划目标尚未完成：本轮要求「生成写入预览」，但只提交了"
    "供阅读的结果卡，没有编译出待确认的写入计划。结果卡只是中间成果。"
    "若目标区域与范围已明确，请改用 submit_result_plan 提交同一结果引用与目标；"
    "若目标信息不足，请先向用户澄清；若本机工具或编译不支持，请如实说明"
    "「分析完成，写入预览未完成」，不要声称任务已成功。"
)


class LocalResultIncomplete(ValueError):
    """A *premature* local-result submission that does not complete the goal.

    Distinct from :class:`ResultSubmissionError`: this is not a refusal of an
    invalid proposal, it is a signal that the recognised goal is not yet met and
    the model may continue within the remaining budget.
    """

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


def normalize_planning_goal(value: Any) -> PlanningGoal:
    """Coerce the model-declared goal. Unknown/missing → ``read_only``.

    Defaulting *towards* ``read_only`` is deliberate: it never manufactures a
    write requirement the model did not recognise. A model that genuinely needs
    a write preview declares ``write_preview`` explicitly.
    """
    return "write_preview" if str(value or "").strip() == "write_preview" else "read_only"


def evaluate_local_result_submission(
    goal: PlanningGoal,
    *, bound_goal: PlanningGoal | None = None,
) -> None:
    """Decide whether a ``submit_local_result`` submission completes the goal.

    Raises :class:`LocalResultIncomplete` when the declared goal is a write
    preview — silently returning a card would mark a partially-done task as a
    full success. Read-only goals pass through untouched, so只读任务仍然直接返回
    结果，不会被强迫生成写入计划。
    """
    # A terminal tool may not downgrade the goal established by intent routing.
    if bound_goal == "write_preview" or goal == "write_preview":
        raise LocalResultIncomplete(
            WRITE_PREVIEW_INCOMPLETE_CODE, WRITE_PREVIEW_INCOMPLETE_MESSAGE
        )


def build_result_plan_response(
    item: RegisteredResult,
    request: PlanRequest,
    *,
    target: Any,
    include_header: bool = True,
) -> PlanResponse:
    """Compile the proposal into a whitelist plan that still awaits confirmation.

    All validation (source version, target authorization, range shape, host
    limits, formula text) stays in :func:`compile_result_plan`; refusals are
    re-raised as :class:`ResultSubmissionError` so both adapters see one code.
    """
    target = target if isinstance(target, dict) else {}
    try:
        plan = compile_result_plan(
            item,
            request,
            target_sheet=str(target.get("sheet", "")),
            target_range=str(target.get("range", "")),
            include_header=bool(include_header),
        )
    except ResultPlanError as error:
        # 固定错误码 + 本地校验文案（只含工作表名/区域等非敏感信息）。
        raise ResultSubmissionError(error.code, str(error)[:200]) from error
    except (ValueError, TypeError, KeyError) as error:
        raise ResultSubmissionError(
            INVALID_TARGET_CODE, INVALID_TARGET_MESSAGE
        ) from error
    return PlanResponse(provider="model", plan=plan)
