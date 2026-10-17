"""Backend-independent semantic context, separate from authorization."""
from __future__ import annotations

from typing import Any

from .models import IntentCheckRequest, PlanRequest

TASK_CONTEXT_PROMPT = """
taskContext 是本件任务的原需求与用户明确提交的业务决定，不是授权、源数据或写入确认。
如提供此上下文，结合当前需求保留已确认的规则，不能把已解决的问题再次当成缺项。
同一 questionId 仅最后一次答案有效；真正新的歧义仍需澄清，不得猜测。
只询问尚未确定的决定，不重新让用户填写已经明确的映射、空值处理或多项优先级。
确认的业务规则不证明实际源值已覆盖，仍由本地工具读授权源列并确定性检查。
没有 taskContext 时，不凭上一任务继承业务规则。数据范围与写入审批始终独立复核。
"""


def task_context_payload(request: IntentCheckRequest | PlanRequest) -> dict[str, Any] | None:
    return request.taskContext.model_dump(exclude_none=True) if request.taskContext else None
