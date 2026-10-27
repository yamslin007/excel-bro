"""来源变化 = 本次规划的**终态**：一个判据、一句话、一处定义。

为什么单独一个模块：``source_changed`` 有两处生产产生点（桥的 :meth:`deliver`
在宿主回执里认出源指纹不符；桥在同一任务里自行比对绑定指纹），也有两处消费点
（工具桥必须停止再派发给宿主；SDK 规划循环必须停下来）。如果每处各写一句
"来源变了"，四处就会各自漂移 —— 用户看到的话、模型收到的码、账本记的码
必须来自同一处。

**这不是"把所有工具错误都改成终止"**：

- ``source_changed`` —— 本次规划的**前提**没了。旧来源的读取、计算、提交都不能
  继续；继续下去只可能产出一份"看起来成功"的旧结果。**终态**。
- ``result_ref_invalid`` —— 前提仍在（来源没变、授权没变），只是这一次的结果引用
  过期了。模型**应当**重新读算再提交（A1-E2 已用真实模型验证过这条路）。
- ``timeout`` / ``busy`` / 暂时网络错误 —— 现场没变，重试是正当的。

把后两类也做成终态，等于废掉"重算一次就能成功"的恢复能力。
"""

from __future__ import annotations

from typing import Any

# 桥回执里给模型的固定码（与前端共享判据 `packages/core/src/sdkToolBridge.ts` 同名）。
SOURCE_CHANGED_RECEIPT_CODE = "source_changed"

# 轮次终态码：账本、SSE 错误帧与接口错误体都用它。
SOURCE_CHANGED_TURN_CODE = "SOURCE_CHANGED"

# 用户可见文案。**唯一一份**：服务端返回、两端面板原样显示，不在前端再写一遍。
SOURCE_CHANGED_STOP_MESSAGE = (
    "来源已变化，当前任务已停止：没有继续读取来源、没有继续计算，也没有写入。"
    "请按最新数据重新发起需求。"
)


class SourceChangedDuringPlanning(RuntimeError):
    """本次规划因来源变化而终止（由生产守卫确认，不是模型判断）。

    继承 ``RuntimeError`` 而**不是** ``LLMResponseError``：这不是模型/网络失败，
    不能被映射成"连接错误、可重试"——那会让用户以为再点一次就能好，而实际上
    现场仍然是变过的那一版数据。
    """

    code = SOURCE_CHANGED_TURN_CODE

    def __init__(self, message: str = SOURCE_CHANGED_STOP_MESSAGE) -> None:
        super().__init__(message)


def is_source_changed_receipt(receipt: Any) -> bool:
    """桥回执 / 工具回执是不是"来源已变化"这条固定码。"""
    return (
        isinstance(receipt, dict)
        and receipt.get("errorCode") == SOURCE_CHANGED_RECEIPT_CODE
    )
