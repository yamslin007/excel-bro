"""开发验收专用的「结果引用过期」精确故障入口（A1-E1）。

为什么必须落在服务端：结果正文只在本机易失登记器（``local_results``）里，模型与
面板都只拿得到不透明句柄 ``res_<32 hex>``。要让"旧引用被拒"成为**真实读取路径**上
的事实（而不是 mock 或一句 UI 文案），故障只能发生在引用被解析的那一刻。

精确位置 —— "已登记、尚未首次写入"：

1. 用户在面板显式启用 → :func:`arm`（仅开发模式可用，单次，限时）；
2. 该工作簿的下一次规划任务打开工具桥 → :func:`bind_task` 把启用**绑定到这个
   ``task_id``**（绑定由可信的服务端代码做，请求里没有任务字段）；
3. 该任务**第一**条 Python 结果登记 → :func:`bind_python_result` 把故障**绑定到
   这个句柄**；
4. 该句柄第一次被 ``resolve_result_ref`` 使用 → :func:`take_expiry` 交回 True，
   调用方把它从登记器里删掉并给出固定拒绝码。计划编译不出来 ⇒ 任何宿主写入都
   不可能发生，也不存在"复用缓存正文"的余地（句柄已经不在了）。

刻意不做什么：

* 不改 ``agentToolBridge.resultTtlSeconds``，不缩短/延长生产结果寿命；
* 不清空全部结果、不整任务丢弃 —— 只动被绑定的那一个句柄；
* 不借用"第二步写入后中断"入口，也不碰 ``recoveryProbe`` 的状态；
* 不把任务或引用作为可传参数暴露出去（模型无从指定别的任务/别的句柄）。

开关：环境变量 ``EXCEL_BRO_RESULT_REF_FAULT=1``。默认关闭；关闭时本模块所有写操作
都是**空操作**，即使有调用方硬调也不会让任何引用失效。
"""

from __future__ import annotations

import os
import threading
import time
from dataclasses import dataclass

from . import dev_fault_target
from .capabilities import capability_int

ENV_FLAG = "EXCEL_BRO_RESULT_REF_FAULT"


class FaultDisabled(RuntimeError):
    """开发故障入口未开启（生产默认），或启用时无法唯一确定目标。"""


@dataclass(frozen=True)
class _Arm:
    host_id: str
    workbook_name: str
    #: 永远是**可核验**身份（见 `dev_fault_target.is_trusted_identity`）。
    workbook_identity: str
    #: `declared`（调用方显式给的）或 `observed`（服务端只见过这一个）。
    identity_source: str
    expires_at: float

    @property
    def identity(self) -> str:
        return self.workbook_identity


_lock = threading.RLock()
_arm: _Arm | None = None
_bound_task: str | None = None
_doomed: tuple[str, str] | None = None
_consumed = False


def _now() -> float:
    """Monotonic clock; tests replace this to exercise the window."""
    return time.monotonic()


def enabled() -> bool:
    return os.environ.get(ENV_FLAG, "").strip() == "1"


def ttl_seconds() -> int:
    return capability_int("devResultRefFault", "ttlSeconds")


def reset() -> None:
    global _arm, _bound_task, _doomed, _consumed
    with _lock:
        _arm = None
        _bound_task = None
        _doomed = None
        _consumed = False


def arm(*, host_id: str, workbook_name: str, workbook_identity: str) -> dict:
    """启用一次（面板按钮背后的动作）。未开启开发模式时直接拒绝。

    与 `dev_source_changed_fault` 同一身份边界：只给名字不算"唯一确定目标"，
    省略身份时必须由服务端"只见过这一个身份"来补全，否则拒绝。
    """
    global _arm, _bound_task, _doomed, _consumed
    if not enabled():
        raise FaultDisabled("结果引用过期入口未开启（需要开发模式）。")
    if not host_id or not workbook_name:
        raise FaultDisabled("缺少宿主或工作簿身份，无法绑定本次验收。")
    try:
        identity, identity_source = dev_fault_target.resolve_identity(
            host_id=host_id, workbook_name=workbook_name, declared=workbook_identity
        )
    except dev_fault_target.IdentityUnavailable as error:
        raise FaultDisabled(str(error)) from error
    with _lock:
        _arm = _Arm(
            host_id=host_id,
            workbook_name=workbook_name,
            workbook_identity=identity,
            identity_source=identity_source,
            expires_at=_now() + ttl_seconds(),
        )
        _bound_task = None
        _doomed = None
        _consumed = False
    return status()


def cancel() -> dict:
    reset()
    return status()


def bind_task(
    *, task_id: str, host_id: str, workbook_name: str, workbook_identity: str
) -> bool:
    """把启用绑定到**这一**个规划任务。返回是否真的绑定。

    判据是**可核验身份相等**，不是名字相等：另一本同名簿、重开后的另一个未保存
    会话都不会绑；任务侧只报了名字（`name:` 回退值）时也不绑 —— 那证明不了是
    同一本簿。
    """
    global _arm, _bound_task
    if not enabled() or not task_id:
        return False
    with _lock:
        arm_record = _arm
        if arm_record is None or _consumed or _bound_task is not None:
            # 已经绑给某一次任务：后面的任务不许把这把"单发"抢过去。
            return False
        if _now() >= arm_record.expires_at:
            return False
        target = (workbook_identity or "").strip()
        if not dev_fault_target.is_trusted_identity(target):
            return False
        if (
            arm_record.host_id != host_id
            or arm_record.workbook_name != workbook_name
            or arm_record.identity != target
        ):
            return False
        _bound_task = task_id
        return True


def bind_python_result(task_id: str, handle: str) -> bool:
    """把故障绑定到该任务**第一条** Python 结果的句柄上。"""
    global _doomed
    if not enabled() or not handle:
        return False
    with _lock:
        if _consumed or _bound_task is None or _bound_task != task_id:
            return False
        if _doomed is not None:
            # 单次生效：已经绑过一条，后续结果不再受影响。
            return False
        _doomed = (task_id, handle)
        return True


def take_expiry(task_id: str, handle: str) -> bool:
    """该引用此刻是否必须失效。返回 True 时调用方负责把它从登记器里删掉。

    只会对**被绑定的那一个** ``(task, handle)`` 返回一次 True；之后故障视为已消费，
    别的引用一律不受影响。
    """
    global _doomed, _consumed, _arm, _bound_task
    if not enabled():
        return False
    with _lock:
        if _doomed is None or _doomed != (task_id, handle):
            return False
        _doomed = None
        _arm = None
        _bound_task = None
        _consumed = True
        return True


def status() -> dict:
    """面板显示用的固定形状。不含句柄，避免把归属凭据泄进 UI。"""
    with _lock:
        if not enabled():
            phase = "off"
            remaining = 0
        elif _consumed:
            phase = "consumed"
            remaining = 0
        elif _arm is None:
            phase = "idle"
            remaining = 0
        elif _now() >= _arm.expires_at:
            phase = "expired"
            remaining = 0
        elif _doomed is not None or _bound_task is not None:
            phase = "bound"
            remaining = max(0, int(_arm.expires_at - _now()))
        else:
            phase = "armed"
            remaining = max(0, int(_arm.expires_at - _now()))
        return {
            "enabled": enabled(),
            "phase": phase,
            "expiresInSeconds": remaining,
            "taskBound": _bound_task is not None,
            "refBound": _doomed is not None,
            "identityBound": _arm is not None,
            "identitySource": _arm.identity_source if _arm is not None else "none",
        }
