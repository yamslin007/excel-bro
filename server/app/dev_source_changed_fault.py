"""开发验收专用的「来源变化」精确注入（A1-E3）。

为什么必须有这个入口：``source_changed`` 在生产上只有两个**真实**产生点 ——
(1) 宿主回执里的源指纹与绑定指纹不符；(2) 桥在同一任务里自行比对绑定指纹。
两者都要求**模型在来源变化之后又读了一次**. 真实宿主里没人能保证模型会再读一次，
于是"来源变化 → 本次规划终止"这条行为在原生环境里只能碰运气触发 —— 靠运气不叫验收。

做法与 ``dev_result_ref_fault`` 是同一套纪律（2026-10-25 第二批收紧）：

1. 仅开发模式可开（``EXCEL_BRO_SOURCE_CHANGED_FAULT=1``）；开关关闭时下面所有写
   操作都是**空操作**，即使调用方硬调也不会改变任何行为；
2. 启用必须**说清是哪本簿**：宿主 + 工作簿名 + **可核验的工作簿身份**。只给名字
   不算"唯一确定目标"—— 两个文件夹里各叫 `测试.xlsx` 的两本簿、两个都叫「工作簿1」
   的未保存簿，在"宿主 + 名字"上完全一样；谁先开桥谁把它拿走，被抢的那一轮莫名
   停住，而真正要验的簿一次都看不到停止。身份可以显式给（``declared``），也可以
   由服务端"只见过这一个身份"补全（``observed``，有证据才补）；一个都没见过、
   或同一个名字见过多个身份 ⇒ **拒绝启用**并说明原因；
3. 任务归属由**服务端在开桥那一刻**绑定（``bind_task``）：规划请求已经确定了
   taskId / 宿主 / 工作簿，归属不交给"谁先说话"。此后只有**那一次任务**的调用
   能命中，另一本同样匹配的簿、同一个簿的另一个任务都抢不走；任务侧只报了名字
   （旧客户端，身份是 ``name:`` 回退值）时**不绑** —— 那证明不了是同一本簿；
4. 单次 + 限时（默认 120 秒），过期自动作废；可显式取消；
5. 命中点是桥的**下一次**工具调用：直接产出生产那份固定回执 ``source_changed``，
   **不派发给宿主**（所以也不会读表），随后走生产停止链路（交回模型 → 停桥）。

刻意**不做**的事：

* 不伪造一次"读表成功"，不改来源指纹，不放宽任何复核；
* 不清空结果、不改预算、不碰恢复检查点；
* 开关关闭时与生产行为逐字节一致（测试 ``test_dev_fault_is_off_by_default`` 钉住）。

**注入能证明什么、不能证明什么**：它证明"守卫一旦给出 ``source_changed``，这条
响应链真的停、旧入口真的消失、宿主写入为 0"；它**不能**代替"真实来源变化被检测
出来"（那要求模型在改数之后再读一次，见 ``docs/SDK_A1_E2_RECOVERY_CHAIN_2026-10-25.md``
与 A1-E3 原生验收包 §4）。

用法（原生验收）：启动后端前设环境变量，然后在发起需求**之前**调用
``POST /api/dev/source-changed-fault/arm``，Body 必须带宿主与工作簿名。
"""

from __future__ import annotations

import os
import threading
import time
from dataclasses import dataclass

from . import dev_fault_target

ENV_FLAG = "EXCEL_BRO_SOURCE_CHANGED_FAULT"
DEFAULT_TTL_SECONDS = 120


class FaultDisabled(RuntimeError):
    """开发故障入口未开启（生产默认），或启用时缺少必要绑定。"""


@dataclass(frozen=True)
class _Arm:
    host_id: str
    workbook_name: str
    #: 永远是**可核验**身份（`url:` / `wps-path:` / `wps-session:` / `unsaved:`）。
    workbook_identity: str
    #: `declared`（调用方显式给的）或 `observed`（服务端只见过这一个）。
    identity_source: str
    expires_at: float


_lock = threading.RLock()
_arm_state: _Arm | None = None
_bound_task: str | None = None
_consumed = False


def enabled() -> bool:
    return os.getenv(ENV_FLAG, "").strip() == "1"


def arm(
    *,
    host_id: str = "",
    workbook_name: str = "",
    workbook_identity: str = "",
    ttl_seconds: int = DEFAULT_TTL_SECONDS,
) -> dict[str, object]:
    """启用一次，绑定到**这一本**工作簿。

    ``host_id`` / ``workbook_name`` 必填；``workbook_identity`` 可省略，但省略时
    服务端必须"只见过这一个名字下的一个身份"才肯补全（``observed``）。否则抛
    :class:`FaultDisabled` 并说清原因 —— 不靠先到先得猜目标。
    """

    if not enabled():
        raise FaultDisabled("当前服务没有开启来源变化验收入口")
    host = host_id.strip()
    name = workbook_name.strip()
    if not host or not name:
        raise FaultDisabled("缺少宿主或工作簿名，无法把这次启用绑定到具体的簿")
    try:
        identity, identity_source = dev_fault_target.resolve_identity(
            host_id=host, workbook_name=name, declared=workbook_identity
        )
    except dev_fault_target.IdentityUnavailable as error:
        raise FaultDisabled(str(error)) from error
    with _lock:
        global _arm_state, _bound_task, _consumed
        _arm_state = _Arm(
            host_id=host,
            workbook_name=name,
            workbook_identity=identity,
            identity_source=identity_source,
            expires_at=time.monotonic() + max(1, int(ttl_seconds)),
        )
        _bound_task = None
        _consumed = False
    return status()


def cancel() -> dict[str, object]:
    if not enabled():
        raise FaultDisabled("当前服务没有开启来源变化验收入口")
    with _lock:
        global _arm_state, _bound_task
        _arm_state = None
        _bound_task = None
    return status()


def reset() -> None:
    """仅供测试：清空启用状态（**不看开关**，也不产生任何生产影响）。"""

    global _arm_state, _bound_task, _consumed
    with _lock:
        _arm_state = None
        _bound_task = None
        _consumed = False


def status() -> dict[str, object]:
    with _lock:
        if not enabled():
            return {
                "enabled": False,
                "phase": "off",
                "expiresInSeconds": 0,
                "hostBound": False,
                "workbookBound": False,
                "taskBound": False,
                "identityBound": False,
                "identitySource": "none",
                "seenTargets": [],
            }
        seen = dev_fault_target.seen_targets()
        task_bound = _bound_task is not None
        if _arm_state is None:
            return {
                "enabled": True,
                "phase": "consumed" if _consumed else "idle",
                "expiresInSeconds": 0,
                "hostBound": False,
                "workbookBound": False,
                "taskBound": False,
                "identityBound": False,
                "identitySource": "none",
                "seenTargets": seen,
            }
        remaining = _arm_state.expires_at - time.monotonic()
        if remaining <= 0:
            return {
                "enabled": True,
                "phase": "expired",
                "expiresInSeconds": 0,
                "hostBound": True,
                "workbookBound": True,
                "taskBound": task_bound,
                "identityBound": True,
                "identitySource": _arm_state.identity_source,
                "seenTargets": seen,
            }
        return {
            "enabled": True,
            "phase": "bound" if task_bound else "armed",
            "expiresInSeconds": int(remaining),
            "hostBound": True,
            "workbookBound": True,
            "taskBound": task_bound,
            "identityBound": True,
            "identitySource": _arm_state.identity_source,
            "seenTargets": seen,
        }


def bind_task(
    *, task_id: str, host_id: str, workbook_name: str, workbook_identity: str
) -> bool:
    """把启用绑定到**这一**个规划任务，返回是否真的绑定。

    调用点是可信路径（``TaskToolBridge.open``）：那一刻 taskId / 宿主 / 工作簿
    都已经由服务端确定，用户与模型都无从指定"注入给哪一次任务"。

    判据是**可核验身份相等**，不是名字相等：另一本同名簿、重开后的另一个未保存
    会话，都会因为身份不同而不绑。任务侧只报了名字（``name:`` 回退值）时不绑 ——
    那时连"是不是同一本簿"都证明不了，按名字绑就是张冠李戴。
    """

    global _bound_task
    if not enabled() or not task_id:
        return False
    with _lock:
        armed = _arm_state
        if armed is None or _consumed or _bound_task is not None:
            # 已经绑给某一次任务：后面的任务不许把这把"单发"抢过去。
            return False
        if time.monotonic() >= armed.expires_at:
            return False
        if armed.host_id != (host_id or "").strip():
            return False
        if armed.workbook_name != (workbook_name or "").strip():
            return False
        target = (workbook_identity or "").strip()
        if not dev_fault_target.is_trusted_identity(target):
            return False
        if target != armed.workbook_identity:
            return False
        _bound_task = task_id
        return True


def take(task_id: str) -> bool:
    """命中一次就交回 True 并作废；开关关闭或不是被绑定的那次任务时恒 False。

    只认 ``bind_task`` 绑过的那一个 taskId：另一本簿、同一个簿的另一个任务拿到
    的都是"什么都没发生"（各自照常派给宿主），不会被别人的验收停住。
    """

    global _arm_state, _bound_task, _consumed
    if not enabled():
        return False
    with _lock:
        if _arm_state is None or _bound_task is None or _bound_task != task_id:
            return False
        if time.monotonic() >= _arm_state.expires_at:
            _arm_state = None
            _bound_task = None
            return False
        _arm_state = None
        _bound_task = None
        _consumed = True
        return True
