"""开发验收故障入口的**目标身份**判据（A1-E3 第三批）。

两个开发入口（`dev_source_changed_fault`、`dev_result_ref_fault`）都要回答同一个
问题：用户说的"这本簿"到底是哪一本？以前两者都允许只报宿主 + 工作簿名，于是
**同名不同簿**（两个文件夹里各一份 `测试.xlsx`，或两个都叫「工作簿1」的未保存簿）
落进同一个键：谁先开桥谁把这次单发拿走 —— 被抢的那一轮真的停在半路，而要验的
那一本一次都看不到停止。探针实证见
`outputs/verify-a1e3-round3/fault-identity-before.txt`。

本模块只做一件事：把"目标"收敛到**可核验的工作簿身份**。

* 可核验身份 = `url:` / `wps-path:` / `wps-session:` / `unsaved:` 前缀。
  `name:<名字>` 是"客户端没报身份"时的回退值，**它本身就是按名字认簿**，
  不能拿来当绑定依据；
* 启用时给了身份 ⇒ 用它（`declared`）；
* 没给 ⇒ 只接受"服务端见过、且这个名字下只有一个身份"的情形（`observed`），
  这是**证据**，不是猜测；
* 一个都没见过、或见过多个 ⇒ 拒绝启用并说明原因，绝不用"先到先得"补上。

刻意不做的事：
* 不猜路径、不按名字模糊匹配、不把"最近一次"当默认值；
* 不持久化（内存里有界环形表，最多 16 条），开关关闭时不记录任何东西；
* 不暴露任务 ID（这里只认工作簿，任务归属仍由开桥那一刻绑定）。
"""

from __future__ import annotations

import threading

#: 可核验的宿主工作簿身份前缀（与两端 `packages/core/src/locator.ts`、
#: WPS 的 `wps-session:` 约定一致）。
TRUSTED_IDENTITY_PREFIXES: tuple[str, ...] = (
    "url:",
    "wps-path:",
    "wps-session:",
    "unsaved:",
)

#: 服务端留存的目标数量上限（够人看清"到底是哪一本"，又不至于变成一本账）。
MAX_SEEN_TARGETS = 16


class IdentityUnavailable(RuntimeError):
    """无法把这次启用唯一地落到一本工作簿上。"""


_lock = threading.RLock()
_seen: list[tuple[str, str, str]] = []


def is_trusted_identity(identity: str | None) -> bool:
    """这个字符串能不能作为"这本簿"的依据。

    `name:` 前缀（以及空串、任意自造串）一律不算：那是按显示名认簿，正是本模块
    要消灭的东西。
    """
    value = (identity or "").strip()
    if not value:
        return False
    return value.startswith(TRUSTED_IDENTITY_PREFIXES)


def reset() -> None:
    """仅供测试与"关掉开发开关"时清空：不产生任何生产影响。"""
    with _lock:
        _seen.clear()


def note_target(*, host_id: str, workbook_name: str, workbook_identity: str) -> bool:
    """记下"服务端见到过这本簿"。

    调用点是可信路径（`TaskToolBridge.open`，规划请求已确定宿主 / 名字 / 身份）。
    只登记可核验身份；`name:` 回退值不登记 —— 它证明不了是哪一本。
    """
    host = (host_id or "").strip()
    name = (workbook_name or "").strip()
    identity = (workbook_identity or "").strip()
    if not host or not name or not is_trusted_identity(identity):
        return False
    entry = (host, name, identity)
    with _lock:
        if entry in _seen:
            # 已经在表里：挪到最前（"最近见过"），不增条数。
            _seen.remove(entry)
        _seen.insert(0, entry)
        del _seen[MAX_SEEN_TARGETS:]
    return True


def seen_targets(*, host_id: str = "", workbook_name: str = "") -> list[dict[str, str]]:
    """已经见过的目标（最近的在前）；可按宿主 / 名字过滤。

    给**人**看的：验收时要显式传身份，就从这里抄，而不是猜路径格式。
    """
    host = (host_id or "").strip()
    name = (workbook_name or "").strip()
    with _lock:
        items = list(_seen)
    return [
        {
            "hostId": item_host,
            "workbookName": item_name,
            "workbookIdentity": item_identity,
        }
        for item_host, item_name, item_identity in items
        if (not host or item_host == host) and (not name or item_name == name)
    ]


def _identities_for(host_id: str, workbook_name: str) -> list[str]:
    return [
        item["workbookIdentity"]
        for item in seen_targets(host_id=host_id, workbook_name=workbook_name)
    ]


def resolve_identity(
    *, host_id: str, workbook_name: str, declared: str = ""
) -> tuple[str, str]:
    """返回 ``(身份, 来源)``；来源是 ``declared`` 或 ``observed``。

    无法唯一确定目标时抛 :class:`IdentityUnavailable`，消息里说清缺什么、怎么办。
    """
    host = (host_id or "").strip()
    name = (workbook_name or "").strip()
    value = (declared or "").strip()
    if value:
        if not is_trusted_identity(value):
            raise IdentityUnavailable(
                "给定的工作簿身份不在可核验范围：需要 url: / wps-path: / "
                "wps-session: / unsaved: 前缀，不能退回只认名字。"
            )
        return value, "declared"
    seen = _identities_for(host, name)
    if not seen:
        raise IdentityUnavailable(
            "缺少可核验的工作簿身份：只给宿主与工作簿名无法区分同名的工作簿。"
            "请显式传 workbookIdentity；不知道值时，先在这本簿上发一条只读需求，"
            "再用 GET /api/dev/source-changed-fault 的 seenTargets 取。"
        )
    if len(seen) > 1:
        raise IdentityUnavailable(
            f"工作簿名「{name}」在服务端见过 {len(seen)} 个不同的身份，"
            "无法唯一确定这次启用针对哪一本；请显式传 workbookIdentity。"
        )
    return seen[0], "observed"
