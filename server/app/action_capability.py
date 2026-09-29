"""宿主 × 版本 × 动作 的可信能力上界与交集（方案 §4.2/§4.3）。

与 `packages/core/src/actionCapability.ts` 镜像；两侧一致性由
`config/action_capability_cases.json` 的共享夹具锁住，不靠人工对读源码。

为什么这一层存在：客户端自报的可用动作是**请求体里的一段字符串数组**，
不能当权威。服务端必须按可信配置先算出该宿主/版本的动作**上界**，
再与客户端集合求交集 —— 交集天然满足 `A ∩ B ⊆ A`，
所以"客户端只能收窄、不能扩大"是运算本身保证的，不需要额外的越权检查。

反向也要挡：客户端漏报（空集合）得到空集，不是"没说就放开"。
空集意味着没有动作可规划，规划器明确拒绝，而不是生成一份之后必被拒的计划。
"""

from __future__ import annotations

from typing import Iterable, Literal, Mapping

from pydantic import BaseModel, Field

HostId = Literal["excel-office", "wps-et", "folder"]

CapabilityState = Literal["supported", "unverified", "unsupported"]


class ActionCapability(BaseModel):
    """单个动作的可用状态。`unverified` 与 `unsupported` 都不进计划。"""

    state: CapabilityState
    reason: str | None = None


class HostCapabilityMatrix(BaseModel):
    """一个宿主的能力表。`versions` 的键是版本前缀。"""

    enabled: bool = True
    versions: dict[str, dict[str, ActionCapability]] = Field(default_factory=dict)


class CapabilityMatrix(BaseModel):
    """可信能力矩阵：宿主标识 → 该宿主的能力。未列出的宿主 = 未知宿主 = 空上界。"""

    hosts: dict[str, HostCapabilityMatrix] = Field(default_factory=dict)


_KNOWN_HOSTS = ("excel-office", "wps-et", "folder")


def parse_host_id(value: str | None) -> HostId | None:
    """解析宿主身份。三者都不满足时返回 `None`，交由调用方拒绝，不猜成任何一端。"""
    host = (value or "").strip().lower()
    if host in _KNOWN_HOSTS:
        return host  # type: ignore[return-value]
    return None


def matches_version_prefix(version: str, prefix: str) -> bool:
    """版本前缀匹配：相等或后接 `.`。

    **不用裸 `startswith`** —— 那会让 `"12.0"` 命中 `"12.01"`，
    把两个不同版本当成同一个。
    """
    return version == prefix or version.startswith(f"{prefix}.")


def resolve_host_action_bound(
    matrix: CapabilityMatrix,
    host_id: str | None,
    host_version: str | None,
) -> dict[str, ActionCapability]:
    """算出一个宿主/版本的**动作上界**。

    未知宿主、未知版本、宿主开关关闭都返回**空表**，不是"全部放开"。
    版本前缀可能互相重叠（例如同时列了 `"12"` 与 `"12.0"`），取**最长的**匹配，
    让更具体的条目覆盖更宽的条目；否则结果会依赖字典的键顺序。
    """
    host = parse_host_id(host_id)
    if host is None:
        return {}
    entry = matrix.hosts.get(host)
    if entry is None or entry.enabled is False:
        return {}
    version = (host_version or "").strip()
    if not version:
        return {}
    matched = sorted(
        (prefix for prefix in entry.versions if matches_version_prefix(version, prefix)),
        key=len,
        reverse=True,
    )
    if not matched:
        return {}
    return dict(entry.versions[matched[0]])


def intersect_action_capabilities(
    bound: Mapping[str, ActionCapability],
    claimed: Iterable[str] | None,
) -> set[str]:
    """服务端上界 ∩ 客户端自报集合。

    结果里**只保留上界中为 `supported` 的动作**。上界里的 `unverified` /
    `unsupported` 即便客户端自报为可用也不会进入结果 ——
    客户端没有资格把一个未验收的动作变成可规划的动作。
    """
    allowed: set[str] = set()
    if claimed is None:
        return allowed
    claimed_set = set(claimed)
    if not claimed_set:
        return allowed
    for action_type, capability in bound.items():
        if capability.state != "supported":
            continue
        if action_type not in claimed_set:
            continue
        allowed.add(action_type)
    return allowed


def resolve_allowed_action_types(
    matrix: CapabilityMatrix,
    host_id: str | None,
    host_version: str | None,
    claimed: Iterable[str] | None,
) -> set[str]:
    """一步到位：宿主/版本 → 上界 → 与自报集合求交。"""
    return intersect_action_capabilities(
        resolve_host_action_bound(matrix, host_id, host_version),
        claimed,
    )


def action_capability_key(action_type: str, variant: str | None = None) -> str:
    """动作键的规范化：`clearRange` 按 `applyTo` 分变体，键形如 `clearRange:contents`。"""
    if action_type == "clearRange" and variant:
        return f"{action_type}:{variant}"
    return action_type


def is_action_allowed_by_keys(
    allowed: set[str],
    action_type: str,
    variant: str | None = None,
) -> bool:
    """判定一个具体动作（含变体）是否在交集内。

    变体规则：**父动作在白名单里但没有对应变体条目时，视为未验收。**
    不能因为 `clearRange` 出现过就放开 `clearRange:formats` ——
    那正是"一个动作名笼统宣称全部支持"的失误（对齐指导 §2 末段）。
    """
    if action_type == "clearRange" and variant:
        return f"{action_type}:{variant}" in allowed
    return action_type in allowed


def matrix_from_config(payload: Mapping[str, object]) -> CapabilityMatrix:
    """从 `capabilities.json` 的形态构造矩阵。

    这是**给现有配置导出等价视图**的入口（方案 §4.2）：物理 schema 迁移单独验收，
    当前配置仍是 `wpsWrite.*` 形态，由此函数转成矩阵。
    """
    hosts: dict[str, HostCapabilityMatrix] = {}
    for host_id, raw in payload.items():
        if not isinstance(raw, Mapping):
            continue
        versions: dict[str, dict[str, ActionCapability]] = {}
        for prefix, actions in (raw.get("versions") or {}).items():
            if not isinstance(actions, Mapping):
                continue
            versions[str(prefix)] = {
                str(action): ActionCapability.model_validate(dict(capability))
                for action, capability in actions.items()
            }
        hosts[str(host_id)] = HostCapabilityMatrix(
            enabled=bool(raw.get("enabled", True)),
            versions=versions,
        )
    return CapabilityMatrix(hosts=hosts)
