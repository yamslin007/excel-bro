"""宿主 × 版本 × 动作 能力矩阵的跨语言共享夹具。

与 `packages/core/src/actionCapability.test.ts` 加载**同一份**
`config/action_capability_cases.json`。两侧各自断言同一批输入 → 期望输出，
因此"两端算的不是同一件事"会在任一侧单独变红时暴露，而不是靠人工对读源码。

（`test_safety_mirror.py` 用的是"抓源码文本"的办法 —— 那对清单有效，
对逐项判定的逻辑无效。）
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from server.app.action_capability import (
    ActionCapability,
    CapabilityMatrix,
    HostCapabilityMatrix,
    intersect_action_capabilities,
    is_action_allowed_by_keys,
    matches_version_prefix,
    parse_host_id,
    resolve_allowed_action_types,
    resolve_host_action_bound,
)

CASES_PATH = (
    Path(__file__).resolve().parents[2] / "config" / "action_capability_cases.json"
)


def _cases() -> dict:
    return json.loads(CASES_PATH.read_text(encoding="utf-8"))


def _matrix() -> CapabilityMatrix:
    """把夹具里的矩阵翻译成 Python 侧模型。

    夹具刻意用 JSON 的嵌套形态（`versions.<prefix>.<action>.state`），
    而不是 Python 特有的结构 —— 同一份文件 TS 侧直接读，两侧字面一致。
    """
    raw = _cases()["matrix"]
    hosts: dict[str, HostCapabilityMatrix] = {}
    for host_id, entry in raw.items():
        versions: dict[str, dict[str, ActionCapability]] = {}
        for prefix, actions in entry.get("versions", {}).items():
            versions[prefix] = {
                action: ActionCapability.model_validate(capability)
                for action, capability in actions.items()
            }
        hosts[host_id] = HostCapabilityMatrix(
            enabled=entry.get("enabled", True),
            versions=versions,
        )
    return CapabilityMatrix(hosts=hosts)


def test_version_prefix_cases() -> None:
    for case in _cases()["versionPrefixCases"]:
        assert (
            matches_version_prefix(case["version"], case["prefix"]) is case["expected"]
        ), f"{case['version']} vs prefix {case['prefix']}"


def test_host_id_cases() -> None:
    for case in _cases()["hostIdCases"]:
        assert parse_host_id(case["input"]) == case["expected"], case


def test_bound_cases() -> None:
    matrix = _matrix()
    for case in _cases()["boundCases"]:
        bound = resolve_host_action_bound(
            matrix, case["hostId"], case["hostVersion"]
        )
        assert sorted(bound.keys()) == sorted(case["expectedActions"]), case


def test_intersection_cases() -> None:
    matrix = _matrix()
    for case in _cases()["intersectionCases"]:
        allowed = resolve_allowed_action_types(
            matrix,
            case["hostId"],
            case["hostVersion"],
            case["claimed"],
        )
        assert sorted(allowed) == sorted(case["expectedAllowed"]), case["name"]


def test_action_key_cases() -> None:
    allowed = {"clearRange:contents", "writeValues"}
    for case in _cases()["actionKeyCases"]:
        assert (
            is_action_allowed_by_keys(
                allowed, case["actionType"], case["variant"]
            )
            is case["expected"]
        ), case


def test_unknown_host_is_empty_not_open() -> None:
    """红线：未知宿主不能变成"全部放开"。

    这是整层最重要的一条 —— 一旦这里反过来，构造过的请求就能让规划器
    生成任何动作，而所有下游检查都以为它已经过闸。
    """
    matrix = _matrix()
    for host in [None, "", "office-online", "wps", "Excel-Office-Web"]:
        assert resolve_host_action_bound(matrix, host, "12.0") == {}
        assert resolve_allowed_action_types(matrix, host, "12.0", ["writeValues"]) == set()


def test_claimed_set_can_never_expand_the_bound() -> None:
    """红线：客户端自报只能收窄。"""
    matrix = _matrix()
    bound = resolve_host_action_bound(matrix, "wps-et", "12.0")
    bound_supported = {
        key for key, capability in bound.items() if capability.state == "supported"
    }
    # 自报"全部动作都存在"时，结果仍然只是上界中 supported 的那部分
    everything = list(bound.keys()) + [
        "createPivotTable",
        "addImage",
        "deleteWorksheet",
        "setHyperlink",
    ]
    allowed = intersect_action_capabilities(bound, everything)
    assert allowed == bound_supported
    assert allowed <= set(bound.keys())


def test_client_claim_cannot_promote_unverified_or_unsupported() -> None:
    """红线：`unverified` / `unsupported` 不因自报而进入可规划集合。"""
    matrix = _matrix()
    bound = resolve_host_action_bound(matrix, "wps-et", "12.0")
    assert bound["writeFormulas"].state == "unverified"
    assert bound["createChart"].state == "unsupported"
    allowed = resolve_allowed_action_types(
        matrix, "wps-et", "12.0", ["writeFormulas", "createChart"]
    )
    assert allowed == set()


@pytest.mark.parametrize("claimed", [None, [], ()])
def test_empty_claim_yields_empty_not_open(claimed) -> None:
    """红线：客户端漏报得到空集，不是"没说就放开"。"""
    matrix = _matrix()
    assert resolve_allowed_action_types(matrix, "wps-et", "12.0", claimed) == set()


def test_missing_version_yields_empty_bound() -> None:
    matrix = _matrix()
    for version in [None, "", "   "]:
        assert resolve_host_action_bound(matrix, "wps-et", version) == {}


def test_disabled_host_yields_empty_bound() -> None:
    matrix = _matrix()
    assert resolve_host_action_bound(matrix, "folder", "1.0") == {}


def test_overlapping_prefixes_pick_longest_match() -> None:
    matrix = CapabilityMatrix(
        hosts={
            "wps-et": HostCapabilityMatrix(
                versions={
                    "12": {"writeValues": ActionCapability(state="unverified")},
                    "12.0": {"writeValues": ActionCapability(state="supported")},
                }
            )
        }
    )
    assert resolve_host_action_bound(matrix, "wps-et", "12.0")[
        "writeValues"
    ].state == "supported"


def test_clear_range_variant_does_not_open_siblings() -> None:
    """红线：`clearRange` 出现不等于全部变体放开。"""
    allowed = {"clearRange:contents"}
    assert is_action_allowed_by_keys(allowed, "clearRange", "contents")
    for variant in ("formats", "all", "hyperlinks"):
        assert not is_action_allowed_by_keys(allowed, "clearRange", variant)
    # 缺失变体信息时不按父动作名放开
    assert not is_action_allowed_by_keys({"clearRange"}, "clearRange", "contents")
