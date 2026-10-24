"""`config/capabilities.json` 的 `hostCapabilities` 与 `ExcelAction` 清单的覆盖锁定。

这一条不是形式检查，它挡住的是**最容易静默扩大授权**的一种改动：
新增一个动作类型后，如果只加进协议而忘记在能力矩阵里登记，
"没登记"就会在某些实现下退化成"默认放开"。

本测试与 `apps/excel-addin/src/contracts.ts` 的 `ExcelAction` 成员逐一比对。
`contracts.ts` 是 TS 源码，Python 侧用文本抓取 —— 这与
`test_safety_mirror.py` 用的是同一种办法，对"清单"有效。
"""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]
CONTRACTS = ROOT / "apps" / "excel-addin" / "src" / "contracts.ts"
CAPABILITIES = ROOT / "config" / "capabilities.json"


def _declared_action_types() -> list[str]:
    """从 `contracts.ts` 抓出 `ExcelAction` 联合的所有 `type: "..."`。

    只扫 `export type ExcelAction =` 到下一个顶层 `export` 声明之间的块，
    不扫 `VerificationCriterion`（它也有 `type:` 字段，混进来会得到假失败）。
    """
    text = CONTRACTS.read_text(encoding="utf-8")
    start = text.index("export type ExcelAction =")
    after = text.index("\nexport ", start + 1)
    block = text[start:after]
    found: list[str] = []
    for match in re.finditer(r'type:\s*"([A-Za-z][A-Za-z0-9]*)"', block):
        if match.group(1) not in found:
            found.append(match.group(1))
    return found


def _matrix() -> dict:
    return json.loads(CAPABILITIES.read_text(encoding="utf-8"))["hostCapabilities"]


def test_contract_declares_37_actions() -> None:
    """数字是当前事实，不是上限。改动协议时应同步更新本断言并在记录里说明。"""
    assert len(_declared_action_types()) == 37


@pytest.mark.parametrize("host_id", ["excel-office", "wps-et", "folder"])
def test_every_action_is_registered_for_host(host_id: str) -> None:
    """每个宿主、每个版本都必须**显式**登记全部动作。

    漏登记等于让该动作在该宿主上的状态取决于实现细节 ——
    这正是"未知即放开"的入口。
    """
    declared = set(_declared_action_types())
    matrix = _matrix()[host_id]
    for version, actions in matrix["versions"].items():
        registered = {key.split(":", 1)[0] for key in actions}
        assert declared == registered, (
            f"{host_id}@{version} 未覆盖：缺少 {sorted(declared - registered)}，"
            f"多出 {sorted(registered - declared)}"
        )


def test_matrix_has_no_unknown_action() -> None:
    declared = set(_declared_action_types())
    for host_id, entry in _matrix().items():
        for version, actions in entry["versions"].items():
            for key, capability in actions.items():
                base = key.split(":", 1)[0]
                assert base in declared, f"{host_id}@{version} 登记了未知动作「{key}」"
                assert capability["state"] in {"supported", "unverified", "unsupported"}
                if capability["state"] != "supported":
                    # unverified / unsupported 必须带原因：两者都不进计划，
                    # 区别只在给用户的说明，没原因的说明等于没有说明。
                    assert capability.get("reason"), (
                        f"{host_id}@{version} 的「{key}」状态为 {capability['state']} 但未说明原因"
                    )


def test_wps_bound_does_not_exceed_the_write_allowlist() -> None:
    """WPS 上界必须等于范围执行器与固定原生检查点执行器的并集。

    两处都是"WPS 上允许写什么"的真值来源，一旦不一致，
    规划前交集与执行预检就会各说各话：用户在预览里看到能执行，
    点确认却被执行器拒绝（或更糟，反过来）。
    """
    config = json.loads(CAPABILITIES.read_text(encoding="utf-8"))
    range_actions = set(config["wpsWrite"]["allowedActions"])
    checkpoint_actions = set(config["worksheetCheckpoint"]["allowedActions"])
    # 检查点白名单里同时有两类东西：计划动作（`deleteWorksheet`）与事后操作
    # （`restoreWorksheet`，原地撤销本次删除）。后者不会出现在任何计划里，
    # 因此不参与"宿主能力矩阵的动作上界"这条比对。
    assert checkpoint_actions == {"deleteWorksheet", "restoreWorksheet"}
    assert "restoreWorksheet" not in range_actions
    assert "restoreWorksheet" not in set(_declared_action_types()), (
        "原地撤销是执行后的恢复操作，不是计划动作；把它塞进 ExcelAction "
        "会让模型有机会在计划里提议一次撤销。"
    )
    assert not range_actions & checkpoint_actions
    allowlist = range_actions | (checkpoint_actions & set(_declared_action_types()))
    bound = {
        key
        for key, capability in config["hostCapabilities"]["wps-et"]["versions"]["12.0"].items()
        if capability["state"] == "supported"
    }
    assert bound == allowlist, (
        f"能力上界与写入白名单不一致：仅在上界 {sorted(bound - allowlist)}；"
        f"仅在白名单 {sorted(allowlist - bound)}"
    )


def test_never_verified_wps_version_prefixes_stay_empty() -> None:
    """未取得可靠产品版本信号前，正式版本名单必须为空。

    这不是"忘了填"，是一条刻意保持的状态：`Application.Version="12.0"`
    是开发调试通道的临时放行，不能变成生产白名单。
    """
    config = json.loads(CAPABILITIES.read_text(encoding="utf-8"))
    assert config["wpsWrite"]["verifiedVersionPrefixes"] == []
    assert config["wpsWrite"]["developmentVersionPrefixes"] == ["12.0"]
