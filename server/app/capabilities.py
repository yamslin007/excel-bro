from __future__ import annotations

import json
import math
import os
from functools import lru_cache
from pathlib import Path
from typing import Any


CAPABILITIES_PATH = Path(
    os.getenv(
        "EXCEL_BRO_CAPABILITIES_PATH",
        str(Path(__file__).resolve().parents[2] / "config" / "capabilities.json"),
    )
)


@lru_cache(maxsize=1)
def capabilities() -> dict[str, Any]:
    return json.loads(CAPABILITIES_PATH.read_text(encoding="utf-8"))


def capability_int(section: str, key: str) -> int:
    value = capabilities()[section][key]
    # bool 是 int 子类：JSON 的 true/false 会被 isinstance(int) 误判为合法整数配额，
    # 必须显式剔除，否则 True 会冒充 1、把限额静默变成"几乎无限制"。
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise RuntimeError(f"能力配置 {section}.{key} 必须是正整数")
    return value


def capability_float(section: str, key: str) -> float:
    value = capabilities()[section][key]
    if isinstance(value, bool) or not isinstance(value, (int, float)) or value < 0:
        raise RuntimeError(f"能力配置 {section}.{key} 必须是非负数值")
    # NaN/Infinity 会绕过一切上限比较（NaN 与任何值比较都为 False），必须拒绝，
    # 否则限额可被静默放大为"无限"。非有限数一律视为非法配置。
    if not math.isfinite(value):
        raise RuntimeError(f"能力配置 {section}.{key} 必须是有限数值")
    return float(value)


def capability_optional_int(section: str, key: str) -> int | None:
    """null uses the provider default; explicit caps must be positive integers.

    Missing/invalid configuration is an error, not an accidental unlimited setting.
    """
    value = capabilities()[section][key]
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise RuntimeError(f"能力配置 {section}.{key} 必须是 null 或正整数")
    return value


def capability_text(section: str, key: str) -> str:
    value = capabilities()[section][key]
    if not isinstance(value, str) or not value:
        raise RuntimeError(f"能力配置 {section}.{key} 必须是非空字符串")
    return value


def capability_bool(section: str, key: str, default: bool = False) -> bool:
    """读取布尔开关。缺失时用默认值（安全默认=False），存在但非 bool 时报错。"""
    value = capabilities().get(section, {}).get(key)
    if value is None:
        return default
    if not isinstance(value, bool):
        raise RuntimeError(f"能力配置 {section}.{key} 必须是布尔值")
    return value


def model_timeout_seconds() -> float:
    """模型请求超时。环境变量 AI_TIMEOUT_SECONDS 优先，否则用 llm.timeoutSeconds。

    统一从这里读取，避免 planner / intent 各自硬编码默认值（架构原则 #10）。
    环境变量为空或非法时回退到配置默认，绝不让部署环境的坏值直接中断请求。
    """
    override = os.getenv("AI_TIMEOUT_SECONDS")
    if override:
        try:
            parsed = float(override)
            if parsed > 0:
                return parsed
        except ValueError:
            pass
    return capability_float("llm", "timeoutSeconds")
