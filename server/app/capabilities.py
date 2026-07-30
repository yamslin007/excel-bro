from __future__ import annotations

import json
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
    if not isinstance(value, int) or value <= 0:
        raise RuntimeError(f"能力配置 {section}.{key} 必须是正整数")
    return value


def capability_float(section: str, key: str) -> float:
    value = capabilities()[section][key]
    if isinstance(value, bool) or not isinstance(value, (int, float)) or value < 0:
        raise RuntimeError(f"能力配置 {section}.{key} 必须是非负数值")
    return float(value)


def capability_text(section: str, key: str) -> str:
    value = capabilities()[section][key]
    if not isinstance(value, str) or not value:
        raise RuntimeError(f"能力配置 {section}.{key} 必须是非空字符串")
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
