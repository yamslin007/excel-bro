from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path
from typing import Any


CAPABILITIES_PATH = (
    Path(__file__).resolve().parents[2] / "config" / "capabilities.json"
)


@lru_cache(maxsize=1)
def capabilities() -> dict[str, Any]:
    return json.loads(CAPABILITIES_PATH.read_text(encoding="utf-8"))


def capability_int(section: str, key: str) -> int:
    value = capabilities()[section][key]
    if not isinstance(value, int) or value <= 0:
        raise RuntimeError(f"能力配置 {section}.{key} 必须是正整数")
    return value


def capability_text(section: str, key: str) -> str:
    value = capabilities()[section][key]
    if not isinstance(value, str) or not value:
        raise RuntimeError(f"能力配置 {section}.{key} 必须是非空字符串")
    return value
