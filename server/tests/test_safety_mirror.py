"""前后端安全黑名单的镜像一致性。

formulaSafety.ts 顶部就写着「前端镜像 server/app/safety.py」，但这个镜像关系全靠
人工维护：两边各有一份函数名黑名单和一组注入载体正则，唯一共享的只有
config/capabilities.json 里的开关。任何一次单边增删都会造成「前端拦、后端放」
或反之的漂移，而这种漂移不会让任何现有测试变红。

这里用最笨的办法把两侧的名单和正则源码抠出来逐条比对 —— 正是因为笨，所以下次
只改一边时它一定会响。
"""

from __future__ import annotations

import re
from pathlib import Path

from server.app import safety

FORMULA_SAFETY_TS = (
    Path(__file__).resolve().parents[2]
    / "apps"
    / "excel-addin"
    / "src"
    / "formulaSafety.ts"
)


def _ts_source() -> str:
    return FORMULA_SAFETY_TS.read_text(encoding="utf-8")


def _normalize_regex(source: str) -> str:
    """抹掉 JS 正则字面量里对 `/` 的强制转义，其余原样比较。

    `/^(?:\\\\|\\/\\/)/` 和 Python 的 `^(?:\\\\|//)` 是同一个正则，差异只因为 JS
    字面量必须转义分隔符。除此之外不做任何等价化 —— 语义差异必须暴露出来。
    """
    return source.replace("\\/", "/")


def test_dangerous_function_blacklists_match() -> None:
    ts_names = re.findall(r'\{\s*name:\s*"([A-Z]+)"', _ts_source())
    py_names = [name for name, _ in safety._DANGEROUS_FORMULA_PATTERNS]

    assert ts_names == py_names, (
        "危险公式函数黑名单前后端不一致：单边增删会让一侧拦、另一侧放。"
        f" 前端={ts_names} 后端={py_names}"
    )


def test_dangerous_function_patterns_match() -> None:
    ts_patterns = dict(
        re.findall(r'\{\s*name:\s*"([A-Z]+)",\s*pattern:\s*/(.+?)/i\s*\}', _ts_source())
    )
    py_patterns = {
        name: pattern.pattern for name, pattern in safety._DANGEROUS_FORMULA_PATTERNS
    }

    assert {
        name: _normalize_regex(source) for name, source in ts_patterns.items()
    } == py_patterns


def test_injection_vector_patterns_match() -> None:
    """DDE / UNC / 外部工作簿 / 超链接前缀这几条不在函数名黑名单里，最容易漏改。"""
    source = _ts_source()
    pairs = {
        "DDE_PATTERN": safety._DDE_PATTERN,
        "UNC_PREFIX_PATTERN": safety._UNC_PREFIX_PATTERN,
        "EXTERNAL_WORKBOOK_REF_PATTERN": safety._EXTERNAL_WORKBOOK_REF_PATTERN,
        "HYPERLINK_NETWORK_PREFIX_PATTERN": safety._HYPERLINK_NETWORK_PREFIX_PATTERN,
        "FILE_SCHEME_PATTERN": safety._FILE_SCHEME_PATTERN,
    }
    for name, compiled in pairs.items():
        match = re.search(rf"^const {name} = /(.+?)/i?;$", source, re.MULTILINE)
        assert match is not None, f"前端缺少 {name}"
        assert _normalize_regex(match.group(1)) == compiled.pattern, (
            f"{name} 前后端正则不一致"
        )


def test_hyperlink_switch_reads_the_shared_config() -> None:
    """两侧都必须从 config/capabilities.json 读开关，不能各自硬编码默认值。"""
    source = _ts_source()

    assert 'capabilities as { safety?: { allowHyperlink?: boolean } }' in source
    assert 'safety?.allowHyperlink === true' in source
    assert safety.hyperlink_allowed() is False
