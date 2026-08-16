"""
本地 AI 服务的安全护栏：拦截 Excel 公式注入与超限内嵌内容。

服务端所有会落到工作簿的 action 都经过 pydantic 校验器（models.py），
AI 生成的公式在 rule_generator.py 再兜一道。这里只提供纯函数判定，
便于前后端测试与复用。
"""

from __future__ import annotations

import re

from .capabilities import capability_bool, capability_int


# ---------------------------------------------------------------------------
# 危险公式函数
# ---------------------------------------------------------------------------
# 这些函数不只是算数：WEBSERVICE/IMPORT*/FILTERXML 会向外部 URL 发请求
# （可把单元格内容编码后外传）；HYPERLINK 是钓鱼载体；EXEC/CALL/REGISTER
# 是旧式可执行宏；=cmd| 是 DDE 命令注入。
_DANGEROUS_FORMULA_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("WEBSERVICE", re.compile(r"\bWEBSERVICE\s*\(", re.IGNORECASE)),
    ("HYPERLINK", re.compile(r"\bHYPERLINK\s*\(", re.IGNORECASE)),
    ("IMPORTXML", re.compile(r"\bIMPORTXML\s*\(", re.IGNORECASE)),
    ("IMPORTDATA", re.compile(r"\bIMPORTDATA\s*\(", re.IGNORECASE)),
    ("IMPORTHTML", re.compile(r"\bIMPORTHTML\s*\(", re.IGNORECASE)),
    ("FILTERXML", re.compile(r"\bFILTERXML\s*\(", re.IGNORECASE)),
    ("EXEC", re.compile(r"\bEXEC\s*\(", re.IGNORECASE)),
    ("CALL", re.compile(r"\bCALL\s*\(", re.IGNORECASE)),
    ("REGISTER", re.compile(r"\bREGISTER\s*\(", re.IGNORECASE)),
)

# DDE：=cmd|'/c calc'!A0 这类老式命令注入，前缀匹配即可。
_DDE_PATTERN = re.compile(r"^\s*=\s*cmd\s*[|'`]", re.IGNORECASE)

# 外部工作簿 / UNC 引用：函数名黑名单查不到这类写法，但打开工作簿或点击超链接时
# 会让 Excel/Windows 对 \\host\share 发起 SMB 连接，泄露 NTLM 哈希（经典钓鱼手法）。
#  `\\` 是 UNC 路径前缀；[file.xlsx] 是外部工作簿引用语法。
_UNC_PREFIX_PATTERN = re.compile(r"\\\\")
_EXTERNAL_WORKBOOK_REF_PATTERN = re.compile(
    r"\[[^\]\s]+\.(?:xlsx|xlsm|xlsb|xls|csv|xlw)\]", re.IGNORECASE
)
# 超链接地址的共享路径前缀（Windows 把 \\ 和 // 都解析为 UNC）与 file: 协议。
_HYPERLINK_NETWORK_PREFIX_PATTERN = re.compile(r"^(?:\\\\|//)", re.IGNORECASE)
_FILE_SCHEME_PATTERN = re.compile(r"^file:", re.IGNORECASE)


def hyperlink_allowed() -> bool:
    """是否放行 HYPERLINK 公式函数（safety.allowHyperlink）。默认拦截。"""
    return capability_bool("safety", "allowHyperlink", default=False)


def dangerous_formula(
    formula: str, *, allow_hyperlink: bool | None = None
) -> str | None:
    """公式含危险函数/注入载体时返回命中名称，否则返回 None。

    对整段公式做搜索（不只开头），能覆盖 IF(WEBSERVICE(...)) 这类嵌套。
    allow_hyperlink=None 时按配置决定；显式传 True/False 可覆盖（供测试）。
    """
    if not formula:
        return None
    if allow_hyperlink is None:
        allow_hyperlink = hyperlink_allowed()
    for name, pattern in _DANGEROUS_FORMULA_PATTERNS:
        if name == "HYPERLINK" and allow_hyperlink:
            continue
        if pattern.search(formula):
            return name
    if _DDE_PATTERN.match(formula):
        return "DDE"
    if _UNC_PREFIX_PATTERN.search(formula):
        return "UNC"
    if _EXTERNAL_WORKBOOK_REF_PATTERN.search(formula):
        return "EXTERNAL_REF"
    return None


def dangerous_hyperlink_address(address: str) -> str | None:
    """超链接地址含命令注入载体时返回命中名称，否则返回 None。"""
    if not address:
        return None
    if address.lstrip().startswith("="):
        return "DDE"
    if _DDE_PATTERN.match(address):
        return "DDE"
    stripped = address.lstrip()
    if _HYPERLINK_NETWORK_PREFIX_PATTERN.match(stripped):
        return "UNC"
    if _FILE_SCHEME_PATTERN.match(stripped):
        return "FILE"
    return None


# ---------------------------------------------------------------------------
# 内嵌图片大小
# ---------------------------------------------------------------------------

def max_image_bytes() -> int:
    return capability_int("images", "maxBytes")


def base64_within_image_limit(base64_text: str) -> bool:
    """base64 文本解码后是否在图片字节上限内（不实际解码，按长度估算）。"""
    max_bytes = max_image_bytes()
    # base64 文本长度 ≈ ceil(n/3)*4；用上界判断，留 1 个空余块。
    max_base64_len = ((max_bytes // 3) + 1) * 4
    return len(base64_text) <= max_base64_len
