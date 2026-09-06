from __future__ import annotations

import re
from dataclasses import dataclass


_VARIANT_TRANSLATION = str.maketrans(
    {
        "刪": "删",
        "清": "清",
        "除": "除",
        "餘": "余",
        "僅": "仅",
        "頁": "页",
        "錶": "表",
    }
)

_SHEET_NOUN = r"(?:工作表|子表|sheet(?:页|表)?)"
_DELETE_VERB = r"(?:删除|清除|移除)"
_TRAILING_PUNCTUATION = r"[。.!！?？]*"

_EXCEPT_KEEPER_PATTERN = re.compile(
    rf"^\s*(?:请\s*)?{_DELETE_VERB}\s*"
    rf"(?:除(?:了)?\s*)?(?P<keeper>.+?)\s*"
    rf"(?:以外|之外|外)\s*(?:的)?\s*"
    rf"(?:所有|全部)?\s*(?:其他|其它|其余|剩余)?\s*"
    rf"(?:所有|全部)?\s*"
    rf"{_SHEET_NOUN}\s*{_TRAILING_PUNCTUATION}\s*$",
    re.IGNORECASE,
)
_KEEP_THEN_DELETE_PATTERN = re.compile(
    rf"^\s*(?:请\s*)?(?:只|仅)?\s*(?:保留|留下)\s*"
    rf"(?P<keeper>.+?)\s*[,，;；]?\s*(?:并|然后|再)?\s*"
    rf"{_DELETE_VERB}\s*(?:所有|全部)?\s*"
    rf"(?:其余|其他|其它|剩余)\s*(?:所有|全部)?\s*(?:的)?\s*"
    rf"{_SHEET_NOUN}\s*{_TRAILING_PUNCTUATION}\s*$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class BulkWorksheetDeletionIntent:
    keeper_name: str


def _clean_keeper_name(value: str) -> str:
    return value.strip().strip("\"'“”‘’`「」『』【】，,;； ")


def parse_bulk_worksheet_deletion(
    prompt: str,
) -> BulkWorksheetDeletionIntent | None:
    """Parse an explicit request to delete every worksheet except one.

    This parser intentionally requires both a delete verb and a worksheet noun.
    It must not turn a request to clear cell contents into worksheet deletion.
    """

    normalized = prompt.translate(_VARIANT_TRANSLATION)
    for pattern in (_EXCEPT_KEEPER_PATTERN, _KEEP_THEN_DELETE_PATTERN):
        match = pattern.fullmatch(normalized)
        if not match:
            continue
        keeper_name = _clean_keeper_name(match.group("keeper"))
        if keeper_name:
            return BulkWorksheetDeletionIntent(keeper_name=keeper_name)
    return None
