"""Untrusted matrix validation only: no code, coordinates, plans or persistence."""
import json
import math

from .capabilities import capability_int


def encode_matrix_result(result) -> tuple[bytes, int, int]:
    # Exact, independent validation. Never trust worker field names or Pydantic
    # scalar coercion as proof that an output is safe to materialize in Excel.
    if type(result) is not dict or set(result) != {"kind", "values"} or result["kind"] != "matrix":
        raise ValueError()
    values = result["values"]
    rows = len(values) if type(values) is list else 0
    if not 1 <= rows <= capability_int("pythonResult", "maxRows"):
        raise ValueError()
    columns = len(values[0]) if type(values[0]) is list else 0
    if (not 1 <= columns <= capability_int("pythonResult", "maxColumns")
        or rows*columns > capability_int("pythonResult", "maxCells")):
        raise ValueError()
    for row in values:
        if type(row) is not list or len(row) != columns:
            raise ValueError()
        for value in row:
            if value is None or type(value) is bool:
                continue
            if type(value) is int and abs(value) <= 9007199254740991:
                continue
            if (type(value) is float and math.isfinite(value)
                and (not value.is_integer() or abs(value) <= 9007199254740991)):
                continue
            if type(value) is str and len(value.encode("utf-16-le")) // 2 <= capability_int("pythonResult", "maxCellChars"):
                continue
            raise ValueError()
    encoded = json.dumps(result, ensure_ascii=False, allow_nan=False, separators=(",", ":")).encode("utf-8")
    if len(encoded) > capability_int("pythonResult", "maxBytes"):
        raise ValueError()
    return encoded, rows, columns
