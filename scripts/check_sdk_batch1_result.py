"""第一批合成案例的**事后业务核对**（只读工件，不参与产品运行）。

用途：真实模型 / SDK / Docker 那次运行只证明"通道把计划带回来了"，
`businessAcceptance=not_checked` 提醒必须另行核对结果值。本脚本读取
`outputs/*/synthetic-result.json`，按 `config/online_link_scenario.json` 的
期望语义逐项断言：

- 汇总金额：金卡 / 银卡 / 未匹配 / 总计；
- 异常明细必须**带来源表与行号**：会员表空编号、会员表重复编号、订单未匹配；
- 结果矩阵必须声明 `complete=True`（不靠省略表示"没有记录"）。

它不把期望值写进任何请求，也不生成结果；只在运行结束后核对工件。
期望值来自场景夹具（测试专用），运行时代码不得引用本文件。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCENARIO = json.loads((ROOT / "config" / "online_link_scenario.json").read_text(encoding="utf-8"))


def _text(value: object) -> str:
    return "" if value is None else str(value)


def _has_source_detail(
    rows: list[list[object]], table: str, row_number: int, *needles: str
) -> bool:
    """存在一行能定位到「哪张表、哪一行」，且说明含全部片段。

    两种合法写法都接受：
    - 列分开写：来源表列恰为 table、行号列恰为 row_number；
    - 同一格写定位：`会员表!A5` / `会员表第5行`。

    不能用裸子串匹配行号：`订单表!A5 在会员表中无对应编号` 既含“会员表”又含 `5`，
    会把“会员表空编号明细缺失”误判成已覆盖。
    """
    locator = re.compile(
        re.escape(table)
        + r"\s*[!！]?\s*(?:第\s*)?(?:[A-Z]{1,3})?"
        + str(row_number)
        + r"(?:\s*行)?"
    )
    for row in rows:
        cells = [_text(cell) for cell in row]
        joined = " ".join(cells)
        if not all(needle in joined for needle in needles):
            continue
        separated = (
            any(cell.strip() == table for cell in cells)
            and any(cell.strip() == str(row_number) for cell in cells)
        )
        if separated or any(locator.search(cell) for cell in cells):
            return True
    return False


def _amount_for(rows: list[list[object]], label: str) -> float | None:
    """在包含 label 的行里取第一个数值（排除"行号"这类编号列）。"""
    for row in rows:
        if not any(label in _text(cell) for cell in row):
            continue
        numbers = [
            float(cell)
            for cell in row
            if isinstance(cell, (int, float)) and not isinstance(cell, bool)
        ]
        if numbers:
            return max(numbers)
    return None


def check(path: Path, mode: str) -> list[str]:
    data = json.loads(path.read_text(encoding="utf-8"))
    rows = data.get("rows") or []
    failures: list[str] = []

    def require(condition: bool, message: str) -> None:
        if not condition:
            failures.append(message)

    expected = (
        SCENARIO["baselineExpected"] if mode == "baseline" else SCENARIO["afterCausalChangeExpected"]
    )
    expected_amounts = {
        row[0]: row[1] for row in expected[1:] if isinstance(row[1], (int, float))
    }

    require(data.get("complete") is True, "结果矩阵未声明 complete=True")
    require(bool(data.get("sourceSheets")), "结果缺少 sourceSheets 归属")

    unmatched = _amount_for(rows, "未匹配")
    require(
        unmatched == float(expected_amounts["未匹配"]),
        f"未匹配金额应为 {expected_amounts['未匹配']}，实际 {unmatched}",
    )
    for level in ("金卡", "银卡"):
        actual = _amount_for(rows, level)
        require(
            actual == float(expected_amounts[level]),
            f"{level}金额应为 {expected_amounts[level]}，实际 {actual}",
        )

    total_rows = [row for row in rows if any("总计" in _text(cell) for cell in row)]
    totals = [
        float(cell)
        for row in total_rows
        for cell in row
        if isinstance(cell, (int, float)) and not isinstance(cell, bool)
    ]
    require(bool(totals) and max(totals) == 1000.0, f"金额总计应为 1000，实际 {totals}")

    # 异常明细：必须能在结果里定位到"哪张表、哪一行"。
    require(
        _has_source_detail(rows, "会员表", 5, "编号为空"),
        "缺少会员表空编号明细（会员表第 5 行）",
    )
    require(
        _has_source_detail(rows, "会员表", 3, "7", "重复"),
        "缺少会员表重复编号明细（会员表第 3 行 键=7）",
    )
    require(
        _has_source_detail(rows, "订单表", 5, "404"),
        "缺少订单未匹配明细（订单表第 5 行 键=404）",
    )
    return failures


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("result", type=Path, help="synthetic-result.json 路径")
    parser.add_argument("--mode", choices=("baseline", "causal"), required=True)
    args = parser.parse_args()
    failures = check(args.result, args.mode)
    if failures:
        print(f"业务核对未通过（{args.mode}）：")
        for failure in failures:
            print(f" - {failure}")
        return 1
    print(f"业务核对通过（{args.mode}）：{args.result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
