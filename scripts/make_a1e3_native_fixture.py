"""生成 A1-E3 S1「注入 source_changed 后停止」原生验收用的独立测试工作簿。

纪律（与 ``make_format_acceptance_fixture.py`` 一致）：
- 只产出**独立测试工作簿**，绝不触碰任何业务工作簿；
- 每轮使用新目录，已有夹具拒绝覆盖，避免破坏已打开文件与证据；
- 一份 .xlsx 同时用于 Excel 与 WPS（两端都读得动 xlsx）。

夹具设计（对齐 ``docs/A1_E3_NATIVE_ACCEPTANCE_2026-10-25.md`` §0.1）：
- 三张表：`来源`、`目标`、`汇总`（S1 只用到前两张，`汇总` 留空兜底）。
- ``来源``：B1 表头「金额」，B2:B4 = 111 / 250 / 333。
- ``目标``：B1 表头「金额」，B2:B4 = -101 / -202 / -303（执行前初值）。
- 验收需求「读取来源表 B2:B4，将每个值乘以 2，写入目标表 B2:B4，来源保持不变」
  在**未注入停止**的正常执行下会把目标写成 222 / 500 / 666；注入停止后**来源与目标
  都必须保持测试前值**（111/250/333 与 -101/-202/-303）。

产物：
- ``outputs/a1e3-native/a1e3-验收.xlsx``（母本）
- ``outputs/a1e3-native/a1e3-验收-副本.xlsx``（验收实际打开的副本；每场景重新打开）

> 母本与副本同目录、同内容，只为让「每个场景从副本重新打开」这一步有据可循，
> 不要求脚本在验收过程中再写文件（避免覆盖被宿主打开占用的工作簿）。
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from openpyxl import Workbook

BASE = Path(__file__).resolve().parent.parent / "outputs" / "a1e3-native"
MASTER = BASE / "a1e3-验收.xlsx"
COPY = BASE / "a1e3-验收-副本.xlsx"

SOURCE_VALUES = [111, 250, 333]
TARGET_VALUES = [-101, -202, -303]
HEADER = "金额"


def build(path: Path) -> None:
    if path.exists():
        raise FileExistsError(f"Refusing to overwrite fixture: {path}")
    workbook = Workbook()
    source = workbook.active
    source.title = "来源"
    target = workbook.create_sheet("目标")
    workbook.create_sheet("汇总")

    source["B1"] = HEADER
    target["B1"] = HEADER
    for row, value in enumerate(SOURCE_VALUES, start=2):
        source.cell(row=row, column=2, value=value)
    for row, value in enumerate(TARGET_VALUES, start=2):
        target.cell(row=row, column=2, value=value)

    workbook.save(path)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", type=Path, default=BASE)
    directory = parser.parse_args().out.resolve()
    directory.mkdir(parents=True, exist_ok=True)
    master = directory / MASTER.name
    copy = directory / COPY.name
    if master.exists() or copy.exists():
        raise FileExistsError("Fixture exists; use a new run directory")
    build(master)
    build(copy)
    print(f"Master: {master}")
    print(f"Test copy: {copy}")
    print(f"来源 B2:B4 = {SOURCE_VALUES}")
    print(f"目标 B2:B4 = {TARGET_VALUES}（执行前初值）")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
