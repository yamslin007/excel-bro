"""生成 Excel「未完成执行 → 重开发现 → 只续剩余动作 → 审阅收尾」的验收夹具。

为什么需要夹具：真机验收不能再靠"手点暂停碰时机"。这条链路要的是**确定性**现场：
一张已保存的工作簿、两张表、来源格有值、目标格是已知的旧值，再由 DEV 面板的中断探针
在第 2 步写入回读之后、完成回执之前断掉。夹具只负责"工作簿长成什么样"，不碰账本、
不碰宿主状态。

用法：
    .venv/Scripts/python.exe scripts/make_excel_recovery_fixture.py
    .venv/Scripts/python.exe scripts/make_excel_recovery_fixture.py --out <path.xlsx>

产物默认写到 `outputs/excel-recovery-fixture-20261029/Excel恢复夹具-20261029.xlsx`，
并把"本次验收要用的那一句话"和逐格期望打印出来。**不覆盖已存在的文件**：夹具一旦
被上一次验收改过（写入/撤销），再覆盖会让"原始值"失真；需要重建时显式 `--force`。
"""

from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import Workbook

SOURCE_SHEET = "综合票房"
TARGET_SHEET = "汇总表"
SPARE_SHEET = "备用表"

# 三格独立单格写入：探针要求正好 3 个互不相同的单格 writeValues。
ROWS = (2, 3, 4)


def build(path: Path) -> None:
    book = Workbook()
    source = book.active
    source.title = SOURCE_SHEET
    source["A1"] = "影片"
    source["B1"] = "票房"
    for index, row in enumerate(ROWS):
        source[f"A{row}"] = f"影片{index + 1}"
        source[f"B{row}"] = (index + 1) * 111  # 111 / 222 / 333

    target = book.create_sheet(TARGET_SHEET)
    target["A1"] = "影片"
    target["B1"] = "票房"
    for index, row in enumerate(ROWS):
        target[f"A{row}"] = f"影片{index + 1}"
        target[f"B{row}"] = -101 * (index + 1)  # -101 / -202 / -303（撤销时应当回到这三个值）

    spare = book.create_sheet(SPARE_SHEET)
    spare["A1"] = "备注"
    spare["A2"] = "这张表只用来验证「授权范围可以收窄/撤销」，不要写它。"

    path.parent.mkdir(parents=True, exist_ok=True)
    book.save(path)


def describe(path: Path) -> None:
    print(f"夹具：{path}")
    print(f"工作表：{SOURCE_SHEET} / {TARGET_SHEET} / {SPARE_SHEET}")
    print(f"来源：{SOURCE_SHEET}!B2:B4 = 111 / 222 / 333")
    print(f"目标旧值：{TARGET_SHEET}!B2:B4 = -101 / -202 / -303")
    print()
    print("在面板里发这一句（字面，不要改写）：")
    print(f"  把「{SOURCE_SHEET}」的 B2、B3、B4 三格值分别写到「{TARGET_SHEET}」的 B2、B3、B4")
    print()
    print("期望：模型给出 3 个独立的单格 writeValues；DEV 面板已启用「下一次第二步中断」时，")
    print(f"      第 2 步写完并回读后中断 —— {TARGET_SHEET} 应停在 111 / 222 / -303，第 3 步未执行。")


def main() -> int:
    parser = argparse.ArgumentParser(description="生成 Excel 恢复验收夹具")
    default = Path("outputs") / "excel-recovery-fixture-20261029" / "Excel恢复夹具-20261029.xlsx"
    parser.add_argument("--out", type=Path, default=default)
    parser.add_argument("--force", action="store_true", help="已存在时覆盖（会丢掉上次验收留下的改动）")
    options = parser.parse_args()

    path: Path = options.out
    if path.exists() and not options.force:
        print(f"已存在，未改动：{path}（需要重建加 --force）")
        describe(path)
        return 0
    build(path)
    describe(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
