"""生成 M1-M10 实机验收用的独立测试工作簿。

可反复运行，每次覆盖重建 —— M5 写入、M7 撤销、M8 冲突、M9 跨表都会改数据，
清单要求「临时工作簿，不使用真实业务数据」，所以需要一个确定的复位手段，
而不是手工准备一份 xlsx 然后祈祷它没被改脏。

产物落在 .tmp/（已在 .gitignore 里），脚本本身进版本库。

夹具的三处刻意设计：
- 「甲」出现两次 —— 模型逐格抄写时重复值最容易错位，查表算出来的不会。
- 第 6 行留真空格（不是空串）—— 走 blankLabel 分支，也用于验收撤销后
  恢复出的是真空格还是空串（§1.1）。
- 标签用甲/乙/丙这类无业务含义的占位符 —— 避免以后读日志的人以为这批
  取值和真实业务语义有关。
"""

from pathlib import Path

from openpyxl import Workbook

OUTPUT = Path(__file__).resolve().parent.parent / ".tmp" / "wps_mapping_fixture.xlsx"

SOURCE_ROWS = [
    ("X001", "甲"),
    ("X002", "乙"),
    ("X003", "甲"),
    ("X004", "丙"),
    ("X005", None),
]


def build() -> None:
    workbook = Workbook()

    data = workbook.active
    data.title = "数据"
    data.append(["单号", "类别", "归属"])
    for order_no, category in SOURCE_ROWS:
        data.append([order_no, category, None])

    # M9 的目标表：从「数据」的 B 列映射写到这里的 H 列，验收跨表时
    # 写入只落在计划声明的那张表上。表头铺到 H 列，让 columnCount 覆盖目标区域。
    summary = workbook.create_sheet("汇总")
    summary.append(["单号", "占位2", "占位3", "占位4", "占位5", "占位6", "占位7", "归属"])
    for order_no, _ in SOURCE_ROWS:
        summary.append([order_no, None, None, None, None, None, None, None])

    # 存在但不勾选，用于 M1 验收「授权集合仅含勾选的表」。
    workbook.create_sheet("不授权")

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(OUTPUT)
    print(f"已生成 {OUTPUT}")
    print(f"  数据  A1:C6  源值在 B2:B6（甲/乙/甲/丙/空），C 列待写")
    print(f"  汇总  A1:H6  H 列待写（M9）")
    print(f"  不授权  空表，用于验收授权边界")


if __name__ == "__main__":
    build()
