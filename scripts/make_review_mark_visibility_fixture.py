"""生成「审阅标记可见性」**双端实机验收**用的专用测试工作簿。

纪律：
- 只产出**独立测试工作簿**，绝不触碰任何业务工作簿。
- 可反复运行、每次覆盖重建，保证验收有确定的复位手段。
- 一份 .xlsx 同时用于 **Excel** 与 **WPS**（两端都读得动 xlsx）。

夹具设计（对应「浅黄底色 vs 只边框」两条呈现路径）：

第一区 A1:C4（写值标记 → 应见浅黄底色 + 金黄边框）：
- A1:C1 表头；A2:C4 是纯文本/数字，无填充、无用户条件格式。
- 用它跑一次「写值」（例如把 B 列数量 +1、C 列重算），执行后整块区域应被
  浅黄 `#FFF3B0` 底色 + 金黄 `#B8860B` 边框标出 —— 不再像普通网格线。

第二区 E1:F4（已有业务填充 → 标记不得覆盖）：
- E2 已有**浅蓝色**直接填充；F2 已有**绿色**直接填充。
- 这些是"用户自己的业务填充"，标记打上去后必须仍在原位、可辨认。

第三区 H1:I4（用户条件格式 → 保留顺序/语义）：
- H2:I2 有一条**用户自建条件格式**（值 > 100 时红底）；用它验证"标记不覆盖、
  不重排用户规则"，确认后只清本次规则、用户规则仍在。

第四区 K1:L3（改填充动作 → 应只见边框、不叠底色）：
- 空表，跑一次「把 K2:L2 填充成黄色」，执行后该区域应**只有金黄边框**，
  实际填充色（黄色）可见、不被审阅底色盖掉。

产物：outputs/review-mark-visibility-2026-10-17/审阅标记验收.xlsx
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

BASE = Path(__file__).resolve().parent.parent / "outputs" / "review-mark-visibility-2026-10-17"
OUTPUT = BASE / "审阅标记验收.xlsx"

# 期望的标记配色（与 packages/core/src/reviewMarkStyle.ts 逐字一致）。
REVIEW_FILL = "FFF3B0"  # 浅黄临时底色
REVIEW_BORDER = "B8860B"  # 醒目金黄边框

# 用户自建的"业务填充"与"业务条件格式"，用来验证标记不覆盖它们。
BUSINESS_FILL_BLUE = "DDEBF7"
BUSINESS_FILL_GREEN = "E2EFDA"


def build() -> None:
    BASE.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "审阅标记验收"

    # ---- 第一区 A1:C4：写值标记（浅黄底色 + 边框）----
    ws["A1"] = "写值验收区（执行后应见浅黄底色 + 金黄边框）"
    ws["A2"], ws["B2"], ws["C2"] = "编号", "数量", "金额"
    for row, (code, qty, amt) in enumerate([
        ("P001", 10, 100), ("P002", 20, 200), ("P003", 30, 300)
    ], start=3):
        ws.cell(row=row, column=1, value=code)
        ws.cell(row=row, column=2, value=qty)
        ws.cell(row=row, column=3, value=amt)

    # ---- 第二区 E1:F4：已有业务填充，标记不得覆盖 ----
    ws["E1"] = "已有业务填充（标记不得覆盖、要仍可辨认）"
    ws["E2"], ws["F2"] = "浅蓝填充", "绿色填充"
    ws["E2"].fill = PatternFill("solid", fgColor=BUSINESS_FILL_BLUE)
    ws["F2"].fill = PatternFill("solid", fgColor=BUSINESS_FILL_GREEN)
    ws["E3"], ws["F3"] = 100, 200
    ws["E4"], ws["F4"] = 300, 400

    # ---- 第三区 H1:I4：用户条件格式，保留顺序/语义 ----
    ws["H1"] = "用户条件格式（值>100 红底；标记不得覆盖或重排）"
    ws["H2"], ws["I2"] = "阈值区", "数值"
    ws["H3"], ws["I3"] = "a", 120
    ws["H4"], ws["I4"] = "b", 80
    # 用户自建规则：I3:I4 值 > 100 → 红底（FCE4E4）。
    ws.conditional_formatting.add(
        "I3:I4",
        CellIsRule(operator="greaterThan", formula=["100"],
                   fill=PatternFill("solid", fgColor="FCE4E4")),
    )

    # ---- 第四区 K1:L3：改填充动作（应只见边框、不叠底色）----
    ws["K1"] = "改填充验收区（执行 setFill 后应只见边框、黄色填充可见）"
    ws["K2"], ws["L2"] = "目标A", "目标B"
    ws["K3"], ws["L3"] = None, None

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["H"].width = 16
    ws.column_dimensions["I"].width = 10
    ws.column_dimensions["K"].width = 12
    ws.column_dimensions["L"].width = 12

    wb.save(OUTPUT)
    print(f"已生成 {OUTPUT}")


if __name__ == "__main__":
    build()
