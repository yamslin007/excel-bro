"""生成格式动作**双端实机验收**用的专用测试工作簿。

纪律：
- 只产出**独立测试工作簿**，绝不触碰任何业务工作簿。
- 可反复运行、每次覆盖重建，保证验收有确定的复位手段。
- 一份 .xlsx 同时用于 **Excel** 与 **WPS**（两端都读得动 xlsx）。

夹具的刻意设计：

第一区 A1:E5（Task #53 三项）：
- A. 加粗＋水平居中的目标区域含**文本/数字/公式**；
- B. 混合字号：同一列里 B2..B5 的字号各不相同（12/16/11/14），
     统一改字号后必须能逐格恢复**原始混合字号**；
- C. 每一行行高各不相同（15/20/25/30/35），用于验证"改行高影响整行"。

第二区 G1:L5（第 19 轮「撤销按动作实际影响拆分恢复」）：
- D. G2：**已有**删除线 + 上标 + 图案色 + 斜线边框 —— 加粗再撤销后这四样必须还在，
     而加粗必须消失（这是"clear(Formats) 之后要把非本动作属性的当前值并回去"的实机判据）；
- E. G3：加粗**之后**再手工改文字方向/收缩适应/保护 —— 撤销**不得**覆盖这些后续手工修改；
- F. 第 4 行（含 G4:L4 **条件格式**规则）：只改行高再撤销 —— 不得因"区域含条件格式"被拒，
     也不得调用任何单元格清格式操作，条件格式规则必须原样；
- G. G5：**非 Normal 命名样式** + 直接蓝色填充覆盖 —— 加粗撤销后样式名与填充色都要还原
     （`clear(Formats)` 会把命名样式打回 Normal，所以必须捕获并写回）。

产物：outputs/format-acceptance-2026-10-14/格式验收.xlsx

> 上一轮（2026-10-13）的那份被宿主打开占用时无法覆盖写入，故**另起一个目录**；
> 旧文件保持原样不动（绝不为了写入而去结束用户正在用的宿主进程）。
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side

BASE = Path(__file__).resolve().parent.parent / "outputs" / "format-acceptance-2026-10-14"
OUTPUT = BASE / "格式验收.xlsx"

HEADER = ["编号", "名称", "数量", "金额（公式）", "备注（公式）"]

# 第二区（撤销语义）常量
UNDO_TITLE = "撤销验收区（G2:L5）"
PRESERVE_TEXT = "保留我"  # G2：删除线+上标+图案色+斜线边框
LATER_EDIT_TEXT = "后续改我"  # G3：加粗后再手工改文字方向/收缩/保护
CF_NUMBERS = [120, 80, 150, 60]  # G4:L4：条件格式目标格
CF_RULE = "G4:L4"
STYLE_CELL_TEXT = "样式我"  # G5：非 Normal 命名样式 + 直接填充覆盖
NAMED_STYLE = "验收样式"

# (编号, 名称, 数量, 金额公式, 备注公式, 名称列字号)
ROWS = [
    ("X001", "甲", 100, "=C2*1.5", '=B2&"元"', 12),
    ("X002", "乙", 200, "=C3*1.5", '=B3&"元"', 16),
    ("X003", "丙", 300, "=C4*1.5", '=B4&"元"', 11),
    ("X004", "丁", 400, "=C5*1.5", '=B5&"元"', 14),
]

# 每个工作表行高度（第 n 行 → 第 n 个值）。
ROW_HEIGHTS = [15, 20, 25, 30, 35]


def build() -> None:
    workbook = Workbook()

    data = workbook.active
    data.title = "格式验收"

    # 表头（加粗 + 居中，作为"改前已有格式"的锚点，撤销后应回到这个状态）。
    data.append(HEADER)
    for column in range(1, len(HEADER) + 1):
        cell = data.cell(row=1, column=column)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # 数据行：文本（A/B）、数字（C）、公式（D/E）。
    for index, (code, name, qty, amount_formula, note_formula, name_size) in enumerate(ROWS, start=2):
        data.cell(row=index, column=1, value=code)
        name_cell = data.cell(row=index, column=2, value=name)
        # 混合字号：同一列里逐行不同 —— 撤销必须逐格恢复这套原始混合字号。
        name_cell.font = Font(size=name_size)
        data.cell(row=index, column=3, value=qty)
        data.cell(row=index, column=4, value=amount_formula)
        data.cell(row=index, column=5, value=note_formula)

    # 每行行高各不相同 —— 验证"改行高影响整行"与撤销恢复。
    for row_index, height in enumerate(ROW_HEIGHTS, start=1):
        data.row_dimensions[row_index].height = height

    # ── 第二区：撤销语义（第 19 轮）─────────────────────────────────────────
    data["G1"] = UNDO_TITLE
    data["G1"].font = Font(bold=True)

    # D. G2：已有删除线 + 上标 + 图案色（lightUp 图案，前色=图案线色）+ 右下斜线边框。
    #    （下标/上标在 xlsx 里由 vertAlign 二选一表达，这里用上标；验收时按实际显示核对。）
    preserve = data["G2"]
    preserve.value = PRESERVE_TEXT
    preserve.font = Font(strike=True, vertAlign="superscript", size=12, color="FF0000")
    preserve.fill = PatternFill(patternType="lightUp", fgColor="00B0F0", bgColor="FFFF00")
    preserve.border = Border(
        diagonal=Side(style="thin", color="FF00FF"), diagonalDown=True
    )

    # E. G3：加粗之后由用户在宿主里手工改文字方向/收缩适应/保护，撤销不得覆盖。
    later = data["G3"]
    later.value = LATER_EDIT_TEXT
    later.alignment = Alignment(horizontal="left")

    # F. 第 4 行的 G4:L4 加条件格式规则（>100 变绿）——注意整行的行高动作会覆盖到这些格。
    for offset, value in enumerate(CF_NUMBERS):
        data.cell(row=4, column=7 + offset, value=value)
    data.conditional_formatting.add(
        CF_RULE,
        CellIsRule(
            operator="greaterThan",
            formula=["100"],
            fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
            font=Font(color="006100"),
        ),
    )

    # G. G5：非 Normal 命名样式 + 直接蓝色填充覆盖。
    style = NamedStyle(name=NAMED_STYLE)
    style.font = Font(size=14, italic=True, color="0000FF")
    workbook.add_named_style(style)
    style_cell = data["G5"]
    style_cell.value = STYLE_CELL_TEXT
    style_cell.style = NAMED_STYLE
    style_cell.fill = PatternFill(start_color="3399FF", end_color="3399FF", fill_type="solid")

    # 列宽给一点余量，便于肉眼核对。
    data.column_dimensions["A"].width = 10
    data.column_dimensions["B"].width = 10
    data.column_dimensions["C"].width = 10
    data.column_dimensions["D"].width = 14
    data.column_dimensions["E"].width = 14
    for column in ("G", "H", "I", "J", "K", "L"):
        data.column_dimensions[column].width = 12

    BASE.mkdir(parents=True, exist_ok=True)
    workbook.save(OUTPUT)

    print(f"已生成 {OUTPUT}")
    print("  表「格式验收」")
    print("  第一区 A1:E5（Task #53）")
    print("  - 文本：A2:A5（编号）、B2:B5（名称）")
    print("  - 数字：C2:C5（数量 100/200/300/400）")
    print("  - 公式：D2:D5（=Cn*1.5）、E2:E5（=Bn&\"元\"）")
    print("  - 混合字号：B2=12 / B3=16 / B4=11 / B5=14")
    print("  - 行高：第1..5行 = 15 / 20 / 25 / 30 / 35")
    print("  - 表头第1行：加粗 + 居中")
    print("  第二区 G1:L5（撤销按动作实际影响拆分恢复）")
    print(f"  - G2「{PRESERVE_TEXT}」：删除线 + 上标 + 图案色 + 斜线边框（场景 D）")
    print(f"  - G3「{LATER_EDIT_TEXT}」：加粗后再手工改文字方向/收缩/保护（场景 E）")
    print(f"  - G4:L4 = {CF_NUMBERS}，条件格式规则「>100 变绿」在 {CF_RULE}（场景 F）")
    print(f"  - G5「{STYLE_CELL_TEXT}」：命名样式「{NAMED_STYLE}」+ 蓝色直接填充（场景 G）")


if __name__ == "__main__":
    build()
