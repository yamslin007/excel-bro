"""在线连续集成用的固定分析代码（由固定的模型替身作为工具参数提交）。

纪律：
- 这份代码**只**读取容器内的 ``input``（宿主真实回执登记出来的数据集），
  不读文件、不联网、不引用任何测试常量。
- 结果矩阵由它**算出**，不是从任何 JSON 夹具里读出来的期望值。
- 规则（按需求约定）：会员键先 trim 再匹配、保留前导零；空键不参与关联；
  重复键保留第一次出现；订单**左连接**会员；按关联到的会员等级汇总金额；
  并给出可定位的空键/重复键/未匹配明细。
"""

members = input["datasets"][0]["values"]
orders = input["datasets"][1]["values"]

member_header, *member_body = members
member_key = member_header.index("会员编号")
member_level = member_header.index("等级")

clean = {}
empty_rows = []
duplicate_rows = []
for offset, row in enumerate(member_body):
    sheet_row = offset + 2
    raw = "" if row[member_key] is None else str(row[member_key])
    key = raw.strip()
    if key == "":
        empty_rows.append(sheet_row)
        continue
    if key in clean:
        duplicate_rows.append((sheet_row, key))
        continue
    clean[key] = row[member_level]

order_header, *order_body = orders
order_key = order_header.index("会员编号")
order_amount = order_header.index("金额")

totals = {}
unmatched = []
for offset, row in enumerate(order_body):
    sheet_row = offset + 2
    raw = "" if row[order_key] is None else str(row[order_key])
    key = raw.strip()
    if key == "":
        continue
    amount = 0 if row[order_amount] is None else row[order_amount]
    level = clean.get(key)
    if level is None:
        level = "未匹配"
        unmatched.append((sheet_row, key))
    totals[level] = totals.get(level, 0) + amount

unmatched_detail = "、".join(
    "订单表第%d行 键=%s" % (row, key) for row, key in sorted(unmatched)
)

rows = []
for level in sorted(totals):
    rows.append([level, totals[level], unmatched_detail if level == "未匹配" else ""])
rows.append([
    "空键数",
    len(empty_rows),
    "、".join("会员表第%d行" % row for row in empty_rows),
])
rows.append([
    "重复键数",
    len(duplicate_rows),
    "、".join("会员表第%d行 键=%s" % (row, key) for row, key in duplicate_rows),
])

result = {"kind": "matrix", "values": [["会员等级", "金额合计", "明细"]] + rows}
