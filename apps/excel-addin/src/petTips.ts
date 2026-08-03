export interface PetTip {
  id: string;
  category: string;
  title: string;
  detail: string;
  shortcut?: string;
}

// 高频且实用的 Excel 操作技巧，供格仔在气泡里轮流展示。
// 按主题分组，新增技巧直接往数组里加即可，顺序不影响逻辑。
export const PET_TIPS: PetTip[] = [
  // —— 录入与编辑 ——
  {
    id: "wrap-in-cell",
    category: "录入编辑",
    title: "单元格内换行",
    detail: "在同一个单元格里想分行输入，按这个组合键就能手动换行。",
    shortcut: "Alt + Enter"
  },
  {
    id: "fill-selection",
    category: "录入编辑",
    title: "选区一次填满",
    detail: "先选好区域再输入内容，用这个组合键可把内容一次填入所有选中单元格。",
    shortcut: "Ctrl + Enter"
  },
  {
    id: "fill-down",
    category: "录入编辑",
    title: "向下复制上一格",
    detail: "选中某格后按下，会把正上方单元格的内容快速填到当前格。",
    shortcut: "Ctrl + D"
  },
  {
    id: "fill-right",
    category: "录入编辑",
    title: "向右复制左一格",
    detail: "把左侧单元格内容快速填到当前格，横向填充很方便。",
    shortcut: "Ctrl + R"
  },
  {
    id: "edit-cell",
    category: "录入编辑",
    title: "直接进入编辑",
    detail: "选中单元格后按 F2，光标会进入格内末尾，不用双击即可修改。",
    shortcut: "F2"
  },
  {
    id: "insert-today",
    category: "录入编辑",
    title: "快速插入当前日期",
    detail: "不用手打日期，一个组合键就能填入今天的日期。",
    shortcut: "Ctrl + ;"
  },
  {
    id: "insert-now",
    category: "录入编辑",
    title: "快速插入当前时间",
    detail: "需要记录时间点时，这个组合键直接填入现在的时间。",
    shortcut: "Ctrl + Shift + ;"
  },
  {
    id: "fill-down-double-click",
    category: "录入编辑",
    title: "双击填充到底",
    detail: "选中带公式的单元格，双击右下角的填充柄，会自动向下填充到相邻数据末尾。"
  },
  {
    id: "flash-fill",
    category: "录入编辑",
    title: "快速填充识别规律",
    detail: "手动示范一两个想要的结果，按下后 Excel 会按规律自动补全整列，拆分文本、提取数字都好用。",
    shortcut: "Ctrl + E"
  },
  {
    id: "repeat-action",
    category: "录入编辑",
    title: "重复上一步操作",
    detail: "刚做过的动作（如设置格式、插入行）想再来一次，按 F4 即可重复。",
    shortcut: "F4"
  },

  // —— 导航与选择 ——
  {
    id: "jump-edge",
    category: "导航选择",
    title: "跳到数据边缘",
    detail: "按住 Ctrl 再按方向键，光标会直接跳到当前连续数据的边界，长表格里超省时间。",
    shortcut: "Ctrl + 方向键"
  },
  {
    id: "select-to-edge",
    category: "导航选择",
    title: "快速选到数据边缘",
    detail: "再多按一个 Shift，就能从当前位置一直选取到数据边界。",
    shortcut: "Ctrl + Shift + 方向键"
  },
  {
    id: "select-all-data",
    category: "导航选择",
    title: "选中整个数据区",
    detail: "点数据区内任意格，按下会先选中当前表格，再按一次选中整张工作表。",
    shortcut: "Ctrl + A"
  },
  {
    id: "goto-a1",
    category: "导航选择",
    title: "一键回到 A1",
    detail: "不管翻到哪，按下立刻回到工作表左上角的 A1 单元格。",
    shortcut: "Ctrl + Home"
  },
  {
    id: "goto-last",
    category: "导航选择",
    title: "跳到数据末尾",
    detail: "按下直接定位到有数据的最后一格（右下角），快速了解表格范围。",
    shortcut: "Ctrl + End"
  },
  {
    id: "select-column",
    category: "导航选择",
    title: "选中整列 / 整行",
    detail: "Ctrl + 空格选中当前整列，Shift + 空格选中当前整行。",
    shortcut: "Ctrl / Shift + 空格"
  },
  {
    id: "next-sheet",
    category: "导航选择",
    title: "切换工作表",
    detail: "在多张工作表间快速跳转，不用鼠标点标签。",
    shortcut: "Ctrl + PgUp / PgDn"
  },

  // —— 格式设置 ——
  {
    id: "quick-format-cells",
    category: "格式设置",
    title: "打开单元格格式",
    detail: "想调整数字格式、边框或对齐，这个组合键直接打开格式设置窗口。",
    shortcut: "Ctrl + 1"
  },
  {
    id: "format-bold",
    category: "格式设置",
    title: "加粗 / 斜体 / 下划线",
    detail: "选中后 Ctrl+B 加粗、Ctrl+I 斜体、Ctrl+U 下划线，和 Word 一致。",
    shortcut: "Ctrl + B / I / U"
  },
  {
    id: "format-currency",
    category: "格式设置",
    title: "一键货币格式",
    detail: "选中数字按下，立即套用带千分位和货币符号的格式。",
    shortcut: "Ctrl + Shift + 4"
  },
  {
    id: "format-percent",
    category: "格式设置",
    title: "一键百分比格式",
    detail: "选中数字按下，立即转成百分比显示。",
    shortcut: "Ctrl + Shift + 5"
  },
  {
    id: "format-painter",
    category: "格式设置",
    title: "复制格式刷",
    detail: "点格式刷再刷目标区域即可套用格式；双击格式刷可连续刷多处。"
  },

  // —— 公式与函数 ——
  {
    id: "auto-sum",
    category: "公式函数",
    title: "一键自动求和",
    detail: "选中一列或一行数字下方的空格，按下即可自动生成求和公式。",
    shortcut: "Alt + ="
  },
  {
    id: "toggle-reference",
    category: "公式函数",
    title: "切换绝对/相对引用",
    detail: "编辑公式时选中单元格引用，反复按可在 A1、$A$1、A$1、$A1 之间切换。",
    shortcut: "F4"
  },
  {
    id: "show-formulas",
    category: "公式函数",
    title: "显示所有公式",
    detail: "想检查公式而不是结果，按下会把整张表切换成显示公式本身，再按一次还原。",
    shortcut: "Ctrl + `"
  },
  {
    id: "vlookup-xlookup",
    category: "公式函数",
    title: "查找匹配用 XLOOKUP",
    detail: "新版建议用 XLOOKUP 代替 VLOOKUP，支持向左查找、找不到时给默认值，更省心。"
  },
  {
    id: "sumif",
    category: "公式函数",
    title: "按条件求和 SUMIF",
    detail: "只想统计满足某条件的数据，用 SUMIF(区域, 条件, 求和区域) 一步搞定。"
  },
  {
    id: "iferror",
    category: "公式函数",
    title: "用 IFERROR 屏蔽报错",
    detail: "把公式包一层 IFERROR(公式, \"\")，出错时显示空白或提示，表格更干净。"
  },

  // —— 数据处理 ——
  {
    id: "format-as-table",
    category: "数据处理",
    title: "套用表格样式",
    detail: "把普通区域一键变成带筛选和样式的智能表格，方便后续引用。",
    shortcut: "Ctrl + T"
  },
  {
    id: "toggle-filter",
    category: "数据处理",
    title: "一键开关筛选",
    detail: "在数据区域内按下，即可为标题行快速添加或移除筛选按钮。",
    shortcut: "Ctrl + Shift + L"
  },
  {
    id: "remove-duplicates",
    category: "数据处理",
    title: "删除重复值",
    detail: "选中数据后，在「数据」选项卡点「删除重复值」，可按指定列去重。"
  },
  {
    id: "pivot-table",
    category: "数据处理",
    title: "数据透视表汇总",
    detail: "面对大量明细想快速分类汇总，插入数据透视表，拖字段即可出统计结果。"
  },
  {
    id: "freeze-panes",
    category: "数据处理",
    title: "冻结标题行",
    detail: "「视图 → 冻结窗格」可让首行/首列滚动时保持可见，看长表不迷路。"
  },
  {
    id: "text-to-columns",
    category: "数据处理",
    title: "分列拆分文本",
    detail: "一列里挤了多项内容，用「数据 → 分列」按分隔符或固定宽度拆成多列。"
  },

  // —— 行列与视图 ——
  {
    id: "insert-row-col",
    category: "行列视图",
    title: "插入行 / 列",
    detail: "选中整行或整列后按下可插入；Ctrl + - 则删除选中的行列。",
    shortcut: "Ctrl + Shift + +"
  },
  {
    id: "hide-row-col",
    category: "行列视图",
    title: "隐藏行 / 列",
    detail: "Ctrl+9 隐藏选中行，Ctrl+0 隐藏选中列，需要时再取消隐藏。",
    shortcut: "Ctrl + 9 / 0"
  },
  {
    id: "autofit-width",
    category: "行列视图",
    title: "自动适应列宽",
    detail: "选中列后双击列标题右边界，列宽会自动适配最长内容。"
  },
  {
    id: "zoom",
    category: "行列视图",
    title: "缩放视图",
    detail: "按住 Ctrl 滚动鼠标滚轮，可快速放大或缩小工作表显示比例。",
    shortcut: "Ctrl + 滚轮"
  },
  {
    id: "new-sheet",
    category: "行列视图",
    title: "快速新建工作表",
    detail: "一个组合键立即插入一张新的空白工作表。",
    shortcut: "Shift + F11"
  }
];

// 按当前索引取一条技巧，索引会自动回绕，越界也安全。
export function tipAt(index: number): PetTip {
  const count = PET_TIPS.length;
  const safe = ((index % count) + count) % count;
  return PET_TIPS[safe];
}
