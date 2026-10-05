# 宽幅双网格变更审阅 —— 指导文件

> 2026-10-03 退役说明：用户已要求删除网页前后对比。本文中的 Excel review 页面、WPS 全幅网页双网格及 core 网格投影模块均已移除。本文仅保留为历史记录，不得按旧指导重新接入。WPS 使用原生独立副本并排审阅；Excel 同类原生能力尚待实现。

> **2026-09-30 WPS 方向已被用户修正**：本文的「WPS Web 宽窗内双 HTML 网格」是历史探索，
> 不再是 WPS 当前实现目标。当前目标是 WPS 原生并排两个**独立工作簿**：左原表、右副本，
> 原对话仍在任务窗格；见 [WPS 原生并排审阅](./WPS_NATIVE_SIDE_BY_SIDE_REVIEW_2026-09-30.md)。
> Excel 独立 Web 审阅窗口的记录仍可参考，不能用本文旧 WPS 表格当成现状。

> 目标：让左侧宽工作区同时展示「修改前 / 修改后」两张**同坐标**表格画布，
> 右侧原对话持续可见；双表格支持同步滚动、变化高亮、跳转、有限上下文与公式栏。
> WPS 是否能用独立窗口实现这三栏同屏，仍须真机证明。
>
> **本文档的口径**：只写**已在源码里查证过**的事实，以及**已经跑出结果**的验证。
> 凡是还没实机验证的，一律标注「待实机验证」，不靠文案伪装成已具备。
>
> 起草日期：2026-09-29。对应仓库状态：批次 A/B/C 已落地（见
> `docs/DUAL_VIEW_CHANGE_REVIEW_GUIDE_2026-09-29.md`），宽幅视图为**新一批**。

---

## 0. 两宿主可行性：先验证，再动手

宽幅视图的前提是"能打开一个足够宽的窗口，并把当前会话的快照安全传进去"。
两个宿主的机制**完全不同**，必须先分别验证，**不允许**先写 UI 再倒推。

### 0.1 Excel（Office.js 当前工作簿模式）—— ✅ 可行，有同构成例

| 能力 | 证据 | 结论 |
|---|---|---|
| 独立窗口 | `apps/excel-addin/src/App.tsx:5502 openFocusWindow()` → `Office.context.ui.displayDialogAsync("focus.html", {height:93,width:100,displayInIframe:false})`（`5531`） | ✅ 已有 `focus.html` 全套先例，改个页面即可 |
| 会话快照传递 | `DialogApi 1.2` 能力检测 `canMessageChild`（`5519-5523`）→ `dialog.messageChild()` 主通道；`DialogMessageReceived` 处理 `focus-ready`（`5551`）/`focus-close`（`5558`） | ✅ 双向消息通道现成 |
| 降级通道 | `localStorage` + `FOCUS_PAYLOAD_STORAGE_KEY = "excel-bro.focus.v1"`（`5515`） | ⚠️ **本批不得复用**（见 §2 约束一） |
| 宽窗口样式 | `apps/excel-addin/src/focus.css` —— `body{min-width:720px}` | ✅ 可直接沿用同款做法 |

**结论**：Excel 侧**可行**，无需新宿主 API，只需新增一个 `review.html` 页面 + 侧栏入口。

### 0.2 WPS（WPS 加载项）—— ⚠️ 待实机 PoC 实测

**之前的错误结论已修正**：早前把 WPS 判为"不具备独立宽窗口"，理由是
`apps/wps-addin/src/wpsHost.ts` 里的 `WpsApplication` **类型**没有声明 `displayDialogAsync`。
这个理由不成立 —— **类型未声明 ≠ 宿主没有这个能力**。WPS 官方在基础接口里提供了
`Application.ShowDialog` 与 `CreateWebDialog`，必须**实机探测**才能下结论。

**已就绪的 PoC 基础设施**（只读、不含任何工作表数据，可在任何工作簿安全运行）：

| 文件 | 作用 |
|---|---|
| `apps/wps-addin/src/wpsHost.ts` | 当前 WPS 官方文档的 `ShowDialog` 返回 `boolean`、宽高使用**逻辑像素**；兼容性类型允许旧宿主返回句柄，但不能把 `true` 当句柄 |
| `apps/wps-addin/src/wpsWideViewPoc.ts` | `probeWideViewDialog(application)` 三态探针：`requested` / `missing` / `error`。先试 `ShowDialog`，未接受时才试 `CreateWebDialog`；**不自动关闭**，让人实际观察窗口宽度与关闭行为 |
| `apps/wps-addin/poc-dialog.html` + `src/pocDialog.tsx` + `src/pocDialogEntry.tsx` | 只读测试页：**不加载工作簿数据、不访问 `ActiveWorkbook`**，只量自己的 `innerWidth/devicePixelRatio` 并显示出来 |
| `apps/wps-addin/src/WideViewPocPanel.tsx` | 任务窗格内探针面板，路由 `#/wide-view-poc`（接线见 `src/main.tsx:23`） |
| `apps/wps-addin/src/WpsAgentPanel.tsx` | **推荐实机入口**：开发模式下原聊天窗格「更多 → 探测独立宽窗口」，结果只活当前 React 内存；同时显示与弹窗握手的状态 |
| `apps/wps-addin/src/wpsWideViewPoc.test.ts` | 单测锁定请求结果不等于实机可用、逻辑像素与不自动关闭等契约 |
| `apps/wps-addin/src/wpsWideViewChannelPoc.ts` | 随机 nonce 的 `BroadcastChannel` 双向握手；**仅发 nonce，不传单元格或聊天内容**，是否可跨 WPS WebView 通信仍须实机验证 |

**已在自动化里钉死的行为**（`wpsWideViewPoc.test.ts`，7 条全过）：
- 无 API → `missing`，`error` 里必须出现 `"ShowDialog"`；
- `ShowDialog` 返回 `true` → `requested`、`openedWith="ShowDialog"`，但没有句柄也没有自动关闭；标题含"变更审阅"、`bModal=false`、宽高 = 目标逻辑像素；
- `ShowDialog` 抛错 → `error`，保留原始错误文本；
- 旧宿主若返回句柄也不自动关闭，`closed=null` 不得解释成"无法关闭"；
- `ShowDialog` 返回 `false` → 继续尝试 `CreateWebDialog`，得到可显示句柄才尝试 `Visible`/`Show`；
- 宿主版本被如实回报（失败时用来定位"哪个版本缺这个 API"）。

2026-09-30 纠错：早期探针把 `ShowDialog` 的 `boolean` 返回值当句柄，且把官方所说的
逻辑像素又乘 `devicePixelRatio`，两处都会把实机结论带偏。现已依据
[WPS 官方 `Application.ShowDialog` 文档](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/addin-api/Application/member/ShowDialog)
修正。`requested` 只说明请求被宿主接受，**不是**独立宽窗口可用的验收结论。

### 0.3 PoC 实测记录（部分已填）

> **这一节只填真实 WPS 截图能证明的事实。**宽度已实测；跨窗口通信与目标布局仍待真机验证。

2026-09-30 收到的首次截图是在普通浏览器标签中直接打开
`127.0.0.1:3889/index.html#/wide-view-poc`，显示「当前页面没有连接到 WPS 表格宿主」。
该截图只证明**浏览器不是 WPS 任务窗格**，不能算 `ShowDialog` 失败。
因此新增开发模式下从真实聊天窗格「更多」直接启动探针的入口；待按此入口重测。

同日第二张真机截图证明：从真实 WPS 任务窗格已打开只读 Web 对话框，页面报告
`innerWidth=1281 CSS px`、`innerHeight=861 CSS px`、`devicePixelRatio=2`。
这关闭了「本机宿主根本打不开宽 Web 窗口」这一阻碍，**没有**证明目标三栏布局：
当前浮窗居中覆盖了几乎整个原聊天窗格，右侧只露出窄边；关窗返回尚无截图证据。
这不是目标布局的阻碍，因为用户要的是
**同一个宽视图内部的双表格＋右侧对话**，不是把浮窗与原任务窗格硬拼在一起。
因此下一关卡转为：先验证不传数据的跨窗口握手，再设计不落盘、不经 URL 传矩阵的
会话内通道，并在宽窗口内复用完整聊天交互；无通道证据前不得把结果卡接到新窗口。
已接入随机 nonce 的 `BroadcastChannel` 握手探针，**自动化通过不等于 WPS 真机连通**。

- 宿主版本：
- 实际探测到的 API：
- 实际打开方式（`ShowDialog` / `CreateWebDialog`）：
- 窗口句柄成员（`ShowDialog` 官方返回值为布尔值，不应伪造句柄）：
- 实测宽度：1281 CSS px；高度：861 CSS px；DPR：2。
- 弹窗与任务窗格的「跨窗口内存通信」状态（两侧都应显示「已双向连通」）：待实机复核。
- 能否关闭并返回任务窗格：
- 现象/异常原文：

**判定规则**：
- `requested` 且人眼确认窗口确实显示、CSS 宽度达到可用范围 → **仅证明宽窗口可打开**；
  再确认弹窗与原任务窗格两侧均显示「已双向连通」、能关闭返回，才可继续开发
  会话内传输。最终还必须在**同一宽窗口内**验收「双表格＋右侧对话」同时可见且可操作；
  普通浏览器标签页或独立的 `#/wide-view-poc` 路由不能提供宿主证据；
- 握手失败 → 记录弹窗和任务窗格的两侧状态及宿主版本；**不能**把普通浏览器中
  `BroadcastChannel` 可用的自动化结果当作 WPS WebView 连通证据。保留已明确标注的
  「窗格内全幅」过渡入口，但不得称它为目标三栏布局；
- 探针失败时，**必须如实报错**，不许偷偷回退到 `localStorage` 快照或窗格内浮层 ——
  那等于把"没有这个能力"粉饰成"能打开"。

> 2026-09-30 后续决定：用户已明确接受**先实施任务窗格内全幅双网格**，
> 此决定不依赖 PoC 成败，亦不改变上面针对「独立窗口能力」的实机取证要求。
> 当前实现不是偷偷回退：界面明确写「窗格内全幅」，独立窗口仍标待实机验证。

**怎么跑**：
```
npm run dev:wps
```
在 **WPS 表格**中点击功能区「Excel Bro → 打开 Excel Bro」，确认真实聊天窗格已出现；
然后点聊天头部「更多 → 探测独立宽窗口」。**不要**把 `127.0.0.1:3889` 直接粘到
Edge/Chrome 地址栏——普通浏览器没有 WPS 宿主对象，只会报「宿主不可用」，这不是
宽窗口能力的测试结果。把面板上的
**宿主版本 / 存在的 API / 实际打开方式 / 句柄成员**，以及宽窗口显示的
**CSS 宽度 / 两侧的跨窗口内存通信状态 / 能否手动关闭并返回**如实回报。
`#/wide-view-poc` 仍保留供开发者排查，但它会替换聊天内容，不用于同屏验收。
vite 控制台会同步打印 `[wps-bridge]` 探针日志（通道见 `apps/wps-addin/vite.config.ts`
的 `/__wps_probe` 处理）。

---

## 1. 宽视图的形态

> **2026-09-30 实机纠偏：当前 WPS 实现不符合本节最终目标。** 用户截图证明
> `ReviewWorkspace` 仍被关在右侧任务窗格内，左右两张只是狭窄的 E 列小网格；
> 左边仍是宿主原工作表。先前把「占满任务窗格」称作「双视图已实现」是错误结论。
> 用户确认目标是：**左侧宽工作区并排两张带行列与已授权上下文的表格画布，
> 右侧对话区同时保留**。图 2 是空间结构参考，不代表可以复制其写入按钮或
> 把别的产品的能力视作 WPS 宿主事实。

### 1.1 聊天侧栏 —— 只留摘要 + 入口

侧栏**不再渲染逐格表格**。保留：
- 一句摘要（复用 `changeReviewSummary(review)`：定位 + 计数，**不含任何单元格内容**）；
- 一个「打开对照视图」按钮。

现有 `ChangeReviewTable` 的渲染点（`apps/excel-addin/src/App.tsx:8134-8143`）改为渲染摘要 + 入口。

### 1.2 宽视图 —— 两张同坐标网格

- 左右两张网格：**修改前** / **修改后**；
- **同坐标**：第 i 行第 j 列在两边的地址必须相同（行号列 + 列表头都要对齐）；
- 同步滚动：左右两格网格共用一个纵向滚动，横向同步；
- 变化高亮：改动格在两边的对应位置都高亮；
- 跳转：从摘要/列表点某处 → 两侧同时滚到该格；
- 有限上下文：每处改动展示**上下左右各 N 格**的上下文，`N` 有明确上限（见 §4）；
- 公式栏：底部固定一条公式栏（见 §1.3）。

### 1.3 长公式不进单元格

- **单元格内只渲染值**，长公式**不得**在单元格里全文换行撑破行高；
- 单元格内公式一律 `nowrap` + 溢出截断（截断只影响显示，值本身不丢）；
- 选中公式格后，**固定公式栏**显示三项：
  1. 完整公式原文；
  2. 计算值；
  3. 该格若被本次写入覆盖原公式 → 「将覆盖公式 / 已覆盖公式」提示。
- 值未采集时：值栏显示 `"显示值未采集"`（来自 `cellDisplayText()`，`unresolved=true`），
  **绝不**把公式原文塞进值栏顶替。

### 1.4 已写入视图没有任何"保存/放弃"

- 计划预览阶段（`status: "ready"`）标**「预计、尚未写入」**；
- 执行结果阶段（`applied` 等）**只使用实际回读**，不得照搬计划值；
- 已写入视图里**不得出现「保存」「放弃」按钮** —— 宽视图是**只读**的。

### 1.5 修正后的桌面布局验收线（WPS / Excel）

```text
┌──────────────────────────────宽工作区───────────────────────────┬────对话区────┐
│  修改前：带行号/列标的表格画布  ║  修改后：同坐标的表格画布       │  原对话继续可见 │
│  已授权且实际采集的周边行列     ║  已授权且实际回读的周边行列      │  进度、摘要、输入 │
│  同步滚动、同格选中、变更高亮   ║  同步滚动、同格选中、变更高亮    │  均不被遮住     │
└─────────────────────────────────────────────────────────────────┴─────────────┘
```

- **不是**右侧任务窗格里再把 E2:E6 分成左右两张小表；主工作区必须有足够宽度，
  且对话区同屏。若 WPS Web 对话框盖住对话区，此路径也**不算通过**。
- 真实来源：`RangeChangeReview.rows` 目前只含**目标矩形**的最多 `maxRows` 条明细，
  不含 A:D 等相邻列；要达到图 2 的上下文，必须在用户授权内增加受限的
  **修改前上下文快照**与**修改后实际回读快照**，并受 `capabilities.changeReview`
  行列/格数上限约束。不能把宿主当前表当作「修改前」，不能凭计划值伪造「修改后」。
- 视口外可虚拟化；未采集区域要明确标「未采集」，不能画成空单元格。
- 宿主可行性分两道取证：① 能否开真正足宽的非模态 Web 窗口；② 它是否能与
  右侧任务窗格对话**同时可见且不遮挡**。只有第一道通过，不代表满足用户目标。
- 在上述布局与快照边界实机验收之前，保留现有窄窗入口作为**过渡实现**，
  不得再在交付报告里称其为目标双视图已完成。

---

## 2. 两条硬约束（本轮新增，优先级最高）

### 约束一：完整矩阵不得走 localStorage 降级

`focus` 窗口的 `localStorage` 降级（`FOCUS_PAYLOAD_STORAGE_KEY`）**不得照搬**到宽幅审阅：
那里装的是**含原始单元格值的完整对照矩阵**，落 `localStorage` 等于把工作表内容持久化。

- 消息传递可用 → 只走会话内存通道（Excel: `dialog.messageChild()`；WPS: 视 PoC 结果定）；
- 消息传递**不可用** → **先如实报错**，不静默降级。
- 与既有边界一致：完整矩阵不进 `WpsUndoRecord`（落 PluginStorage）、不进 `ChatMessage`（落 localStorage）、
  不进任何出站请求；只在会话内内存里活着。

### 约束二：`ChangeRow[]` 只有改动格，不得虚构上下文

`buildRangeChangeReview()` 产出的 `ChangeRow[]` 是**目标矩形内**的逐格明细
（包括未改动格，且会按 `maxRows` 截断），**不携带目标矩形外**相邻单元格的值。因此：

- **不允许**凭它"脑补"上下左右相邻格 —— 那会画出并不存在的单元格内容；
- **不允许**把"未采集"的格子画成空白 —— 空白会被误读成"这里是空的"；
- 二选一：
  1. **使用已采集的受限上下文** —— 即真正再读一次目标区域周边的有限值（读取须走本地确定性读取、
     受 `capabilities.changeReview` 限额约束、且仍在用户勾选的工作表内）；
  2. **明确显示稀疏变更网格** —— 只画改动格，并**显式说明**"仅显示本次改动单元格，未展示相邻上下文"。

两种都不做，就等于欺骗用户。

---

## 3. 核心层改造清单

### 核心 A —— `packages/core/src/changeReviewView.ts`（新增）

纯函数，无宿主 API：

- `toGrid(rows: ChangeRow[]): ChangeGrid` —— 把稀疏的 `ChangeRow[]` 还原成
  `{minRow, maxRow, minCol, maxCol}` 的坐标网格（**只含改动格**，不补空）；
- `alignGrids(before, after): AlignedGridPair` —— 左右同坐标对齐（同一地址在两边的坐标必须一致）；
- `highlightSet(rows)` —— 变化高亮集（按地址索引）；
- `contextWindow(rows, {up, down, left, right, maxCells})` —— 有限上下文窗口，
  **上限有明确数字**（见 §4），超出即 `truncated`；
- `sparseMode(rows)` —— 无已采集上下文时的稀疏模式标记（UI 据此显示说明文案）。

**要求**：`toGrid` 与 `contextWindow` 都**不许**凭空产生行/列；没有采集到的格子就是"不存在"，
不是"空"。

### 核心 B —— 单元格渲染与公式栏

- 单元格组件：只渲染值，`nowrap`，超长截断；
- 公式栏组件：显示完整公式 + 计算值 + 覆盖提示；
- 二者共用 `cellDisplayText()`，"显示值未采集"文案统一来自 core。

---

## 4. 限额与虚拟化

限额**优先放** `config/capabilities.json` 的 `changeReview` 段（现有：
`maxRows:200`、`maxTotalCells:20000`、`captureMaxCells:20000`、`previewMaxRows:5`、
`previewMaxColumns:5`），**不放前端硬编码**。

宽视图新增需要的上限（**待定稿，须写进 capabilities**）：
- 上下文半径：上/下/左/右各几格 → 需明确数字，且有 `maxCells` 兜底；
- 视口外不渲染（虚拟滚动），只渲染可视行；
- **不扫描整本工作簿**：读取范围一律 ⊆ 本次已写入的目标区域 ⊆ 用户勾选的工作表；
  上下文只在目标区域**周边有限扩展**，不得扩散到整表。

超限时：`truncated = true`，顶部一句「已展示 X / Y，其余 Z 项省略」，**不静默裁掉**。

---

## 5. 两宿主的接线差异

| | Excel | WPS |
|---|---|---|
| 打开宽视图 | `review.html` + `ReviewWorkspace.tsx` + `review.css`，`displayDialogAsync` 打开（`displayInIframe:false`） | 用户于 2026-09-30 明确同意先做**任务窗格内全幅双网格**；独立宽窗口仍待 PoC 实机结果，不能混称 |
| 传递快照 | 主通道 `dialog.messageChild()`（`DialogApi 1.2`）；**不用** localStorage 降级 | 同一个 React 会话内由 `changeReviewsRef` 直接传给共用 `ReviewWorkspace`；**不用** URL/hash、聊天历史或 localStorage 装矩阵 |
| 撤销语义 | Excel 独立撤销（`undo.maxEntries`）不变 | WPS 独立撤销（`undoWpsExecution`）不变 |
| 视觉 checks | `Office.context.requirements.isSetSupported` → Excel 式能力勾选 | **不得伪造** Excel 式 checks（见 §6） |

---

## 6. 明确不做

- **不为视觉对齐伪造 WPS 的 Excel 式 checks** —— WPS 没有 `isSetSupported` 那套，
  不要为了跟 Excel 长得一样就捏造一个。
- 不改 `AnalysisPlan` 白名单、不动授权边界、不放开任何 `unverified` 动作。
- 不新增脚本 / VBA / 宏 / 外部程序执行。
- 不把完整前后矩阵写进 URL、日志或聊天历史。
- **不持续轮询工作簿**；**不读整本工作簿**。
- 不承诺跨关闭可撤销。
- 不提交 Git、不生成安装包、不操作真实工作簿。

---

## 7. 验收

### 7.1 自动化（本批要求）

- core：`changeReviewView` 纯函数单测（含"不虚构上下文""稀疏模式"两条专门用例）；
- Excel：`ReviewWorkspace` 组件用例 + 侧栏摘要入口用例；
- WPS：先验收用户同意的窗格内全幅双网格（摘要入口、双网格、公式栏、返回对话、隐私）；独立窗口仍按 PoC 结果另行评估；
- 回归：`npm run test:core` / `test:addin` / `test:wps` / `build:addin` / `build:wps` /
  `python -m pytest server/tests -q` / `python -m compileall -q server/app`。

### 7.2 实机验收（须单独报告，不得与自动化混为一谈）

- Excel 实机：打开宽视图、左右同坐标、同步滚动、高亮、跳转、公式栏、长公式不撑破单元格；
- WPS 实机：先验收窗格内全幅的左右同坐标、同步滚动、公式栏、返回对话；独立窗口验收仍以 PoC 通过为前提；
- **两端各提供相同状态的截图**；
- 「自动化通过」与「实机验收通过」**分开报告**，不许用前者替后者背书。

### 7.3 关键用例

| # | 场景 | 断言要点 |
|---|---|---|
| 1 | 长公式不进单元格 | 单元格内不出现换行撑高的公式原文；公式只在底部公式栏完整显示 |
| 2 | 同坐标对齐 | 改动格在两边的行列坐标一致；行号列/列表头对齐 |
| 3 | 稀疏模式 | 无已采集上下文时只画改动格，且**显式说明**未展示相邻上下文；不出现被脑补的相邻格 |
| 4 | 显示值未采集 | 值栏文案为"显示值未采集"，且不含公式原文 |
| 5 | 计划预览 | `status:"ready"` → 标"预计、尚未写入"，不出现"成功/已验证" |
| 6 | 实际回读 | 执行结果显示的是**回读值**，不是计划值（构造二者不同时断言取回读） |
| 7 | 已写入视图 | DOM 里不存在「保存」「放弃」按钮 |
| 8 | 授权边界 | 读取范围 ⊆ 已写入区域 ⊆ 勾选工作表；不扫描整本工作簿 |
| 9 | 限额 | 超 `maxRows`/上下文上限 → `truncated=true` 且提示带真实 X/Y |
| 10 | 隐私 | 完整矩阵不出现在 localStorage 快照、ChatMessage、任何请求体里 |

### 7.4 当前状态（2026-09-29 实测记录）

**自动化 —— 已通过：**

| 套件 | 命令 | 结果 |
|---|---|---|
| core | 在 `packages/core` 下跑 vitest | 9 文件 / **132 passed**（含 `changeReviewView.test.ts` 15 条） |
| Excel | 在 `apps/excel-addin` 下跑 vitest | 54 文件 / **554 passed**（含摘要卡 5 条） |
| WPS | 在 `apps/wps-addin` 下跑 vitest | 16 文件 / **180 passed**（含探针 6 条） |
| server | `pytest server/tests -q` | **577 passed** |
| 构建 | `build:addin` / `build:wps` | 均通过（`tsc --noEmit` 干净，`review.html` 已进产物） |

> ⚠️ **vitest 必须在各自 workspace 目录下跑**。从仓库根运行会因相对路径与
> test 环境差异炸出 25 个**假失败**，那不是真回归。

**实机验收 —— 尚未进行：**

- Excel 实机：**未做**（需在真实 Excel 打开 `review.html` 验证）。
- WPS 实机：**未做**。窗格内全幅已由用户明确接受并接线，仍需在真实 WPS 检查视觉与交互；独立窗口依旧必须先跑 §0.3 的 PoC。
- 两端同状态截图：**未提供**。

**本批实际交付的代码：**

- core：`packages/core/src/changeReviewView.ts`（新）+ `changeReviewView.test.ts`（新）+ `index.ts`（导出）。
- Excel：`review.html`、`src/review.tsx`、`src/ReviewWorkspace.tsx`、`src/review.css`、
  `src/reviewState.ts`（均为新）、`src/components/ChangeReviewSummaryCard.tsx`（新）+ 测试，
  `src/App.tsx`（改：`openReviewWindow` + 渲染点换成摘要卡）、`vite.config.ts`（改：加 review 入口）、
  `src/styles.css`（改：摘要卡样式）。
- WPS：`src/wpsWideViewPoc.ts`、`src/WideViewPocPanel.tsx`、`src/pocDialog.tsx`、
  `src/pocDialogEntry.tsx`、`poc-dialog.html`、`src/wpsWideViewPoc.test.ts`（均为新），
  `src/wpsHost.ts`、`src/main.tsx`、`src/styles.css`（改）。
  2026-09-30 补充：WPS 已从消息内窄表改为仅含定位/计数的摘要卡；点击后由共用
  `ReviewWorkspace` 占满当前任务窗格，返回后仍是原对话。`changeReviewsRef` 中的实际回读
  矩阵只在内存里传递，撤销后仍释放明细。`wpsFullPaneReview.test.tsx` 验证绝对坐标、
  稀疏格「未采集」、公式栏与多区域切换；`WpsAgentPanel.test.tsx` 验证真实入口、
  失败状态与不落盘。**自动化已通过，宿主实机待验；这不是独立宽窗口。**

**2026-09-30 本轮验证（与上面的 09-29 历史表分开）：**
`test:core` 171/11、`test:addin` 581/58、`test:wps` 240/21，
`build:addin`、`build:wps` 均通过（含 TypeScript 检查）。
返回对话恢复原滚动位置也有 WPS 面板用例。未改服务端代码，本轮未重跑 pytest；
WPS/Excel 宿主实机视觉与同步滚动尚未验证，不能由 jsdom 绿灯代替。

---

## 8. 参考（本文档引用到的源码位置）

- Excel focus 窗口先例：`apps/excel-addin/src/App.tsx:5502-5574`（`displayDialogAsync` 在 `5531`，
  `DialogApi 1.2` 检测 `5519`，localStorage 降级 `5515`）；
- Excel 侧栏 `ChangeReviewTable` 渲染点：`apps/excel-addin/src/App.tsx:8134-8143`；
- 共用组件：`apps/excel-addin/src/components/ChangeReviewTable.tsx`；
- core 数据模型：`packages/core/src/changeReview.ts`；
- WPS 探针：`apps/wps-addin/src/wpsWideViewPoc.ts`、`src/WideViewPocPanel.tsx`、
  `src/pocDialog.tsx`、`poc-dialog.html`、`src/main.tsx:23`；
- WPS 宿主类型：`apps/wps-addin/src/wpsHost.ts:53-88`；
- WPS 任务窗格宽度：`apps/wps-addin/src/ribbon.ts:4-5`（`DEFAULT_TASK_PANE_WIDTH=1080`）；
- 限额：`config/capabilities.json` 的 `changeReview` 段。

> 行号对应当前工作区状态；改动后请同步更新本文档，**不要引用已经漂移的旧行号**。
