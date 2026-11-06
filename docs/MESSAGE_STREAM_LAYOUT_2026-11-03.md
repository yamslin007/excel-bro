# 消息流排版：一段回复一颗抬头（2026-11-03）

> 起因是一张并排截图：左 ExcelMaster、右 Excel Bro。用户的判断是「EB 的上下左右留白间隙太大」，
> 并在第一次只调数值之后回了一句「你没有感觉不一样吗」。这份记录说明**看到的差异是什么**、
> **改在哪里**、**哪些还没验**。

## 1. 症状（截图上能指出来的三件）

| # | 现象 | 根因 |
| --- | --- | --- |
| 1 | 同一段执行回复里出现**两条**「Excel Bro」 | 引擎把一段回复拆成多条消息落盘（过程卡一条、结论一条、结果卡一条），渲染层**每条**都印一次抬头 |
| 2 | 过程条是一条**通栏空条**（左边「已完成 · 381ms」，右边「查看过程」，中间全空） | 三步骤执行不是 `isCompactMessageProcess`，于是那不省抬头的分支正好落在过程行上，过程行与结论行各占一个视觉块 |
| 3 | 正文、卡片、气泡左右贴边不齐 | 卡片有 12~14px 内边距，正文没有；同一列里两种左边界差 12px |

用户的词是「留白间隙太大」，但把留白**量下来**发现方向是反的（同一张截图，scale ≈ 1.76，1px 边框 = 2 原生像素）：

| | 面板边 → 内容左界 | 卡片边框 → 卡内文字 | 正文字号 / 行距 |
| --- | --- | --- | --- |
| ExcelMaster | ≈ 21px | ≈ 22px | 更大 |
| Excel Bro（改前） | 12px | 12~14px | 14px / 1.62 |

也就是说 **ExcelMaster 的绝对留白更大**。EB 之所以显得"散"，是因为它**没有分组**：
平铺消息靠空隙分组，空隙叠加上卡片自身的内边距；而且抬头重复，一条回复被读成三件事。
所以这一轮不动"某个数值调小"，改的是结构。

## 2. 改了什么

### ① 署名判据收到共享层

唯一出处 `packages/core/src/messageRows.ts`，导出 `messageRowPresentation(messages)`，逐行给
`{ author, groupStart }`。两端（`App.tsx`、`WpsAgentPanel.tsx`）只负责把消息映射成入参、
并给出本端的徽标文案（`messageAuthorBadge`）。三条不变量写在文件头：

- **A 组内只印一次**；两种**都写明了的**徽标互相不同时才再印（`Agent` 的过程后面紧跟
  「需求确认」）。"本来没有徽标、后来才标出一个"**不算**换出处 —— WPS 的徽标口径是
  "没有过程/计划/验证时才标 Agent"，按"徽标变了就再印"会让每一对「过程行（无徽标）→
  结论行（Agent）」都多一颗抬头，正是要消除的样子。
- **B 空过程行让位**：单步消息理解（`isCompactMessageProcess`）没有文字、下一行就是结论时，
  名字印在结论行上（旧行为，保留；反向锚点是一条真实回归）。
- **C 换段重新署**：「继续」开新回复段（`replySegments.ts`），两段之间可能没有用户消息。
  只按"上一条是不是助手"分组，新段第一行会被当成续行而省掉名字 —— 用户看到的是一段
  凭空冒出来的进度。

`groupStart` 同时是间距的分级依据（下面的 `data-group`），所以"分组"只有一份判断。

### ② 一条内容列

`.message-stream` 左右 `12px → 16px`：正文、卡片边框、用户气泡落在同一条左边界上，
卡片自己的内边距留在**卡片里面**。**卡片的水平内边距一个都没改** —— 它与
`.activity-log-card` 共用对齐契约（`styles.test.ts`「运行态卡与已完成卡的水平内边距一致」），
两态切换时同一句话左右跳动比多 2px 留白难看得多。

### ③ 两级行距

`.message-row` `margin-bottom: 10px → 8px`，组首（`data-group="start"`）再加 `padding-top: 4px`：

- 组内（过程卡 → 结论 → 结果卡）：**8**；
- 组间（一段回复之间、用户消息之前）：**12**。

组间那 4px 用 `padding-top` 而**不是** `margin-top`：相邻块级兄弟的外边距会折叠（8 与 4 折成 8），
垫上去等于没垫。`styles.test.ts` 里有一条反向断言专防这件事被"顺手改回去"。

## 3. 验证到哪一档

| 面 | 结果 |
| --- | --- |
| 判据单测（`packages/core/src/messageRows.test.ts`） | 15 条通过：组内一次、让位、换段、徽标规则、跳过行、空数组 |
| 样式契约（`apps/excel-addin/src/styles.test.ts`） | 29 条通过（原 27 条 + 内容列/两级行距 2 条） |
| Excel 宿主用例 | 新增"这句问话之后的回复只印一次抬头"（欢迎语是独立一组，只能按最后一条用户消息之后计数） |
| WPS 宿主用例 | 新增"一段回复一颗抬头"：从 localStorage 播种 三段消息（过程卡 + 结论），验 `data-group` 与抬头数量 |
| 原生窗格（真实 Excel / WPS） | **未验**。这一轮只有 jsdom 结构与浏览器量级，没有在真实任务窗格里点过 |

### 已知既有失败（不是这次引入的）

`apps/excel-addin/src/messageFirstRouting.test.tsx` 有 4 条失败，全部落在「接回/补收事件」
家族（`running 接回处理读表回执` ×2、`继续只补收未见事件`、`continuing re-opens the local tool bridge`），
表现为点「核对任务状态」之后等不到 `收到真实读表回执后完成`。

判断依据：**只把署名渲染改回旧写法**（其余改动全部保留）后这 4 条照样失败；
把本轮改动的 5 个文件整体收进 stash 后，该文件是 12 条失败（HEAD 的 App.tsx 还缺前一轮
审阅状态持久化的实现，所以那次的绝对数字不能直接比）。结论只到"与本次排版改动无关"。

## 4. 复看

```bash
npx vitest run packages/core/src/messageRows.test.ts
npx vitest run apps/excel-addin/src/styles.test.ts
npx vitest run apps/excel-addin/src/messageFirstRouting.test.tsx -t "ordinary message"
npm run test --workspace excel-bro-wps -- -t "one author label per reply"
```

真机复看时只盯三处：① 一段执行回复里「Excel Bro」只出现一次；② 正文左边界与过程卡左边界
是否在同一条竖线上；③ 过程卡与结论之间的空隙是否明显小于它与上一条用户消息之间。
