# 主流程：澄清 → 重新授权 → 生成预览

> 适用范围：Excel（Office.js）与 WPS（JSAPI）两条通道，外加两种模型后端
> （原模型路径 / Agent SDK）。
> 目标：用户说需求 → 本端按**真实数据**判断缺什么 → 只补必要信息一次 → 保留答案 →
> 生成副本预览 → 用户确认回写。不再进入「重新确认范围 → 重复追问 → 再次失效」的循环。

## 1. 三段状态必须分开管

| 状态 | 谁产生 | 存在哪 | 不许被谁冒充 |
|---|---|---|---|
| 任务语义（原需求、已确认口径） | 用户 | `taskResume` confirmations | 不被"重新授权"清空 |
| 澄清问题（ID、归因、待答/已答） | 实现按**真实读取**结果 | 澄清消息 + 面板内存 | 不被授权动作写成"已回答" |
| 授权上下文（已提交范围、版本、宿主身份） | 用户点「确认范围」 | `authorizationRevision` / 身份指纹 | 不产生答案 |
| 执行证据（源快照、指纹、回读） | 执行时 | 内存 / 撤销记录 | 不跨授权复用 |

历史上前三段共用一个 `clarification.answer`：点一下重新确认范围，卡片就被写成
"已回答：已重新确认数据范围并继续"，既覆盖真实答案，又让写入资格看起来已具备。

## 2. 澄清必须由**真实阻塞条件**驱动

- 意图阶段只有语义与结构，没有数据行。此时只能问语义问题，**不能问"空白怎么办"**。
- 授权之后本地读到源值，才可能发现三类缺口，各自一组互斥选项：

| `reason` | 触发条件 | 选项 |
|---|---|---|
| `unmapped` | 已按声明分隔符拆开、拆出的项缺映射（归因确定） | 补充类别映射 |
| `separator_or_mapping` | 整格没命中、又像是用了未声明的分隔写法（归因不确定） | 补分隔符 / 补映射，用户选 |
| `blank` | **真的读到空白格**且空白策略未定 | 保留原值 / 清空 / 写指定值 / 写空串 |

- **没读到就不许声称有空白**：`blankCount` 只能来自真实读取。源区域读不到时报"没读到"，
  不得编造"发现 N 格空白"。
- 归因不确定的规则（`rule.clarification`）在**读源值之前**就分流，不白读一次源列。

## 3. 空白的四种处理互不冒充

| 策略 | 落地动作 | 与相邻策略的差别 |
|---|---|---|
| `keep` | 该位置**不在任何动作的 range 里** | 不是读回原值再写回；目标格的公式、值、数字格式、条件格式都不动 |
| `clear` | `clearRange` + `applyTo: "contents"` | 真空单元格（`ISBLANK()` 为真） |
| `value`（值可为 `""`） | `writeValues` 写指定值 | 写空串留下常量空字符串（`ISBLANK()` 为假） |
| `undecided` | 尚未决定 | 只有真的读到空白才升级为澄清；静默折成默认标签被禁止 |

协议用结构化 `BlankPolicy`（`apps/excel-addin/src/contracts.ts` ↔ `server/app/models.py`）。
旧 `blankLabel` 非空串等价于 `{kind:"value"}`，语义无损；旧协议里空串本就被拒，
折成 `undecided` 不改变任何曾经合法的旧计划。

跳过位置过多时按 `capabilities.columnMapping.maxSplitActions`（64）拒绝并说明原因，
**不降级成写空值**。WPS 的公式通道只在目标是一整块连续写入时启用 —— 公式覆盖区域内
每一格，表达不了"这一格别动"。

## 4. 重新授权：保留答案，不重问已解决的问题

1. 授权版本变化（含**收窄**）→ 丢弃旧快照、旧计划与写入资格，`pending` 澄清失效。
2. 用户明确重新确认范围 → `authorizationRevision` 递增；
   - 有真实答案 → 保留答案，状态推进到 `resolved`（`clarificationAfterReauthorization`）；
   - 没有答案 → 落到 `continued`，**绝不编造答案**；
   - 说明文案 `REAUTHORIZATION_NOTICE` 只进消息流，不写进任何卡片的 `answer`。
3. 保留原需求与已确认口径，重新读取并生成**新预览**；写入前仍要再确认一次。
4. 不要求旧宿主对象引用相等 —— 新包装对象是 `unknown`，不是"用户换了表"。
5. 存在未核实的历史写入时，走既有只读核对，**不得**通过重授权直接重写。

## 5. 模型与本端工具的分工

- 模型：映射、拆分、优先级这类**语义**判断。
- 本端：真实读取、异常发现、确定性计算、复核。
- 只发送隐私策略允许的摘要；原始行与目标矩阵不出设备、不落盘。
- 结构化澄清同时覆盖原模型路径与 Agent SDK 路径 —— 两条后端共用同一份
  `map_column_values` 工具契约（SDK 从 `excel_agent` 导入，不另起炉灶）。

## 6. 证据入口

- `packages/core/src/blankPolicy.ts`、`columnMapping.ts`（`splitTargetActions`）、
  `clarificationAnswers.ts`、`taskResume.ts`
- `apps/excel-addin/src/columnMappingPlan.ts`、`components/LocalMappingResolution.tsx`、
  `App.tsx`（`recordExcelBlankPolicyGap` / `resolveExcelLocalMapping`）
- `apps/wps-addin/src/wpsReadonlyAgent.ts`、`wpsColumnMappingPlan.ts`、
  `WpsAgentPanel.tsx`、`wpsClarificationVerdict.ts`
- `server/app/column_mapping.py`、`server/app/excel_agent.py`、`server/app/models.py`
- 共享夹具：`config/column_mapping_cases.json`（core 与 Python 同表）
- 端到端：`apps/wps-addin/src/wpsLocalPlanOutbound.test.tsx`（真实 `fetch` 拦截 +
  localStorage 负向断言 + 用户点选到新预览）

## 7. 尚未覆盖 / 实机待验

- 面板重开后"未答完的澄清"跨会话恢复：当前重开一律落到 `unconfirmed`，保留原需求与答案。
- 空白策略在「编辑计划 JSON」里只进不出。
- 真机上"有空白 → 选保留原值 → 目标格公式与格式确实未变"；真宿主返回新包装对象时
  重新授权到新预览的完整时序。
