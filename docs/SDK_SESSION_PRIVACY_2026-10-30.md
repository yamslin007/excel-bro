# U3 隐私方案（草案，待决定；2026-10-30）

> 用途：把"要不要开 SDK 原生会话持久化"这个问题变成可拍板的选项。
> **本轮不动代码**：`--no-session-persistence` 保持开启，不新增持久化，**也不删**
> 已有的 30 份转录（删不删是你的决定，不是这轮的运维动作）。

## 1. 先核实：那 30 份文件到底是什么（本机实测）

目录：`C:\Users\jia\AppData\Local\ExcelBro\ClaudeAgentSdk`
（`CLAUDE_AGENT_SDK_STATE_DIR`，`claude_sdk_agent.py:81` 创建；默认值来自启动脚本）

| 项 | 实测值 |
| --- | --- |
| 目录内文件总数 | 33（30 份 `.jsonl` 转录 + `.claude.json` + 1 个 backup + `.last-cleanup`） |
| 总体积 | 1,507,757 字节（约 1.44 MiB） |
| 转录时间跨度 | **2026-10-01 14:17 → 2026-10-05 18:52**，之后**没有新增** |
| 记录总行数 | 612 行 |
| 目录 ACL | owner `AC\jia` + `SYSTEM` + `Administrators`，**继承已切断**（仅 3 条 ACE） |

**为什么 10-05 之后就没有了**：`--no-session-persistence` 是提交 `a8ae64cb`（2026-10-07）引入的，
落在 `claude_sdk_agent.py:201` 的 `extra_args`，由 3 条测试钉死
（`test_claude_sdk_agent.py:263/282`、`test_model_connection_purposes.py:127`）。
**30 份转录全部产生在这个开关之前**，不是现在的代码写出来的。

### 1.1 内容类别（按结构统计，没有抄正文）

| 类别 | 数量 | 说明 |
| --- | --- | --- |
| `attachment` | 277 | CLI 侧渲染元数据：`cwd` / `entrypoint` / `userType` / `version` / `rendered` |
| `assistant` | 100 | 模型回复（含计划与说明文字） |
| `queue-operation` / `last-prompt` / `atis-latch` / `cost-state` / `system` | 60 / 45 / 46 / 29 / 6 | CLI 自己的运行态记录 |
| `user` | 49 | **30 条是用户原话字符串**（当时输入框里打的那句），19 条是 `tool_result` |

工具调用**只有 4 个** excel_bro MCP 工具：`map_column_values` ×9、`submit_clarification` ×7、
`submit_answer` ×2、`submit_plan` ×1。没有 `Read`/`Bash`/`Write`/`WebFetch`
（`tools=[]` + `disallowed_tools` + `strict_mcp_config` 保证）。

还有两条只读结论：

- **0 行**包含 `C:\` 形式的绝对路径 ⇒ 转录里**没有**工作簿的磁盘路径。
- **16 行**含 image 标记 ⇒ 用户当时附的截图（base64 image part）在转录里。
- **实质隐私内容在这里**：`map_column_values` 的入参是列值样本、`submit_plan` 的入参含计划里的
  单元格值 ⇒ **这 30 份文件可能含工作表数据**（不是"只有聊天文字"）。

`.claude.json`（569 B）只有 CLI 身份与版本迁移标记：`machineID` / `userID` / `firstStartTime` /
`firstStartVersion` / `pluginUsage`，**不含工作簿内容**。
`.last-cleanup` 是 ISO 时间戳 `2026-10-28T11:38:51Z`：CLI 的清理例程那天跑过，但 10-01 的转录
**仍在** ⇒ CLI 侧的保留策略并不保证把旧的收掉。

### 1.2 谁能读（实际读取路径）

| 主体 | 依据 |
| --- | --- |
| **产品服务端**（`server/app`） | **不读**。`_state_dir()` 只把它喂给 `CLAUDE_CONFIG_DIR` 和 `cwd`（`claude_sdk_agent.py:173`），并强制"仓库外绝对路径"，防止读到仓库里的 `CLAUDE.md` |
| CLI 自己 | 会碰这个目录（`.last-cleanup` 是它的痕迹）；`--no-session-persistence` 时**不写**新转录、也不 resume |
| 本机同用户进程 / SYSTEM / Administrators | **读得到**。ACL 只挡了其它普通用户 |

**一个必须写下来的缺口**：`claude_sdk_agent.py:81` 的 `path.mkdir(mode=0o700)` 在 Windows 上
**不生效**（`mode` 只影响只读位，不改 ACL）。现在那三条 ACE 是"本来就在"的，**不是产品保证**。
也就是说：换台机器、换个路径，这个目录未必还有同样的限制。

## 2. 三个选项

| 选项 | 做什么 | 拿到什么 | 代价 / 风险 |
| --- | --- | --- | --- |
| **A（建议先做，不等于 U3 完成）** | 保持 `--no-session-persistence`；补一个**显式**入口：列出/删除历史转录；README 写清"目录在哪、谁能读、留多久"；顺手把目录 ACL 显式设成同上三条 | 用户能自己收拾；"谁能读"变成产品说了算 | 不解决"会话恢复"这个能力 |
| **B** | 关掉 `--no-session-persistence`，直接用 CLI 原生 resume | 续聊最省事 | 转录**必然**落盘且含单元格值；历史由 CLI 的保留策略做主。**不建议** |
| **C** | 开持久化，但**换目录**（每次执行独立、执行完即删）+ 显式 ACL + 保留期上限（如 24h）+ 恢复走"执行级"而不是"CLI 会话级" | 有恢复能力，落盘窗口可控 | 要写清理与生命周期判据；与 `resume` 的语义边界要重新定义 |

**建议**：先做 A 的三件事（它们是 B/C 共同的前置，且都不需要新决定）；真有"续聊已关闭面板的会话"
需求时再评估 C。**B 不建议**——它把"转录含单元格值"从历史问题变成每轮都在发生的事。

## 3. 需要你拍板的三句话

1. 那 30 份历史转录：**留着**、还是**由产品提供一个删除入口后自行清掉**？（我不擅自删）
2. 目录 ACL：是否同意"产品启动时显式设置、并在设置失败时拒绝启动"（而不是像现在这样靠系统默认）？
3. 若将来要恢复能力，是否接受**C**（执行级、短保留、独立目录）而不接受 **B**？
