# WPS 专注窗口（2026-11-03）

> 目标：从 WPS 输入框左下角 EB 菜单点「专注窗口」，打开真实独立窗口查看当前工作簿对应的
> 对话、结果与工具说明，能关闭返回原任务窗格。以 Excel 现有专注窗口为能力基准，**不**新增
> 第二套模型调用、执行器或写入入口，不扩展成双表审阅系统。

## 1. 现状核对（动手前）

| 层 | Excel（已完整） | WPS（本轮前） |
| --- | --- | --- |
| 共享协议 `packages/core/src/focusProtocol.ts` | 共用 | 共用（信封/校验/子状态机齐全） |
| 共享展示 `FocusWorkspaceView` | 共用 | 共用 |
| 子窗口页面 | `FocusWorkspace.tsx` + Office `DialogApi` | `WpsFocusWorkspace.tsx` + `wpsFocusChannel.ts`（**已存在，无测试**） |
| 父窗格入口 | `openFocusWindow` 已接线 | **菜单按钮 `disabled`，无任何父窗格逻辑** |

结论：WPS 缺的是**父窗格侧**（打开、通信、心跳、关闭、切换检测）+ **生产构建多页** +
**通道层与父窗格自动化测试**。这三处是本轮交付。

## 2. 实际使用的 WPS 窗口 API 与通信方式

- **打开**：`Application.ShowDialog` 优先，`Application.CreateWebDialog` 兜底（`wpsHost.ts`
  的 `WpsShowDialog` 类型）。`ShowDialog(url, caption, width, height, bModal)` 当前官方签名
  返回 `boolean`；`CreateWebDialog` 可能返回句柄。**`boolean true` 只代表请求被接受**，不当作
  窗口句柄，也不当作"窗口已显示"。
- **通信**（`wpsFocusChannel.ts`）：WPS 没有 Office 的 `dialog.messageChild`，父子窗口是同一
  加载项源下的两个 WebView。通道按顺序尝试、宿主自报名字：
  1. `BroadcastChannel`（channel 名带 sessionId，非固定频道名）；
  2. 同源 `localStorage` 邮箱（按 sessionId 隔离、带序号与发送时间、250ms 轮询）。
  无论哪条，消息都是共享协议信封（`channel`/`version`/`kind`/`sessionId`/`seq`/`sentAt`/
  `workbook`/`snapshot`），缺一不可，校验失败拒收并记录原因。
- **就绪**：只有子窗口 `hello` → 父窗格回 `snapshot`（带会话标识）→ 父会话进入 `ready` 才算
  握手完成。`focusOpening` 只表示"正在打开"。

## 3. 修改的文件

| 文件 | 改动 |
| --- | --- |
| `apps/wps-addin/src/WpsAgentPanel.tsx` | 新增父窗格侧：`focusOpening`/`focusSessionRef`/`focusDialogRef`/`focusWorkbookRef` state、`closeFocusWindow`/`buildFocusSnapshot`/`openFocusWindow`、工作簿切换 `closing` effect、卸载收尾 effect；菜单「专注窗口」从 `disabled` 改为 `onClick=openFocusWindow` |
| `apps/wps-addin/vite.config.ts` | `build.rollupOptions.input` 增加 `focus` 与 `poc-dialog`，多页构建让 `focus.html` 进生产产物 |
| `apps/wps-addin/src/wpsFocusChannel.test.ts` | 新增：BroadcastChannel 握手、邮箱回退、通道不可用、握手超时、失联 tick、closing、非法消息拒绝 |
| `apps/wps-addin/src/wpsFocusIntegration.test.tsx` | 新增：菜单打开、连续点击去重、无对话框能力、ShowDialog 抛错保留对话、零隐式审阅处置、卸载清理通道 |
| `docs/ARCHITECTURE.md` | 补「WPS 专注窗口（父窗格接线）」小节 |
| 本文件 | 验收记录 |

**未改动**：`packages/core/src/focusProtocol.ts`（协议不变，无 TS/Python 定义需同步）、
`FocusWorkspaceView.tsx`、`WpsFocusWorkspace.tsx`、`wpsFocusChannel.ts`（通道层原实现保留，
只补了测试）、`AppRoutes.tsx`（`focus.html` 是独立入口，不走 hash 路由，专注窗口与 DEV 诊断页
不冲突）。

## 4. 测试与构建

| 项 | 结果 |
| --- | --- |
| `wpsFocusChannel.test.ts` | 7 通过 |
| `wpsFocusIntegration.test.tsx` | 6 通过 |
| WPS 全量 `npm run test:wps` | 见构建日志（新增 13 条，其余未破坏） |
| Excel 全量 `npm run test:addin` | 见构建日志（共享层未改，应无破坏） |
| `npm run build:wps` | 通过，`dist/focus.html` 与 `dist/poc-dialog.html` 均已产出（相对路径 `./assets/…`） |
| `tsc --noEmit`（WPS） | 通过 |

## 5. 已实现 / 未实现 / 实机待验

### WPS

**已实现（自动化证据）**
- 菜单「专注窗口」接通，`ShowDialog`/`CreateWebDialog` 打开 `focus.html?session=…`。
- 父/子握手、快照同步、心跳、失联（`stale`）、`closing`/`closed`、非法消息拒绝。
- 连续点击只 `sync` 不重建；卸载清理通道；工作簿切换 `closing` 收尾。
- 零模型调用、零工作簿写入、零隐式审阅处置。
- 生产多页构建产出 `focus.html`。

**实机待验（当前环境无法做实机，以下必须在真实 WPS 里逐项过）**
1. `Application.ShowDialog` / `CreateWebDialog` 是否真的在当前 WPS 版本存在、能否打开到
   1100×760 可用尺寸（`wpsWideViewPoc.ts` 的探针只到"requested"，不等于真机已显示）。
2. `BroadcastChannel` 是否能**跨两个 WPS WebView** 投递（浏览器里能通 ≠ 宿主里能通）；
   不能时同源 localStorage 邮箱是否被两个 WebView 共享。
3. 打开 → 内容正确 → 同步 → 关闭返回 → 再次打开的全流程。
4. 父窗格刷新/关闭/切换工作簿后，子窗口是否明确进入失联/过期，不冒充当前现场。

### Excel

**已实现**：Office `DialogApi 1.2` 完整专注窗口（打开、`messageChild` 通信、心跳、关闭、
DialogApi<1.2 的会话隔离快照兜底）。本轮未改 Excel 侧任何文件。

**未实现 / 实机待验**：本轮无关；Excel 专注窗口的既有真机验证状态沿用
`docs/DEVELOPMENT.md` 的记载（`DialogApi 1.2` 可用/不可用两条路径）。

## 6. 剩余问题

- **跨 WebView BroadcastChannel 未实机验证**：这是 WPS 专注窗口最不确定的一点。若宿主不支持，
  邮箱通道是唯一回退，而它依赖"两个 WebView 共享同源 localStorage"——这同样需要实机确认。
  两条都不可用时，`startWpsFocusParentSession` 会如实报 `transport-unavailable` 并在任务窗格
  说明，子窗口停在"未连接"，**不会**伪装成已同步。
- **ShowDialog 关窗能力**：`WpsWebDialog.Close` 只在 `CreateWebDialog` 返回句柄时可用；
  `ShowDialog` 返回 `boolean` 时没有句柄可关，`closeFocusWindow` 只能靠 `window.close()` 或
  用户手关（`WpsFocusWorkspace` 的关闭按钮已处理"关不掉就停在已结束"）。实机需确认
  `ShowDialog` 场景下点「返回任务窗格」能否真正关窗。
- 上述两项**未实机 = 未验证**，不写"仅待实机"含糊过去，也不宣称整体已完成。

## 7. 实机验收步骤（真实 WPS）

1. `npm run build:wps` 后按 `docs/DEVELOPMENT.md` 部署加载项并刷新任务窗格。
2. 输入框左下角 EB 菜单 → 点「专注窗口」→ 确认独立窗口打开、尺寸可用、顶部显示
   "已同步"与工作簿名。
3. 在任务窗格发一条消息，确认专注窗口随后同步；窗口内切换「对话/工具」与不同历史对话。
4. 点「返回任务窗格」关窗，确认原任务窗格仍可用；再次打开正常。
5. 打开专注窗口后，在任务窗格切换工作簿或刷新窗格，确认专注窗口明确显示失联/已结束。
6. 全程确认：没有模型被调用、工作簿未被写入、待审阅结果未被自动确认或撤销。
