# 计划编辑器独立窗口方案

## 背景

当前计划编辑器在任务窗格内全屏显示，受限于窗格宽度（约 700px）。用户希望像「专注窗口」一样，打开独立的大窗口进行编辑。

## 目标

使用 Office.js Dialog API 创建独立的计划编辑器窗口，提供更大的编辑空间和更好的专注体验。

---

## 技术方案

### 架构设计

```
┌─────────────────┐                    ┌──────────────────────┐
│   App.tsx       │                    │  plan-editor.html    │
│  (任务窗格)      │ ◄──── Dialog API ───► │  (独立窗口)           │
│                 │      消息通信         │                      │
│ - planEditor    │                    │ - PlanEditorDialog   │
│ - openDialog()  │                    │ - 接收初始数据        │
│ - 接收保存结果   │                    │ - 编辑 JSON          │
│                 │                    │ - 校验 + 保存         │
└─────────────────┘                    └──────────────────────┘
```

### 通信协议

**父窗口 → 子窗口**（初始化数据）
```typescript
{
  type: "plan-editor-init",
  payload: {
    plan: AnalysisPlan,           // 待编辑的计划
    target: PlanEditorTarget,     // 编辑目标（消息/工具）
    workbook: WorkbookSnapshot    // 工作簿快照（用于指纹重算）
  }
}
```

**子窗口 → 父窗口**（就绪信号）
```typescript
{
  type: "plan-editor-ready"
}
```

**子窗口 → 父窗口**（保存成功）
```typescript
{
  type: "plan-editor-save",
  payload: {
    plan: AnalysisPlan,
    target: PlanEditorTarget,
    updatedTools?: SavedTool[],   // 工具列表（编辑工具时）
    closeAfterSave?: boolean      // 「保存」按钮触发时父窗口关闭对话框；Ctrl+S 不传
  }
}
```

**子窗口 → 父窗口**（取消关闭）
```typescript
{
  type: "plan-editor-close"
}
```

---

## 实施步骤

### P1：创建独立 HTML 页面（1.5h）

**文件：** `apps/excel-addin/plan-editor.html`

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>编辑执行脚本 - Excel Bro</title>
  <link rel="stylesheet" href="/src/styles.css">
</head>
<body>
  <div id="plan-editor-root"></div>
  <script type="module" src="/src/plan-editor-entry.tsx"></script>
</body>
</html>
```

**文件：** `apps/excel-addin/src/plan-editor-entry.tsx`

```tsx
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import PlanEditorDialog from "./components/PlanEditorDialog";

const root = document.getElementById("plan-editor-root");
if (root) {
  createRoot(root).render(
    <StrictMode>
      <PlanEditorDialog />
    </StrictMode>
  );
}
```

**任务清单：**
- [ ] 创建 `plan-editor.html`
- [ ] 创建 `plan-editor-entry.tsx`
- [ ] 更新 `vite.config.ts`，添加新入口
- [ ] 验证构建产物包含 `plan-editor.html`

---

### P2：实现 PlanEditorDialog 组件（2.5h）

**文件：** `apps/excel-addin/src/components/PlanEditorDialog.tsx`

**功能要求：**

1. **Office.js 初始化**
   - 调用 `Office.onReady()`
   - 监听 `DialogParentMessageReceived` 事件
   - 发送 `plan-editor-ready` 信号

2. **状态管理**
   ```typescript
   type DialogPhase = "loading" | "editing" | "validating" | "error";
   
   interface DialogState {
     phase: DialogPhase;
     draft: string;           // JSON 文本
     errorMessage: string;
     payload: InitPayload | null;
   }
   ```

3. **UI 布局**
   - 顶部：标题 + 关闭按钮
   - 中部：全宽 textarea（语法高亮可选）
   - 底部：错误提示 + 「放弃修改」+「保存」按钮
   - 加载态：显示「正在准备编辑器…」

4. **编辑逻辑**
   - 接收初始数据后，格式化为 `JSON.stringify(plan, null, 2)`
   - 实时更新 `draft` 状态
   - 清除错误提示

5. **保存逻辑**（复用现有 `usePlanEditor.save` 逻辑）
   - 本地 JSON 解析 + `assertEditablePlan` 校验
   - 工具模式：调用 `updateToolPlanTemplate`
   - 指纹重算（`toolSchemaFingerprintForSnapshot`）
   - 发送 `plan-editor-save` 消息到父窗口
   - 点「保存」按钮保存后关闭窗口；Ctrl+S 提示「已保存」并留在窗口

6. **取消逻辑**
   - 发送 `plan-editor-close` 消息
   - 关闭窗口

**任务清单：**
- [ ] 创建 `PlanEditorDialog.tsx` 组件
- [ ] 实现 Office.js 消息监听
- [ ] 实现状态管理（useState）
- [ ] 实现 UI 布局（参考 `PlanJsonEditor.tsx`）
- [ ] 实现保存逻辑（复用 `usePlanEditor.ts` 代码）
- [ ] 实现取消逻辑
- [ ] 错误处理（网络失败、解析失败、校验失败）

---

### P3：父窗口集成 Dialog API（1.5h）

**文件：** `apps/excel-addin/src/App.tsx`

**改动点：**

1. **添加 Dialog 引用**
   ```typescript
   const planEditorDialogRef = useRef<Office.Dialog | null>(null);
   const [planEditorOpening, setPlanEditorOpening] = useState(false);
   ```

2. **创建 `openPlanEditorDialog` 函数**（参考 `openFocusWindow`）
   ```typescript
   function openPlanEditorDialog(plan: AnalysisPlan, target: PlanEditorTarget) {
     if (!isRunningInExcel()) {
       appendMessage({
         role: "system",
         text: "计划编辑器需要从 Excel 任务窗格打开。"
       });
       return;
     }
     
     const payload: PlanEditorInitPayload = {
       type: "plan-editor-init",
       payload: { plan, target, workbook }
     };
     
     // 存储到 localStorage（兜底）
     try {
       localStorage.setItem("EB_PLAN_EDITOR_PAYLOAD", JSON.stringify(payload));
     } catch {}
     
     // DialogApi 1.2 检测
     const canMessageChild = Office.context.requirements.isSetSupported("DialogApi", "1.2");
     
     // 如果已打开，直接更新数据
     if (planEditorDialogRef.current) {
       if (canMessageChild) {
         planEditorDialogRef.current.messageChild(JSON.stringify(payload));
       }
       return;
     }
     
     setPlanEditorOpening(true);
     const url = new URL("plan-editor.html", window.location.href).toString();
     
     Office.context.ui.displayDialogAsync(
       url,
       { height: 90, width: 85, displayInIframe: false },
       (result) => {
         if (result.status !== Office.AsyncResultStatus.Succeeded) {
           setPlanEditorOpening(false);
           appendMessage({
             role: "system",
             text: `无法打开编辑器：${result.error.message}`
           });
           return;
         }
         
         const dialog = result.value;
         planEditorDialogRef.current = dialog;
         
         // 监听子窗口消息
         dialog.addEventHandler(Office.EventType.DialogMessageReceived, (event) => {
           if (!("message" in event)) return;
           try {
             const message = JSON.parse(event.message);
             
             if (message.type === "plan-editor-ready" && canMessageChild) {
               dialog.messageChild(JSON.stringify(payload));
               setPlanEditorOpening(false);
             }
             
             if (message.type === "plan-editor-save") {
               handlePlanEditorSave(message.payload);
               dialog.close();
               planEditorDialogRef.current = null;
             }
             
             if (message.type === "plan-editor-close") {
               dialog.close();
               planEditorDialogRef.current = null;
               setPlanEditorOpening(false);
             }
           } catch {}
         });
         
         // 监听关闭事件
         dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
           planEditorDialogRef.current = null;
           setPlanEditorOpening(false);
         });
       }
     );
   }
   ```

3. **实现 `handlePlanEditorSave`**（复用 `usePlanEditor.onSave` 回调）
   ```typescript
   function handlePlanEditorSave(payload: {
     plan: AnalysisPlan;
     target: PlanEditorTarget;
     updatedTools?: SavedTool[];
   }) {
     const { plan, target, updatedTools } = payload;
     
     if (target.kind === "message") {
       setMessages((prev) =>
         prev.map((msg) =>
           msg.id === target.messageId ? { ...msg, plan } : msg
         )
       );
       appendMessage({
         role: "system",
         text: "已更新计划脚本。"
       });
     }
     
     if (target.kind === "savedTool" && updatedTools) {
       setTools(updatedTools.filter((t) => t.kind === "workflow"));
       setQueryTools(updatedTools.filter((t) => t.kind === "query"));
       appendMessage({
         role: "system",
         text: "已更新工具脚本。"
       });
     }
   }
   ```

4. **替换现有调用**
   - 找到 `planEditor.open()` 的两处调用（消息编辑 + 工具编辑）
   - 改为 `openPlanEditorDialog(plan, target)`

5. **移除旧组件**
   - 删除 `<PlanJsonEditor>` 组件的渲染
   - 保留 `usePlanEditor.ts`（可能用于非 Excel 环境）

**任务清单：**
- [ ] 添加 Dialog 引用和状态
- [ ] 实现 `openPlanEditorDialog` 函数
- [ ] 实现 `handlePlanEditorSave` 函数
- [ ] 替换消息编辑按钮的调用
- [ ] 替换工具编辑按钮的调用
- [ ] 移除 `<PlanJsonEditor>` 渲染

---

### P4：样式调整（0.5h）

**文件：** `apps/excel-addin/src/styles.css`

**样式说明（2026-08-25 UI 重构后）：**

- `.json-editor` 系列：共享 JSON 代码编辑器（语法高亮叠加层、行号、实时校验、
  复制/格式化、深/浅代码皮肤），`styles.css` 根变量提供浅色皮肤默认值，
  `.json-editor.skin-dark` 覆盖为固定暗色。
- `.plan-editor-dialog-*`：独立窗口布局。头部含关闭按钮、标题、面包屑、
  「已修改」徽标和操作按钮；中部两栏（AI 对话 + JSON 编辑器），聊天栏可折叠
  成窄侧栏、可通过分隔条拖动宽度。
- `.plan-editor-modal`：放弃未保存修改时的确认弹窗。
- 错误提示分两层：JSON 编辑过程中的语法错误由 `.json-editor-error` 就地显示；
  结构校验/AI 调用失败由 `.plan-editor-dialog-error` 显示在窗口顶部。

**任务清单：**
- [ ] 添加独立窗口样式
- [ ] 验证在不同主题下的显示效果
- [ ] 调整按钮、输入框样式与现有设计一致

---

### P5：构建配置更新（0.5h）

**文件：** `apps/excel-addin/vite.config.ts`

**改动：** 添加新的入口点

```typescript
export default defineConfig({
  // ...existing config
  build: {
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "taskpane.html"),
        focus: resolve(__dirname, "focus.html"),
        planEditor: resolve(__dirname, "plan-editor.html") // 新增
      }
    }
  }
});
```

**任务清单：**
- [ ] 更新 `vite.config.ts`
- [ ] 执行 `npm run build:addin`
- [ ] 验证 `dist` 目录包含 `plan-editor.html`
- [ ] 验证 HMR 在开发模式正常工作

---

### P6：测试与验收（1.5h）

**单元测试（可选）**
- [ ] `PlanEditorDialog` 组件渲染测试
- [ ] 消息通信 mock 测试
- [ ] 保存逻辑测试（复用 `usePlanEditor.test.ts`）

**集成测试（手动）**

1. **从消息编辑计划**
   - [ ] 执行一个计划，获得 AI 响应
   - [ ] 点击「编辑脚本」按钮
   - [ ] 验证打开独立窗口（90% 高度、85% 宽度）
   - [ ] 验证显示正确的 JSON 内容
   - [ ] 修改 JSON（如改标题）
   - [ ] 点击「保存」
   - [ ] 验证窗口关闭，任务窗格显示「已更新计划脚本」
   - [ ] 验证消息中的 `plan` 已更新

2. **从工具编辑计划**
   - [ ] 打开「我的工具」
   - [ ] 选择一个工具，查看详情
   - [ ] 点击「编辑脚本」
   - [ ] 验证打开独立窗口
   - [ ] 修改 JSON
   - [ ] 点击「保存」
   - [ ] 验证工具列表已更新

3. **错误处理**
   - [ ] 输入非法 JSON，验证错误提示
   - [ ] 输入不符合 schema 的 JSON，验证错误提示
   - [ ] 白名单结构校验失败，验证错误提示
   - [ ] 点击「放弃修改」，验证窗口关闭且无副作用

4. **边界情况**
   - [ ] 快速连续点击「编辑脚本」，验证只打开一个窗口
   - [ ] 已打开窗口时再次点击，验证数据刷新
   - [ ] 手动关闭窗口（点击 X），验证状态重置

5. **兼容性**
   - [ ] Excel for Windows（Desktop）
   - [ ] Excel for Mac（Desktop）
   - [ ] Excel Online（浏览器）

**性能验收**
- [ ] 窗口打开延迟 < 500ms
- [ ] JSON 解析 + 校验 < 200ms

---

## 回退方案

如果 Dialog API 遇到兼容性问题或实现困难，可以回退到当前方案（任务窗格内全屏编辑器）。

**保留文件：**
- `usePlanEditor.ts`
- `PlanJsonEditor.tsx`

**回退步骤：**
1. 恢复 `<PlanJsonEditor>` 组件渲染
2. 恢复 `planEditor.open()` 调用
3. 删除 Dialog 相关代码

---

## 时间估算

| 阶段 | 预计时间 | 关键风险 |
|---|---|---|
| P1：HTML 页面 | 1.5h | Vite 多入口配置 |
| P2：Dialog 组件 | 2.5h | Office.js 消息通信 |
| P3：父窗口集成 | 1.5h | 状态同步复杂度 |
| P4：样式调整 | 0.5h | 主题兼容性 |
| P5：构建配置 | 0.5h | HMR 调试 |
| P6：测试验收 | 1.5h | 跨平台兼容性 |
| **总计** | **8h** | |

---

## 参考资料

- 现有实现：`FocusWorkspace.tsx`（专注窗口）
- Office.js Dialog API：https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins
- 现有编辑器：`PlanJsonEditor.tsx` + `usePlanEditor.ts`
- 校验端点：`server/app/main.py` 的 `/plan/validate`

---

## 后续优化（可选）

1. **语法高亮**：集成轻量级 JSON 编辑器（如 `react-simple-code-editor`）
2. **快捷键**：Ctrl+S 保存，Esc 关闭
3. **差异对比**：显示修改前后的 diff
4. **模板库**：提供常见动作的代码片段
5. **实时校验**：输入时显示 JSON 语法错误（本地）

已实现（2026-08-25 UI 重构）：轻量 JSON 语法高亮 + 行号（textarea 叠加层，未引入
额外依赖）、Ctrl+S 快捷保存（提示「已保存」，窗口保持打开）、保存按钮保存后关闭
窗口、编辑时本地实时校验、复制/格式化、未保存离开确认、AI 思考中动画、聊天栏
折叠与宽度拖动。

---

**状态：** 待实施  
**优先级：** P2（用户体验优化）  
**预计完成时间：** 1 个工作日  
**最后更新：** 2026-08-22
