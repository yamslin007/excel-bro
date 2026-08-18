# 工作簿文件名（另存为改名）实时同步问题

> **问题**：用户通过"另存为"功能改变工作簿文件名后，插件顶部显示的工作簿名称没有实时更新。

---

## 📋 问题描述

### 复现步骤

1. 打开 Excel Bro 插件，当前文件名为 "Book1.xlsx"
2. 插件顶部显示："Book1.xlsx"
3. 用户执行 **文件 → 另存为**，保存为 "数据分析.xlsx"
4. 插件顶部仍然显示："Book1.xlsx"，没有更新为 "数据分析.xlsx"

### 预期行为

- 工作簿文件名改变后，插件应立即同步更新显示的文件名

---

## 🔍 当前实现分析

### 工作簿名称获取方式（`excel.ts:648`）

```typescript
const snapshot: WorkbookSnapshot = {
  name: workbookNameFromDocumentUrl(Office.context.document.url),
  // ...
};
```

### 名称提取逻辑（`workbookIdentity.ts:1-14`）

```typescript
export function workbookNameFromDocumentUrl(documentUrl?: string | null): string {
  const raw = documentUrl?.trim();
  if (!raw) return "未保存的工作簿";

  const withoutQuery = raw.split(/[?#]/, 1)[0].replace(/\\/g, "/");
  const encodedName = withoutQuery.slice(withoutQuery.lastIndexOf("/") + 1);
  if (!encodedName) return "未保存的工作簿";

  try {
    return decodeURIComponent(encodedName);
  } catch {
    return encodedName;
  }
}
```

### 问题所在

1. **工作簿名称在扫描时静态获取**
   - `captureWorkbookStructure` 调用时读取 `Office.context.document.url`
   - 获取到的是**当时**的文件名
   - 存储到 `workbook.name` 中

2. **"另存为"后 URL 变化，但没有触发重新扫描**
   - Office.js **没有提供** `onDocumentUrlChanged` 或类似事件
   - `watchWorkbookStructureChanges` 监听的是工作表变化，不包括文档 URL

3. **结构缓存会保留旧名称**
   - `captureWorkbookStructure` 有缓存机制
   - 缓存命中时直接返回旧快照（包括旧文件名）
   - 只有在结构变化（工作表增删改）时才失效缓存

---

## 🎯 解决方案

### 方案 1：监听 `document.settings` 变化（不可行）

Office.js 没有提供监听文档 URL 变化的事件。`document.settings.addHandlerAsync` 只能监听自定义设置的变化，无法监听文件名。

❌ **此方案不可行**

---

### 方案 2：定期轮询 `document.url`（不推荐）

**实现思路**：
- 每隔 N 秒检查 `Office.context.document.url` 是否变化
- 如果变化，触发重新扫描

**缺点**：
- 性能开销（持续轮询）
- 延迟响应（取决于轮询间隔）
- 不优雅

❌ **不推荐**

---

### 方案 3：在 `captureWorkbookStructure` 时总是读取最新 URL（推荐）

**核心思路**：
- 即使命中缓存，也更新 `workbook.name` 为最新值
- 确保每次扫描时文件名都是最新的

**改动位置**：`excel.ts:664-687`

**实现代码**：

```typescript
export async function captureWorkbookStructure(
  dataSheetNames?: string[]
): Promise<WorkbookSnapshot> {
  const cacheKey = JSON.stringify([...(dataSheetNames ?? [])].sort());
  if (structureCacheEnabled && structureCache?.key === cacheKey) {
    const cached = structuredClone(structureCache.snapshot);
    
    // 🆕 更新工作簿名称（即使命中缓存也读取最新）
    cached.name = workbookNameFromDocumentUrl(Office.context.document.url);
    
    // 结构缓存只在工作表数据/集合变化时失效，切换活动工作表不会失效。
    // 命中缓存时用一次轻量选区读取刷新活动表与选区，避免拿过期值构建
    // 指纹，导致确认期与运行期不一致的误判。
    const selection = await captureSelectionContext();
    cached.activeWorksheet = selection.activeWorksheet;
    cached.selectedRange = selection.selectedRange;
    
    // 活动表不在缓存的工作表集合里，说明集合已变（如新建/删除表）但缓存
    // 尚未失效；此时放弃缓存走全量重扫，避免 activeWorksheet 悬空导致
    // 前端构造出空 sheets 被后端拒绝。
    if (
      !cached.worksheets.some(
        (sheet) => sheet.name === selection.activeWorksheet
      )
    ) {
      structureCache = null;
      return captureWorkbookStructure(dataSheetNames);
    }
    return cached;
  }
  return Excel.run(async (context) => {
    // ... 原有逻辑
  });
}
```

**优点**：
- 简单直接，只需添加一行代码
- 性能开销极小（只是读取一个字符串）
- 确保文件名始终最新

**工作原理**：
- 用户"另存为"后，下次调用 `captureWorkbookStructure` 时
- `Office.context.document.url` 已经是新文件路径
- 即使命中缓存，也会用新 URL 提取新文件名
- UI 自动显示新名称

---

### 方案 4：在所有使用 `workbook.name` 的地方实时读取（备选）

**实现思路**：
- 不在快照中存储 `workbook.name`
- 每次显示时动态读取 `Office.context.document.url`

**优点**：
- 100% 实时

**缺点**：
- 需要修改多处代码
- `workbook.name` 在后端日志、工具参数等多处使用
- 改动范围大

❌ **改动过大，不推荐**

---

## ✅ 推荐实现方案

**方案 3：在 `captureWorkbookStructure` 缓存命中时更新文件名**

这是最简单、最有效的方案：
- 只需在 `excel.ts:674` 附近添加一行代码
- 确保每次扫描时文件名都是最新的
- 对现有逻辑影响最小

---

## 📂 涉及文件

| 文件 | 改动内容 | 行数位置 |
|------|---------|---------|
| `apps/excel-addin/src/excel.ts` | 在 `captureWorkbookStructure` 缓存命中分支中，添加一行更新 `cached.name` | 约 674 行 |

---

## 🧪 验收标准

1. **另存为改名同步**
   - ✅ 在 Excel 中执行"文件 → 另存为"，保存为新文件名
   - ✅ 插件顶部显示的文件名立即更新（下次扫描时）
   - ✅ 工作表列表、对话记录等不受影响

2. **首次扫描**
   - ✅ 打开新文件时，文件名正确显示

3. **性能验证**
   - ✅ 添加的一行代码不影响性能
   - ✅ 缓存机制正常工作

---

## 🚀 实现步骤

1. **修改 `excel.ts`**
   ```typescript
   // 在第 674 行附近（cached 返回之前）添加：
   cached.name = workbookNameFromDocumentUrl(Office.context.document.url);
   ```

2. **测试验证**
   - 运行 `npm run test:addin` 确保测试通过
   - 手动测试"另存为"场景

3. **提交代码**
   - 提交信息：`fix: 另存为改名后实时同步工作簿文件名`

---

## 📝 技术细节

### 为什么缓存命中时也要更新名称？

**原因**：
- 结构缓存的失效条件是**工作表结构变化**（增删改工作表、数据变化）
- "另存为"不改变工作表结构，所以**不会触发缓存失效**
- 如果不在缓存命中时更新名称，用户会一直看到旧文件名
- 添加这一行后，即使缓存命中，文件名也是最新的

### `Office.context.document.url` 的更新时机

- "另存为"完成后，`Office.context.document.url` **立即更新**为新文件路径
- 不需要等待事件或轮询
- 下次读取时自动就是新 URL

---

## 🎯 触发更新的时机

虽然我们在 `captureWorkbookStructure` 中更新了文件名，但需要触发扫描才能让 UI 看到新名称。

### 当前会触发扫描的场景

1. **用户发送消息时**（`sendMessage` → `scan`）
2. **用户切换工作表选择后发送**
3. **用户点击"刷新"按钮**（如果有）

### "另存为"后的行为

- 用户"另存为"后，通常会继续进行操作
- 下次发送消息或切换数据范围时，会触发扫描
- 扫描时就会读取到新文件名

### 如果需要立即更新（可选增强）

可以在 `watchWorkbookStructureChanges` 的 `handler` 中添加扫描触发：

```typescript
const handler = async (args?: SheetChangeEventArgs) => {
  invalidateWorkbookStructureCache();
  const sheetName = // ...
  bumpSheetDataEpoch(sheetName);
  
  // 🆕 触发扫描以更新 UI（包括文件名）
  onInvalidated?.();
};
```

但这会在**任何变化**时都触发扫描，可能过于频繁。

---

## 💡 用户体验说明

修复后的行为：
1. 用户"另存为"保存新文件
2. 用户继续操作（如输入消息、切换工作表）
3. 插件重新扫描工作簿
4. 文件名自动更新为新名称

**注意**：不是"另存为"瞬间就更新，而是**下次扫描时**更新。这是合理的，因为：
- Office.js 没有提供文档 URL 变化事件
- 持续轮询性能开销大
- 用户"另存为"后通常会继续操作，自然会触发扫描

---

**文档版本**：v1.0  
**创建日期**：2026-08-18  
**优先级**：🔥 高（影响用户体验）  
**状态**：已实现 ✅
