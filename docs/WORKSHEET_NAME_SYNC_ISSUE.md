# 工作表名称实时同步问题

> **问题**：用户在 Excel 中重命名工作表后，插件内显示的工作表名称没有实时更新。

---

## 📋 问题描述

### 复现步骤

1. 打开 Excel Bro 插件
2. 插件读取并显示当前工作簿的工作表列表（如"Sheet1"）
3. 在 Excel 中将工作表重命名（如 "Sheet1" → "数据表"）
4. 插件中的工作表名称仍然显示为"Sheet1"，没有自动更新

### 预期行为

- 工作表名称改变后，插件应立即同步更新显示的名称

---

## 🔍 当前实现分析

### 监听机制（`excel.ts:130-194`）

当前代码已经实现了工作表变化监听：

```typescript
export async function watchWorkbookStructureChanges(
  onInvalidated?: () => void
): Promise<() => void> {
  const idToName = new Map<string, string>();
  
  const handler = (args: any) => {
    const sheetName = args?.worksheetId && idToName.has(args.worksheetId)
      ? idToName.get(args.worksheetId) ?? null
      : null;
    bumpSheetDataEpoch(sheetName);
    onInvalidated?.();
  };
  
  const worksheets = context.workbook.worksheets;
  
  if (supportsCollectionChanges) {
    worksheets.onChanged.add(handler);  // ✅ 监听工作表内容变化
  }
  
  if (supportsCollectionLifecycle) {
    worksheets.onAdded.add(handler);    // ✅ 监听工作表添加
    worksheets.onDeleted.add(handler);  // ✅ 监听工作表删除
  }
  
  // 初始化 id -> name 映射
  worksheets.load("items/name,items/id");
  await context.sync();
  for (const sheet of worksheets.items) {
    idToName.set(sheet.id, sheet.name);
  }
}
```

### 问题所在

1. **监听了变化，但没有更新映射**
   - `worksheets.onChanged` 会触发 `handler`
   - `handler` 调用 `onInvalidated()`，触发缓存失效
   - **但 `idToName` 映射没有更新**
   - 下次查询时仍使用旧的名称映射

2. **缺少 `onNameChanged` 事件监听**
   - Office.js 提供了 `worksheets.onNameChanged` 事件（ExcelApi 1.17+）
   - 当前代码没有监听该事件
   - 工作表重命名时，应该更新 `idToName` 映射

3. **`onInvalidated` 只触发缓存失效**
   - 在 `App.tsx:1250` 调用：
     ```typescript
     dispose = await watchWorkbookStructureChanges(() => {
       invalidateStructureCache();
     });
     ```
   - 只清空了结构缓存，但没有触发重新扫描
   - 用户界面不会自动刷新工作表列表

---

## 🎯 解决方案

### 方案 1：监听 `onNameChanged` 事件并更新映射（推荐）

**改动位置**：`excel.ts:130-194`

**实现思路**：
1. 检测是否支持 `onNameChanged` 事件（ExcelApi 1.17+）
2. 添加 `worksheets.onNameChanged` 监听
3. 在 `onNameChanged` 回调中：
   - 更新 `idToName` 映射（从旧名称更新到新名称）
   - 调用 `onInvalidated()` 触发界面刷新

**代码示例**：
```typescript
// 检测 API 支持
const supportsNameChanged =
  Office.context.requirements.isSetSupported("ExcelApi", "1.17");

// 添加监听
if (supportsNameChanged) {
  worksheets.onNameChanged.add(async (event: Excel.WorksheetNameChangedEventArgs) => {
    // 更新 id -> name 映射
    const worksheetId = event.worksheetId;
    const newName = event.nameAfter;
    idToName.set(worksheetId, newName);
    
    // 触发界面刷新
    onInvalidated?.();
  });
}
```

**优点**：
- 精确监听重命名事件
- 自动更新映射，无需重新扫描整个工作簿
- 性能最优

**缺点**：
- 需要 ExcelApi 1.17+（Excel 2019/2021/365）
- 旧版本不支持（需要降级方案）

---

### 方案 2：所有变化事件都重新扫描工作表列表

**改动位置**：`excel.ts:130-194` + `App.tsx:1250`

**实现思路**：
1. 在 `watchWorkbookStructureChanges` 的 `handler` 中：
   - 调用 `onInvalidated()` 后
   - 重新加载 `worksheets.items` 并更新 `idToName` 映射
2. 在 `App.tsx` 的回调中：
   - 除了清空缓存，还触发 `scan()` 重新读取工作簿

**代码示例**：
```typescript
const handler = async (args: any) => {
  // 重新加载工作表列表
  const worksheets = context.workbook.worksheets;
  worksheets.load("items/name,items/id");
  await context.sync();
  
  // 更新映射
  idToName.clear();
  for (const sheet of worksheets.items) {
    idToName.set(sheet.id, sheet.name);
  }
  
  // 触发界面刷新
  onInvalidated?.();
};
```

**优点**：
- 兼容所有版本
- 确保名称始终最新

**缺点**：
- 性能较差（每次变化都重新扫描）
- 在 `onChanged` 高频触发时可能影响性能

---

### 方案 3：在 UI 层面手动刷新

**改动位置**：`App.tsx`、`useWorkbookContext.ts`

**实现思路**：
1. 在 `onInvalidated` 回调中：
   - 清空缓存
   - 调用 `scan()` 重新读取工作簿结构
2. UI 自动使用最新的工作簿快照

**代码示例**：
```typescript
// App.tsx:1250
dispose = await watchWorkbookStructureChanges(async () => {
  invalidateStructureCache();
  // 重新扫描工作簿
  await scan({ announce: false });
});
```

**优点**：
- 实现简单
- 兼容所有版本
- 确保 UI 显示的是最新数据

**缺点**：
- 每次变化都全量扫描（包括读取字段结构）
- 性能开销较大

---

## ✅ 推荐实现方案

**组合方案**：方案 1（优先） + 方案 2（降级）

```typescript
export async function watchWorkbookStructureChanges(
  onInvalidated?: () => void
): Promise<() => void> {
  const idToName = new Map<string, string>();
  
  const handler = (args: any) => {
    const sheetName = args?.worksheetId && idToName.has(args.worksheetId)
      ? idToName.get(args.worksheetId) ?? null
      : null;
    bumpSheetDataEpoch(sheetName);
    onInvalidated?.();
  };
  
  // 🆕 重命名事件处理器
  const nameChangedHandler = async (event: Excel.WorksheetNameChangedEventArgs) => {
    // 更新映射：旧名称 -> 新名称
    idToName.set(event.worksheetId, event.nameAfter);
    bumpSheetDataEpoch(event.nameAfter);
    onInvalidated?.();
  };
  
  const supportsCollectionChanges = 
    Office.context.requirements.isSetSupported("ExcelApi", "1.9");
  const supportsCollectionLifecycle = 
    Office.context.requirements.isSetSupported("ExcelApi", "1.7");
  const supportsNameChanged = 
    Office.context.requirements.isSetSupported("ExcelApi", "1.17");
  
  const watched = await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    
    if (supportsCollectionChanges) {
      worksheets.onChanged.add(handler);
    }
    
    if (supportsCollectionLifecycle) {
      worksheets.onAdded.add(handler);
      worksheets.onDeleted.add(handler);
    }
    
    // 🆕 监听工作表重命名事件
    if (supportsNameChanged) {
      worksheets.onNameChanged.add(nameChangedHandler);
    }
    
    worksheets.load("items/name,items/id");
    await context.sync();
    
    idToName.clear();
    for (const sheet of worksheets.items) {
      idToName.set(sheet.id, sheet.name);
    }
    
    // ... 其他初始化代码
    
    return worksheets.items.map((sheet) => sheet.name);
  });
  
  return () => {
    void Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      
      if (supportsCollectionChanges) {
        worksheets.onChanged.remove(handler);
      }
      
      if (supportsCollectionLifecycle) {
        worksheets.onAdded.remove(handler);
        worksheets.onDeleted.remove(handler);
      }
      
      // 🆕 移除重命名监听
      if (supportsNameChanged) {
        worksheets.onNameChanged.remove(nameChangedHandler);
      }
      
      await context.sync();
    });
  };
}
```

---

## 📂 涉及文件

| 文件 | 改动内容 | 行数位置 |
|------|---------|---------|
| `apps/excel-addin/src/excel.ts` | 1. 添加 `supportsNameChanged` 检测<br>2. 添加 `nameChangedHandler` 处理器<br>3. 监听 `worksheets.onNameChanged`<br>4. 在清理函数中移除监听 | 约 130-194 行 |
| `apps/excel-addin/src/App.tsx`（可选） | 在 `onInvalidated` 回调中触发 `scan()`（如果需要强制刷新 UI） | 约 1250 行 |

---

## 🧪 验收标准

1. **工作表重命名同步**
   - ✅ 在 Excel 中重命名工作表
   - ✅ 插件中的工作表名称立即更新
   - ✅ 工作表选择状态保持（通过 ID 关联，而非名称）

2. **向后兼容**
   - ✅ 在支持 `onNameChanged` 的版本（Excel 2019+）使用新事件
   - ✅ 在不支持的旧版本中降级为定期轮询或手动刷新

3. **性能验证**
   - ✅ 重命名操作不触发全量工作簿扫描
   - ✅ 只更新名称映射，不重新读取字段结构

---

## 🚀 实现步骤

1. **修改 `excel.ts`**
   - 添加 `onNameChanged` 事件监听
   - 实现 `nameChangedHandler` 更新映射
   - 添加清理逻辑

2. **测试验证**
   - 运行 `npm run test:addin` 确保测试通过
   - 手动测试：重命名工作表，检查插件是否同步更新

3. **（可选）增强 UI 刷新**
   - 如果名称更新后 UI 没有自动刷新
   - 在 `App.tsx` 的 `onInvalidated` 中调用 `scan()`

4. **提交代码**
   - 提交信息：`fix: 监听工作表重命名事件以实时同步名称`

---

## 📝 技术细节

### Office.js API 版本要求

| 事件 | API 版本 | Excel 版本 |
|------|---------|-----------|
| `onChanged` | 1.9 | Excel 2019+ |
| `onAdded` / `onDeleted` | 1.7 | Excel 2016+ |
| `onNameChanged` | 1.17 | Excel 2019/2021/365 |

### `onNameChanged` 事件参数

```typescript
interface WorksheetNameChangedEventArgs {
  worksheetId: string;  // 工作表 ID（不变）
  nameAfter: string;    // 新名称
  // 注意：没有 oldName，需要从 idToName 映射中查找
}
```

---

**文档版本**：v1.0  
**创建日期**：2026-08-18  
**优先级**：🔥 高（影响用户体验）  
**状态**：已实现 ✅
