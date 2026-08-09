# EB 函数系统设计

## 1. 概述

EB 函数系统允许用户创建可在 Excel 单元格中使用的自定义函数。核心特性：

- **离线可用**：规则存储在 Excel 文件内，本地执行，无需联网
- **动态更新**：修改规则定义，所有引用自动生效
- **依赖管理**：自动检测循环依赖，按拓扑排序刷新
- **规则保护**：多重备份，防止误删

## 2. 用户体验

### 创建规则
```
用户：帮我创建一个"客户评分"规则
      参数：购买金额、购买次数、最近活跃天数
      计算：金额每1000元1分 + 购买次数每次2分 + 30天内活跃加10分

AI：已创建规则，可以在单元格中使用：
    =EB("客户评分", B2, C2, D2)
```

### 使用规则
```excel
=EB("隐形字符清理", A1)
=EB("空白检测", A1)
=EB("提取数字", A1)
=EB("客户评分", B2, C2, D2)
```

### 修改规则
```
用户：修改"客户评分"，购买次数改为每次3分

AI：已更新规则，点击"刷新所有 EB 函数"应用到表格
```

## 3. 规则存储格式

### 3.1 存储位置

**主存储**：隐藏工作表 `#EB_RULES`（井号开头，Excel 界面不显示）
**备份**：隐藏工作表 `#EB_RULES_BACKUP`（每次修改时自动备份）
**本地备份**：`%LOCALAPPDATA%/Excel Bro/rules-backup/`（Windows）或 `~/Library/Application Support/Excel Bro/rules-backup/`（macOS）

### 3.2 规则定义（JSON）

```typescript
interface EBRule {
  // 基础信息
  id: string;                    // UUID，唯一标识
  name: string;                  // 规则名称（用户可见）
  description: string;           // 规则描述
  category: 'builtin' | 'user';  // 预制或用户自定义
  
  // 参数定义
  params: EBParam[];             // 参数列表
  
  // 执行逻辑
  logic: string;                 // JavaScript 表达式
  compiled: string;              // 编译后的函数体
  
  // 依赖关系
  dependencies: string[];        // 依赖的其他规则 ID
  
  // 元数据
  createdAt: string;             // ISO 8601
  updatedAt: string;             // ISO 8601
  version: number;               // 版本号
  readonly: boolean;             // 是否只读（预制规则）
}

interface EBParam {
  name: string;                  // 参数名
  type: 'text' | 'number' | 'any'; // 参数类型
  required: boolean;             // 是否必填
  description?: string;          // 参数说明
}
```

### 3.3 存储示例

隐藏工作表 `#EB_RULES` 的格式：

| A (id) | B (name) | C (definition JSON) |
|--------|----------|---------------------|
| uuid-1 | 隐形字符清理 | `{...}` |
| uuid-2 | 空白检测 | `{...}` |
| uuid-3 | 提取数字 | `{...}` |
| uuid-4 | 客户评分 | `{...}` |

### 3.4 预制规则定义

```json
{
  "id": "builtin-clean-invisible",
  "name": "隐形字符清理",
  "description": "清除 CHAR(160)、零宽空格、控制字符",
  "category": "builtin",
  "params": [
    {
      "name": "text",
      "type": "text",
      "required": true,
      "description": "要清理的文本"
    }
  ],
  "logic": "cleanInvisibleChars(text)",
  "compiled": "function(text) { return String(text).replace(/[\\u0000-\\u001F\\u007F-\\u009F]/g, '').replace(/[\\u00A0\\u1680\\u2000-\\u200B\\u202F\\u205F\\u3000]/g, ' ').replace(/\\uFEFF/g, '').replace(/[\\u200E\\u200F\\u202A-\\u202E]/g, '').trim().replace(/\\s+/g, ' '); }",
  "dependencies": [],
  "createdAt": "2026-08-08T00:00:00Z",
  "updatedAt": "2026-08-08T00:00:00Z",
  "version": 1,
  "readonly": true
}
```

## 4. Custom Functions 实现

### 4.1 函数注册

```typescript
// apps/excel-addin/src/functions/functions.ts
CustomFunctions.associate("EB", {
  invoke: async (ruleName: string, ...args: any[]) => {
    const rule = await loadRule(ruleName);
    if (!rule) {
      throw new Error(`规则 "${ruleName}" 不存在`);
    }
    return executeRule(rule, args);
  }
});
```

### 4.2 manifest.xml 配置

需要在 manifest.xml 中声明 Custom Functions：

```xml
<ExtensionPoint xsi:type="CustomFunctions">
  <Script>
    <SourceLocation resid="Functions.Script.Url"/>
  </Script>
  <Page>
    <SourceLocation resid="Functions.Page.Url"/>
  </Page>
  <Metadata>
    <SourceLocation resid="Functions.Metadata.Url"/>
  </Metadata>
</ExtensionPoint>
```

### 4.3 函数元数据（functions.json）

```json
{
  "functions": [
    {
      "id": "EB",
      "name": "EB",
      "description": "执行 Excel Bro 自定义规则",
      "parameters": [
        {
          "name": "ruleName",
          "description": "规则名称",
          "type": "string"
        },
        {
          "name": "args",
          "description": "规则参数",
          "type": "any",
          "repeating": true,
          "optional": true
        }
      ],
      "result": {
        "type": "any"
      }
    }
  ]
}
```

## 5. 依赖管理

### 5.1 依赖提取

```typescript
function extractDependencies(logic: string): string[] {
  const ebCallPattern = /EB\s*\(\s*["']([^"']+)["']/g;
  const dependencies: string[] = [];
  let match;
  
  while ((match = ebCallPattern.exec(logic)) !== null) {
    dependencies.push(match[1]);
  }
  
  return [...new Set(dependencies)];
}
```

### 5.2 循环检测

```typescript
function detectCycle(rules: Map<string, EBRule>): string[] | null {
  const visited = new Set<string>();
  const recStack = new Set<string>();
  
  function dfs(ruleId: string, path: string[]): string[] | null {
    if (recStack.has(ruleId)) {
      return [...path, ruleId]; // 返回循环路径
    }
    if (visited.has(ruleId)) {
      return null;
    }
    
    visited.add(ruleId);
    recStack.add(ruleId);
    path.push(ruleId);
    
    const rule = rules.get(ruleId);
    if (rule) {
      for (const dep of rule.dependencies) {
        const cycle = dfs(dep, [...path]);
        if (cycle) return cycle;
      }
    }
    
    recStack.delete(ruleId);
    return null;
  }
  
  for (const ruleId of rules.keys()) {
    const cycle = dfs(ruleId, []);
    if (cycle) return cycle;
  }
  
  return null;
}
```

### 5.3 拓扑排序

```typescript
function topologicalSort(rules: Map<string, EBRule>): string[] {
  const sorted: string[] = [];
  const visited = new Set<string>();
  
  function visit(ruleId: string) {
    if (visited.has(ruleId)) return;
    visited.add(ruleId);
    
    const rule = rules.get(ruleId);
    if (rule) {
      for (const dep of rule.dependencies) {
        visit(dep);
      }
    }
    
    sorted.push(ruleId);
  }
  
  for (const ruleId of rules.keys()) {
    visit(ruleId);
  }
  
  return sorted;
}
```

## 6. 规则保护机制

### 6.1 自动备份

```typescript
async function saveRule(rule: EBRule): Promise<void> {
  await Excel.run(async (context) => {
    // 1. 保存到主存储
    const mainSheet = getOrCreateHiddenSheet(context, '#EB_RULES');
    writeRuleToSheet(mainSheet, rule);
    
    // 2. 备份到 #EB_RULES_BACKUP
    const backupSheet = getOrCreateHiddenSheet(context, '#EB_RULES_BACKUP');
    writeRuleToSheet(backupSheet, rule);
    
    await context.sync();
  });
  
  // 3. 备份到本地目录
  await backupToLocalFile(rule);
}
```

### 6.2 规则恢复

```typescript
async function checkAndRestore(): Promise<void> {
  await Excel.run(async (context) => {
    const workbook = context.workbook;
    workbook.worksheets.load('items/name');
    await context.sync();
    
    const hasMain = workbook.worksheets.items
      .some(ws => ws.name === '#EB_RULES');
    const hasBackup = workbook.worksheets.items
      .some(ws => ws.name === '#EB_RULES_BACKUP');
    
    if (!hasMain && hasBackup) {
      // 从备份恢复
      await restoreFromBackup(context);
      showNotification('规则已从备份恢复');
    } else if (!hasMain && !hasBackup) {
      // 尝试从本地文件恢复
      const restored = await restoreFromLocalFile();
      if (restored) {
        showNotification('规则已从本地备份恢复');
      } else {
        showWarning('规则已丢失且无备份，EB 函数将失效');
      }
    }
  });
}
```

## 7. 规则管理界面

### 7.1 界面结构

```
┌─ EB 函数 ────────────────────┐
│ 🔍 搜索规则...                │
│                               │
│ 📦 预制规则（3）              │
│   ✓ 隐形字符清理              │
│     =EB("隐形字符清理", A1)   │
│   ✓ 空白检测                  │
│   ✓ 提取数字                  │
│                               │
│ 🛠️ 我的规则（2）              │
│   ✓ 客户评分     [编辑] [删除]│
│     =EB("客户评分", A1, B1)   │
│   ✓ 会员折扣                  │
│                               │
│ [+ 创建新规则]                │
│ [🔄 刷新所有 EB 函数]         │
│ [📤 导出] [📥 导入]           │
└───────────────────────────────┘
```

### 7.2 创建规则流程

```
1. 用户在对话中描述需求
2. AI 生成规则定义
3. 显示预览（名称、参数、示例公式）
4. 用户确认
5. 保存到隐藏工作表 + 备份
6. 注册到 Custom Functions runtime
7. 显示使用示例
```

## 8. 技术约束

### 8.1 Custom Functions 限制

- 运行在独立 runtime，不能访问 UI
- 不能调用 Office.js 的 Excel API
- 异步函数需要返回 Promise
- 函数必须是纯函数（无副作用）

### 8.2 安全约束

- 不执行任意代码，只执行编译后的白名单函数
- 用户自定义规则需要 AI 生成并验证
- 禁止访问外部网络
- 禁止访问文件系统

## 9. 实现阶段

### 阶段 1：基础设施
- 规则存储格式
- 隐藏工作表读写
- 备份恢复机制

### 阶段 2：预制规则
- 3 个预制规则实现
- Custom Functions 注册
- 本地执行引擎

### 阶段 3：依赖管理
- 依赖提取
- 循环检测
- 拓扑排序

### 阶段 4：规则管理
- 规则列表界面
- 创建/编辑/删除
- 导入/导出

### 阶段 5：AI 集成
- 对话中创建规则
- AI 生成逻辑
- 规则测试器

## 10. 与现有系统集成

- **不冲突**：EB 函数独立于现有的 AnalysisPlan 系统
- **可互补**：用户可以在对话中说"用我的 EB 函数计算..."
- **共享存储**：使用相同的本地持久化策略
- **共享 AI**：使用相同的模型接口生成规则
