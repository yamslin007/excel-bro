# SDK 改造快速参考卡

> 停用提示（2026-10-06）：本页旧配置/代码速查不可用于实施，不能以 cwd 或 Bash 权限实现安全 Python。当前指导为 [隔离 Python 设计](./SDK_ISOLATED_PYTHON_DESIGN_2026-10-06.md) 与 [DeepSeek 实施提示词](./DEEPSEEK_ISOLATED_PYTHON_PROMPT_2026-10-06.md)；运行器未实现、未启用。正文只作历史追溯。

## 🎯 核心改动（5 分钟看完）

### 1. 工具配置改动

**位置：** `server/app/claude_sdk_agent.py` 第 74-90 行

```python
# ❌ 改造前
tools=[],
disallowed_tools=["Read", "Bash", "Write", ...],
allowed_tools=["mcp__excel_bro__submit_*"]

# ✅ 改造后
tools=["Read", "Bash"],
allowed_tools=["Read", "Bash", "mcp__excel_bro__*"],
mcp_servers={"excel_bro": _create_excel_mcp_server()},
cwd=workspace_dir,  # 新增：独立工作目录
```

---

### 2. 创建 MCP Server

**位置：** `server/app/claude_sdk_agent.py` 新增函数

```python
def _create_excel_mcp_server():
    """创建 MCP 工具"""
    sdk = _sdk()
    tools = []
    
    # 工具 1: 读取授权表
    @sdk.tool("read_authorized_sheet", "读取已授权的工作表数据", {...})
    async def read_authorized_sheet(sheet_name: str, sample: bool = False):
        if sheet_name not in authorized_sheets:
            return {"error": "未授权"}
        
        data = request.sheetsData[sheet_name]
        if sample and len(data) > 1000:
            data = data[:1000]
        
        return {"content": [{"type": "text", "text": json.dumps({
            "sheet": sheet_name,
            "data": data
        })}]}
    
    tools.append(read_authorized_sheet)
    
    # 工具 2: 提交计划
    @sdk.tool("submit_plan", "提交最终操作计划", {...})
    async def submit_plan(plan: dict):
        nonlocal terminal
        terminal = AssistantResponse(kind="plan", plan=plan)
        return {"content": [{"type": "text", "text": "✅ 计划已提交"}]}
    
    tools.append(submit_plan)
    
    # 工具 3: 提交回答
    @sdk.tool("submit_answer", "提交只读回答", {...})
    async def submit_answer(message: str):
        nonlocal terminal
        terminal = AssistantResponse(kind="answer", message=message)
        return {"content": [{"type": "text", "text": "✅ 回答已提交"}]}
    
    tools.append(submit_answer)
    
    return sdk.create_sdk_mcp_server("excel_bro", "1.0.0", tools=tools)
```

---

### 3. 前端数据读取

**Excel: `apps/excel-addin/src/officeDataReader.ts`**

```typescript
export async function readSheetDataOffice(sheetName: string): Promise<any[]> {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const usedRange = sheet.getUsedRange();
    usedRange.load(['values']);
    await context.sync();
    
    const values = usedRange.values;
    const headers = values[0];
    const rows = values.slice(1);
    
    return rows.map(row => {
      const obj: Record<string, any> = {};
      headers.forEach((h, i) => { obj[String(h)] = row[i]; });
      return obj;
    });
  });
}
```

**WPS: `apps/wps-addin/src/wpsDataReader.ts`**

```typescript
export async function readSheetDataWps(sheetName: string): Promise<any[]> {
  const sheet = Application.Worksheets.Item(sheetName);
  const usedRange = sheet.UsedRange;
  const values = usedRange.Value2;
  
  const headers = values[0];
  const rows = values.slice(1);
  
  return rows.map(row => {
    const obj: Record<string, any> = {};
    headers.forEach((h, i) => { 
      obj[String(h)] = normalizeWpsValue(row[i]); 
    });
    return obj;
  });
}

function normalizeWpsValue(value: any): any {
  if (value === null || value === undefined) return null;
  if (typeof value === 'number') return value;
  if (typeof value === 'string') return value;
  if (typeof value === 'boolean') return value;
  if (value instanceof Date) return value.toISOString().split('T')[0];
  return String(value);
}
```

**发送数据: `apps/*/src/api.ts`**

```typescript
// 读取所有授权表的数据
const sheetsData: Record<string, any[]> = {};
for (const sheet of workbook.worksheets) {
  if (isOfficeJs()) {
    sheetsData[sheet.name] = await readSheetDataOffice(sheet.name);
  } else {
    sheetsData[sheet.name] = await readSheetDataWps(sheet.name);
  }
}

// 发送请求
const response = await fetch('/api/turn', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    prompt,
    workbook,
    sheetsData,  // 🆕 新增
    modelId,
  })
});
```

---

### 4. 后端接收数据

**位置：** `server/app/models.py`

```python
class PlanRequest(BaseModel):
    prompt: str
    workbook: WorkbookSnapshot
    modelId: str | None = None
    
    # 🆕 新增
    sheetsData: dict[str, list[dict]] | None = None
```

---

## 📋 实施检查清单

### 第 1 天：基础改造

- [ ] 修改 `_options()` 释放工具
- [ ] 实现 `_create_excel_mcp_server()`
- [ ] 实现 3 个 MCP 工具（read/submit_plan/submit_answer）
- [ ] 更新 system prompt
- [ ] 创建工作目录逻辑

### 第 2 天：前端数据读取

- [ ] Excel: 实现 `readSheetDataOffice()`
- [ ] WPS: 实现 `readSheetDataWps()`
- [ ] 统一数据格式（类型转换）
- [ ] 在请求中包含 `sheetsData`
- [ ] 添加数据量限制（max 10000 行）

### 第 3 天：后端集成

- [ ] 更新 `PlanRequest` 模型
- [ ] MCP 工具从 `request.sheetsData` 读取
- [ ] 授权验证
- [ ] 错误处理

### 第 4 天：测试

- [ ] Excel 简单读取测试
- [ ] WPS 简单读取测试
- [ ] Python 分析测试
- [ ] 多表关联测试
- [ ] 错误场景测试

### 第 5 天：优化和文档

- [ ] 性能优化（缓存、限制）
- [ ] 更新 README
- [ ] 更新 ARCHITECTURE.md
- [ ] 添加使用示例

---

## 🔧 常见问题速查

### Q: SDK 启动失败？

```bash
# Windows 需要用特殊脚本启动
powershell -File scripts/start_claude_sdk_server.ps1

# 或添加到 package.json
"scripts": {
  "dev:sdk": "powershell -File scripts/start_claude_sdk_server.ps1"
}
```

### Q: 数据格式不一致？

```typescript
// 确保统一标准化
function standardizeValue(value: any): any {
  if (value === null || value === undefined) return null;
  if (typeof value === 'number') return value;
  if (typeof value === 'string') return value;
  if (typeof value === 'boolean') return value;
  if (value instanceof Date) return value.toISOString().split('T')[0];
  return String(value);
}
```

### Q: 授权验证失败？

```python
# 确保从 workbook.worksheets 获取授权列表
authorized_sheets = [s.name for s in request.workbook.worksheets]

if sheet_name not in authorized_sheets:
    return {"error": f"未授权访问 {sheet_name}，已授权: {authorized_sheets}"}
```

### Q: Python 脚本执行失败？

```python
# 检查用户环境
try:
    import pandas
except ImportError:
    return {"error": "需要安装 pandas: pip install pandas"}

# 添加超时保护
subprocess.run(..., timeout=60)
```

---

## 🎨 System Prompt 关键部分

### 中文版（用于 DeepSeek/Qwen/GLM）

```markdown
## 工作流程

1. **探索阶段**（多轮）
   - 用 `read_authorized_sheet(sheet_name, sample=true)` 查看结构
   - 用 `read_authorized_sheet(sheet_name, sample=false)` 获取完整数据

2. **分析阶段**（可选）
   - 用 `Bash` 运行 Python 脚本
   - 可用库：pandas, numpy, scipy

3. **提交阶段**（必须且仅一次）
   - 修改 Excel: `submit_plan(plan)`
   - 只读查询: `submit_answer(message)`
```

### 英文版（用于 Claude/GPT）

```markdown
## Workflow

1. **Exploration Phase** (multiple rounds)
   - Use `read_authorized_sheet(sheet_name, sample=true)` to check structure
   - Use `read_authorized_sheet(sheet_name, sample=false)` for full data

2. **Analysis Phase** (optional)
   - Use `Bash` to run Python scripts
   - Available: pandas, numpy, scipy

3. **Submission Phase** (must and only once)
   - To modify: `submit_plan(plan)`
   - Read-only: `submit_answer(message)`
```

---

## 📊 测试场景

### 场景 1：简单读取

```
用户: "销售表有多少行？"

期望流程:
1. read_authorized_sheet("销售表", sample=true)
2. submit_answer("共 1,234 行")
```

### 场景 2：复杂分析

```
用户: "分析销售趋势，预测下月"

期望流程:
1. read_authorized_sheet("销售表", sample=true)
2. read_authorized_sheet("销售表", sample=false)
3. Bash: python 分析脚本
4. Read: result.json
5. submit_plan({actions: [...]})
```

### 场景 3：多表关联

```
用户: "合并订单和客户信息"

期望流程:
1. read_authorized_sheet("订单表", sample=false)
2. read_authorized_sheet("客户表", sample=false)
3. Bash: python pandas merge
4. submit_plan({actions: [...]})
```

---

## 🚀 立即可用的改进

### 已完成 ✅

```json
// config/capabilities.json 已修改
{
  "agent.maxWallTimeSeconds": 600,        // 1分钟 → 10分钟
  "taskRecovery.planningHardWallTimeSeconds": 1800,  // 5分钟 → 30分钟
  "folder.sessionTtlSeconds": 86400,      // 30分钟 → 24小时
  "llm.timeoutSeconds": 300,              // 3分钟 → 5分钟
  "turns.ttlSeconds": 86400               // 1小时 → 24小时
}
```

**效果：** 重启服务后立即生效

```bash
npm run dev
```

---

## 📚 完整文档索引

| 文档 | 用途 | 预计阅读时间 |
|------|------|------------|
| `CLAUDE_SDK_TRANSFORMATION_GUIDE.md` | 完整改造指南 | 30 分钟 |
| `SDK_SESSION_PERSISTENCE.md` | 会话持久化方案 | 15 分钟 |
| `TIMEOUT_CONFIG_CHANGES.md` | 配置改动记录 | 5 分钟 |
| `DEEPSEEK_PROMPT_FOR_SDK.md` | DeepSeek 提示词 | 10 分钟 |
| `WORK_LOG_2026-10-06.md` | 今日工作日志 | 10 分钟 |
| **本文档** | 快速参考卡 | **5 分钟** |

---

## 🎯 预期效果

| 维度 | 改造前 | 改造后 |
|------|--------|--------|
| 探索能力 | ❌ 无 | ✅ 多轮 Python 分析 |
| 执行时间 | 5 分钟 | ✅ 30 分钟 |
| 会话保持 | 30 分钟 | ✅ 24 小时 |
| 准确率 | ~70% | ✅ ~90%+ |
| 双宿主 | 分离 | ✅ 统一协议 |

---

## 💡 关键提示

1. **安全边界不变** - 三层防护保持不变
2. **双宿主统一** - Excel 和 WPS 使用相同协议
3. **数据在请求中** - 前端读取后随请求发送
4. **MCP 守卫写入** - 所有修改必须通过白名单
5. **多轮探索** - Agent 可以充分验证后再提交

---

**预计实施时间：3-5 天**
**预计提升：准确率 70% → 90%+**
**立即生效：超时限制已放宽（重启服务）**
