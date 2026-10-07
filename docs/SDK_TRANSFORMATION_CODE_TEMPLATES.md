# SDK 改造代码模板

> 停用提示（2026-10-06）：以下旧模板不可直接复制实施。开放 SDK Bash/Read、以目录代替沙箱或默认发送数据的示例不符合当前批准范围。按 [隔离 Python 设计](./SDK_ISOLATED_PYTHON_DESIGN_2026-10-06.md) 与 [DeepSeek 实施提示词](./DEEPSEEK_ISOLATED_PYTHON_PROMPT_2026-10-06.md) 重新核对源码、实现真实隔离及双端接线。当前运行器未实现、未启用；原文保留仅供追溯。

## 📝 说明

本文件包含所有需要修改和新增的代码，可以直接复制使用。

---

## 1. 后端改造

### 1.1 修改 `server/app/claude_sdk_agent.py`

#### 修改 `_options()` 函数（第 74-90 行附近）

```python
def _options(sdk, connection, system_prompt, **kwargs):
    state_dir = _state_dir()
    
    # 🆕 创建会话专用工作目录
    session_id = kwargs.get("session_id", uuid.uuid4().hex[:8])
    workspace_dir = state_dir / f"workspace_{session_id}"
    workspace_dir.mkdir(parents=True, exist_ok=True)
    
    return sdk.ClaudeAgentOptions(
        model=connection.model,
        system_prompt=system_prompt,
        
        # ✅ 改造：允许 Read 和 Bash
        tools=["Read", "Bash"],
        
        # ⚠️ 仍禁止这些
        disallowed_tools=[
            "Write",      # 不让直接写文件
            "Edit",       # 不让编辑文件
            "WebFetch", "WebSearch",  # 不让联网
            "Agent", "Skill",  # 不让嵌套
        ],
        
        # ✅ 允许内置工具 + MCP 工具
        allowed_tools=["Read", "Bash", "mcp__excel_bro__*"],
        
        # 🆕 提供 MCP Server
        mcp_servers={
            "excel_bro": _create_excel_mcp_server(kwargs.get("request"), kwargs.get("terminal_holder"))
        },
        
        # 🆕 独立工作目录
        cwd=str(workspace_dir),
        
        permission_mode="dontAsk",
        max_turns=kwargs.get("max_turns", 8),
    )
```

#### 新增 `_create_excel_mcp_server()` 函数

```python
def _create_excel_mcp_server(request, terminal_holder):
    """
    创建 Excel Bro 的 MCP Server
    
    Args:
        request: PlanRequest 对象（包含 sheetsData）
        terminal_holder: dict，用于存储终结状态 {"terminal": None}
    
    Returns:
        MCP Server 实例
    """
    sdk = _sdk()
    tools = []
    
    # ========== 工具 1: 列出授权表 ==========
    
    async def list_authorized_sheets():
        """列出所有已授权的工作表"""
        if not request or not request.workbook:
            return {
                "content": [{
                    "type": "text",
                    "text": json.dumps({"error": "工作簿信息不可用"}, ensure_ascii=False)
                }],
                "is_error": True
            }
        
        sheets_info = [
            {
                "name": sheet.name,
                "rowCount": sheet.rowCount,
                "columnCount": sheet.columnCount,
                "hasData": request.sheetsData and sheet.name in request.sheetsData
            }
            for sheet in request.workbook.worksheets
        ]
        
        return {
            "content": [{
                "type": "text",
                "text": json.dumps({
                    "authorized_sheets": sheets_info,
                    "total": len(sheets_info)
                }, ensure_ascii=False)
            }]
        }
    
    tools.append(
        sdk.tool(
            "list_authorized_sheets",
            "列出所有已授权的工作表信息",
            {"type": "object", "properties": {}, "required": []}
        )(list_authorized_sheets)
    )
    
    # ========== 工具 2: 读取授权表数据 ==========
    
    async def read_authorized_sheet(
        sheet_name: str,
        sample: bool = False,
        max_rows: int = 10000
    ):
        """
        读取已授权的工作表数据
        
        Args:
            sheet_name: 工作表名称
            sample: 是否只返回采样数据（前 1000 行）
            max_rows: 最大返回行数
        """
        # 验证授权
        if not request or not request.workbook:
            return {
                "content": [{"type": "text", "text": "❌ 工作簿信息不可用"}],
                "is_error": True
            }
        
        authorized_sheets = [s.name for s in request.workbook.worksheets]
        if sheet_name not in authorized_sheets:
            return {
                "content": [{
                    "type": "text",
                    "text": json.dumps({
                        "error": f"未授权访问表 '{sheet_name}'",
                        "authorized_sheets": authorized_sheets
                    }, ensure_ascii=False)
                }],
                "is_error": True
            }
        
        # 读取数据
        if not request.sheetsData or sheet_name not in request.sheetsData:
            return {
                "content": [{
                    "type": "text",
                    "text": json.dumps({
                        "error": f"表 '{sheet_name}' 的数据未加载",
                        "hint": "请确保前端已发送该表的数据"
                    }, ensure_ascii=False)
                }],
                "is_error": True
            }
        
        data = request.sheetsData[sheet_name]
        total_rows = len(data)
        
        # 采样
        if sample and total_rows > 1000:
            sampled_data = data[:1000]
        elif total_rows > max_rows:
            sampled_data = data[:max_rows]
        else:
            sampled_data = data
        
        return {
            "content": [{
                "type": "text",
                "text": json.dumps({
                    "sheet": sheet_name,
                    "total_rows": total_rows,
                    "returned_rows": len(sampled_data),
                    "sampled": sample or total_rows > max_rows,
                    "columns": list(sampled_data[0].keys()) if sampled_data else [],
                    "data": sampled_data
                }, ensure_ascii=False, default=str)
            }]
        }
    
    tools.append(
        sdk.tool(
            "read_authorized_sheet",
            "读取已授权的工作表数据。只能读取用户明确授权的表。",
            {
                "type": "object",
                "properties": {
                    "sheet_name": {
                        "type": "string",
                        "description": "工作表名称"
                    },
                    "sample": {
                        "type": "boolean",
                        "description": "是否只返回采样数据（前1000行），默认 false",
                        "default": False
                    },
                    "max_rows": {
                        "type": "integer",
                        "description": "最大返回行数，默认 10000",
                        "default": 10000
                    }
                },
                "required": ["sheet_name"]
            }
        )(read_authorized_sheet)
    )
    
    # ========== 工具 3: 提交计划 ==========
    
    async def submit_plan(plan: dict):
        """
        提交最终的操作计划
        
        Args:
            plan: AnalysisPlan 格式的计划对象
        """
        if terminal_holder["terminal"] is not None:
            return {
                "content": [{"type": "text", "text": "❌ 已经提交过结果，不能重复提交"}],
                "is_error": True
            }
        
        try:
            # 复用现有的计划校验
            from .planner import validate_plan
            validated_plan = validate_plan(plan, request)
            
            terminal_holder["terminal"] = AssistantResponse(
                kind="plan",
                plan=validated_plan,
                provider="model"
            )
            
            return {
                "content": [{
                    "type": "text",
                    "text": "✅ 计划已提交，等待用户确认"
                }]
            }
        except Exception as e:
            return {
                "content": [{
                    "type": "text",
                    "text": f"❌ 计划校验失败：{str(e)}"
                }],
                "is_error": True
            }
    
    tools.append(
        sdk.tool(
            "submit_plan",
            "提交最终的操作计划。计划会显示给用户预览和确认。",
            {
                "type": "object",
                "properties": {
                    "plan": {
                        "type": "object",
                        "description": "Excel 操作计划（AnalysisPlan 格式）"
                    }
                },
                "required": ["plan"]
            }
        )(submit_plan)
    )
    
    # ========== 工具 4: 提交回答 ==========
    
    async def submit_answer(message: str):
        """
        提交只读查询的回答
        
        Args:
            message: 回答内容
        """
        if terminal_holder["terminal"] is not None:
            return {
                "content": [{"type": "text", "text": "❌ 已经提交过结果"}],
                "is_error": True
            }
        
        terminal_holder["terminal"] = AssistantResponse(
            kind="answer",
            message=message,
            provider="model"
        )
        
        return {
            "content": [{
                "type": "text",
                "text": "✅ 回答已提交"
            }]
        }
    
    tools.append(
        sdk.tool(
            "submit_answer",
            "提交只读查询的回答（不修改 Excel）",
            {
                "type": "object",
                "properties": {
                    "message": {
                        "type": "string",
                        "description": "回答内容"
                    }
                },
                "required": ["message"]
            }
        )(submit_answer)
    )
    
    # ========== 工具 5: 提交澄清 ==========
    
    async def submit_clarification(clarification: dict):
        """
        提交结构化的澄清问题
        
        Args:
            clarification: 澄清对象
        """
        if terminal_holder["terminal"] is not None:
            return {
                "content": [{"type": "text", "text": "❌ 已经提交过结果"}],
                "is_error": True
            }
        
        terminal_holder["terminal"] = AssistantResponse(
            kind="clarification",
            clarification=clarification,
            provider="model"
        )
        
        return {
            "content": [{
                "type": "text",
                "text": "✅ 澄清问题已提交"
            }]
        }
    
    tools.append(
        sdk.tool(
            "submit_clarification",
            "当信息不足时，提交结构化的澄清问题",
            {
                "type": "object",
                "properties": {
                    "clarification": {
                        "type": "object",
                        "description": "澄清对象"
                    }
                },
                "required": ["clarification"]
            }
        )(submit_clarification)
    )
    
    return sdk.create_sdk_mcp_server(
        name="excel_bro",
        version="1.0.0",
        tools=tools
    )
```

#### 修改 `run_sdk_plan()` 函数

```python
def run_sdk_plan(
    sdk,
    connection: ModelConnection,
    request: "PlanRequest",
    prompt: str | None = None,
    system_prompt: str | None = None,
    checkpoint: dict | None = None,
    **kwargs,
):
    """运行 SDK Agent"""
    
    # ... 现有代码 ...
    
    # 🆕 传递 request 和 terminal_holder
    terminal_holder = {"terminal": None}
    
    options = _options(
        sdk,
        connection,
        system_prompt or AGENT_SYSTEM_PROMPT,
        request=request,  # 🆕
        terminal_holder=terminal_holder,  # 🆕
        **kwargs
    )
    
    # ... 运行 agent ...
    
    # 🆕 从 terminal_holder 获取结果
    if terminal_holder["terminal"] is not None:
        return terminal_holder["terminal"]
    
    # ... 现有的错误处理 ...
```

---

### 1.2 修改 `server/app/models.py`

```python
class PlanRequest(BaseModel):
    prompt: str
    workbook: WorkbookSnapshot
    modelId: str | None = None
    resumeCheckpoint: dict | None = None
    
    # 🆕 新增：前端发送的数据
    sheetsData: dict[str, list[dict]] | None = None
    
    # ... 其他字段 ...
```

---

## 2. 前端改造

### 2.1 Excel - 新增 `apps/excel-addin/src/officeDataReader.ts`

```typescript
/**
 * Office.js 数据读取器
 */

/**
 * 读取工作表的所有数据
 * @param sheetName 工作表名称
 * @param maxRows 最大读取行数，默认 10000
 * @returns 对象数组
 */
export async function readSheetDataOffice(
  sheetName: string,
  maxRows: number = 10000
): Promise<Record<string, any>[]> {
  return await Excel.run(async (context) => {
    try {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const usedRange = sheet.getUsedRange();
      usedRange.load(['values', 'rowCount']);
      await context.sync();
      
      const values = usedRange.values;
      const totalRows = usedRange.rowCount;
      
      if (values.length === 0) {
        return [];
      }
      
      // 第一行是表头
      const headers = values[0].map(h => String(h));
      
      // 数据行（限制最大行数）
      const dataRows = values.slice(1, Math.min(values.length, maxRows + 1));
      
      // 转换为对象数组
      const data = dataRows.map(row => {
        const obj: Record<string, any> = {};
        headers.forEach((header, index) => {
          obj[header] = standardizeValue(row[index]);
        });
        return obj;
      });
      
      console.log(`[Excel] 读取表 "${sheetName}": ${data.length}/${totalRows} 行`);
      
      return data;
    } catch (error) {
      console.error(`[Excel] 读取表 "${sheetName}" 失败:`, error);
      throw error;
    }
  });
}

/**
 * 标准化值（统一类型）
 */
function standardizeValue(value: any): any {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  
  if (typeof value === 'number') {
    return value;
  }
  
  if (typeof value === 'string') {
    return value.trim();
  }
  
  if (typeof value === 'boolean') {
    return value;
  }
  
  // Excel 日期是数字（从 1900-01-01 开始的天数）
  if (typeof value === 'number' && value > 40000 && value < 60000) {
    // 可能是日期（2009-2164 年之间）
    const date = new Date((value - 25569) * 86400 * 1000);
    return date.toISOString().split('T')[0];
  }
  
  return String(value);
}

/**
 * 读取所有授权表的数据
 */
export async function readAllAuthorizedSheetsOffice(
  sheetNames: string[],
  maxRowsPerSheet: number = 10000
): Promise<Record<string, Record<string, any>[]>> {
  const sheetsData: Record<string, Record<string, any>[]> = {};
  
  for (const sheetName of sheetNames) {
    try {
      sheetsData[sheetName] = await readSheetDataOffice(sheetName, maxRowsPerSheet);
    } catch (error) {
      console.error(`读取表 "${sheetName}" 失败，跳过`, error);
      sheetsData[sheetName] = [];
    }
  }
  
  return sheetsData;
}
```

### 2.2 WPS - 新增 `apps/wps-addin/src/wpsDataReader.ts`

```typescript
/**
 * WPS JSAPI 数据读取器
 */

/**
 * 读取工作表的所有数据
 * @param sheetName 工作表名称
 * @param maxRows 最大读取行数，默认 10000
 * @returns 对象数组
 */
export async function readSheetDataWps(
  sheetName: string,
  maxRows: number = 10000
): Promise<Record<string, any>[]> {
  try {
    const sheet = Application.Worksheets.Item(sheetName);
    const usedRange = sheet.UsedRange;
    
    // WPS 返回的是 COM 对象，需要转换
    const values = usedRange.Value2;
    
    if (!values || values.length === 0) {
      return [];
    }
    
    // 第一行是表头
    const headers = values[0].map((h: any) => String(h));
    
    // 数据行（限制最大行数）
    const dataRows = values.slice(1, Math.min(values.length, maxRows + 1));
    
    // 转换为对象数组
    const data = dataRows.map((row: any[]) => {
      const obj: Record<string, any> = {};
      headers.forEach((header, index) => {
        obj[header] = normalizeWpsValue(row[index]);
      });
      return obj;
    });
    
    console.log(`[WPS] 读取表 "${sheetName}": ${data.length}/${values.length - 1} 行`);
    
    return data;
  } catch (error) {
    console.error(`[WPS] 读取表 "${sheetName}" 失败:`, error);
    throw error;
  }
}

/**
 * 标准化 WPS 值（处理 COM 类型）
 */
function normalizeWpsValue(value: any): any {
  // 空值
  if (value === null || value === undefined || value === '') {
    return null;
  }
  
  // 数字
  if (typeof value === 'number') {
    // WPS 日期也是数字（从 1899-12-30 开始）
    if (value > 40000 && value < 60000) {
      // 可能是日期
      const date = new Date((value - 25569) * 86400 * 1000);
      return date.toISOString().split('T')[0];
    }
    return value;
  }
  
  // 字符串
  if (typeof value === 'string') {
    return value.trim();
  }
  
  // 布尔
  if (typeof value === 'boolean') {
    return value;
  }
  
  // Date 对象
  if (value instanceof Date) {
    return value.toISOString().split('T')[0];
  }
  
  // 其他情况转字符串
  return String(value);
}

/**
 * 读取所有授权表的数据
 */
export async function readAllAuthorizedSheetsWps(
  sheetNames: string[],
  maxRowsPerSheet: number = 10000
): Promise<Record<string, Record<string, any>[]>> {
  const sheetsData: Record<string, Record<string, any>[]> = {};
  
  for (const sheetName of sheetNames) {
    try {
      sheetsData[sheetName] = await readSheetDataWps(sheetName, maxRowsPerSheet);
    } catch (error) {
      console.error(`读取表 "${sheetName}" 失败，跳过`, error);
      sheetsData[sheetName] = [];
    }
  }
  
  return sheetsData;
}
```

### 2.3 修改 API 调用

**Excel: `apps/excel-addin/src/api.ts`**

```typescript
import { readAllAuthorizedSheetsOffice } from './officeDataReader';

export async function sendPlanRequest(
  prompt: string,
  workbook: WorkbookSnapshot,
  modelId: string
): Promise<AssistantResponse> {
  
  // 🆕 读取所有授权表的数据
  let sheetsData: Record<string, any[]> | null = null;
  
  // 只有 SDK 模式才需要发送数据
  if (modelId?.includes('sdk')) {
    const sheetNames = workbook.worksheets.map(s => s.name);
    sheetsData = await readAllAuthorizedSheetsOffice(sheetNames);
    
    console.log('[API] 发送数据:', Object.keys(sheetsData).length, '个表');
  }
  
  const response = await fetch('/api/turn', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      prompt,
      workbook,
      modelId,
      sheetsData,  // 🆕
    })
  });
  
  return response.json();
}
```

**WPS: `apps/wps-addin/src/wpsHost.ts`**

```typescript
import { readAllAuthorizedSheetsWps } from './wpsDataReader';

export async function sendPlanRequest(
  prompt: string,
  workbook: WorkbookSnapshot,
  modelId: string
): Promise<AssistantResponse> {
  
  // 🆕 读取所有授权表的数据
  let sheetsData: Record<string, any[]> | null = null;
  
  // 只有 SDK 模式才需要发送数据
  if (modelId?.includes('sdk')) {
    const sheetNames = workbook.worksheets.map(s => s.name);
    sheetsData = await readAllAuthorizedSheetsWps(sheetNames);
    
    console.log('[API] 发送数据:', Object.keys(sheetsData).length, '个表');
  }
  
  const response = await fetch('/api/turn', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      prompt,
      workbook,
      modelId,
      sheetsData,  // 🆕
    })
  });
  
  return response.json();
}
```

---

## 3. System Prompt 更新

### 3.1 在 `server/app/claude_sdk_agent.py` 中更新

```python
ENHANCED_SYSTEM_PROMPT = """
你是 Excel Bro 的智能分析引擎，专门帮助用户分析和处理 Excel/WPS 数据。

## 可用工具

### Excel Bro 工具（必须使用）

1. **list_authorized_sheets()** - 列出所有已授权的工作表

2. **read_authorized_sheet(sheet_name, sample, max_rows)** - 读取工作表数据
   - sheet_name: 表名
   - sample: true 时返回采样（前1000行）
   - max_rows: 最大返回行数（默认10000）

3. **submit_plan(plan)** - 提交最终操作计划（修改 Excel）

4. **submit_answer(message)** - 提交只读查询的回答

5. **submit_clarification(clarification)** - 提交澄清问题

### 内置工具（自由使用）

1. **Read** - 读取临时文件
2. **Bash** - 运行 Python 脚本（可用 pandas, numpy, scipy）

## 工作流程

1. **探索** - 先用 list_authorized_sheets() 和 read_authorized_sheet() 了解数据
2. **分析** - 如果需要，用 Bash 运行 Python 脚本
3. **提交** - 调用 submit_plan/submit_answer/submit_clarification（仅一次）

## 重要规则

✅ 充分探索再决策 - 先采样查看结构，再读取完整数据
✅ 使用 Python 进行复杂计算 - pandas, numpy, scipy
✅ 生成准确计划 - 确保字段名、范围、公式都正确

❌ 不要猜测 - 必须先读取数据确认
❌ 不要直接写文件 - 所有修改通过 submit_plan
❌ 不要访问未授权表 - 只能读取用户授权的表
❌ 不要多次提交 - 只能调用一次 submit_*

## Python 分析示例

```python
import pandas as pd
import json

# 读取数据
with open('data.json') as f:
    data = json.load(f)

df = pd.DataFrame(data['data'])

# 分析
result = df.groupby('类别')['销量'].sum().to_dict()

# 保存
with open('result.json', 'w') as f:
    json.dump(result, f, ensure_ascii=False)
```

记住：准确性优于速度，充分探索后再提交。
"""

# 在 run_sdk_plan() 中使用
system_prompt = system_prompt or ENHANCED_SYSTEM_PROMPT
```

---

## 4. 测试代码

### 4.1 后端测试 `server/tests/test_sdk_transformation.py`

```python
import pytest
from app.claude_sdk_agent import _create_excel_mcp_server
from app.models import PlanRequest, WorkbookSnapshot, WorksheetSnapshot

def test_list_authorized_sheets():
    """测试列出授权表"""
    request = PlanRequest(
        prompt="test",
        workbook=WorkbookSnapshot(
            worksheets=[
                WorksheetSnapshot(name="Sheet1", rowCount=100, columnCount=10),
                WorksheetSnapshot(name="Sheet2", rowCount=50, columnCount=5),
            ]
        ),
        sheetsData={}
    )
    
    terminal_holder = {"terminal": None}
    server = _create_excel_mcp_server(request, terminal_holder)
    
    # 调用 list_authorized_sheets
    # ... 验证返回 ...

def test_read_authorized_sheet():
    """测试读取授权表"""
    request = PlanRequest(
        prompt="test",
        workbook=WorkbookSnapshot(
            worksheets=[WorksheetSnapshot(name="Sales", rowCount=100, columnCount=3)]
        ),
        sheetsData={
            "Sales": [
                {"产品": "A", "销量": 100},
                {"产品": "B", "销量": 200},
            ]
        }
    )
    
    terminal_holder = {"terminal": None}
    server = _create_excel_mcp_server(request, terminal_holder)
    
    # 调用 read_authorized_sheet
    # ... 验证返回数据 ...

def test_unauthorized_access():
    """测试未授权访问"""
    request = PlanRequest(
        prompt="test",
        workbook=WorkbookSnapshot(
            worksheets=[WorksheetSnapshot(name="Sheet1", rowCount=100, columnCount=10)]
        ),
        sheetsData={"Sheet1": []}
    )
    
    terminal_holder = {"terminal": None}
    server = _create_excel_mcp_server(request, terminal_holder)
    
    # 尝试读取未授权的 "Sheet2"
    # ... 验证返回错误 ...
```

### 4.2 前端测试 `apps/excel-addin/src/officeDataReader.test.ts`

```typescript
import { readSheetDataOffice } from './officeDataReader';

describe('officeDataReader', () => {
  it('should read sheet data', async () => {
    // Mock Excel.run
    global.Excel = {
      run: jest.fn(async (callback) => {
        const mockContext = {
          workbook: {
            worksheets: {
              getItem: jest.fn(() => ({
                getUsedRange: jest.fn(() => ({
                  load: jest.fn(),
                  values: [
                    ['产品', '销量'],
                    ['A', 100],
                    ['B', 200],
                  ],
                  rowCount: 3,
                })),
              })),
            },
          },
          sync: jest.fn(),
        };
        return callback(mockContext);
      }),
    } as any;
    
    const data = await readSheetDataOffice('Sheet1');
    
    expect(data).toEqual([
      { '产品': 'A', '销量': 100 },
      { '产品': 'B', '销量': 200 },
    ]);
  });
});
```

---

## 5. 启动脚本

### 5.1 更新 `package.json`

```json
{
  "scripts": {
    "dev": "npm run dev:server & npm run dev:excel & npm run dev:wps",
    "dev:server": "cd server && python -m uvicorn app.main:app --reload --port 8765",
    "dev:server:sdk": "powershell -File scripts/start_claude_sdk_server.ps1",
    "dev:excel": "cd apps/excel-addin && npm run dev",
    "dev:wps": "cd apps/wps-addin && npm run dev"
  }
}
```

### 5.2 README 使用说明

```markdown
## 启动服务

### 普通模式（不使用 SDK）

```bash
npm run dev
```

### SDK 模式（支持复杂分析）

```bash
npm run dev:server:sdk  # 使用特殊脚本启动后端
npm run dev:excel       # 启动 Excel 插件
npm run dev:wps         # 启动 WPS 插件
```

## 模型配置

在 `server/.env` 中配置：

```bash
# 使用 Claude SDK
AI_BASE_URL=https://api.anthropic.com
AI_MODEL=claude-3-5-sonnet-20241022
AI_API_KEY=sk-ant-...

# 或使用 DeepSeek
AI_BASE_URL=https://api.deepseek.com/v1
AI_MODEL=deepseek-reasoner
AI_API_KEY=sk-...
```
```

---

## 📋 检查清单

使用这些代码模板后，检查：

- [ ] 后端：`_options()` 已修改
- [ ] 后端：`_create_excel_mcp_server()` 已添加
- [ ] 后端：`PlanRequest` 添加了 `sheetsData`
- [ ] 前端 Excel：`officeDataReader.ts` 已创建
- [ ] 前端 WPS：`wpsDataReader.ts` 已创建
- [ ] 前端：API 调用包含 `sheetsData`
- [ ] System Prompt 已更新
- [ ] 测试代码已添加
- [ ] 启动脚本已更新
- [ ] README 已更新

---

**所有代码模板准备完毕，可以开始实施了！** 🚀
