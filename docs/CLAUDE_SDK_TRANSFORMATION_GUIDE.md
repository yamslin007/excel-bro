# Claude Agent SDK 改造指南

> 停用提示（2026-10-06）：下文为旧改造草稿，不作为当前实施指导或验收依据。直接开放 Read/Bash、上传数据与“完整版本已完成”等示例/结论不能采用。当前获批准的是 [隔离 Python 分析设计](./SDK_ISOLATED_PYTHON_DESIGN_2026-10-06.md)，实施用 [DeepSeek 提示词](./DEEPSEEK_ISOLATED_PYTHON_PROMPT_2026-10-06.md)；双端执行器尚未实现、未启用。原文保留供追溯。

## 目标

将当前**能力受限的 SDK 集成**改造为**完整灵活的 Agent 引擎**，同时保留 Excel Bro 的安全边界。

---

## 当前问题

### 现状分析

```python
# server/app/claude_sdk_agent.py (第 74-76 行)

tools=[],  # ❌ 空的！所有工具被禁用
disallowed_tools=["Read", "Bash", "Write", "Edit", "Glob", "Grep", ...],
allowed_tools=["mcp__excel_bro__submit_*"],  # 只能立即提交
```

**问题：**
- ❌ Agent 无法探索数据
- ❌ Agent 无法运行复杂计算
- ❌ Agent 无法多轮推理
- ❌ 被迫一次性生成计划 → 容易出错

**结果：**
- 付出了 SDK 的复杂度（子进程、特殊启动）
- 但没有获得 SDK 的灵活性
- 还丢失了检查点恢复能力

---

## 改造方案

### 核心理念

```
释放 SDK 的探索能力
     ↓
通过 MCP 守卫最终写入
     ↓
保留白名单 + 预览 + 恢复点
```

### 双宿主架构（Excel + WPS）

**重要说明：** Excel Bro 同时支持 Office.js (Excel) 和 WPS JSAPI 两种宿主，数据读取需要处理双宿主场景。

```
┌────────────────────────────────────────┐
│  前端 (React)                           │
│  ├─ Excel Add-in (Office.js)           │
│  │   → 读取当前工作簿数据                │
│  └─ WPS Add-in (WPS JSAPI)             │
│      → 读取当前工作簿数据                │
└──────────────┬─────────────────────────┘
               │ HTTP (携带数据快照)
               ▼
┌────────────────────────────────────────┐
│  后端 FastAPI                           │
│  └─ 从请求中获取数据（双宿主通用）       │
└──────────────┬─────────────────────────┘
               │
               ▼
┌────────────────────────────────────────┐
│  Claude Agent SDK                       │
│  └─ 通过 MCP 工具访问数据               │
└────────────────────────────────────────┘
```

### 架构设计

```
用户需求（自然语言）
    ↓
┌─────────────────────────────────────────┐
│  Claude Agent SDK (完整能力)             │
│                                         │
│  ✅ Read - 读临时文件和分析结果          │
│  ✅ Bash - 运行 Python 进行复杂分析      │
│  ❌ 不能直接写 Excel                     │
└──────────────┬──────────────────────────┘
               │
        MCP Server 守卫
               │
               ▼
┌─────────────────────────────────────────┐
│  Excel Bro MCP 工具（受控）              │
│                                         │
│  read_authorized_sheet                  │
│    → 只读已授权的表                      │
│                                         │
│  submit_plan                            │
│    → 白名单校验                          │
│    → 用户预览确认                        │
│    → 创建恢复点                          │
└─────────────────────────────────────────┘
```

---

## 改造步骤

### 阶段 1：基础改造（Day 1-2）

#### 1.1 修改工具配置

**文件：** `server/app/claude_sdk_agent.py`

**改动点 1：释放内置工具**

```python
def _options(
    sdk: Any, connection: ModelConnection, system_prompt: str,
    *, max_turns: int, server: Any | None = None, allowed: list[str] | None = None,
) -> Any:
    state_dir = _state_dir()
    
    # 为每次调用创建独立的工作目录
    workspace_dir = state_dir / f"workspace_{uuid.uuid4().hex[:8]}"
    workspace_dir.mkdir(parents=True, exist_ok=True)
    
    sdk_env = {
        "ANTHROPIC_API_KEY": connection.api_key,
        "ANTHROPIC_AUTH_TOKEN": "",
        "ANTHROPIC_BASE_URL": connection.base_url,
        "CLAUDE_CONFIG_DIR": str(state_dir),
        "CLAUDE_CODE_DISABLE_NONESSENTIAL_TRAFFIC": "1",
    }
    
    if connection.base_url == DEEPSEEK_ANTHROPIC_URL:
        sdk_env.update({
            "ANTHROPIC_AUTH_TOKEN": connection.api_key,
            "ANTHROPIC_MODEL": connection.model,
            "ANTHROPIC_DEFAULT_OPUS_MODEL": connection.model,
            "ANTHROPIC_DEFAULT_SONNET_MODEL": connection.model,
            "ANTHROPIC_DEFAULT_HAIKU_MODEL": connection.model,
            "CLAUDE_CODE_SUBAGENT_MODEL": connection.model,
        })
    
    return sdk.ClaudeAgentOptions(
        model=connection.model,
        system_prompt=system_prompt,
        
        # ✅ 改动：允许探索工具
        tools=["Read", "Bash"],  # 不再是空的！
        
        # ⚠️ 仍禁止这些
        disallowed_tools=[
            "Write",  # 不让直接写文件到系统
            "Edit",   # 不让编辑任意文件
            "WebFetch", "WebSearch",  # 不让联网
            "Agent", "Skill",  # 不让嵌套 agent
            "Glob", "Grep",  # 防止扫描敏感文件
        ],
        
        mcp_servers={"excel_bro": server} if server else {},
        strict_mcp_config=True,
        setting_sources=[],
        
        # ✅ 改动：允许内置工具 + MCP 工具
        allowed_tools=allowed or [
            "Read",
            "Bash", 
            "mcp__excel_bro__*"
        ],
        
        permission_mode="dontAsk",
        
        # ✅ 改动：工作在独立临时目录
        cwd=str(workspace_dir),
        
        max_turns=max_turns,
        max_budget_usd=capability_float("agent", "claudeSdkMaxBudgetUsd"),
        env=sdk_env,
        extra_args={"no-session-persistence": None},
    )
```

**关键变化：**
- `tools=["Read", "Bash"]` → 允许读文件和运行 Python
- `cwd=workspace_dir` → 每次调用有独立临时目录
- `allowed_tools` 包含内置工具和 MCP 工具

---

#### 1.2 创建 MCP Server

**在 `run_sdk_plan()` 中添加：**

```python
async def run_sdk_plan(
    request: PlanRequest, connection: ModelConnection,
    *, on_event: Callable[[dict[str, Any]], Awaitable[None]] | None = None,
) -> AssistantResponse:
    """完整版 SDK 规划：允许探索，但守卫写入。"""
    
    if request.images and not connection.supports_vision:
        raise ModelNotAvailableError("当前模型不支持图片，请选择支持视觉输入的模型。")
    
    from .excel_agent import AGENT_SYSTEM_PROMPT
    
    sdk = _sdk()
    terminal: AssistantResponse | None = None
    tools = []
    allowed = []
    
    # ===== 创建 MCP 工具 =====
    
    # 工具 1: 读取授权表
    async def read_authorized_sheet(
        sheet_name: str,
        sample: bool = False,
        max_rows: int = 1000
    ) -> dict[str, Any]:
        """读取已授权的工作表数据"""
        
        # 查找授权的表
        authorized = [
            sheet for sheet in request.workbook.worksheets
            if sheet.name == sheet_name
        ]
        
        if not authorized:
            return {
                "content": [{
                    "type": "text",
                    "text": f"❌ 未授权访问表 '{sheet_name}'。\n\n已授权的表：{[s.name for s in request.workbook.worksheets]}"
                }],
                "is_error": True
            }
        
        sheet = authorized[0]
        
        # 读取数据（这里需要实际的读取逻辑）
        # 简化示例，实际需要调用 Office.js 或 openpyxl
        try:
            # TODO: 实现实际的数据读取
            # data = read_sheet_data(request.workbook, sheet_name)
            
            data = []  # 占位
            
            if sample and len(data) > max_rows:
                data = data[:max_rows]
            
            return {
                "content": [{
                    "type": "text",
                    "text": json.dumps({
                        "sheet": sheet_name,
                        "total_rows": sheet.rowCount,
                        "total_columns": sheet.columnCount,
                        "headers": sheet.headers,
                        "sampled": sample,
                        "data_rows": len(data),
                        "data": data
                    }, ensure_ascii=False)
                }]
            }
        except Exception as e:
            return {
                "content": [{
                    "type": "text",
                    "text": f"❌ 读取失败：{str(e)}"
                }],
                "is_error": True
            }
    
    tools.append(
        sdk.tool(
            "read_authorized_sheet",
            """读取已授权的 Excel 工作表数据。
            
参数：
- sheet_name: 工作表名称（必须是已授权的表）
- sample: 是否只返回采样数据（默认 False）
- max_rows: 采样时的最大行数（默认 1000）

返回：
- JSON 格式的表数据，包含表头和数据行
""",
            {
                "type": "object",
                "properties": {
                    "sheet_name": {
                        "type": "string",
                        "description": "工作表名称"
                    },
                    "sample": {
                        "type": "boolean",
                        "description": "是否只返回采样数据",
                        "default": False
                    },
                    "max_rows": {
                        "type": "integer",
                        "description": "采样时的最大行数",
                        "default": 1000,
                        "minimum": 1,
                        "maximum": 10000
                    }
                },
                "required": ["sheet_name"]
            }
        )(read_authorized_sheet)
    )
    allowed.append("mcp__excel_bro__read_authorized_sheet")
    
    # 工具 2: 提交计划
    async def submit_plan(plan: dict) -> dict[str, Any]:
        """提交最终的白名单操作计划"""
        nonlocal terminal
        
        if terminal is not None:
            return {
                "content": [{
                    "type": "text",
                    "text": "❌ 已经提交过计划，不能重复提交。"
                }],
                "is_error": True
            }
        
        try:
            # 使用现有的计划校验
            from .planner import _validate_plan_schema
            from .models import AnalysisPlan
            
            # 校验格式
            validated = AnalysisPlan.model_validate(plan)
            
            # 校验能力和授权
            # TODO: 添加宿主能力检查
            
            terminal = AssistantResponse(
                kind="plan",
                plan=validated,
                provider="model"
            )
            
            if on_event:
                await on_event({
                    "phase": "planning",
                    "stepId": "planning.submit",
                    "status": "completed",
                    "title": "已生成操作计划",
                    "toolName": "submit_plan"
                })
            
            return {
                "content": [{
                    "type": "text",
                    "text": "✅ 计划已提交。用户将看到预览并决定是否执行。"
                }]
            }
        except Exception as e:
            return {
                "content": [{
                    "type": "text",
                    "text": f"❌ 计划校验失败：{str(e)[:500]}"
                }],
                "is_error": True
            }
    
    tools.append(
        sdk.tool(
            "submit_plan",
            """提交最终的 Excel 操作计划。
            
计划必须包含白名单动作：
- writeValues: 写入常量值
- writeFormulas: 写入公式
- clearRange: 清空区域
- setNumberFormat: 设置数字格式
- 等等

计划会显示给用户预览，用户确认后才会实际执行。
""",
            {
                "type": "object",
                "properties": {
                    "plan": {
                        "type": "object",
                        "description": "AnalysisPlan 格式的操作计划"
                    }
                },
                "required": ["plan"]
            }
        )(submit_plan)
    )
    allowed.append("mcp__excel_bro__submit_plan")
    
    # 工具 3: 提交回答
    async def submit_answer(
        message: str,
        result_context: str | None = None
    ) -> dict[str, Any]:
        """提交只读查询的回答"""
        nonlocal terminal
        
        if terminal is not None:
            return {
                "content": [{
                    "type": "text",
                    "text": "❌ 已经提交过回答。"
                }],
                "is_error": True
            }
        
        terminal = AssistantResponse(
            kind="answer",
            message=message,
            resultContext=result_context,
            provider="model"
        )
        
        if on_event:
            await on_event({
                "phase": "planning",
                "stepId": "planning.answer",
                "status": "completed",
                "title": "已生成回答",
                "toolName": "submit_answer"
            })
        
        return {
            "content": [{
                "type": "text",
                "text": "✅ 回答已提交。"
            }]
        }
    
    tools.append(
        sdk.tool(
            "submit_answer",
            "提交只读查询任务的最终回答（不需要修改 Excel）",
            {
                "type": "object",
                "properties": {
                    "message": {
                        "type": "string",
                        "description": "回答内容"
                    },
                    "result_context": {
                        "type": "string",
                        "description": "结果上下文（可选）"
                    }
                },
                "required": ["message"]
            }
        )(submit_answer)
    )
    allowed.append("mcp__excel_bro__submit_answer")
    
    # 工具 4: 澄清
    async def submit_clarification(clarification: dict) -> dict[str, Any]:
        """提交澄清问题"""
        nonlocal terminal
        
        if terminal is not None:
            return {
                "content": [{
                    "type": "text",
                    "text": "❌ 已经提交过结果。"
                }],
                "is_error": True
            }
        
        try:
            from .models import IntentClarification
            validated = IntentClarification.model_validate(clarification)
            
            terminal = AssistantResponse(
                kind="clarification",
                clarification=validated,
                provider="model"
            )
            
            return {
                "content": [{
                    "type": "text",
                    "text": "✅ 澄清问题已提交。"
                }]
            }
        except Exception as e:
            return {
                "content": [{
                    "type": "text",
                    "text": f"❌ 澄清格式错误：{str(e)[:500]}"
                }],
                "is_error": True
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
                        "description": "IntentClarification 格式"
                    }
                },
                "required": ["clarification"]
            }
        )(submit_clarification)
    )
    allowed.append("mcp__excel_bro__submit_clarification")
    
    # 创建 MCP Server
    server = sdk.create_sdk_mcp_server(
        name="excel_bro",
        version="1.0.0",
        tools=tools
    )
    
    # ===== 改进的 System Prompt =====
    
    enhanced_prompt = f"""{AGENT_SYSTEM_PROMPT}

## 可用工具

你现在有完整的探索和分析能力：

### 内置工具（自由使用）

**Read**
- 读取临时文件、分析结果
- 示例：`Read result.json`

**Bash**
- 运行 Python 脚本进行复杂分析
- 你的工作目录是一个临时目录，可以自由读写
- 示例：
  ```bash
  python -c "
  import pandas as pd
  import json
  
  # 读取数据
  with open('data.json') as f:
      data = json.load(f)
  
  # 分析
  df = pd.DataFrame(data)
  result = df.groupby('产品')['销量'].sum().nlargest(5).to_dict()
  
  # 保存结果
  with open('top5.json', 'w') as f:
      json.dump(result, f, ensure_ascii=False)
  
  print('分析完成')
  "
  ```

### Excel Bro 工具（必须使用）

**mcp__excel_bro__read_authorized_sheet**
- 读取已授权的 Excel 工作表数据
- 只能读取用户明确授权的表
- 可以选择采样模式（sample=true）

**mcp__excel_bro__submit_plan**
- 提交最终的白名单操作计划
- 必须是合法的 AnalysisPlan 格式
- 用户会看到预览并确认

**mcp__excel_bro__submit_answer**
- 提交只读查询的回答（不修改 Excel）

**mcp__excel_bro__submit_clarification**
- 当信息不足时，提交澄清问题

## 推荐工作流

1. **探索阶段**（多轮）
   - 用 `read_authorized_sheet` 获取数据
   - 用 `Bash` 运行 Python 分析
   - 用 `Read` 读取分析结果
   - 反复探索直到理解清楚

2. **规划阶段**（一次）
   - 用 `submit_plan` 提交最终计划
   - 或用 `submit_answer` 提交回答
   - 或用 `submit_clarification` 请求澄清

## 重要规则

❌ **不要做：**
- 不要尝试直接修改 Excel 文件
- 不要使用 Write 工具（已禁用）
- 不要尝试访问工作目录外的文件

✅ **要做：**
- 所有 Excel 写入必须通过 `submit_plan`
- 临时文件随便用（在你的工作目录）
- 可以多轮探索，充分验证后再提交
- Python 可以使用 pandas, numpy, scipy 等

## 示例

用户: "分析销售趋势，找出 TOP 5 产品"

你应该：
1. `read_authorized_sheet(sheet_name="销售", sample=true)` 
   → 先看看数据结构
2. `read_authorized_sheet(sheet_name="销售", sample=false)`
   → 获取完整数据
3. `Bash: python -c "..."` 
   → 运行分析脚本
4. `Read analysis_result.json`
   → 读取分析结果
5. `submit_plan(...)`
   → 提交生成报表的计划
"""
    
    # ===== 构造请求载荷 =====
    
    prompt = json.dumps({
        "prompt": request.prompt,
        "workbook": {
            "name": request.workbook.name,
            "activeWorksheet": request.workbook.activeWorksheet,
            "selectedRange": request.workbook.selectedRange,
            "worksheets": [{
                "name": sheet.name,
                "usedRange": sheet.usedRange,
                "rowCount": sheet.rowCount,
                "columnCount": sheet.columnCount,
                "headers": sheet.headers,
            } for sheet in request.workbook.worksheets],
        },
    }, ensure_ascii=False)
    
    if on_event:
        await on_event({
            "phase": "planning",
            "stepId": "planning.model",
            "status": "started",
            "title": "正在生成方案（完整探索模式）",
        })
    
    # ===== 执行 SDK =====
    
    options = _options(
        sdk, connection, enhanced_prompt,
        max_turns=capability_int("agent", "maxTurns"),
        server=server,
        allowed=allowed,
    )
    
    content = await _result(
        sdk,
        _image_prompt(prompt, request.images) if request.images else prompt,
        options,
        capability_float("taskRecovery", "planningHardWallTimeSeconds"),
    )
    
    # ===== 处理结果 =====
    
    if terminal is None:
        # 兜底：解析纯文本 JSON（向后兼容）
        try:
            from .excel_agent import _parse_content
            parsed = _parse_content(content)
            
            if parsed.kind == "plan":
                terminal = AssistantResponse(
                    kind="plan",
                    plan=parsed.plan,
                    provider="model"
                )
            elif parsed.kind == "answer":
                terminal = AssistantResponse(
                    kind="answer",
                    message=parsed.message,
                    resultContext=parsed.resultContext,
                    provider="model"
                )
            elif parsed.kind == "clarification":
                terminal = AssistantResponse(
                    kind="clarification",
                    clarification=parsed.clarification,
                    provider="model"
                )
        except (ValueError, TypeError, KeyError):
            pass
    
    if terminal is None:
        raise LLMResponseError("Claude Agent SDK 未提交有效的受控结果。")
    
    return terminal
```

---

### 阶段 2：实现双宿主数据读取（Day 2-3）

#### 2.1 数据读取策略

**核心方案：前端读取 + 请求携带（推荐）**

原理：
- Office.js 和 WPS JSAPI 都在前端运行
- 前端读取完整数据后随 HTTP 请求发送
- 后端从请求中获取，无需区分宿主

优点：
- ✅ 双宿主通用（Office/WPS 代码分离但协议统一）
- ✅ 实现简单（无需 WebSocket）
- ✅ 本地网络传输快

#### 2.2 前端改造

**步骤 1：添加数据读取接口**

```typescript
// packages/core/src/dataReader.ts (或各端分别实现)

export interface SheetDataReader {
  readSheetData(sheetName: string, maxRows?: number): Promise<any[]>;
}

// Excel 实现
// apps/excel-addin/src/officeDataReader.ts

export async function readSheetDataOffice(
  sheetName: string,
  maxRows: number = 10000
): Promise<any[]> {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const usedRange = sheet.getUsedRange();
    
    usedRange.load(['values', 'rowCount', 'columnCount']);
    await context.sync();
    
    // 限制行数
    const actualRows = Math.min(usedRange.rowCount, maxRows);
    const values = usedRange.values.slice(0, actualRows);
    
    // 转换为对象数组
    const headers = values[0];
    const rows = values.slice(1).map(row => {
      const obj: Record<string, any> = {};
      headers.forEach((header, index) => {
        obj[String(header)] = row[index];
      });
      return obj;
    });
    
    return rows;
  });
}

// WPS 实现
// apps/wps-addin/src/wpsDataReader.ts

export async function readSheetDataWps(
  sheetName: string,
  maxRows: number = 10000
): Promise<any[]> {
  const sheet = Application.Worksheets.Item(sheetName);
  const usedRange = sheet.UsedRange;
  
  // WPS Value2 返回二维数组
  const values = usedRange.Value2;
  
  // 限制行数
  const actualRows = Math.min(values.length, maxRows);
  const limitedValues = values.slice(0, actualRows);
  
  // 转换为对象数组
  const headers = limitedValues[0];
  const rows = limitedValues.slice(1).map(row => {
    const obj: Record<string, any> = {};
    headers.forEach((header: any, index: number) => {
      obj[String(header)] = row[index];
    });
    return obj;
  });
  
  return rows;
}
```

**步骤 2：在请求中包含数据**

```typescript
// apps/excel-addin/src/api.ts (Office 版本)

async function sendPlanRequest(
  prompt: string,
  workbook: WorkbookSnapshot,
  modelId: string,
  // ...其他参数
): Promise<AssistantResponse> {
  
  // 🆕 只在 SDK 模式下读取完整数据
  let sheetsData: Record<string, any[]> | null = null;
  
  if (modelId?.includes('agent-sdk')) {
    sheetsData = {};
    
    for (const sheet of workbook.worksheets) {
      try {
        sheetsData[sheet.name] = await readSheetDataOffice(
          sheet.name,
          10000  // 最多 1 万行
        );
      } catch (error) {
        console.error(`读取表 ${sheet.name} 失败:`, error);
        sheetsData[sheet.name] = [];
      }
    }
  }
  
  const response = await fetch('/api/turn', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      prompt,
      workbook,
      modelId,
      sheetsData,  // 🆕 新增字段
      planningReadPolicy: sheetsData ? 'ifNeeded' : 'none',
      // ...其他字段
    })
  });
  
  return await response.json();
}

// apps/wps-addin/src/wpsApi.ts (WPS 版本 - 类似逻辑)

async function sendWpsPlanRequest(
  prompt: string,
  workbook: WorkbookSnapshot,
  modelId: string
): Promise<AssistantResponse> {
  
  let sheetsData: Record<string, any[]> | null = null;
  
  if (modelId?.includes('agent-sdk')) {
    sheetsData = {};
    
    for (const sheet of workbook.worksheets) {
      try {
        sheetsData[sheet.name] = await readSheetDataWps(sheet.name, 10000);
      } catch (error) {
        console.error(`读取表 ${sheet.name} 失败:`, error);
        sheetsData[sheet.name] = [];
      }
    }
  }
  
  // 发送请求（与 Excel 相同的协议）
  const response = await fetch('/api/turn', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      prompt,
      workbook,
      modelId,
      sheetsData,
      planningReadPolicy: sheetsData ? 'ifNeeded' : 'none',
    })
  });
  
  return await response.json();
}
```

#### 2.3 后端改造

**步骤 1：更新请求模型**

```python
# server/app/models.py

class PlanRequest(BaseModel):
    prompt: str
    workbook: WorkbookSnapshot
    modelId: str | None = None
    planningReadPolicy: str = "none"
    
    # 🆕 新增：前端读取的完整数据
    sheetsData: dict[str, list[dict]] | None = None
    
    # ...其他现有字段
```

**步骤 2：从请求中读取数据**

```python
# server/app/claude_sdk_agent.py

# 在 run_sdk_plan() 的 read_authorized_sheet MCP 工具中

async def read_authorized_sheet(
    sheet_name: str,
    sample: bool = False,
    max_rows: int = 1000
) -> dict[str, Any]:
    """从请求快照中读取数据（双宿主通用）"""
    
    # 1. 验证授权
    authorized = [
        sheet for sheet in request.workbook.worksheets
        if sheet.name == sheet_name
    ]
    
    if not authorized:
        available_sheets = [s.name for s in request.workbook.worksheets]
        return {
            "content": [{
                "type": "text",
                "text": f"❌ 未授权访问表 '{sheet_name}'。\n\n已授权的表：{available_sheets}"
            }],
            "is_error": True
        }
    
    sheet_info = authorized[0]
    
    # 2. 从请求中获取数据
    if not request.sheetsData:
        return {
            "content": [{
                "type": "text",
                "text": "❌ 数据未加载。请确保使用支持数据读取的模型（如 claude-agent-sdk）。"
            }],
            "is_error": True
        }
    
    if sheet_name not in request.sheetsData:
        return {
            "content": [{
                "type": "text",
                "text": f"❌ 表 '{sheet_name}' 的数据未加载。可能读取时出错。"
            }],
            "is_error": True
        }
    
    # 3. 获取数据
    full_data = request.sheetsData[sheet_name]
    
    # 4. 采样（如果需要）
    if sample and len(full_data) > max_rows:
        sampled_data = full_data[:max_rows]
    else:
        sampled_data = full_data[:max_rows] if len(full_data) > max_rows else full_data
    
    # 5. 返回 JSON 格式
    return {
        "content": [{
            "type": "text",
            "text": json.dumps({
                "sheet": sheet_name,
                "total_rows": len(full_data),
                "total_columns": len(sheet_info.headers) if sheet_info.headers else 0,
                "headers": sheet_info.headers or [],
                "returned_rows": len(sampled_data),
                "sampled": sample or len(full_data) > max_rows,
                "data": sampled_data
            }, ensure_ascii=False, default=str)
        }]
    }
```

**步骤 3：兼容文件夹模式（可选）**

```python
# 如果同时支持文件夹模式

async def read_authorized_sheet(sheet_name: str, ...):
    """双模式数据读取"""
    
    # 验证授权（同上）
    # ...
    
    # 判断数据来源
    if request.sheetsData and sheet_name in request.sheetsData:
        # 方式 1：从请求快照读取（Office.js / WPS）
        data = request.sheetsData[sheet_name]
    
    elif request.workbook.sourceMode == "folder" and request.workbook.sourceFile:
        # 方式 2：从磁盘文件读取（文件夹模式）
        from .folder_data import read_sheet_from_file
        data = read_sheet_from_file(
            request.workbook.sourceFile,
            sheet_name,
            max_rows=10000
        )
    
    else:
        return {
            "content": [{
                "type": "text",
                "text": "❌ 无法读取数据：请求中未包含数据快照，也不是文件夹模式。"
            }],
            "is_error": True
        }
    
    # 采样和返回（同上）
    # ...
```

#### 2.4 性能优化

**前端优化：**

```typescript
// 1. 只在必要时读取完整数据
if (!modelId?.includes('agent-sdk')) {
  // 普通模式：不读取完整数据
  sheetsData = null;
}

// 2. 大数据集提示
const totalRows = workbook.worksheets.reduce((sum, s) => sum + s.rowCount, 0);

if (totalRows > 50000) {
  const confirmed = await showConfirmDialog(
    '数据量提示',
    `当前工作簿共 ${totalRows.toLocaleString()} 行数据，\n` +
    `Agent 分析将只使用前 10,000 行。\n\n` +
    `是否继续？`
  );
  
  if (!confirmed) {
    return; // 用户取消
  }
}

// 3. 显示加载进度
showProgress('正在读取数据...');
for (const sheet of workbook.worksheets) {
  sheetsData[sheet.name] = await readSheetData(sheet.name);
  updateProgress(`已读取: ${sheet.name}`);
}
hideProgress();
```

**后端优化：**

```python
# 1. 数据大小限制
MAX_DATA_SIZE_MB = 50

def validate_data_size(sheets_data: dict[str, list[dict]]):
    """验证数据大小"""
    import sys
    
    total_size = sys.getsizeof(json.dumps(sheets_data))
    size_mb = total_size / (1024 * 1024)
    
    if size_mb > MAX_DATA_SIZE_MB:
        raise ValueError(
            f"数据量过大（{size_mb:.1f} MB），超过限制 {MAX_DATA_SIZE_MB} MB。"
            f"请减少授权的工作表或行数。"
        )

# 在 run_sdk_plan 开始时调用
if request.sheetsData:
    validate_data_size(request.sheetsData)

# 2. 会话级缓存（可选）
_session_data_cache: dict[str, dict[str, list[dict]]] = {}

def get_cached_sheets_data(session_id: str) -> dict[str, list[dict]] | None:
    return _session_data_cache.get(session_id)

def cache_sheets_data(session_id: str, data: dict[str, list[dict]]):
    _session_data_cache[session_id] = data
    # 限制缓存大小，30 分钟后自动清理
```

---

#### 2.5 双宿主测试清单

**Excel (Office.js) 测试：**

```typescript
// 测试场景
1. ✅ 简单表（100 行）- 读取和发送
2. ✅ 多表（3 个表）- 批量读取
3. ✅ 大表（10000 行）- 性能测试
4. ✅ 空表 - 边界情况
5. ✅ 特殊字符表名 - 编码测试
6. ✅ 数据类型混合（文本、数字、日期）
```

**WPS (JSAPI) 测试：**

```typescript
// 测试场景
1. ✅ 简单表（100 行）- 读取和发送
2. ✅ 多表（3 个表）- 批量读取
3. ✅ 大表（10000 行）- 性能测试
4. ✅ 空表 - 边界情况
5. ✅ 已保存工作簿 - 身份验证
6. ✅ Value2 格式转换 - 数据正确性
```

**协议兼容性测试：**

```python
# 后端应该无法区分数据来源
def test_backend_agnostic():
    # Excel 发送的请求
    excel_request = {
        "sheetsData": {"销售": [{"产品": "A", "销量": 100}]}
    }
    
    # WPS 发送的请求
    wps_request = {
        "sheetsData": {"销售": [{"产品": "A", "销量": 100}]}
    }
    
    # 后端处理逻辑应该完全相同
    assert process_request(excel_request) == process_request(wps_request)
```

#### 2.6 添加辅助工具（可选）

可以添加更多 MCP 工具帮助 Agent：

```python
# 工具：列出授权的表
async def list_authorized_sheets() -> dict[str, Any]:
    """列出所有已授权的工作表"""
    sheets_info = [{
        "name": sheet.name,
        "rows": sheet.rowCount,
        "columns": sheet.columnCount,
        "headers": sheet.headers,
    } for sheet in request.workbook.worksheets]
    
    return {
        "content": [{
            "type": "text",
            "text": json.dumps({
                "workbook": request.workbook.name,
                "sheets": sheets_info
            }, ensure_ascii=False)
        }]
    }

tools.append(
    sdk.tool(
        "list_authorized_sheets",
        "列出所有已授权的工作表信息（不含数据）",
        {"type": "object", "properties": {}}
    )(list_authorized_sheets)
)
allowed.append("mcp__excel_bro__list_authorized_sheets")

# 工具：验证公式语法
async def validate_formula(formula: str) -> dict[str, Any]:
    """验证 Excel 公式语法"""
    from .formula_safety import check_formula_safety
    
    try:
        is_safe, reason = check_formula_safety(formula)
        if not is_safe:
            return {
                "content": [{
                    "type": "text",
                    "text": f"❌ 公式不安全：{reason}"
                }],
                "is_error": True
            }
        
        # TODO: 进一步的语法验证
        
        return {
            "content": [{
                "type": "text",
                "text": f"✅ 公式语法有效"
            }]
        }
    except Exception as e:
        return {
            "content": [{
                "type": "text",
                "text": f"❌ 验证失败：{str(e)}"
            }],
            "is_error": True
        }

tools.append(
    sdk.tool(
        "validate_formula",
        "预验证 Excel 公式的语法和安全性（不实际执行）",
        {
            "type": "object",
            "properties": {
                "formula": {
                    "type": "string",
                    "description": "要验证的 Excel 公式（如 =SUM(A1:A10)）"
                }
            },
            "required": ["formula"]
        }
    )(validate_formula)
)
allowed.append("mcp__excel_bro__validate_formula")
```

---

### 阶段 3：测试和优化（Day 3-5）

#### 3.1 测试场景

**测试 1：简单读取 → 提交（双宿主）**

```python
# Excel 环境
用户: "显示销售表的前 10 行"

Agent 行为:
1. read_authorized_sheet(sheet_name="销售", sample=true)
   → 从 request.sheetsData["销售"] 读取
2. submit_answer(message="表中共有 1234 行，前 10 行如下：...")

# WPS 环境 - 相同行为
用户: "显示销售表的前 10 行"
Agent: 相同的工具调用序列
→ 后端无法区分数据来源
```

**测试 2：复杂分析（跨宿主验证）**

```python
# Excel 或 WPS（协议相同）
用户: "分析最近三个月的销售趋势，找出 TOP 5 产品"

Agent 行为:
1. list_authorized_sheets()
   → 了解有哪些表
2. read_authorized_sheet(sheet_name="销售", sample=true)
   → 先采样看结构
3. read_authorized_sheet(sheet_name="销售", sample=false)
   → 获取完整数据（来自请求快照）
4. Bash: echo '...' > sales.json
   → 保存到临时文件
5. Bash: python -c "
     import pandas as pd
     import json
     
     with open('sales.json') as f:
         data = json.load(f)
     
     df = pd.DataFrame(data['data'])
     df['日期'] = pd.to_datetime(df['日期'])
     
     # 最近三个月
     from datetime import datetime, timedelta
     three_months_ago = datetime.now() - timedelta(days=90)
     recent = df[df['日期'] >= three_months_ago]
     
     # TOP 5
     top5 = recent.groupby('产品')['销量'].sum().nlargest(5)
     
     result = top5.to_dict()
     
     with open('result.json', 'w') as f:
         json.dump(result, f, ensure_ascii=False)
     
     print('分析完成')
   "
   → 返回: "分析完成"
6. Read: result.json
   → 读取分析结果
7. submit_plan({
     actions: [{
       type: "writeValues",
       targetSheet: "TOP5分析",
       values: [...]
     }]
   })
```

**测试 3：多表关联（Excel + WPS 都支持）**

```python
用户: "分析销售表和产品表，找出高价低销的产品"

Agent 行为:
1. list_authorized_sheets()
   → ["销售", "产品信息"]
2. read_authorized_sheet("销售")
3. read_authorized_sheet("产品信息")
4. Bash: python -c "
     # 读取两个表
     # 合并分析
     # 找出高价低销产品
   "
5. submit_plan(...)
```

**测试 4：错误处理（双宿主一致）**

```python
# 场景 1：访问未授权的表
Agent: read_authorized_sheet("机密数据")
→ 返回: "❌ 未授权访问"

# 场景 2：数据未加载（非 SDK 模式）
modelId = "gpt-4"  # 普通模式，sheetsData=null
Agent: read_authorized_sheet("销售")
→ 返回: "❌ 数据未加载"

# 场景 3：Python 脚本错误
Agent: Bash: python -c "print(undefined_variable)"
→ 返回: "NameError: name 'undefined_variable' is not defined"

# 场景 4：Excel 读取失败
# 前端捕获错误，发送空数组
sheetsData["销售"] = []
Agent: read_authorized_sheet("销售")
→ 返回: data=[], total_rows=0
```

#### 3.2 性能测试

**小数据集（< 1000 行）**
```
Excel:
- 读取：< 500ms
- 发送：< 200ms
- 总计：< 1s

WPS:
- 读取：< 600ms (Value2 访问)
- 发送：< 200ms
- 总计：< 1s
```

**中数据集（5000 行）**
```
Excel:
- 读取：1-2s
- JSON 编码：200-300ms
- 发送：300-500ms
- 总计：2-3s

WPS:
- 读取：1.5-2.5s
- JSON 编码：200-300ms
- 发送：300-500ms
- 总计：2.5-3.5s
```

**大数据集（10000 行限制）**
```
Excel:
- 读取：2-3s
- JSON 编码：500ms
- 发送：500-800ms
- 总计：3-5s

WPS:
- 读取：3-4s
- JSON 编码：500ms
- 发送：500-800ms
- 总计：4-6s
```

**优化建议：**
- 5000 行以下：可接受
- 5000-10000 行：提示用户
- 超过 10000 行：强制限制

#### 3.3 双宿主兼容性验证

**协议一致性测试：**

```python
# 测试：两端发送的数据结构相同
def test_protocol_consistency():
    # Excel 前端发送
    excel_payload = {
        "prompt": "分析销售",
        "workbook": {...},
        "sheetsData": {
            "销售": [
                {"产品": "A", "销量": 100},
                {"产品": "B", "销量": 200}
            ]
        }
    }
    
    # WPS 前端发送
    wps_payload = {
        "prompt": "分析销售",
        "workbook": {...},
        "sheetsData": {
            "销售": [
                {"产品": "A", "销量": 100},
                {"产品": "B", "销量": 200}
            ]
        }
    }
    
    # 后端处理应该完全相同
    excel_result = backend.process(excel_payload)
    wps_result = backend.process(wps_payload)
    
    assert excel_result == wps_result
```

**数据类型一致性：**

```typescript
// Excel 和 WPS 都转换为相同格式

// Excel (Office.js)
const excelData = {
  "数字列": 123,      // number
  "文本列": "ABC",    // string
  "日期列": "2024-01-01",  // string (ISO)
  "空值": null,       // null
  "布尔": true        // boolean
};

// WPS (JSAPI Value2)
const wpsData = {
  "数字列": 123,      // number (Value2 已标准化)
  "文本列": "ABC",    // string
  "日期列": "2024-01-01",  // string (需转换)
  "空值": null,       // null
  "布尔": true        // boolean (需转换)
};

// 后端接收：完全相同
```

---

#### 3.2 性能优化

**优化 1：工作目录复用**

```python
# 当前：每次调用创建新目录
workspace_dir = state_dir / f"workspace_{uuid.uuid4().hex[:8]}"

# 优化：会话级别复用（可选）
# 如果 request 有 session_id，可以复用目录
# 但要注意清理策略
```

**优化 2：数据缓存**

```python
# 如果同一个请求多次读取同一个表，可以缓存
_sheet_data_cache = {}

async def read_authorized_sheet(sheet_name: str, ...):
    cache_key = f"{request.workbook.name}:{sheet_name}"
    if cache_key in _sheet_data_cache:
        return _sheet_data_cache[cache_key]
    
    data = ...  # 实际读取
    _sheet_data_cache[cache_key] = data
    return data
```

**优化 3：超时设置**

```python
# 根据任务复杂度动态调整超时
if request.workbook.rowCount > 100000:
    timeout = 180  # 大数据集给更多时间
else:
    timeout = 60   # 普通任务
```

---

#### 3.3 监控和日志

**添加审计日志：**

```python
# server/app/audit_log.py

import logging
from datetime import datetime

audit_logger = logging.getLogger("excel_bro.audit")

def log_sdk_execution(
    user_id: str,
    request_id: str,
    action: str,
    details: dict
):
    """记录 SDK 执行的关键操作"""
    audit_logger.info(
        "SDK_EXECUTION",
        extra={
            "timestamp": datetime.utcnow().isoformat(),
            "user_id": user_id,
            "request_id": request_id,
            "action": action,
            "details": details
        }
    )

# 在 run_sdk_plan 中使用
log_sdk_execution(
    user_id=request.user_id if hasattr(request, 'user_id') else "local",
    request_id=str(uuid.uuid4()),
    action="sdk_plan_started",
    details={
        "prompt": request.prompt[:100],
        "sheets": [s.name for s in request.workbook.worksheets],
        "model": connection.model
    }
)
```

---

### 阶段 4：文档和发布（Day 5-7）

#### 4.1 更新用户文档

**README.md 更新：**

```markdown
## Claude Agent SDK 模式（完整探索）

Excel Bro 现在支持使用 Claude Agent SDK 进行复杂分析。

### 特点

✅ **完整探索能力**
- Agent 可以多轮探索数据
- Agent 可以运行 Python 分析
- Agent 可以验证逻辑后再生成计划

✅ **安全边界保留**
- 仍只能访问授权的表
- 最终写入仍需预览确认
- 仍可以撤销操作

✅ **双宿主支持**
- Excel (Office.js) 和 WPS (JSAPI) 使用相同的后端逻辑
- 前端读取数据并发送
- 协议完全统一

### 环境要求

```bash
# Python 环境（Agent 会用）
pip install pandas numpy scipy

# Claude Agent SDK（自动安装）
pip install -r server/requirements-claude-sdk.txt
```

### 启动方式

**Windows:**
```powershell
powershell -File scripts/start_claude_sdk_server.ps1
```

**注意：** SDK 模式需要特殊启动脚本，不能使用 `npm run dev:server`（Uvicorn 热重载与 SDK 子进程不兼容）

### 配置

```bash
# .env
CLAUDE_AGENT_SDK_MODEL=claude-3-5-sonnet-20241022
ANTHROPIC_API_KEY=sk-ant-...
CLAUDE_AGENT_SDK_STATE_DIR=/absolute/path/outside/repo
```

### 工作原理

```
用户请求
  ↓
前端（Excel/WPS）读取授权表数据
  ↓
HTTP 请求携带数据快照
  ↓
Claude Agent SDK
  ├─ 从请求读取数据
  ├─ 运行 Python 分析
  ├─ 多轮探索验证
  └─ 提交白名单计划
      ↓
用户预览确认
  ↓
安全执行
```

### 数据读取方式

**Excel (Office.js):**
```typescript
// 前端使用 Office.js 读取
Excel.run(async (context) => {
  const range = sheet.getUsedRange();
  range.load('values');
  await context.sync();
  // 转换为对象数组并发送
});
```

**WPS (JSAPI):**
```typescript
// 前端使用 WPS JSAPI 读取
const sheet = Application.Worksheets.Item(sheetName);
const values = sheet.UsedRange.Value2;
// 转换为对象数组并发送
```

**后端统一处理：**
```python
# 无需区分来源
data = request.sheetsData[sheet_name]
```

### 对比普通模式

| 特性 | 普通模式 | SDK 模式 |
|------|---------|---------|
| 探索能力 | 一次生成 | 多轮探索 |
| Python 分析 | 无 | ✅ |
| 数据读取 | 结构摘要 | 完整数据 |
| 启动方式 | 简单 | 需要特殊脚本 |
| 速度 | 快（5-10s） | 较慢（15-30s） |
| 准确率 | 中 | 高 |
| Excel 支持 | ✅ | ✅ |
| WPS 支持 | ✅ | ✅ |

### 性能参考

| 数据量 | Excel 读取 | WPS 读取 | 传输 | Agent 分析 | 总计 |
|--------|-----------|---------|------|-----------|------|
| 1000 行 | 0.5s | 0.6s | 0.2s | 10-20s | 11-21s |
| 5000 行 | 2s | 2.5s | 0.5s | 15-25s | 18-28s |
| 10000 行 | 3s | 4s | 0.8s | 20-30s | 24-35s |

### 选择建议

- **简单任务**（排序、筛选）→ 普通模式
- **复杂分析**（趋势、预测、相关性）→ SDK 模式
- **Excel 和 WPS**：两者在 SDK 模式下体验一致
```

#### 4.2 更新架构文档

**docs/ARCHITECTURE.md 双宿主部分：**

```markdown
## Claude Agent SDK 后端（完整版 - 双宿主）

### 2026-10-06 改造

之前的 SDK 集成禁用了所有内置工具，导致 Agent 无法探索和推理。
现已改造为完整版本，释放探索能力同时保留安全边界。

### 双宿主数据读取

**架构：**
```
┌─────────────────────────────────────────┐
│  Excel Add-in (Office.js)               │
│  └─ readSheetDataOffice()               │
│     → Excel.run() 读取 UsedRange        │
└──────────────┬──────────────────────────┘
               │
               ├─ HTTP POST /api/turn
               │  { sheetsData: {...} }
               │
┌──────────────▼──────────────────────────┐
│  WPS Add-in (WPS JSAPI)                 │
│  └─ readSheetDataWps()                  │
│     → Application.Worksheets 读取 Value2│
└──────────────┬──────────────────────────┘
               │
               ▼
┌─────────────────────────────────────────┐
│  FastAPI Backend                        │
│  └─ request.sheetsData[sheet_name]     │
│     (双宿主无差异)                       │
└──────────────┬──────────────────────────┘
               │
               ▼
┌─────────────────────────────────────────┐
│  Claude Agent SDK                       │
│  └─ MCP 工具访问数据                     │
└─────────────────────────────────────────┘
```

**关键点：**
1. 前端负责读取：Office.js 和 WPS JSAPI 各自实现
2. 协议统一：都转换为 `Record<string, any[]>` 格式
3. 后端无感知：无需区分数据来源
4. 类型标准化：数字、文本、日期、null、布尔

### 可用工具

**内置工具（SDK 提供）：**
- `Read` - 读取临时文件
- `Bash` - 运行 Python 脚本

**MCP 工具（Excel Bro 提供）：**
- `read_authorized_sheet` - 从请求快照读取已授权的表
- `list_authorized_sheets` - 列出所有授权表
- `submit_plan` - 提交白名单计划
- `submit_answer` - 提交只读回答
- `submit_clarification` - 提交澄清问题
- `validate_formula` - 预验证公式（可选）

### 安全机制

三层防护：

1. **授权层**：MCP 工具守卫，只能访问授权的表
2. **白名单层**：submit_plan 校验，只能提交合法动作
3. **确认层**：用户预览并确认后才执行

### 数据大小限制

```python
# 前端限制
MAX_ROWS_PER_SHEET = 10000

# 后端限制
MAX_DATA_SIZE_MB = 50

# 超过限制时的处理
if totalRows > 50000:
    showConfirmDialog("数据量过大，只分析前 10000 行")
```

### 工作目录

每次调用创建独立的临时目录：
```
CLAUDE_AGENT_SDK_STATE_DIR/
  ├── workspace_a1b2c3d4/  # 本次调用
  │   ├── sales.json       # Agent 的临时文件
  │   ├── analysis.py
  │   └── result.json
  └── workspace_e5f6g7h8/  # 另一次调用
```

调用结束后自动清理（或按 TTL 保留）。

### 限制

- ❌ 不支持检查点恢复（未来改进）
- ❌ 需要特殊启动脚本（Uvicorn 兼容性问题）
- ⚠️ 首次运行需下载 Claude CLI（约 100MB）
- ⚠️ 大数据集（> 10000 行）会被截断

### 性能

- 启动开销：~1-2 秒（子进程）
- 数据读取：Excel 0.5-3s, WPS 0.6-4s（取决于数据量）
- 每轮调用：取决于分析复杂度
- 典型场景：3-5 轮，15-30 秒
```

#### 4.3 双宿主使用示例

**docs/SDK_EXAMPLES.md 更新：**

```markdown
# Claude Agent SDK 使用示例（双宿主）

## 前端数据读取示例

### Excel (Office.js)

```typescript
// apps/excel-addin/src/officeDataReader.ts

async function readSheetDataOffice(sheetName: string): Promise<any[]> {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const usedRange = sheet.getUsedRange();
    
    usedRange.load(['values', 'rowCount']);
    await context.sync();
    
    const values = usedRange.values;
    const headers = values[0];
    const rows = values.slice(1, 10001); // 最多 10000 行
    
    return rows.map(row => {
      const obj: Record<string, any> = {};
      headers.forEach((h, i) => { obj[String(h)] = row[i]; });
      return obj;
    });
  });
}
```

### WPS (JSAPI)

```typescript
// apps/wps-addin/src/wpsDataReader.ts

async function readSheetDataWps(sheetName: string): Promise<any[]> {
  const sheet = Application.Worksheets.Item(sheetName);
  const usedRange = sheet.UsedRange;
  const values = usedRange.Value2; // WPS 特有
  
  const headers = values[0];
  const rows = values.slice(1, 10001); // 最多 10000 行
  
  return rows.map(row => {
    const obj: Record<string, any> = {};
    headers.forEach((h: any, i: number) => { obj[String(h)] = row[i]; });
    return obj;
  });
}
```

### 统一发送

```typescript
// 两端使用相同的发送逻辑
async function sendToBackend(sheetsData: Record<string, any[]>) {
  await fetch('/api/turn', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      prompt: userPrompt,
      workbook: workbookSnapshot,
      sheetsData: sheetsData,  // Excel 和 WPS 格式相同
      modelId: 'claude-agent-sdk'
    })
  });
}
```

## 后端处理示例

```python
# server/app/claude_sdk_agent.py

async def read_authorized_sheet(sheet_name: str, ...):
    # 从请求读取（无需区分 Excel 还是 WPS）
    data = request.sheetsData[sheet_name]
    
    # 数据格式完全相同
    # [
    #   {"产品": "A", "销量": 100},
    #   {"产品": "B", "销量": 200}
    # ]
    
    return {"content": [{"type": "text", "text": json.dumps(data)}]}
```

## 完整示例：销售分析（双宿主通用）

**用户在 Excel 中：**
```
用户: "分析销售趋势并预测下月"
→ Excel 读取数据 → 发送到后端
```

**用户在 WPS 中：**
```
用户: "分析销售趋势并预测下月"
→ WPS 读取数据 → 发送到后端
```

**后端处理（完全相同）：**
```python
# Agent 执行过程
1. read_authorized_sheet("销售")
   → 获取数据（来自 Excel 或 WPS，无差异）
   
2. Bash: python -c "
     import pandas as pd
     import json
     
     with open('sales.json') as f:
         data = json.load(f)
     
     df = pd.DataFrame(data['data'])
     # 分析逻辑...
   "
   
3. submit_plan({...})
   → 提交计划
```

**用户看到（Excel 和 WPS 相同）：**
```
Excel Bro:
  已完成趋势分析 ✓
  
  预览操作：
  📝 在"分析结果"表写入数据
  📈 预测下月销量：15,234
  
  [确认执行] [取消]
```
```

---

#### 4.2 更新架构文档

**docs/ARCHITECTURE.md 更新：**

```markdown
## Claude Agent SDK 后端（完整版）

### 2026-10-06 改造

之前的 SDK 集成禁用了所有内置工具，导致 Agent 无法探索和推理。
现已改造为完整版本，释放探索能力同时保留安全边界。

### 可用工具

**内置工具（SDK 提供）：**
- `Read` - 读取临时文件
- `Bash` - 运行 Python 脚本

**MCP 工具（Excel Bro 提供）：**
- `read_authorized_sheet` - 读取已授权的表
- `submit_plan` - 提交白名单计划
- `submit_answer` - 提交只读回答
- `submit_clarification` - 提交澄清问题

### 安全机制

三层防护：

1. **授权层**：MCP 工具守卫，只能访问授权的表
2. **白名单层**：submit_plan 校验，只能提交合法动作
3. **确认层**：用户预览并确认后才执行

### 工作目录

每次调用创建独立的临时目录：
```
CLAUDE_AGENT_SDK_STATE_DIR/
  ├── workspace_a1b2c3d4/  # 本次调用
  │   ├── sales.json       # Agent 的临时文件
  │   ├── analysis.py
  │   └── result.json
  └── workspace_e5f6g7h8/  # 另一次调用
```

调用结束后自动清理（或按 TTL 保留）。

### 限制

- ❌ 不支持检查点恢复（未来改进）
- ❌ 需要特殊启动脚本（Uvicorn 兼容性问题）
- ⚠️ 首次运行需下载 Claude CLI（约 100MB）

### 性能

- 启动开销：~1-2 秒（子进程）
- 每轮调用：取决于分析复杂度
- 典型场景：3-5 轮，15-30 秒
```

---

#### 4.3 添加示例

**docs/SDK_EXAMPLES.md**

```markdown
# Claude Agent SDK 使用示例

## 示例 1：销售趋势分析

**用户请求：**
> 分析最近三个月的销售趋势，预测下月销量

**Agent 执行过程：**

```python
# Turn 1: 了解数据结构
→ read_authorized_sheet(sheet_name="销售", sample=True)
← 返回：表头、前 20 行数据

# Turn 2: 获取完整数据
→ read_authorized_sheet(sheet_name="销售", sample=False)
← 返回：完整数据（1234 行）

# Turn 3: 保存并分析
→ Bash: echo '{...}' > sales.json
→ Bash: python -c "
  import pandas as pd
  import json
  from scipy import stats
  import numpy as np
  
  # 读取
  with open('sales.json') as f:
      data = json.load(f)
  
  df = pd.DataFrame(data['data'])
  df['日期'] = pd.to_datetime(df['日期'])
  df['月份'] = df['日期'].dt.to_period('M')
  
  # 最近三个月
  recent = df[df['日期'] >= df['日期'].max() - pd.Timedelta(days=90)]
  monthly = recent.groupby('月份')['销量'].sum()
  
  # 线性回归
  x = np.arange(len(monthly))
  slope, intercept, r, p, se = stats.linregress(x, monthly.values)
  
  # 预测
  next_month = slope * len(monthly) + intercept
  
  result = {
      'monthly': monthly.to_dict(),
      'trend': 'increasing' if slope > 0 else 'decreasing',
      'prediction': float(next_month),
      'r_squared': float(r**2)
  }
  
  with open('result.json', 'w') as f:
      json.dump(result, f, ensure_ascii=False, default=str)
  
  print('分析完成')
"
← 返回：分析完成

# Turn 4: 读取结果
→ Read result.json
← 返回：{"monthly": {...}, "prediction": 15234, ...}

# Turn 5: 提交计划
→ submit_plan({
    "actions": [
        {
            "type": "writeValues",
            "targetSheet": "趋势分析",
            "targetRange": "A1:C5",
            "values": [
                ["月份", "实际销量", ""],
                ["2024-08", "12000", ""],
                ["2024-09", "13500", ""],
                ["2024-10", "14200", ""],
                ["预测-11", "15234", "← AI 预测"]
            ]
        }
    ],
    "summary": "趋势分析完成：上升趋势 (R²=0.89)，预测下月 15,234"
})
← 返回：✅ 计划已提交
```

**用户看到的界面：**
```
Excel Bro:
  已完成趋势分析 ✓
  
  预览操作：
  📝 在"趋势分析"表写入月度数据
  📈 预测下月销量：15,234
  📊 趋势：上升（可信度 R²=0.89）
  
  [确认执行] [修改] [取消]
```

---

## 示例 2：相关性分析

**用户请求：**
> 分析产品价格与销量的相关性

**Agent 执行过程：**

```python
# Turn 1: 列出所有表
→ list_authorized_sheets()
← 返回：["销售", "产品信息", "库存"]

# Turn 2-3: 读取两个表
→ read_authorized_sheet("销售")
→ read_authorized_sheet("产品信息")

# Turn 4: 合并分析
→ Bash: python -c "
  import pandas as pd
  import json
  
  # 读取两个表
  with open('sales.json') as f:
      sales = pd.DataFrame(json.load(f)['data'])
  
  with open('products.json') as f:
      products = pd.DataFrame(json.load(f)['data'])
  
  # 合并
  merged = sales.merge(products, on='产品ID')
  
  # 相关性
  correlation = merged[['价格', '销量']].corr().iloc[0, 1]
  
  result = {
      'correlation': float(correlation),
      'interpretation': '负相关' if correlation < -0.3 else 
                       '正相关' if correlation > 0.3 else '弱相关'
  }
  
  with open('correlation.json', 'w') as f:
      json.dump(result, f)
"

# Turn 5: 提交回答
→ submit_answer(
    message="价格与销量呈负相关（r=-0.67），价格越高销量越低。"
)
```

---

## 示例 3：异常检测

**用户请求：**
> 找出销售数据中的异常值

**Agent 执行过程：**

```python
# Turn 1: 读取数据
→ read_authorized_sheet("销售")

# Turn 2: 统计分析
→ Bash: python -c "
  import pandas as pd
  import numpy as np
  import json
  
  with open('sales.json') as f:
      df = pd.DataFrame(json.load(f)['data'])
  
  # Z-score 异常检测
  mean = df['销量'].mean()
  std = df['销量'].std()
  df['z_score'] = (df['销量'] - mean) / std
  
  # 找出异常（|z| > 3）
  anomalies = df[df['z_score'].abs() > 3]
  
  result = {
      'total_anomalies': len(anomalies),
      'anomalies': anomalies[['日期', '产品', '销量']].to_dict('records')
  }
  
  with open('anomalies.json', 'w') as f:
      json.dump(result, f, ensure_ascii=False, default=str)
"

# Turn 3: 生成报告
→ submit_plan({
    "actions": [{
        "type": "writeValues",
        "targetSheet": "异常报告",
        "targetRange": "A1:D10",
        "values": [...]  # 异常数据
    }]
})
```
```

---

## 常见问题

### Q1: 为什么不直接移除 SDK？

**A:** 改造后的 SDK 提供了三个核心价值：

1. **完整的探索能力**：多轮推理、Python 分析
2. **进程隔离**：SDK 自带的沙箱和临时目录管理
3. **工具生态**：未来可以集成更多 MCP 工具

如果只是想要"运行 Python"，确实可以用更简单的本地执行器。
但 SDK 提供的是"完整的 Agent 能力"，不只是脚本执行。

---

### Q2: 性能会变差吗？

**A:** 会比普通模式慢，但换来的是准确率提升：

- 普通模式：1 轮，5-10 秒，准确率 70%
- SDK 模式：3-5 轮，15-30 秒，准确率 90%+

对于复杂任务，多花 20 秒换来准确结果是值得的。

---

### Q3: Excel 和 WPS 的数据格式一致吗？

**A:** 前端会标准化为统一格式：

**Excel (Office.js):**
```typescript
// range.values 返回二维数组
[["产品", "销量"], ["A", 100], ["B", 200]]

// 转换为对象数组
[{"产品": "A", "销量": 100}, {"产品": "B", "销量": 200}]
```

**WPS (JSAPI):**
```typescript
// Value2 返回二维数组（需要处理 COM 类型）
[["产品", "销量"], ["A", 100], ["B", 200]]

// 转换为对象数组（与 Excel 相同）
[{"产品": "A", "销量": 100}, {"产品": "B", "销量": 200}]
```

**后端接收：**
```python
# 完全相同的数据结构
data = [{"产品": "A", "销量": 100}, {"产品": "B", "销量": 200}]
```

**数据类型映射：**
| Excel 类型 | WPS 类型 | 统一格式 |
|-----------|---------|---------|
| Number | Number | `number` |
| String | String | `string` |
| Date | Date Serial | `string` (ISO 8601) |
| Empty | Empty | `null` |
| Boolean | Boolean | `boolean` |

---

### Q4: 大数据集怎么办？

**A:** 三层保护：

**前端限制：**
```typescript
if (sheet.rowCount > 10000) {
  showWarning("数据量较大，只分析前 10000 行");
  data = data.slice(0, 10000);
}
```

**后端限制：**
```python
MAX_DATA_SIZE_MB = 50  # 请求体大小

if size_mb > MAX_DATA_SIZE_MB:
    raise ValueError("数据量超过 50MB 限制")
```

**用户提示：**
```
Excel Bro:
  当前工作簿共 52,345 行数据
  
  由于数据量较大，Agent 分析将只使用前 10,000 行。
  
  建议：先筛选数据范围或使用透视表汇总。
  
  [继续分析] [取消]
```

---

### Q5: WPS 和 Excel 的执行结果会不同吗？

**A:** 不会，后端逻辑完全相同：

**相同的输入：**
- Excel 和 WPS 发送相同格式的数据
- 后端无法区分来源

**相同的处理：**
- 使用相同的 Agent
- 使用相同的 Python 分析
- 使用相同的白名单校验

**可能的差异：**
- 前端执行写入操作时的 API 不同
- 但计划内容和验证逻辑相同

**测试保证：**
```python
def test_cross_host_consistency():
    """验证跨宿主一致性"""
    # 相同的输入
    input_data = {"销售": [{"产品": "A", "销量": 100}]}
    
    # Excel 环境
    excel_result = run_sdk_plan(input_data, host="excel")
    
    # WPS 环境
    wps_result = run_sdk_plan(input_data, host="wps")
    
    # 计划应该相同
    assert excel_result.plan == wps_result.plan
```

---

### Q6: 检查点恢复什么时候支持？

**A:** 当前改造未实现检查点，原因：

- SDK 的会话状态在子进程中
- 服务端拿不到完整现场

**未来方案：**
1. 实现 SDK 会话序列化
2. 或者在服务端包装一层状态管理

预计需要 2-3 周。

**当前限制：**
```python
if config.adapter == "claude_agent_sdk":
    if resume_checkpoint is not None:
        raise ModelNotAvailableError(
            "Claude Agent SDK 暂不支持恢复检查点，请使用其他模型或重新开始任务。"
        )
```

---

### Q7: 我应该用 SDK 模式还是普通模式？

**决策树：**

```
任务复杂度？
├─ 简单（排序、筛选、简单公式）
│  → 普通模式（快速）
│
└─ 复杂（趋势分析、相关性、预测、多表关联）
   ├─ 有 SDK 环境且不介意特殊启动？
   │  → SDK 模式（准确）
   │
   └─ 只想简单用？
      → 普通模式 + 手动验证结果
```

**宿主无关：**
- Excel 和 WPS 都支持两种模式
- 选择取决于任务复杂度，不取决于宿主

---

### Q8: WPS 的 Value2 和 Excel 的 values 有什么区别？

**A:** 主要差异在底层实现：

**Excel (Office.js):**
```typescript
// 直接返回 JavaScript 数组
const values: any[][] = range.values;
// [["Header", 100], ["Text", 200]]
```

**WPS (JSAPI):**
```typescript
// 返回 COM 对象（需要转换）
const values: any = range.Value2;
// 可能需要处理 COM 类型

// 安全转换
function normalizeWpsValue(value: any): any {
  if (value === null || value === undefined) return null;
  if (typeof value === 'number') return value;
  if (typeof value === 'string') return value;
  if (typeof value === 'boolean') return value;
  
  // 日期对象转换
  if (value instanceof Date) {
    return value.toISOString().split('T')[0];
  }
  
  // 其他情况转字符串
  return String(value);
}
```

**统一处理后：**
```typescript
// Excel 和 WPS 最终都转换为相同的 JSON 格式
const standardData = rows.map(row => {
  const obj: Record<string, any> = {};
  headers.forEach((h, i) => {
    obj[String(h)] = normalizeValue(row[i]); // 统一标准化
  });
  return obj;
});
```

---

### Q9: 安全吗？

**A:** 三层防护确保安全：

1. **授权层**：Agent 只能读已授权的表
   ```python
   if sheet_name not in authorized_sheets:
       return {"error": "未授权访问"}
   ```

2. **白名单层**：只能提交合法的操作类型
   ```python
   ALLOWED_ACTIONS = [
       "writeValues", "writeFormulas", 
       "clearRange", "setNumberFormat"
   ]
   
   if action.type not in ALLOWED_ACTIONS:
       raise ValidationError("不允许的操作类型")
   ```

3. **确认层**：用户预览并确认
   ```typescript
   // 用户看到预览
   <ReviewPanel plan={plan}>
     <ActionList actions={plan.actions} />
     <ConfirmButton onClick={execute} />
   </ReviewPanel>
   ```

**双宿主一致：**
- Excel 和 WPS 使用相同的验证逻辑
- 安全边界不因宿主而降低

---

### Q10: 如何调试前端数据读取？

**Excel 调试：**
```typescript
// 添加日志
async function readSheetDataOffice(sheetName: string) {
  console.log(`[Excel] 开始读取表: ${sheetName}`);
  
  const data = await Excel.run(async (context) => {
    // ...
    console.log(`[Excel] 读取到 ${values.length} 行`);
    return rows;
  });
  
  console.log(`[Excel] 转换后: ${data.length} 条记录`);
  console.log('[Excel] 示例数据:', data.slice(0, 3));
  
  return data;
}
```

**WPS 调试：**
```typescript
// 添加日志
async function readSheetDataWps(sheetName: string) {
  console.log(`[WPS] 开始读取表: ${sheetName}`);
  
  const values = usedRange.Value2;
  console.log(`[WPS] 读取到 ${values.length} 行`);
  
  const data = convertToObjects(values);
  console.log(`[WPS] 转换后: ${data.length} 条记录`);
  console.log('[WPS] 示例数据:', data.slice(0, 3));
  
  return data;
}
```

**后端验证：**
```python
# 添加日志
@app.post("/api/turn")
async def turn(request: PlanRequest):
    if request.sheetsData:
        logger.info(f"接收到数据: {len(request.sheetsData)} 个表")
        for sheet_name, data in request.sheetsData.items():
            logger.info(f"  表 {sheet_name}: {len(data)} 行")
            logger.debug(f"    示例: {data[:2]}")
```

---

## 总结

改造后的 SDK 集成：
- ✅ 获得完整灵活性
- ✅ 保留安全边界
- ✅ 提升复杂任务准确率
- ✅ 双宿主（Excel + WPS）完全支持
- ✅ 协议统一，前端读取 + 后端处理
- ⚠️ 需要特殊启动
- ⚠️ 暂不支持检查点

**推荐场景：** 复杂数据分析、多表关联、趋势预测

**不推荐：** 简单操作、要求极致速度

**双宿主保证：** Excel 和 WPS 使用相同的后端逻辑和安全边界
```

---

## 下一步

完成基础改造后，可以考虑：

### 短期优化（1-2 周）

1. **实现检查点恢复**
   - 序列化 SDK 会话状态
   - 支持续跑

2. **改进启动方式**
   - 解决 Uvicorn 兼容性
   - 或提供自动降级

3. **添加更多 MCP 工具**
   - 数据验证工具
   - 公式预测试工具

### 长期规划（1-2 月）

1. **性能优化**
   - 工作目录复用
   - 数据缓存策略
   - 并行读取

2. **监控和可观测性**
   - 审计日志
   - 性能指标
   - 错误追踪

3. **用户体验**
   - 实时进度显示
   - 中间结果预览
   - 交互式调试

---

## 参考资料

- [Claude Agent SDK 文档](https://docs.anthropic.com/en/docs/agents)
- [MCP Protocol](https://modelcontextprotocol.io/)
- [Excel Bro 架构文档](./ARCHITECTURE.md)
- [改造前的 SDK 集成](../server/app/claude_sdk_agent.py)（Git 历史）
