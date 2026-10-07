# SDK 安全改造 - 实施代码

> 停用提示（2026-10-06）：下面的代码未经当前边界验收，不可直接复制。任意工具 dict 序列化可能泄露原数据；示例契约、测试连接与实际源码也须复核。请按 [隔离 Python 设计](./SDK_ISOLATED_PYTHON_DESIGN_2026-10-06.md) 和 [DeepSeek 实施提示词](./DEEPSEEK_ISOLATED_PYTHON_PROMPT_2026-10-06.md) 实施。用户仅批准真实隔离中的分析，运行器未实现、未启用；这里不提供完成证据，正文仅作历史追溯。

## 改动文件：server/app/claude_sdk_agent.py

### 改动 1：修复非终结工具返回（核心改造）

**位置：** Line 228-247

**当前代码：**
```python
async def handle(args: dict[str, Any], tool_name: str = name) -> dict[str, Any]:
    nonlocal terminal
    if terminal is not None:
        return {"content": [{"type": "text", "text": "本轮结果已提交。"}], "is_error": True}
    try:
        result = _execute_tool(request, tool_name, args)
    except (ValueError, TypeError, KeyError) as error:
        return {
            "content": [{"type": "text", "text": f"规则或计划未通过校验：{str(error)[:300]}"}],
            "is_error": True,
        }
    if isinstance(result, dict):
        return {"content": [{"type": "text", "text": "非终结工具结果不可用。"}], "is_error": True}
    terminal = result
    if on_event:
        await on_event({
            "phase": "planning", "stepId": "planning.submit", "status": "completed",
            "title": "已生成受控结果", "toolName": tool_name,
        })
    return {"content": [{"type": "text", "text": "结果已通过本端校验，请结束本轮。"}]}
```

**改为：**
```python
async def handle(args: dict[str, Any], tool_name: str = name) -> dict[str, Any]:
    nonlocal terminal
    
    # 如果已经提交终结结果，不允许继续调用
    if terminal is not None:
        return {"content": [{"type": "text", "text": "本轮结果已提交。"}], "is_error": True}
    
    try:
        result = _execute_tool(request, tool_name, args)
    except (ValueError, TypeError, KeyError) as error:
        return {
            "content": [{"type": "text", "text": f"规则或计划未通过校验：{str(error)[:300]}"}],
            "is_error": True,
        }
    
    # ✅ 关键改动：区分只读工具和终结工具
    if isinstance(result, dict):
        # 只读工具：返回查询结果，允许继续探索
        # get_workbook_context, find_fields, read_range 都返回 dict
        return {
            "content": [{
                "type": "text",
                "text": json.dumps(result, ensure_ascii=False, indent=2)
            }]
        }
    
    # 终结工具：记录结果并结束
    # submit_plan, submit_answer, submit_clarification 返回 AssistantResponse
    terminal = result
    if on_event:
        await on_event({
            "phase": "planning",
            "stepId": "planning.submit",
            "status": "completed",
            "title": "已生成受控结果",
            "toolName": tool_name,
        })
    return {"content": [{"type": "text", "text": "结果已通过本端校验，请结束本轮。"}]}
```

---

### 改动 2：放宽 planningReadPolicy 限制（可选但推荐）

**位置：** Line 209-213

**当前代码：**
```python
if request.planningReadPolicy != "none" or request.dataResults:
    raise ModelNotAvailableError(
        "Claude Agent SDK 当前仅支持不上传数据行的表结构规划。"
        "请切换到原模型连接处理需要数据行的请求。"
    )
```

**改为：**
```python
# 允许 SDK 使用快照数据（通过只读工具），但不允许前端上传的 dataResults
if request.dataResults:
    raise ModelNotAvailableError(
        "Claude Agent SDK 不支持前端预计算的 dataResults。"
        "请切换到原模型连接，或让 SDK 使用只读工具探索快照数据。"
    )

# ✅ 移除 planningReadPolicy 限制
# SDK 可以通过 read_range 等只读工具读取快照中的数据行
```

---

### 改动 3：增强 System Prompt

**位置：** Line 254-269（在传给 SDK 之前）

**当前代码：**
```python
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
    "planningReadPolicy": "none",
}, ensure_ascii=False)
```

**在之前添加：**
```python
# 构建增强的 System Prompt
enhanced_system_prompt = f"""{AGENT_SYSTEM_PROMPT}

## SDK 模式特殊说明

你现在运行在 Claude Agent SDK 模式下，可以使用以下工具进行多轮探索：

### 只读工具（可多次调用）

这些工具返回查询结果，不会结束本轮对话，你可以继续调用其他工具：

1. **get_workbook_context()** - 查看工作簿结构
   - 无参数
   - 返回：工作表列表、表头、usedRange、rowCount、columnCount 等结构信息
   - 用途：了解工作簿整体结构，确认有哪些表和字段

2. **find_fields(query)** - 查找字段位置
   - 参数：query (string) - 用户提到的字段名称
   - 返回：匹配的字段位置、列地址、非空数据范围
   - 用途：确认用户提到的字段在哪个表的哪一列

3. **read_range(sheet, range)** - 读取数据范围
   - 参数：
     - sheet (string) - 工作表名称
     - range (string) - Excel 地址，如 "A1:C10"
   - 返回：该范围的值（二维数组）
   - 限制：单次最多读取 100 行 × 26 列
   - 用途：查看具体数据，验证字段内容和格式

### 终结工具（只能调用一次，调用后结束本轮）

这些工具会结束本轮对话，提交最终结果：

- **submit_plan(plan)** - 提交 Excel 操作计划
- **submit_answer(message, resultContext?)** - 提交只读查询的回答
- **submit_clarification(clarification)** - 提交澄清问题
- **map_column_values(...)** - 按映射表填写（也是终结工具）

### 推荐工作流程

对于复杂任务，推荐按以下步骤探索：

1. **查看结构** - 调用 `get_workbook_context()` 了解整体结构
2. **定位字段** - 调用 `find_fields("字段名")` 确认字段位置
3. **采样验证** - 调用 `read_range("Sheet1", "A1:E20")` 查看前几行数据
4. **多次验证** - 如有必要，继续调用只读工具确认数据特征
5. **提交结果** - 确认无误后，调用一次终结工具提交计划或回答

### 示例场景

**场景 1：简单查询**
```
用户: "有哪些产品？"

步骤:
1. get_workbook_context() 
   → 看到 "产品表" 有 "产品名称" 列
2. read_range("产品表", "A1:A100")
   → 读取产品名称列
3. submit_answer("共有 50 个产品：...")
```

**场景 2：复杂统计**
```
用户: "按产品统计销量"

步骤:
1. get_workbook_context()
   → 看到 "销售记录" 表有 "产品"、"数量" 列
2. find_fields("销量")
   → 确认 "数量" 列在 B 列
3. read_range("销售记录", "A1:B20")
   → 采样前 20 行，确认数据格式
4. submit_plan({
     actions: [{
       type: "splitGroupAggregate",
       splitBy: ["产品"],
       aggregations: [{{field: "数量", function: "sum"}}]
     }]
   })
```

### 重要规则

✅ **要做：**
- 先用只读工具确认字段名、数据格式
- 采样数据验证假设（不要读取全部数据）
- 确认无误后再提交计划
- 只读工具可以多次调用，充分探索

❌ **不要：**
- 不要猜测字段名（用 find_fields 确认）
- 不要猜测数据格式（用 read_range 验证）
- 不要一次读取大量数据（采样即可）
- 不要提交未验证的计划
- 终结工具只能调用一次（调用后立即结束）

### 数据隐私

- 所有数据都在用户本地快照中
- 只读工具从快照读取，不上传原始数据
- 你看到的数据不会离开用户设备
- read_range 有行列限制，防止一次读取过多数据
"""

# 然后传给 SDK
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

# 使用增强的 system_prompt
options = _options(
    sdk, connection, 
    enhanced_system_prompt,  # ← 使用增强版本
    # ...
)
```

---

## 完整的修改后代码

**文件：** `server/app/claude_sdk_agent.py`

### 修改区域 1：run_sdk_plan 函数开头

```python
async def run_sdk_plan(
    request: PlanRequest, connection: ModelConnection,
    *, on_event: Callable[[dict[str, Any]], Awaitable[None]] | None = None,
) -> AssistantResponse:
    """No workbook rows; explicitly attached images use native content blocks."""
    # ✅ 改动：只检查 dataResults，不检查 planningReadPolicy
    if request.dataResults:
        raise ModelNotAvailableError(
            "Claude Agent SDK 不支持前端预计算的 dataResults。"
            "请切换到原模型连接，或让 SDK 使用只读工具探索快照数据。"
        )
    if request.images and not connection.supports_vision:
        raise ModelNotAvailableError("当前模型不支持图片，请选择支持视觉输入的模型。")
    from .excel_agent import (
        AGENT_SYSTEM_PROMPT, AGENT_TOOLS_WITH_DATA, _execute_tool, _parse_content,
    )
    
    # ... 继续原有代码
```

### 修改区域 2：handle 函数

```python
# 在 _create_mcp_tools 函数中
def _create_mcp_tools(...):
    sdk = _sdk()
    terminal: AssistantResponse | None = None
    tools = []
    allowed = []
    for definition in AGENT_TOOLS_WITH_DATA:
        function = definition["function"]
        name = function["name"]

        async def handle(args: dict[str, Any], tool_name: str = name) -> dict[str, Any]:
            nonlocal terminal
            
            # 如果已经提交终结结果，不允许继续调用
            if terminal is not None:
                return {"content": [{"type": "text", "text": "本轮结果已提交。"}], "is_error": True}
            
            try:
                result = _execute_tool(request, tool_name, args)
            except (ValueError, TypeError, KeyError) as error:
                return {
                    "content": [{"type": "text", "text": f"规则或计划未通过校验：{str(error)[:300]}"}],
                    "is_error": True,
                }
            
            # ✅ 核心改动：区分只读工具和终结工具
            if isinstance(result, dict):
                # 只读工具（get_workbook_context, find_fields, read_range）
                # 返回查询结果，允许继续探索
                return {
                    "content": [{
                        "type": "text",
                        "text": json.dumps(result, ensure_ascii=False, indent=2)
                    }]
                }
            
            # 终结工具（submit_plan, submit_answer, submit_clarification, map_column_values）
            # 记录结果并结束
            terminal = result
            if on_event:
                await on_event({
                    "phase": "planning",
                    "stepId": "planning.submit",
                    "status": "completed",
                    "title": "已生成受控结果",
                    "toolName": tool_name,
                })
            return {"content": [{"type": "text", "text": "结果已通过本端校验，请结束本轮。"}]}

        tools.append(sdk.tool(name, function["description"], function["parameters"])(handle))
        allowed.append(f"mcp__excel_bro__{name}")
    # ... 继续
```

### 修改区域 3：System Prompt 构建

```python
# 在传给 SDK 之前
def _create_mcp_tools(...):
    # ...
    
    # ✅ 构建增强的 System Prompt
    enhanced_system_prompt = f"""{AGENT_SYSTEM_PROMPT}

## SDK 模式特殊说明

你现在运行在 Claude Agent SDK 模式下，可以使用以下工具进行多轮探索：

### 只读工具（可多次调用）

1. **get_workbook_context()** - 查看工作簿结构
2. **find_fields(query)** - 查找字段位置
3. **read_range(sheet, range)** - 读取数据范围（最多 100 行 × 26 列）

### 终结工具（只能调用一次）

- **submit_plan(plan)** - 提交操作计划
- **submit_answer(message)** - 提交回答
- **submit_clarification(clarification)** - 提交澄清
- **map_column_values(...)** - 按映射表填写

### 推荐流程

1. get_workbook_context() - 查看结构
2. find_fields("字段") - 定位字段
3. read_range("Sheet", "A1:E20") - 采样验证
4. submit_plan/answer/clarification - 提交结果

### 规则

✅ 先确认字段名，别猜测
✅ 采样验证数据格式
✅ 只读工具可多次调用
❌ 终结工具只能调用一次
❌ 不要读取全部数据
"""

    # 构建请求 payload
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
    
    # ✅ 使用增强的 system_prompt
    return server, allowed, terminal, enhanced_system_prompt, prompt
```

---

## 测试代码

### 测试 1：只读工具不报错

```python
# server/tests/test_claude_sdk_agent.py

import pytest
from server.app.claude_sdk_agent import run_sdk_plan
from server.app.models import PlanRequest, WorkbookSnapshot, WorksheetSnapshot

@pytest.mark.asyncio
async def test_readonly_tools_work():
    """测试只读工具不再报错"""
    request = PlanRequest(
        prompt="查看产品列表",
        workbook=WorkbookSnapshot(
            name="测试.xlsx",
            worksheets=[
                WorksheetSnapshot(
                    name="产品表",
                    headers=["产品名称", "价格", "库存"],
                    dataRows=[
                        ["产品A", 100, 50],
                        ["产品B", 200, 30],
                    ],
                    rowCount=3,
                    columnCount=3,
                    usedRange="A1:C3",
                )
            ],
        ),
        planningReadPolicy="structure",  # ← 不再报错
        modelId="claude-3-5-sonnet-20241022",
    )
    
    connection = ModelConnection(
        id="test",
        model="claude-3-5-sonnet-20241022",
        adapter="claude_agent_sdk",
    )
    
    result = await run_sdk_plan(request, connection)
    
    # 预期：成功返回，不报 "非终结工具结果不可用"
    assert result.kind in ("answer", "plan", "clarification")
```

### 测试 2：多轮探索

```python
@pytest.mark.asyncio
async def test_multi_turn_exploration():
    """测试多轮探索：get_workbook_context → find_fields → read_range → submit_answer"""
    request = PlanRequest(
        prompt="有哪些产品，价格是多少？",
        workbook=WorkbookSnapshot(
            name="测试.xlsx",
            worksheets=[
                WorksheetSnapshot(
                    name="产品表",
                    headers=["产品名称", "价格", "库存"],
                    dataRows=[
                        ["产品A", 100, 50],
                        ["产品B", 200, 30],
                        ["产品C", 150, 20],
                    ],
                    rowCount=4,
                    columnCount=3,
                    usedRange="A1:C4",
                )
            ],
        ),
        planningReadPolicy="structure",
        modelId="claude-3-5-sonnet-20241022",
    )
    
    connection = ModelConnection(
        id="test",
        model="claude-3-5-sonnet-20241022",
        adapter="claude_agent_sdk",
    )
    
    # 记录工具调用
    tool_calls = []
    
    async def on_event(event):
        if event.get("toolName"):
            tool_calls.append(event["toolName"])
    
    result = await run_sdk_plan(request, connection, on_event=on_event)
    
    # 预期：多次调用只读工具，最后调用终结工具
    assert len(tool_calls) >= 2  # 至少 2 轮
    assert "get_workbook_context" in tool_calls or "find_fields" in tool_calls
    assert result.kind == "answer"
```

---

## 回滚方案

如果改造后出现问题，恢复原始代码：

```python
# 恢复 handle 函数
if isinstance(result, dict):
    return {"content": [{"type": "text", "text": "非终结工具结果不可用。"}], "is_error": True}

# 恢复 planningReadPolicy 检查
if request.planningReadPolicy != "none" or request.dataResults:
    raise ModelNotAvailableError(...)
```

---

## 总结

### 核心改动

**只需改 3 处：**
1. 放宽 `planningReadPolicy` 检查
2. 修改 `handle` 函数（区分只读/终结工具）
3. 增强 System Prompt

### 改动规模

- **代码行数：** ~50 行
- **文件数：** 1 个
- **测试：** 2-3 个测试用例
- **风险：** 低（只读工具，无新风险）

### 预期效果

- ✅ 支持多轮受控探索
- ✅ 字段名准确率提升
- ✅ 安全边界保持
- ✅ 不引入新风险

---

**可以安全实施。**
