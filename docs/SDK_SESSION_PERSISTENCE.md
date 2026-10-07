# SDK 会话持久化方案

> 历史草稿提示（2026-10-06）：本页不是启用 SDK 会话持久化的授权。当前仍保持 no-session-persistence；隔离 Python 的输入/输出/带数据代码不能写入 SDK 转录或恢复账本。新能力按 [隔离 Python 设计](./SDK_ISOLATED_PYTHON_DESIGN_2026-10-06.md) 与 [DeepSeek 实施提示词](./DEEPSEEK_ISOLATED_PYTHON_PROMPT_2026-10-06.md) 实施，执行器尚未实现、未启用。会话持久化需另外审查，正文仅保留追溯。

## 问题

当前的超时限制对数据分析任务来说太短：

```json
// config/capabilities.json

"folder": {
  "sessionTtlSeconds": 1800,  // ← 30分钟会话过期
}

"taskRecovery": {
  "planningHardWallTimeSeconds": 300,  // ← 5分钟硬超时
}

"agent": {
  "maxWallTimeSeconds": 60.0,  // ← 1分钟 Agent 超时
}
```

**影响：**
- 复杂分析任务可能超过 5 分钟
- 用户需要重新开始
- SDK 状态无法恢复
- 30 分钟后文件夹会话失效

---

## 解决方案

### 方案 A：延长超时 + 持久化状态（推荐）

#### A1. 调整超时配置

```json
// config/capabilities.json

{
  "folder": {
    "maxFiles": 100,
    "maxSessions": 20,
    "sessionTtlSeconds": 86400,  // ✅ 改为 24 小时
    "outputFileName": "Excel Bro 结果.xlsx"
  },
  
  "agent": {
    "maxTurns": 8,
    "claudeSdkMaxBudgetUsd": 1.0,
    "maxTotalTokens": null,
    "maxWallTimeSeconds": 600.0,  // ✅ 改为 10 分钟
    "maxConsecutiveErrors": 2,
    "maxCacheHitsPerQuery": 3,
    "maxToolPathLoopLength": 4,
    "warningTurn": 5,
    "maxReadRows": 200,
    "maxReadColumns": 30,
    "maxOutputTokens": null,
    "repairMaxOutputTokens": null
  },
  
  "taskRecovery": {
    "wpsOwnerLeaseSeconds": 120,
    "maxRegisteredDirectories": 50,
    "operationRetentionDays": 30,
    "maxCheckpointBytes": 67108864,
    "planningReminderSeconds": 60,
    "planningExtensionSeconds": 60,
    "planningHardWallTimeSeconds": 1800,  // ✅ 改为 30 分钟
    "planningMaxExtensions": 4
  },
  
  "llm": {
    "maxRetries": 5,
    "retryBaseDelaySeconds": 1.5,
    "retryMaxDelaySeconds": 45.0,
    "timeoutSeconds": 300.0  // ✅ 改为 5 分钟（单次请求）
  },
  
  "turns": {
    "maxActive": 500,
    "ttlSeconds": 86400  // ✅ 改为 24 小时
  },
  
  "undo": {
    "maxEntries": 5,
    "maxTotalCells": 400000,
    "formatCaptureMaxCells": 20000,
    "folderBackupRetentionDays": 7,
    "folderBackupMaxBytes": 1073741824,
    "folderBackupMinFreeBytes": 67108864,
    "folderBackupOpenTokenTtlSeconds": 86400  // ✅ 改为 24 小时
  }
}
```

#### A2. SDK 状态持久化

**当前问题：**
```python
# server/app/claude_sdk_agent.py

# SDK 状态在子进程中，服务端重启后丢失
options = sdk.ClaudeAgentOptions(
    # 状态目录
    cwd=str(workspace_dir),
    # 但没有持久化机制
)
```

**改进方案：**

```python
# server/app/claude_sdk_agent.py

def _get_persistent_workspace(session_id: str) -> Path:
    """为每个会话创建持久的工作目录"""
    state_dir = _state_dir()
    
    # 使用会话 ID 作为目录名（而非随机 UUID）
    workspace_dir = state_dir / f"session_{session_id}"
    workspace_dir.mkdir(parents=True, exist_ok=True)
    
    # 记录会话信息
    session_info_file = workspace_dir / "session_info.json"
    if not session_info_file.exists():
        session_info = {
            "session_id": session_id,
            "created_at": datetime.now().isoformat(),
            "last_accessed": datetime.now().isoformat()
        }
        session_info_file.write_text(json.dumps(session_info, indent=2))
    else:
        # 更新访问时间
        session_info = json.loads(session_info_file.read_text())
        session_info["last_accessed"] = datetime.now().isoformat()
        session_info_file.write_text(json.dumps(session_info, indent=2))
    
    return workspace_dir


def _options(
    sdk: Any, 
    connection: ModelConnection, 
    system_prompt: str,
    session_id: str,  # 🆕 新增参数
    **kwargs
) -> Any:
    state_dir = _state_dir()
    
    # 🆕 使用持久化工作目录
    workspace_dir = _get_persistent_workspace(session_id)
    
    sdk_env = {
        "ANTHROPIC_API_KEY": connection.api_key,
        # ...其他环境变量
    }
    
    return sdk.ClaudeAgentOptions(
        model=connection.model,
        system_prompt=system_prompt,
        tools=["Read", "Bash"],
        disallowed_tools=[...],
        mcp_servers={"excel_bro": server} if server else {},
        
        # ✅ 持久化工作目录（会话级别）
        cwd=str(workspace_dir),
        
        permission_mode="dontAsk",
        max_turns=max_turns,
        max_budget_usd=capability_float("agent", "claudeSdkMaxBudgetUsd"),
        env=sdk_env,
        extra_args={"no-session-persistence": None},
    )


async def run_sdk_plan(
    request: PlanRequest, 
    connection: ModelConnection,
    *, 
    on_event: Callable[[dict[str, Any]], Awaitable[None]] | None = None,
) -> AssistantResponse:
    """完整版 SDK 规划：允许探索，但守卫写入。"""
    
    # 🆕 生成或获取会话 ID
    session_id = request.sessionId if hasattr(request, 'sessionId') else str(uuid.uuid4())
    
    # ...其他代码
    
    options = _options(
        sdk, 
        connection, 
        enhanced_prompt,
        session_id=session_id,  # 🆕 传入会话 ID
        max_turns=capability_int("agent", "maxTurns"),
        server=server,
        allowed=allowed,
    )
    
    # ...
```

#### A3. 会话清理策略

```python
# server/app/claude_sdk_agent.py

def cleanup_old_sessions(max_age_days: int = 7):
    """清理过期的 SDK 会话目录"""
    state_dir = _state_dir()
    now = datetime.now()
    
    for session_dir in state_dir.glob("session_*"):
        if not session_dir.is_dir():
            continue
        
        session_info_file = session_dir / "session_info.json"
        if not session_info_file.exists():
            # 没有会话信息，删除
            shutil.rmtree(session_dir, ignore_errors=True)
            continue
        
        try:
            session_info = json.loads(session_info_file.read_text())
            last_accessed = datetime.fromisoformat(session_info["last_accessed"])
            
            age = now - last_accessed
            if age.days > max_age_days:
                logger.info(f"清理过期会话: {session_dir.name} (闲置 {age.days} 天)")
                shutil.rmtree(session_dir, ignore_errors=True)
        except Exception as e:
            logger.error(f"清理会话失败 {session_dir.name}: {e}")


# 定期清理（例如在启动时或定时任务）
@app.on_event("startup")
async def startup_cleanup():
    """服务启动时清理过期会话"""
    try:
        cleanup_old_sessions(max_age_days=7)
    except Exception as e:
        logger.error(f"启动清理失败: {e}")
```

---

### 方案 B：按需续期机制

**思路：** 让前端定期"心跳"来续期会话

#### B1. 前端心跳

```typescript
// apps/excel-addin/src/sessionKeepAlive.ts

export class SessionKeepAlive {
  private sessionId: string;
  private intervalId: NodeJS.Timeout | null = null;
  
  constructor(sessionId: string) {
    this.sessionId = sessionId;
  }
  
  start() {
    // 每 5 分钟发送一次心跳
    this.intervalId = setInterval(async () => {
      try {
        await fetch('/api/session/keepalive', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ sessionId: this.sessionId })
        });
        console.log(`[KeepAlive] 会话 ${this.sessionId} 已续期`);
      } catch (error) {
        console.error('[KeepAlive] 续期失败:', error);
      }
    }, 5 * 60 * 1000); // 5 分钟
  }
  
  stop() {
    if (this.intervalId) {
      clearInterval(this.intervalId);
      this.intervalId = null;
    }
  }
}

// 使用
const keepAlive = new SessionKeepAlive(sessionId);
keepAlive.start();

// 页面卸载时停止
window.addEventListener('beforeunload', () => {
  keepAlive.stop();
});
```

#### B2. 后端续期接口

```python
# server/app/main.py

@app.post("/api/session/keepalive")
async def keep_session_alive(body: dict):
    """续期会话"""
    session_id = body.get("sessionId")
    
    if not session_id:
        raise HTTPException(status_code=400, detail="缺少 sessionId")
    
    # 更新会话访问时间
    state_dir = Path(os.getenv("CLAUDE_AGENT_SDK_STATE_DIR", ".claude-sdk-state"))
    workspace_dir = state_dir / f"session_{session_id}"
    
    if not workspace_dir.exists():
        raise HTTPException(status_code=404, detail="会话不存在")
    
    session_info_file = workspace_dir / "session_info.json"
    
    if session_info_file.exists():
        session_info = json.loads(session_info_file.read_text())
        session_info["last_accessed"] = datetime.now().isoformat()
        session_info_file.write_text(json.dumps(session_info, indent=2))
        
        return {"status": "ok", "sessionId": session_id}
    else:
        raise HTTPException(status_code=404, detail="会话信息不存在")
```

---

### 方案 C：数据库持久化（完整方案）

**最彻底的方案：** 将 SDK 状态存入数据库

```python
# server/app/sdk_session_store.py

import sqlite3
from pathlib import Path
from datetime import datetime
import json

class SDKSessionStore:
    """SDK 会话持久化存储"""
    
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self._init_db()
    
    def _init_db(self):
        """初始化数据库"""
        conn = sqlite3.connect(str(self.db_path))
        conn.execute("""
            CREATE TABLE IF NOT EXISTS sdk_sessions (
                session_id TEXT PRIMARY KEY,
                workspace_path TEXT NOT NULL,
                created_at TEXT NOT NULL,
                last_accessed TEXT NOT NULL,
                status TEXT NOT NULL,
                metadata TEXT
            )
        """)
        
        conn.execute("""
            CREATE TABLE IF NOT EXISTS sdk_checkpoints (
                checkpoint_id TEXT PRIMARY KEY,
                session_id TEXT NOT NULL,
                turn_number INTEGER NOT NULL,
                state_data TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY (session_id) REFERENCES sdk_sessions(session_id)
            )
        """)
        
        conn.commit()
        conn.close()
    
    def create_session(
        self, 
        session_id: str, 
        workspace_path: str,
        metadata: dict | None = None
    ):
        """创建新会话"""
        conn = sqlite3.connect(str(self.db_path))
        now = datetime.now().isoformat()
        
        conn.execute(
            """
            INSERT INTO sdk_sessions 
            (session_id, workspace_path, created_at, last_accessed, status, metadata)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                session_id,
                workspace_path,
                now,
                now,
                "active",
                json.dumps(metadata or {})
            )
        )
        
        conn.commit()
        conn.close()
    
    def update_access_time(self, session_id: str):
        """更新访问时间"""
        conn = sqlite3.connect(str(self.db_path))
        now = datetime.now().isoformat()
        
        conn.execute(
            "UPDATE sdk_sessions SET last_accessed = ? WHERE session_id = ?",
            (now, session_id)
        )
        
        conn.commit()
        conn.close()
    
    def get_session(self, session_id: str) -> dict | None:
        """获取会话信息"""
        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        
        cursor = conn.execute(
            "SELECT * FROM sdk_sessions WHERE session_id = ?",
            (session_id,)
        )
        
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return dict(row)
        return None
    
    def cleanup_old_sessions(self, max_age_days: int = 7) -> int:
        """清理过期会话"""
        conn = sqlite3.connect(str(self.db_path))
        cutoff = datetime.now().timestamp() - (max_age_days * 86400)
        
        cursor = conn.execute(
            """
            SELECT session_id, workspace_path 
            FROM sdk_sessions 
            WHERE datetime(last_accessed) < datetime(?, 'unixepoch')
            """,
            (cutoff,)
        )
        
        rows = cursor.fetchall()
        
        for row in rows:
            session_id, workspace_path = row
            
            # 删除工作目录
            workspace = Path(workspace_path)
            if workspace.exists():
                shutil.rmtree(workspace, ignore_errors=True)
            
            # 删除数据库记录
            conn.execute(
                "DELETE FROM sdk_sessions WHERE session_id = ?",
                (session_id,)
            )
        
        conn.commit()
        count = len(rows)
        conn.close()
        
        return count


# 使用
session_store = SDKSessionStore(
    Path(os.getenv("CLAUDE_AGENT_SDK_STATE_DIR")) / "sessions.db"
)
```

---

## 推荐实施方案

### 第 1 步：立即改（5 分钟）

**修改 `config/capabilities.json`：**

```json
{
  "folder": {
    "sessionTtlSeconds": 86400  // 30分钟 → 24小时
  },
  "agent": {
    "maxWallTimeSeconds": 600.0  // 1分钟 → 10分钟
  },
  "taskRecovery": {
    "planningHardWallTimeSeconds": 1800  // 5分钟 → 30分钟
  },
  "llm": {
    "timeoutSeconds": 300.0  // 3分钟 → 5分钟
  },
  "turns": {
    "ttlSeconds": 86400  // 1小时 → 24小时
  },
  "undo": {
    "folderBackupOpenTokenTtlSeconds": 86400  // 5分钟 → 24小时
  }
}
```

**效果：**
- ✅ 立即生效
- ✅ 无需代码改动
- ✅ 支持长时间任务

**风险：**
- ⚠️ 内存占用增加（会话不及时清理）
- ⚠️ 需要手动清理旧会话

---

### 第 2 步：持久化工作目录（1 天）

**实现方案 A1 + A2：**
1. 修改 `_options()` 使用会话 ID 而非随机 UUID
2. 添加 `session_info.json` 跟踪
3. 添加清理函数

**效果：**
- ✅ SDK 工作目录可恢复
- ✅ 支持断点续跑（如果 SDK 支持）
- ✅ 自动清理旧会话

---

### 第 3 步：前端心跳（可选，1 天）

**实现方案 B：**
1. 前端添加 `SessionKeepAlive` 类
2. 后端添加 `/api/session/keepalive` 接口

**效果：**
- ✅ 用户活跃时会话不过期
- ✅ 用户关闭页面后自动清理

---

### 第 4 步：数据库持久化（可选，2-3 天）

**实现方案 C：**
1. 创建 `SDKSessionStore` 类
2. 集成到现有代码

**效果：**
- ✅ 完整的会话管理
- ✅ 可查询历史会话
- ✅ 支持检查点恢复

---

## 测试场景

### 场景 1：长时间分析任务

```
用户：分析 5 万行数据的复杂趋势

旧行为：
  5 分钟后 → 超时 ❌
  用户需要重新开始

新行为：
  Agent 运行 15 分钟 → 成功完成 ✅
  工作目录保留 24 小时
```

### 场景 2：会话恢复

```
用户：开始一个复杂任务 → 关闭浏览器 → 重新打开

旧行为：
  会话丢失 ❌
  需要重新开始

新行为（方案 A）：
  工作目录保留 → 可以查看之前的分析结果 ✅
  
新行为（方案 C）：
  完整恢复会话状态 → 继续执行 ✅
```

### 场景 3：自动清理

```
旧会话闲置 7 天

行为：
  自动清理工作目录 ✅
  释放磁盘空间 ✅
```

---

## 配置建议

### 本地开发

```json
{
  "folder": {"sessionTtlSeconds": 86400},  // 24小时
  "agent": {"maxWallTimeSeconds": 600},     // 10分钟
  "taskRecovery": {"planningHardWallTimeSeconds": 1800}  // 30分钟
}
```

### 生产环境

```json
{
  "folder": {"sessionTtlSeconds": 604800},  // 7天
  "agent": {"maxWallTimeSeconds": 1200},    // 20分钟
  "taskRecovery": {"planningHardWallTimeSeconds": 3600}  // 1小时
}
```

### 严格限制（多用户 SaaS）

```json
{
  "folder": {"sessionTtlSeconds": 3600},    // 1小时
  "agent": {"maxWallTimeSeconds": 300},     // 5分钟
  "taskRecovery": {"planningHardWallTimeSeconds": 600}  // 10分钟
}
```

---

## 总结

**立即可做（推荐）：**
1. 修改 `config/capabilities.json` 延长超时
2. 重启服务 → 立即生效

**后续优化（1-2 天）：**
1. 持久化工作目录
2. 添加清理策略
3. （可选）前端心跳

**完整方案（2-3 天）：**
1. 数据库持久化
2. 检查点恢复
3. 会话管理界面
