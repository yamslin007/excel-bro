# 超时配置改动记录

## 修改日期：2026-10-06

## 问题
原有的超时配置对数据分析任务来说太短，导致：
- 30 分钟后会话过期
- 5 分钟硬超时无法完成复杂分析
- SDK 无法进行充分探索

## 改动内容

### 1. 文件夹会话 TTL
```json
"folder.sessionTtlSeconds": 1800 → 86400
```
- **旧值：** 30 分钟
- **新值：** 24 小时
- **影响：** 文件夹模式的工作簿会话保持 24 小时

### 2. Agent 最大执行时间
```json
"agent.maxWallTimeSeconds": 60.0 → 600.0
```
- **旧值：** 1 分钟
- **新值：** 10 分钟
- **影响：** Agent 循环可以运行更长时间

### 3. 规划硬超时
```json
"taskRecovery.planningHardWallTimeSeconds": 300 → 1800
```
- **旧值：** 5 分钟
- **新值：** 30 分钟
- **影响：** 规划阶段最多可以运行 30 分钟

### 4. LLM 请求超时
```json
"llm.timeoutSeconds": 180.0 → 300.0
```
- **旧值：** 3 分钟
- **新值：** 5 分钟
- **影响：** 单次 LLM API 调用超时时间

### 5. Turn 状态 TTL
```json
"turns.ttlSeconds": 3600 → 86400
```
- **旧值：** 1 小时
- **新值：** 24 小时
- **影响：** Turn 状态缓存保持 24 小时

### 6. 备份 Token TTL
```json
"undo.folderBackupOpenTokenTtlSeconds": 300 → 86400
```
- **旧值：** 5 分钟
- **新值：** 24 小时
- **影响：** 备份访问令牌有效期延长

## 预期效果

### ✅ 支持的场景

**场景 1：长时间数据分析**
```
旧：5 分钟超时 → 任务中断
新：30 分钟硬限制 → 完成复杂分析
```

**场景 2：多轮探索**
```
旧：1 分钟 Agent 超时 → 无法充分探索
新：10 分钟 Agent 时间 → 3-5 轮深入分析
```

**场景 3：会话恢复**
```
旧：30 分钟后会话失效 → 需重新开始
新：24 小时会话保持 → 可以继续工作
```

### ⚠️ 需要注意

**内存占用：**
- 会话和 Turn 状态保留更久
- 可能增加服务器内存占用
- 建议监控内存使用情况

**磁盘占用：**
- 备份和工作目录保留更久
- 7 天后自动清理（`folderBackupRetentionDays`）
- SDK 工作目录需要手动清理策略

**用户体验：**
- 长任务不会意外中断 ✅
- 可以关闭浏览器后继续 ✅
- 但也意味着错误任务会占用资源更久 ⚠️

## 回滚方案

如果需要回滚到原有配置：

```json
{
  "folder": {"sessionTtlSeconds": 1800},
  "agent": {"maxWallTimeSeconds": 60.0},
  "taskRecovery": {"planningHardWallTimeSeconds": 300},
  "llm": {"timeoutSeconds": 180.0},
  "turns": {"ttlSeconds": 3600},
  "undo": {"folderBackupOpenTokenTtlSeconds": 300}
}
```

## 下一步优化

参考 `docs/SDK_SESSION_PERSISTENCE.md`：

1. **持久化工作目录**（1 天）
   - 使用会话 ID 而非随机 UUID
   - 记录会话信息
   - 自动清理过期会话

2. **前端心跳机制**（1 天）
   - 定期续期活跃会话
   - 用户关闭页面自动清理

3. **数据库持久化**（2-3 天）
   - 完整会话管理
   - 检查点恢复
   - 历史查询

## 测试清单

- [ ] 启动服务验证配置加载
- [ ] 运行超过 5 分钟的分析任务
- [ ] 验证会话在 24 小时内有效
- [ ] 监控内存和磁盘使用
- [ ] 测试备份令牌有效期
- [ ] 验证自动清理机制（7 天）

## 相关文档

- `docs/SDK_SESSION_PERSISTENCE.md` - 持久化完整方案
- `docs/CLAUDE_SDK_TRANSFORMATION_GUIDE.md` - SDK 改造指南
- `config/capabilities.json` - 配置文件
