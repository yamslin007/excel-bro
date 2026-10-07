# Git 提交指南 - SDK 安全改造方案

## 提交内容

### 本次提交包含

1. ✅ **配置优化**（已完成，立即生效）
   - `config/capabilities.json` - 6 处超时参数调整

2. ✅ **安全改造方案文档**（设计完成，待实施）
   - 完整指南
   - 实施代码
   - 工作日志
   - 文档索引

3. ⚠️ **废弃文档标记**（错误方案，请勿使用）

---

## 提交方式

### 方案 A：分两次提交（推荐）

#### 第 1 次提交：配置优化（立即提交）

```bash
# 只提交配置改动
git add config/capabilities.json

git commit -m "config: 优化超时配置以支持长时间任务

改动：
- folder.sessionTtlSeconds: 1800 → 86400 (30分钟 → 24小时)
- agent.maxWallTimeSeconds: 60 → 600 (1分钟 → 10分钟)
- taskRecovery.planningHardWallTimeSeconds: 300 → 1800 (5分钟 → 30分钟)
- llm.timeoutSeconds: 180 → 300 (3分钟 → 5分钟)
- turns.ttlSeconds: 3600 → 86400 (1小时 → 24小时)
- undo.folderBackupOpenTokenTtlSeconds: 300 → 86400 (5分钟 → 24小时)

效果：
- 支持更长时间的复杂分析任务
- 会话保持更长时间，减少中断
- 立即生效（重启服务）

参考：docs/TIMEOUT_CONFIG_CHANGES.md

Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>"
```

#### 第 2 次提交：改造方案文档（设计文档）

```bash
# 提交正确方案文档
git add docs/SDK_SAFE_TRANSFORMATION_GUIDE.md
git add docs/SDK_SAFE_TRANSFORMATION_CODE.md
git add docs/WORK_LOG_2026-10-06_CORRECTED.md
git add docs/SDK_TRANSFORMATION_START_HERE.md

# 提交废弃文档标记
git add docs/TIMEOUT_CONFIG_CHANGES.md
git add docs/SDK_SESSION_PERSISTENCE.md

git commit -m "docs: SDK 安全改造方案（设计阶段）

背景：
当前 SDK 模式下，只读工具（get_workbook_context, find_fields, read_range）
返回 dict 时被直接拒绝，导致无法多轮探索，字段名准确率低。

设计方案：
- 修改 server/app/claude_sdk_agent.py 的 handle 函数
- 区分只读工具（返回 dict）和终结工具（返回 AssistantResponse）
- 只读工具返回 JSON，允许继续探索
- 终结工具结束本轮

预期效果：
- 支持 3-5 轮受控探索
- 字段名准确率提升 70% (50% → 85%)
- 保持安全边界不变（不开放 Bash/Write/联网）
- 保持数据隐私（不自动上传整表）

文档：
- docs/SDK_SAFE_TRANSFORMATION_GUIDE.md - 完整指南
- docs/SDK_SAFE_TRANSFORMATION_CODE.md - 实施代码
- docs/WORK_LOG_2026-10-06_CORRECTED.md - 工作日志
- docs/SDK_TRANSFORMATION_START_HERE.md - 文档索引

注意：
本次提交为设计文档，代码改动将在后续实施验证后提交。

感谢：
感谢 DeepSeek 审查指出初版方案的安全问题（开放 Bash、自动上传整表）
本方案基于受控探索思路重新设计，保持安全边界。

参考：docs/SDK_TRANSFORMATION_START_HERE.md

Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>"
```

---

### 方案 B：一次性提交

```bash
# 提交所有改动
git add config/capabilities.json
git add docs/SDK_SAFE_TRANSFORMATION_GUIDE.md
git add docs/SDK_SAFE_TRANSFORMATION_CODE.md
git add docs/WORK_LOG_2026-10-06_CORRECTED.md
git add docs/SDK_TRANSFORMATION_START_HERE.md
git add docs/TIMEOUT_CONFIG_CHANGES.md
git add docs/SDK_SESSION_PERSISTENCE.md

git commit -m "feat: SDK 模式改造设计与配置优化

## 配置优化（立即生效）

改动：config/capabilities.json (6 处)
- 会话 TTL: 30分钟 → 24小时
- Agent 超时: 1分钟 → 10分钟
- 规划超时: 5分钟 → 30分钟
- LLM 超时: 3分钟 → 5分钟

效果：支持长时间任务，减少中断

## SDK 改造方案（设计阶段）

问题：
- 只读工具返回 dict 被直接拒绝
- 无法多轮探索
- 字段名准确率低 (~50%)

方案：
- 修改 handle 函数区分只读/终结工具
- 只读工具返回 JSON，允许继续
- 支持 3-5 轮受控探索

预期：
- 字段名准确率 +70% (50% → 85%)
- 安全边界保持不变
- 数据隐私保持不变

文档：
- docs/SDK_TRANSFORMATION_START_HERE.md - 从这里开始
- docs/SDK_SAFE_TRANSFORMATION_GUIDE.md - 完整指南
- docs/SDK_SAFE_TRANSFORMATION_CODE.md - 实施代码

注意：
- 本次为设计文档，代码改动待实施验证后提交
- 初版方案有安全问题已废弃（开放 Bash/整表上传）
- 当前方案基于受控探索，保持安全边界

Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>"
```

---

## 废弃文档的处理

### 选项 1：不提交（推荐）

```bash
# 废弃文档不提交到仓库
# 只保留本地以供参考历史错误

# 添加到 .gitignore（如果需要）
echo "docs/CLAUDE_SDK_TRANSFORMATION_GUIDE.md" >> .gitignore
echo "docs/SDK_TRANSFORMATION_CODE_TEMPLATES.md" >> .gitignore
echo "docs/SDK_TRANSFORMATION_DELIVERY_PACKAGE.md" >> .gitignore
echo "docs/SDK_TRANSFORMATION_QUICK_REFERENCE.md" >> .gitignore
echo "docs/SDK_TRANSFORMATION_README.md" >> .gitignore
echo "docs/DEEPSEEK_PROMPT_FOR_SDK.md" >> .gitignore
echo "docs/FINAL_SUMMARY_2026-10-06.md" >> .gitignore
echo "docs/GIT_COMMIT_GUIDE.md" >> .gitignore
```

### 选项 2：删除后提交

```bash
# 删除废弃文档
rm docs/CLAUDE_SDK_TRANSFORMATION_GUIDE.md
rm docs/SDK_TRANSFORMATION_CODE_TEMPLATES.md
rm docs/SDK_TRANSFORMATION_DELIVERY_PACKAGE.md
rm docs/SDK_TRANSFORMATION_QUICK_REFERENCE.md
rm docs/SDK_TRANSFORMATION_README.md
rm docs/DEEPSEEK_PROMPT_FOR_SDK.md
rm docs/FINAL_SUMMARY_2026-10-06.md
rm docs/GIT_COMMIT_GUIDE.md

# 如果已经提交过，需要从 git 历史删除
git rm docs/CLAUDE_SDK_TRANSFORMATION_GUIDE.md
# ...
```

---

## 后续实施提交（Day 1-3）

### Day 1：代码改造

```bash
# 实施改造后
git add server/app/claude_sdk_agent.py
git add server/tests/test_claude_sdk_agent.py

git commit -m "feat: 支持 SDK 模式多轮只读探索

改动：
- server/app/claude_sdk_agent.py:
  - handle 函数区分只读/终结工具 (line 228-247)
  - 放宽 planningReadPolicy 限制 (line 209-213)
  - 增强 System Prompt (~50 行)
  
- server/tests/test_claude_sdk_agent.py:
  - 测试只读工具不报错
  - 测试多轮探索

效果：
- ✅ 只读工具（get_workbook_context, find_fields, read_range）不再报错
- ✅ 支持 3-5 轮探索
- ✅ 字段名准确率实测提升
- ✅ 安全边界保持不变

测试：
python -m pytest server/tests/test_claude_sdk_agent.py -v

参考：docs/SDK_SAFE_TRANSFORMATION_CODE.md

Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>"
```

### Day 2-3：文档更新

```bash
git add README.md
git add docs/ARCHITECTURE.md
git add docs/DEVELOPMENT.md

git commit -m "docs: 更新 SDK 模式说明

更新：
- README.md - SDK 模式能力说明
- ARCHITECTURE.md - 架构图更新（只读工具流程）
- DEVELOPMENT.md - 开发者指南

Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>"
```

---

## 推荐提交流程

```bash
# 第 1 步：立即提交配置优化
git add config/capabilities.json
git commit -m "config: 优化超时配置..."

# 第 2 步：提交设计文档
git add docs/SDK_*.md docs/WORK_LOG_2026-10-06_CORRECTED.md
git commit -m "docs: SDK 安全改造方案（设计阶段）..."

# 第 3 步：实施后提交代码（Day 1）
git add server/app/claude_sdk_agent.py server/tests/...
git commit -m "feat: 支持 SDK 模式多轮只读探索..."

# 第 4 步：更新文档（Day 2-3）
git add README.md docs/ARCHITECTURE.md docs/DEVELOPMENT.md
git commit -m "docs: 更新 SDK 模式说明..."
```

---

## 检查清单

### 提交前

- [ ] 配置改动已测试（重启服务验证）
- [ ] 文档链接正确（相互引用）
- [ ] 没有提交废弃文档
- [ ] commit message 清晰
- [ ] 包含 Co-Authored-By 行

### 实施前

- [ ] 已读 SDK_TRANSFORMATION_START_HERE.md
- [ ] 已读 SDK_SAFE_TRANSFORMATION_GUIDE.md
- [ ] 已读 SDK_SAFE_TRANSFORMATION_CODE.md
- [ ] 理解为什么错误方案被废弃

### 实施后

- [ ] 运行测试通过
- [ ] 手动验证只读工具
- [ ] 验证字段名准确率
- [ ] 验证安全边界不变
- [ ] 更新相关文档

---

## 注意事项

### 重要提醒

1. **配置改动立即生效** - 重启服务后即可使用
2. **代码改动需验证** - 实施前先测试
3. **废弃文档勿提交** - 避免误导他人
4. **分阶段提交** - 配置、设计、实施、文档分开

### 回滚方案

如果实施后出现问题：

```bash
# 回滚代码改动
git revert <commit-hash>

# 或手动恢复
# 参考 SDK_SAFE_TRANSFORMATION_CODE.md 的回滚方案
```

---

## 总结

### 本次提交

- ✅ 配置优化（6 处，立即生效）
- ✅ 安全改造方案（设计文档）
- ✅ 工作日志（记录过程）
- ✅ 文档索引（从这里开始）

### 后续提交

- Day 1: 代码改造
- Day 2: 测试验证
- Day 3: 文档更新

### 预期收益

- 字段名准确率 +70%
- 支持多轮探索
- 安全边界不变

---

**准备好了就提交！** 🚀
