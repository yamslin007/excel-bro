# Codex 任务：背景图片功能优化

## 背景

用户上传了女朋友的照片作为背景图，但效果不理想。经过多次迭代调整，现在已经解决了以下问题：

1. ✅ 背景图片显示（之前被遮罩覆盖，图片灰蒙蒙）
2. ✅ 菜单栏可见性（之前被 z-index 问题遮挡）
3. ✅ 图层层次结构（确定了哪些元素需要半透明背景，哪些保持透明）

## 当前实现状态

### 图层结构（从底到顶）

1. **背景图层**（z-index: 0）- 用户上传的背景图片，100% 不透明度显示
2. **消息容器** `.message-content` - 透明
3. **消息标题** `.message-author` - 透明（"Excel Bro 需求确认"）
4. **消息文字** `.message-text` - 92% 白色半透明背景 + 模糊效果
5. **用户气泡** `.user .message-text` - 95% 绿色半透明背景
6. **输入框** `.composer-box` - 92% 白色半透明背景
7. **头部** `.chat-header` - 95% 白色半透明背景

### 相关文件

**`apps/excel-addin/src/hooks/useTheme.ts:105-110`**
```typescript
const DEFAULT_SETTINGS: ThemeSettings = {
  preset: "default",
  opacity: 100,  // 默认 100%，让背景图完整显示
  autoMask: false,  // 默认关闭遮罩
  hasCustomBackground: false
};
```

**`apps/excel-addin/src/styles.css:5001-5066`**
```css
.chat-shell {
  position: relative;
  overflow: hidden;
}

.chat-shell::before,
.chat-shell::after {
  content: "";
  position: absolute;
  inset: 0;
  pointer-events: none;
}

.chat-shell::before {
  z-index: 0;
  background-color: var(--color-background);
  background-image: var(--theme-pattern, none);
  background-position: center;
  background-repeat: repeat;
  background-size: 60px 60px;
  transition: opacity 0.3s ease;
}

.chat-shell::after {
  z-index: 1;
  display: none;
  background: rgba(0, 0, 0, 0.25);
  transition: opacity 0.3s ease;
}

.chat-shell[data-has-background="true"]::before {
  background-image: var(--custom-background-image);
  background-position: center;
  background-repeat: no-repeat;
  background-size: cover;
  opacity: var(--custom-background-opacity);
  z-index: 0;
}

.chat-shell[data-has-background="true"][data-auto-mask="true"]::after {
  display: block;
}

/* 确保对话内容在背景图和遮罩之上 */
.chat-shell .conversation-wrapper,
.chat-shell .composer-box,
.chat-shell .chat-header {
  position: relative;
  z-index: 2;
}

/* 有背景图片时的图层策略：
   - 消息容器保持透明
   - 消息标题保持透明
   - 消息文字、用户气泡加半透明背景
   - 输入框和头部加半透明背景 */

.chat-shell[data-has-background="true"] .message-text {
  background: rgba(255, 255, 255, 0.92);
  backdrop-filter: blur(8px);
  padding: 12px 16px;
  border-radius: 12px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
}

.chat-shell[data-has-background="true"] .user .message-text {
  background: rgba(47, 107, 71, 0.95);
  backdrop-filter: blur(8px);
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.12);
}

.chat-shell[data-has-background="true"] .composer-box {
  background: rgba(255, 255, 255, 0.92);
  backdrop-filter: blur(8px);
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
}

.chat-shell[data-has-background="true"] .chat-header {
  background: rgba(255, 255, 255, 0.95);
  backdrop-filter: blur(12px);
  box-shadow: 0 1px 4px rgba(0, 0, 0, 0.06);
}

.chat-shell[data-has-background="true"] .conversation-controls {
  background: rgba(255, 255, 255, 0.92);
  backdrop-filter: blur(8px);
}
```

## 需要你优化的内容

用户觉得当前效果还不够理想，需要你从以下方面进行优化：

### 1. 视觉美感优化

**问题**：
- 半透明卡片的圆角、阴影、模糊效果可能需要调整
- 白色半透明背景的不透明度可能需要微调
- 用户气泡的绿色半透明可能不够协调
- 消息标题透明显示在复杂背景上，可读性可能不足

**建议方向**：
- 调整 `border-radius`、`box-shadow` 让卡片更精致
- 微调半透明度（90%、92%、95% 等）找到最佳平衡点
- 考虑给透明的消息标题加文字阴影（`text-shadow`）提升可读性
- 检查是否还有其他元素（如意图确认、工具结果等）需要加半透明背景

### 2. 响应式和边界情况

**需要检查**：
- 滚动时背景图是否固定（`background-attachment: fixed` vs `scroll`）
- 长消息内容是否正常显示（padding 是否合适）
- 其他消息类型（意图确认 `.intent-clarification`、工具结果 `.tool-result` 等）是否也需要半透明背景
- 主题切换时（温暖橙、宁静蓝、活力绿）背景图是否协调

### 3. 性能和兼容性

**需要确认**：
- `backdrop-filter: blur()` 在 Excel 加载项的 WebView 中是否支持
- 如果不支持，需要提供降级方案（纯半透明背景，不带模糊）

### 4. 用户体验细节

**建议添加**：
- 平滑过渡动画（`transition: all 0.3s ease`）
- 悬停效果（hover 时卡片略微提升或改变不透明度）
- 确保所有交互元素（按钮、链接等）在背景图上也清晰可见

## 实施要求

1. **只修改 CSS**，不要改动 HTML 结构或 TypeScript 逻辑
2. **保持当前的图层策略**：消息容器和标题透明，内容卡片半透明
3. **测试所有消息类型**：普通消息、意图确认、工具结果、工具抽屉等
4. **提供前后对比**：说明你改了什么，为什么这样改

## 验收标准

✅ 背景图片清晰完整显示  
✅ 所有文字清晰可读（包括透明标题）  
✅ 卡片视觉精致，有层次感  
✅ 用户气泡和助手消息协调统一  
✅ 所有交互元素（按钮、输入框等）正常可用  
✅ 主题切换时效果良好  
✅ 性能流畅，无卡顿

## Git 提交信息

```
fix: 优化背景图片功能的视觉效果和用户体验
```

---

## 附加说明

用户的女朋友会看到这个界面，所以视觉美感很重要！请确保：
- 半透明效果柔和自然
- 卡片和背景图和谐融合
- 文字清晰易读
- 整体感觉温馨、精致

加油！💪✨
