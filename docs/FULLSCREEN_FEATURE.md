# 全屏功能添加 - 2026-08-24

## 改动内容

为**专注窗口**和**编辑脚本窗口**添加全屏按钮，允许用户将窗口最大化以获得更大的编辑/查看空间。

### 技术实现

使用浏览器的 **Fullscreen API**：
- `document.documentElement.requestFullscreen()` - 进入全屏
- `document.exitFullscreen()` - 退出全屏
- 监听 `fullscreenchange` 事件更新按钮状态

### 修改文件

**专注窗口：**
- `src/FocusWorkspace.tsx` (+30 行)
  - 添加 `isFullscreen` 状态
  - 添加 `toggleFullscreen` 函数
  - 添加全屏按钮到 header
  - 监听 `fullscreenchange` 事件

- `src/focus.css` (+10 行)
  - 添加 `.focus-fullscreen` 样式
  - hover 效果

**编辑脚本窗口：**
- `src/components/PlanEditorDialog.tsx` (+35 行)
  - 添加 `isFullscreen` 状态
  - 添加 `toggleFullscreen` 函数
  - 添加全屏按钮到 header
  - 监听 `fullscreenchange` 事件

- `src/styles.css` (+25 行)
  - 添加 `.plan-editor-dialog-header-actions` 样式
  - 按钮组布局（全屏 + 关闭）

### UI 设计

**按钮位置：** header 右侧，全屏按钮在关闭按钮左边
**图标：** ⛶（全屏图标，进入和退出使用相同图标）
**样式：** 与现有按钮风格一致（圆角、边框、hover 效果）

### 用户体验

- **点击全屏按钮** → 窗口占据整个屏幕
- **再次点击** 或 **按 ESC 键** → 退出全屏
- **tooltip 提示**："全屏" / "退出全屏"
- **状态同步**：按钮文本根据当前状态自动更新

### 构建结果

```
✓ built in 983ms
focus.css      +0.13 kB (5.82 kB)
styles.css     +0.48 kB (101.20 kB)
planEditor.js  +0.62 kB (3.95 kB)
focus.js       +0.52 kB (6.43 kB)
```

### 浏览器兼容性

Fullscreen API 支持：
- ✅ Chrome/Edge (Chromium)
- ✅ Firefox
- ✅ Safari
- ✅ Excel Online（WebView2）

---

**状态：** 已完成，构建通过  
**待测试：** 在 Excel 中验证全屏功能正常工作  
**完成时间：** 2026-08-24
