# Notion 风格视觉优化方案

> **目标**：在保持 Fluent Design 基础架构的前提下，借鉴 Notion 的视觉风格，提升 Excel Bro 插件的界面精致度和用户体验。

---

## 📋 设计原则

### Notion 风格的核心特点

1. **暖调灰色阶**
   - 背景不是纯白或冷灰，而是带米色调的暖灰（#F7F6F3）
   - 文字是深棕灰（#37352F），而非纯黑或冷灰
   - 整体视觉更柔和、温暖

2. **清晰的悬停反馈**
   - 列表项悬停时整行浅灰高亮（`rgba(0,0,0,0.03)`）
   - 过渡动画流畅（150ms ease）
   - 点击态有轻微的背景色加深

3. **统一的圆角**
   - 按钮、卡片、输入框使用统一的中等圆角（6-8px）
   - 不是极端的胶囊形（999px），也不是直角
   - 头像等特殊元素可以保持圆形

4. **精准的间距节奏**
   - 使用 8px 网格系统（8px、12px、16px、24px）
   - 行高舒适（body 1.5-1.6，标题 1.2-1.3）
   - 留白充足但不浪费空间

5. **统一的图标风格**
   - Outline 风格，线条粗细 1.5px
   - 尺寸统一（16px 或 20px）
   - 与文字对齐精准

---

## 🎨 色彩方案

### 当前色彩（Fluent Design）

```css
/* 背景 */
--color-background: #ffffff;
--color-surface: #f5f5f5;

/* 文字 */
--color-text-primary: #242424;
--color-text-secondary: #605e5c;

/* 品牌色（绿色） */
--color-brand-bg: #e9f4ec;
--color-brand-text: #34714a;
```

### 目标色彩（Notion 风格）

```css
/* 背景 - 暖调灰 */
--color-background: #FFFFFF;          /* 主背景保持纯白 */
--color-background-subtle: #F7F6F3;   /* 次级背景（侧边栏、卡片） */
--color-surface: #FAFAF9;             /* 悬浮元素背景 */
--color-surface-hover: #F1F0EE;       /* 悬停态背景 */

/* 文字 - 深棕灰 */
--color-text-primary: #37352F;        /* 主要文字 */
--color-text-secondary: #787774;      /* 次要文字 */
--color-text-tertiary: #9B9A97;       /* 辅助文字（占位符、禁用） */

/* 品牌色（保留绿色但微调） */
--color-brand-bg: #EDF3EC;            /* 稍微暖一点的绿 */
--color-brand-text: #2F6B47;          /* 深绿 */
--color-brand-hover: #E3EBE2;         /* 悬停态 */

/* 边框 */
--color-border: #E9E9E7;              /* 默认边框 */
--color-border-strong: #D3D2CF;       /* 强调边框 */

/* 阴影（保持轻量） */
--shadow-sm: 0 1px 2px rgba(0, 0, 0, 0.04);
--shadow-md: 0 2px 4px rgba(0, 0, 0, 0.06);
```

---

## 📐 圆角规范

### 当前圆角

```css
/* 按钮、标签 */
border-radius: 999px;  /* 极端胶囊形 */

/* 卡片、输入框 */
border-radius: 8px;    /* 混用不统一 */
```

### 目标圆角（统一规范）

```css
/* 圆角变量 */
--radius-sm: 4px;      /* 小元素（标签、徽章） */
--radius-md: 8px;      /* 中等元素（按钮、输入框、卡片） */
--radius-lg: 12px;     /* 大元素（模态框、抽屉） */
--radius-full: 50%;    /* 圆形（头像、图标按钮） */

/* 应用规则 */
- 按钮、输入框、列表项：8px
- 消息气泡：8px（不再是 999px）
- 标签（如"AI 模式"）：4px
- 头像、格仔（宠物）：50%
- 工具抽屉、设置面板：12px
```

**注意**：格仔（宠物乌龟）🐢 保持圆形，不受此规范影响。

---

## 🎯 具体改动清单

### 1. 全局色彩变量（`styles.css` 顶部）

**位置**：`apps/excel-addin/src/styles.css:1-100`

**改动**：
- 添加新的 Notion 风格色彩变量
- 保留原有变量名，只修改色值
- 添加悬停态、阴影变量

```css
:root {
  /* 背景 */
  --color-background: #FFFFFF;
  --color-background-subtle: #F7F6F3;
  --color-surface: #FAFAF9;
  --color-surface-hover: #F1F0EE;
  
  /* 文字 */
  --color-text-primary: #37352F;
  --color-text-secondary: #787774;
  --color-text-tertiary: #9B9A97;
  
  /* 品牌色 */
  --color-brand-bg: #EDF3EC;
  --color-brand-text: #2F6B47;
  --color-brand-hover: #E3EBE2;
  
  /* 边框 */
  --color-border: #E9E9E7;
  --color-border-strong: #D3D2CF;
  
  /* 阴影 */
  --shadow-sm: 0 1px 2px rgba(0, 0, 0, 0.04);
  --shadow-md: 0 2px 4px rgba(0, 0, 0, 0.06);
  
  /* 圆角 */
  --radius-sm: 4px;
  --radius-md: 8px;
  --radius-lg: 12px;
  --radius-full: 50%;
  
  /* 间距（8px 网格） */
  --spacing-xs: 4px;
  --spacing-sm: 8px;
  --spacing-md: 12px;
  --spacing-lg: 16px;
  --spacing-xl: 24px;
}
```

---

### 2. 主容器背景

**位置**：`styles.css` 中的 `.app-container`、`.conversation-container`

**改动**：
```css
.app-container {
  background: var(--color-background);  /* 纯白 */
}

.conversation-container {
  background: var(--color-background-subtle);  /* 暖灰 #F7F6F3 */
}
```

---

### 3. 消息样式优化

#### 3.1 消息气泡圆角

**位置**：`styles.css:2800-2850`（`.message` 相关）

**改动**：
```css
/* 助手消息（左侧，无背景） */
.assistant .message-text {
  background: transparent;
  /* 无需圆角 */
}

/* 用户消息（右侧，绿色气泡） */
.user .message-text {
  background: var(--color-brand-bg);
  border-radius: var(--radius-md);  /* 从 999px 改为 8px */
  padding: var(--spacing-sm) var(--spacing-md);
}
```

#### 3.2 消息作者标签

**位置**：`styles.css:2824-2829`

**改动**：
```css
.message-author em {
  border-radius: var(--radius-sm);  /* 从 999px 改为 4px */
  background: var(--color-brand-bg);
  color: var(--color-brand-text);
  padding: 2px 6px;
  font-size: var(--font-size-caption);
  font-style: normal;
}
```

---

### 4. 按钮样式

**位置**：`styles.css` 中的 `button`、`.send-button`、`.tool-action-button` 等

**改动**：
```css
button {
  border-radius: var(--radius-md);  /* 统一 8px */
  background: var(--color-surface);
  color: var(--color-text-primary);
  border: 1px solid var(--color-border);
  padding: var(--spacing-sm) var(--spacing-lg);
  font-size: var(--font-size-body);
  cursor: pointer;
  transition: background 150ms ease, border-color 150ms ease, transform 150ms ease;
}

button:hover {
  background: var(--color-surface-hover);
  border-color: var(--color-border-strong);
  transform: translateY(-1px);  /* 轻微上浮 */
}

button:active {
  transform: translateY(0);
}

/* 主操作按钮（发送、确认等） */
.primary-button {
  background: var(--color-brand-text);
  color: #FFFFFF;
  border: none;
}

.primary-button:hover {
  background: #265336;  /* 稍深的绿 */
}
```

---

### 5. 输入框样式

**位置**：`styles.css` 中的 `textarea`、`input`

**改动**：
```css
textarea,
input[type="text"] {
  border-radius: var(--radius-md);
  background: var(--color-surface);
  color: var(--color-text-primary);
  border: 1px solid var(--color-border);
  padding: var(--spacing-sm) var(--spacing-md);
  font-size: var(--font-size-body);
  transition: border-color 150ms ease, box-shadow 150ms ease;
}

textarea:focus,
input[type="text"]:focus {
  outline: none;
  border-color: var(--color-brand-text);
  box-shadow: 0 0 0 2px var(--color-brand-bg);  /* 绿色聚焦环 */
}

textarea::placeholder,
input[type="text"]::placeholder {
  color: var(--color-text-tertiary);
}
```

---

### 6. 列表项悬停态

**位置**：`styles.css` 中的工具列表、工作表列表

**改动**：
```css
/* 工具列表项 */
.tool-item {
  padding: var(--spacing-sm) var(--spacing-md);
  border-radius: var(--radius-md);
  cursor: pointer;
  transition: background 150ms ease;
}

.tool-item:hover {
  background: var(--color-surface-hover);  /* Notion 式整行高亮 */
}

/* 工作表列表项 */
.worksheet-item {
  padding: var(--spacing-sm) var(--spacing-md);
  border-radius: var(--radius-md);
  cursor: pointer;
  transition: background 150ms ease;
}

.worksheet-item:hover {
  background: var(--color-surface-hover);
}

.worksheet-item.active {
  background: var(--color-brand-bg);
  color: var(--color-brand-text);
}
```

---

### 7. 卡片样式（计划、工具详情等）

**位置**：`styles.css` 中的 `.plan-card`、`.tool-card` 等

**改动**：
```css
.card {
  background: var(--color-surface);
  border: 1px solid var(--color-border);
  border-radius: var(--radius-md);
  padding: var(--spacing-lg);
  box-shadow: var(--shadow-sm);
}

.card:hover {
  box-shadow: var(--shadow-md);
  border-color: var(--color-border-strong);
}
```

---

### 8. 工具抽屉（Tool Drawer）

**位置**：`styles.css` 中的 `.tool-drawer`

**改动**：
```css
.tool-drawer {
  background: var(--color-background);
  border-radius: var(--radius-lg) var(--radius-lg) 0 0;  /* 顶部圆角 */
  box-shadow: 0 -4px 12px rgba(0, 0, 0, 0.08);
  padding: var(--spacing-xl);
}

.tool-drawer-header {
  font-size: var(--font-size-subtitle);
  font-weight: 600;
  color: var(--color-text-primary);
  margin-bottom: var(--spacing-lg);
}
```

---

### 9. 顶部状态栏

**位置**：`styles.css` 中的 `.header`、`.status-bar` 等

**改动**：
```css
.header {
  background: var(--color-surface);
  border-bottom: 1px solid var(--color-border);
  padding: var(--spacing-md) var(--spacing-lg);
}

.workbook-name {
  font-size: var(--font-size-body);
  font-weight: 600;
  color: var(--color-text-primary);
}

.status-indicator {
  font-size: var(--font-size-caption);
  color: var(--color-text-secondary);
}
```

---

### 10. 图标统一

**位置**：所有 SVG 图标

**要求**：
- 线条粗细统一为 `stroke-width="1.5"`
- 尺寸统一为 16px 或 20px
- 颜色使用 `currentColor` 继承文字色
- Outline 风格，不用 filled 图标

**示例**：
```tsx
<svg width="16" height="16" viewBox="0 0 16 16" fill="none">
  <path
    d="M8 1v14M1 8h14"
    stroke="currentColor"
    strokeWidth="1.5"
    strokeLinecap="round"
  />
</svg>
```

---

## 📂 涉及文件

| 文件 | 改动内容 | 优先级 |
|------|---------|--------|
| `apps/excel-addin/src/styles.css` | 全局色彩变量、圆角变量、间距变量 | 🔥 高 |
| `apps/excel-addin/src/styles.css` | 消息样式（气泡圆角、标签圆角） | 🔥 高 |
| `apps/excel-addin/src/styles.css` | 按钮样式（圆角、悬停态、动画） | 🔥 高 |
| `apps/excel-addin/src/styles.css` | 输入框样式（圆角、聚焦态） | 🔥 高 |
| `apps/excel-addin/src/styles.css` | 列表项悬停态（工具、工作表） | 🔥 高 |
| `apps/excel-addin/src/styles.css` | 卡片样式 | 🟡 中 |
| `apps/excel-addin/src/styles.css` | 工具抽屉样式 | 🟡 中 |
| `apps/excel-addin/src/styles.css` | 顶部状态栏样式 | 🟡 中 |
| 所有 SVG 图标文件 | 统一线条粗细和尺寸 | 🟢 低 |

---

## 🧪 验收标准

### 视觉验收

1. **色彩统一**
   - ✅ 背景色使用暖灰色阶（#F7F6F3、#FAFAF9）
   - ✅ 文字色使用深棕灰（#37352F）
   - ✅ 品牌绿色微调为暖色调

2. **圆角统一**
   - ✅ 按钮、输入框、列表项：8px
   - ✅ 标签（如"AI 模式"）：4px
   - ✅ 消息气泡：8px（不再是 999px 胶囊形）
   - ✅ 头像、格仔：50%（圆形）

3. **悬停反馈**
   - ✅ 列表项悬停时整行浅灰高亮
   - ✅ 按钮悬停时轻微上浮（`translateY(-1px)`）
   - ✅ 过渡动画流畅（150ms ease）

4. **间距节奏**
   - ✅ 使用 8px 网格系统（8px、12px、16px、24px）
   - ✅ 行高舒适（body 1.6，标题 1.3）

5. **图标一致性**
   - ✅ 线条粗细统一（1.5px）
   - ✅ 尺寸统一（16px 或 20px）
   - ✅ Outline 风格

---

### 功能验收

1. **不影响现有功能**
   - ✅ 所有按钮可点击
   - ✅ 输入框可输入
   - ✅ 列表项可选择
   - ✅ 工具抽屉可展开/收起

2. **性能无影响**
   - ✅ CSS 动画流畅（60fps）
   - ✅ 页面加载速度不变

3. **可访问性保持**
   - ✅ 色彩对比度符合 WCAG AA 标准
   - ✅ 聚焦态清晰可见
   - ✅ 键盘导航正常

---

## 🎨 Notion 风格核心：悬停反馈示例

**Notion 的列表项悬停效果**是其视觉体验的核心，以下是完整的实现代码：

```css
/* 基础列表项 */
.list-item {
  display: flex;
  align-items: center;
  padding: 6px 8px;
  border-radius: 4px;
  cursor: pointer;
  transition: background 150ms ease;
  user-select: none;
}

/* 悬停态 */
.list-item:hover {
  background: rgba(0, 0, 0, 0.03);  /* Notion 的经典悬停色 */
}

/* 点击态 */
.list-item:active {
  background: rgba(0, 0, 0, 0.05);
}

/* 选中态 */
.list-item.selected {
  background: var(--color-brand-bg);
  color: var(--color-brand-text);
}

.list-item.selected:hover {
  background: var(--color-brand-hover);
}
```

**应用到 Excel Bro**：
- 工具列表（Tool List）
- 工作表列表（Worksheet List）
- 模型列表（Model List）
- 设置项列表

---

## 📝 实现步骤

### 阶段 1：基础色彩和圆角（30 分钟）

1. **修改全局变量**
   - 在 `styles.css` 顶部添加新的色彩和圆角变量
   - 测试变量是否生效（检查浏览器开发者工具）

2. **替换现有色值**
   - 全局搜索 `#ffffff`、`#f5f5f5` 等旧色值
   - 替换为新的变量引用
   - 全局搜索 `border-radius: 999px`
   - 替换为 `var(--radius-md)` 或 `var(--radius-sm)`

3. **测试构建**
   ```bash
   npm run build:addin
   npm run lint
   ```

---

### 阶段 2：悬停态和动画（20 分钟）

1. **添加列表项悬停态**
   - 工具列表（`.tool-item`）
   - 工作表列表（`.worksheet-item`）
   - 其他可交互列表

2. **添加按钮悬停动画**
   - `transform: translateY(-1px)`
   - `transition: 150ms ease`

3. **测试交互**
   - 鼠标悬停时是否有明显反馈
   - 动画是否流畅（60fps）

---

### 阶段 3：细节打磨（20 分钟）

1. **统一图标样式**
   - 检查所有 SVG 图标
   - 确保 `stroke-width="1.5"`、尺寸统一

2. **检查间距节奏**
   - 确保使用 `var(--spacing-*)` 变量
   - 不要出现硬编码的 `padding: 10px` 等

3. **测试深色模式（如果有）**
   - 如果支持深色模式，确保新色彩在深色模式下也合理

---

### 阶段 4：全量测试（10 分钟）

1. **视觉回归测试**
   - 打开所有主要界面（对话、工具管理、模型管理）
   - 截图对比改动前后
   - 确保没有遗漏的地方

2. **功能测试**
   - 发送消息
   - 选择工作表
   - 打开工具抽屉
   - 切换模型
   - 确保所有交互正常

3. **性能测试**
   - 打开浏览器性能面板
   - 检查动画是否流畅（60fps）
   - 检查页面加载速度

---

## 🚀 渐进式迁移策略

如果担心一次性改动太大，可以分阶段迁移：

### 第一批（高优先级，影响最大）
1. 全局色彩变量（背景、文字）
2. 按钮圆角和悬停态
3. 消息气泡圆角

### 第二批（中优先级，提升体验）
1. 列表项悬停态
2. 输入框聚焦态
3. 卡片样式

### 第三批（低优先级，锦上添花）
1. 图标统一
2. 细节动画优化
3. 深色模式适配（如果有）

---

## 💡 注意事项

1. **保持格仔（宠物乌龟）🐢 的圆形**
   - 格仔是情感元素，必须保持可见且圆形
   - 不要因为统一圆角而改成 8px

2. **不要过度使用阴影**
   - Notion 风格的阴影很轻（`0 1px 2px rgba(0,0,0,0.04)`）
   - 不要用厚重的 Material Design 阴影

3. **保持可访问性**
   - 新色彩的对比度必须符合 WCAG AA 标准
   - 聚焦态必须清晰可见
   - 键盘导航不能受影响

4. **测试窄窗口**
   - Excel 任务窗格最窄可能只有 300px
   - 确保新样式在窄窗口下不会错位

5. **兼容 Office 主题**
   - 用户可能使用深色主题的 Excel
   - 确保插件在不同主题下都可用

---

## 📊 预期效果

### 改动前
- 背景：纯白 / 冷灰
- 文字：纯黑 / 冷灰
- 按钮：999px 胶囊形
- 悬停：无明显反馈
- 整体感觉：冷淡、不够精致

### 改动后
- 背景：暖灰色阶（#F7F6F3）
- 文字：深棕灰（#37352F）
- 按钮：8px 圆角，悬停上浮
- 悬停：整行浅灰高亮
- 整体感觉：温暖、精致、清晰

---

**文档版本**：v1.0  
**创建日期**：2026-08-18  
**预计实现时间**：1.5 小时  
**优先级**：🔥 高（显著提升用户体验）  
**状态**：待实现 ⏳
