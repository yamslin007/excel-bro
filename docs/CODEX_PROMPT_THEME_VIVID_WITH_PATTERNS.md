# Codex 任务：主题系统完整修复（明亮配色 + 专属图案）

## 设计理念

1. **色彩**：明亮清新有活力，但饱和度适中不刺眼
2. **图案**：每个主题专属的可爱小图案，增加趣味性
3. **配套**：完整的 14 变量色彩系统，所有元素协调统一
4. **层级**：背景图片、图案、遮罩、内容四层清晰分离

---

## 问题 1：输入框双边框问题（高优先级）

### 问题描述

点击输入框后，出现两个绿色边框，视觉上很突兀。

### 问题定位

在 `styles.css` 中有两处聚焦态样式冲突：

**第一处**：`styles.css:5411-5414` - 外层容器聚焦态
```css
.composer-box:focus-within {
  border-color: var(--color-brand-text);  /* 绿色边框 */
  box-shadow: none;
}
```

**第二处**：`styles.css:4907-4912` - 内层 textarea 聚焦态
```css
.sheet-search-row input:focus,
.composer textarea:focus {
  outline: none;
  border-color: var(--color-brand-text);  /* 又一个绿色边框 */
  box-shadow: none;
}
```

### 解决方案

修改 `styles.css:4907-4912`：

**改为**：
```css
.sheet-search-row input:focus {
  outline: none;
  border-color: var(--color-brand-text);
  box-shadow: none;
}

.composer textarea:focus {
  outline: none;
  border-color: transparent;  /* 保持透明，只显示外层边框 */
  box-shadow: none;
}
```

---

## 问题 2：完整的色彩系统 + 专属图案（高优先级）

### 色彩设计理念

**默认 Notion**：保持当前暖灰调，专业克制  
**温暖橙**：像日落余晖、焦糖拿铁，温暖明亮  
**宁静蓝**：像晴朗天空、清澈海洋，清新舒适  
**活力绿**：像春天嫩叶、清新薄荷，生机勃勃

### 完整的主题配置

**位置**：`apps/excel-addin/src/hooks/useTheme.ts:16-58`

**替换为**：

```typescript
interface PresetPalette {
  // 背景层次（4 层）
  background: string;
  backgroundSubtle: string;
  surface: string;
  surfaceHover: string;
  // 文字层次（3 层）
  textPrimary: string;
  textSecondary: string;
  textTertiary: string;
  // 品牌色（3 层）
  brandBg: string;
  brandText: string;
  brandHover: string;
  // 边框（2 层）
  border: string;
  borderStrong: string;
  // 阴影（2 层）
  shadowSm: string;
  shadowMd: string;
  // 专属图案（SVG data URI）
  pattern: string;
}

const PRESETS: Record<ThemePreset, PresetPalette> = {
  default: {
    // 默认 Notion 风格（保持不变）
    background: "#ffffff",
    backgroundSubtle: "#f7f6f3",
    surface: "#fafaf9",
    surfaceHover: "#f1f0ee",
    textPrimary: "#37352f",
    textSecondary: "#787774",
    textTertiary: "#9b9a97",
    brandBg: "#edf3ec",
    brandText: "#2f6b47",
    brandHover: "#e3ebe2",
    border: "#e9e9e7",
    borderStrong: "#d3d2cf",
    shadowSm: "0 1px 2px rgba(0, 0, 0, 0.04)",
    shadowMd: "0 2px 4px rgba(0, 0, 0, 0.06)",
    pattern: "none"
  },
  "warm-orange": {
    // 温暖橙 - 日落余晖 🌅
    background: "#fffcf7",
    backgroundSubtle: "#fff5e8",
    surface: "#ffefd9",
    surfaceHover: "#ffe6c7",
    textPrimary: "#5c3d2e",
    textSecondary: "#8b6f5c",
    textTertiary: "#b39a87",
    brandBg: "#ffe4c4",
    brandText: "#d97739",
    brandHover: "#ffd9b0",
    border: "#f5d9b8",
    borderStrong: "#e8c89f",
    shadowSm: "0 1px 2px rgba(217, 119, 57, 0.08)",
    shadowMd: "0 2px 4px rgba(217, 119, 57, 0.12)",
    // 小太阳图案（SVG）
    pattern: `url("data:image/svg+xml,%3Csvg width='60' height='60' xmlns='http://www.w3.org/2000/svg'%3E%3Ccircle cx='10' cy='10' r='3' fill='%23d97739' opacity='0.04'/%3E%3Cpath d='M10 5 L10 6 M10 14 L10 15 M5 10 L6 10 M14 10 L15 10 M7 7 L7.7 7.7 M12.3 12.3 L13 13 M13 7 L12.3 7.7 M7.7 12.3 L7 13' stroke='%23d97739' opacity='0.04' stroke-width='0.5'/%3E%3Ccircle cx='50' cy='50' r='3' fill='%23d97739' opacity='0.03'/%3E%3C/svg%3E")`
  },
  "calm-blue": {
    // 宁静蓝 - 晴朗天空 ☁️
    background: "#f9fbff",
    backgroundSubtle: "#f0f6ff",
    surface: "#e3f0ff",
    surfaceHover: "#d4e8ff",
    textPrimary: "#2c3e5c",
    textSecondary: "#5a6b8a",
    textTertiary: "#8a9bb5",
    brandBg: "#d4e8ff",
    brandText: "#3b7dd6",
    brandHover: "#c2deff",
    border: "#c7dff7",
    borderStrong: "#aecef0",
    shadowSm: "0 1px 2px rgba(59, 125, 214, 0.08)",
    shadowMd: "0 2px 4px rgba(59, 125, 214, 0.12)",
    // 小云朵图案（SVG）
    pattern: `url("data:image/svg+xml,%3Csvg width='80' height='60' xmlns='http://www.w3.org/2000/svg'%3E%3Cellipse cx='15' cy='20' rx='8' ry='5' fill='%233b7dd6' opacity='0.04'/%3E%3Cellipse cx='20' cy='18' rx='6' ry='4' fill='%233b7dd6' opacity='0.04'/%3E%3Cellipse cx='10' cy='21' rx='5' ry='3' fill='%233b7dd6' opacity='0.04'/%3E%3Cellipse cx='55' cy='45' rx='7' ry='4' fill='%233b7dd6' opacity='0.03'/%3E%3C/svg%3E")`
  },
  "vivid-green": {
    // 活力绿 - 春天嫩叶 🌱
    background: "#f7fffb",
    backgroundSubtle: "#edfff4",
    surface: "#dfffea",
    surfaceHover: "#ceffd9",
    textPrimary: "#1e4d3b",
    textSecondary: "#4a7562",
    textTertiary: "#7a9f8e",
    brandBg: "#c8f5d9",
    brandText: "#2d9b63",
    brandHover: "#b5f0ca",
    border: "#b8e8ca",
    borderStrong: "#9ddab3",
    shadowSm: "0 1px 2px rgba(45, 155, 99, 0.08)",
    shadowMd: "0 2px 4px rgba(45, 155, 99, 0.12)",
    // 小叶子图案（SVG）
    pattern: `url("data:image/svg+xml,%3Csvg width='60' height='60' xmlns='http://www.w3.org/2000/svg'%3E%3Cpath d='M12 18 Q15 12, 18 18 Q15 15, 12 18' fill='%232d9b63' opacity='0.04'/%3E%3Cpath d='M8 10 Q10 6, 12 10 Q10 8, 8 10' fill='%232d9b63' opacity='0.04'/%3E%3Cpath d='M48 48 Q50 44, 52 48 Q50 46, 48 48' fill='%232d9b63' opacity='0.03'/%3E%3C/svg%3E")`
  }
};
```

### 应用完整色彩系统和图案

**位置**：`apps/excel-addin/src/hooks/useTheme.ts:71-80`

**替换为**：

```typescript
function applyPalette(preset: ThemePreset): void {
  const palette = PRESETS[preset];
  const style = rootStyle();
  
  // 背景层次
  style.setProperty("--color-background", palette.background);
  style.setProperty("--color-background-subtle", palette.backgroundSubtle);
  style.setProperty("--color-surface", palette.surface);
  style.setProperty("--color-surface-hover", palette.surfaceHover);
  
  // 文字层次
  style.setProperty("--color-text-primary", palette.textPrimary);
  style.setProperty("--color-text-secondary", palette.textSecondary);
  style.setProperty("--color-text-tertiary", palette.textTertiary);
  
  // 品牌色
  style.setProperty("--color-brand-bg", palette.brandBg);
  style.setProperty("--color-brand-text", palette.brandText);
  style.setProperty("--color-brand-hover", palette.brandHover);
  
  // 边框
  style.setProperty("--color-border", palette.border);
  style.setProperty("--color-border-strong", palette.borderStrong);
  
  // 阴影
  style.setProperty("--shadow-sm", palette.shadowSm);
  style.setProperty("--shadow-md", palette.shadowMd);
  
  // 专属图案
  style.setProperty("--theme-pattern", palette.pattern);
}
```

---

## 问题 3：背景层级系统（高优先级）

### 层级结构（从底到顶）

```
z-index: -1 → 主题图案层（小图案背景）
z-index: 0  → 自定义背景图层（用户上传的图片）
z-index: 1  → 遮罩层（确保可读性）
z-index: 2  → 内容层（所有文字、按钮等）
```

### CSS 实现

**位置**：`apps/excel-addin/src/styles.css`，在主题彩蛋部分（约 4946 行）添加：

```css
/* 确认主容器名称（.chat-shell 或实际容器类名） */
.chat-shell {
  position: relative;
}

/* 图案层（最底层） */
.chat-shell::before {
  content: '';
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background-color: var(--color-background);
  background-image: var(--theme-pattern, none);
  background-size: 60px 60px;
  background-repeat: repeat;
  pointer-events: none;
  z-index: -1;
}

/* 自定义背景图层（用户上传） */
.chat-shell[data-has-background="true"]::before {
  background-image: var(--custom-background-image);
  background-size: cover;
  background-position: center;
  background-repeat: no-repeat;
  opacity: var(--custom-background-opacity, 0);
  z-index: 0;
}

/* 遮罩层（确保文字可读） */
.chat-shell[data-has-background="true"][data-auto-mask="true"]::after {
  content: '';
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: var(--color-background);
  opacity: 0.6;
  pointer-events: none;
  z-index: 1;
}

/* 内容层提升到最上层 */
.chat-shell > * {
  position: relative;
  z-index: 2;
}
```

**重要**：需要确认实际的主容器类名：
1. 在浏览器中打开插件
2. F12 打开开发者工具
3. 查看 DOM 结构，找到最外层对话容器的 class 名称
4. 如果不是 `.chat-shell`，将上述 CSS 中的 `.chat-shell` 替换为实际类名
5. 同步修改 `useTheme.ts:84, 88, 98` 中的选择器

---

## 问题 4：工具抽屉配色配套（高优先级）

### 需要替换的位置

搜索 `styles.css` 中以下选择器，将硬编码颜色替换为主题变量：

#### 1. 工具抽屉容器

```css
.tool-drawer {
  background: var(--color-background);  /* 替换 #fff */
  border: 1px solid var(--color-border);  /* 替换 #e0e0e0 */
  box-shadow: var(--shadow-md);  /* 替换固定阴影值 */
}
```

#### 2. 工具卡片

```css
.tool-card {
  background: var(--color-surface);  /* 替换 #f9f9f9 */
  border: 1px solid var(--color-border);  /* 替换 #ddd */
}

.tool-card:hover {
  background: var(--color-surface-hover);  /* 替换 #f0f0f0 */
}
```

#### 3. 工具字段和参数

```css
.tool-field,
.tool-parameters input,
.tool-parameters select {
  background: var(--color-surface);
  border: 1px solid var(--color-border);
  color: var(--color-text-primary);
}

.tool-parameters input:focus,
.tool-parameters select:focus {
  border-color: var(--color-brand-text);
}
```

#### 4. 工具按钮

```css
.tool-drawer .secondary-button {
  background: var(--color-surface);
  border: 1px solid var(--color-border);
  color: var(--color-text-primary);
}

.tool-drawer .secondary-button:hover {
  background: var(--color-surface-hover);
}

.tool-drawer .primary-button {
  background: var(--color-brand-text);
  color: #ffffff;
}

.tool-drawer .primary-button:hover {
  opacity: 0.9;
}
```

#### 5. 工具文字

```css
.tool-drawer h3,
.tool-card h4 {
  color: var(--color-text-primary);
}

.tool-description,
.tool-field-label {
  color: var(--color-text-secondary);
}

.tool-hint,
.tool-placeholder {
  color: var(--color-text-tertiary);
}
```

### 批量替换策略

**搜索正则表达式**：
- 背景色：`background:\s*#(fff|ffffff|f[0-9a-f]{5}|fafafa|f5f5f5)`
- 文字色：`color:\s*#(000|333|666|242424|3[0-9a-f]{5})`
- 边框色：`border.*#(e0e0e0|ddd|ccc|d[0-9a-f]{5})`

**替换规则**：
- 主背景 → `var(--color-background)`
- 次级背景/按钮背景 → `var(--color-surface)`
- 主文字 → `var(--color-text-primary)`
- 次文字 → `var(--color-text-secondary)`
- 边框 → `var(--color-border)`

---

## 实施顺序

### 第一步：修复输入框双边框

修改 `styles.css:4907-4912`，分离 `.composer textarea:focus`。

### 第二步：完善主题色彩系统

1. 修改 `useTheme.ts:16-58`，添加完整的 14 变量 + pattern
2. 修改 `useTheme.ts:71-80`，应用所有变量包括图案

### 第三步：添加背景层级系统

1. 确认主容器类名
2. 在 `styles.css` 添加四层背景样式（图案层、背景层、遮罩层、内容层）
3. 同步 JS 选择器

### 第四步：工具抽屉配色配套

搜索并替换工具相关样式，使用主题变量。

### 第五步：全量测试

```bash
npm run build:addin
npm run lint
npm run test:addin
```

**Excel 中测试**：

1. **输入框聚焦** → 只有一个绿框 ✅
2. **切换到温暖橙** → 整体橙色调，背景有小太阳图案 ✅
3. **切换到宁静蓝** → 整体蓝色调，背景有小云朵图案 ✅
4. **切换到活力绿** → 整体绿色调，背景有小叶子图案 ✅
5. **打开工具抽屉** → 颜色和主界面配套 ✅
6. **上传背景图片** → 图案被覆盖，背景图显示，文字清晰 ✅

---

## Git 提交信息

```
fix: 完善主题系统（明亮配色+专属图案+完整色彩变量+层级修复）
```

---

## 🎨 视觉效果预期

### 默认 Notion
- 暖灰色调，专业克制
- 无图案

### 温暖橙 🌅
- 整体橙色调，像日落余晖
- 背景有淡淡的小太阳图案
- 明亮温暖，不刺眼

### 宁静蓝 ☁️
- 整体蓝色调，像晴朗天空
- 背景有淡淡的小云朵图案
- 清新舒适，不刺眼

### 活力绿 🌱
- 整体绿色调，像春天嫩叶
- 背景有淡淡的小叶子图案
- 生机勃勃，不刺眼

**女朋友看到肯定会开心！** 💕✨
