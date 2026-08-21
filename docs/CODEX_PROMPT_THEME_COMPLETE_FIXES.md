# Codex 任务：修复主题面板问题（完整色彩系统 + 输入框双边框）

## 问题 1：输入框双边框问题（高优先级）

### 问题描述

用户反馈：点击输入框后，出现两个绿色边框，视觉上很突兀。

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
.composer textarea:focus {
  outline: none;
  border-color: var(--color-brand-text);  /* 又一个绿色边框 */
  box-shadow: none;
}
```

**第三处**：`styles.css:4914-4916` - textarea 默认样式
```css
.composer textarea {
  border: 1px solid transparent;
}
```

当点击输入框时：
- 外层 `.composer-box` 边框变绿
- 内层 `textarea` 边框从透明变绿
- 结果：两个绿框重叠，看起来像两条线

### 解决方案

**只保留外层容器的聚焦态，移除内层 textarea 的聚焦态边框。**

#### 修改位置 1：删除 textarea 的聚焦态边框

找到 `styles.css:4907-4912`：
```css
.sheet-search-row input:focus,
.composer textarea:focus {
  outline: none;
  border-color: var(--color-brand-text);
  box-shadow: none;
}
```

**改为**（只保留 `.sheet-search-row input:focus`，移除 `.composer textarea:focus`）：
```css
.sheet-search-row input:focus {
  outline: none;
  border-color: var(--color-brand-text);
  box-shadow: none;
}

.composer textarea:focus {
  outline: none;
  border-color: transparent;  /* 保持透明，不显示边框 */
  box-shadow: none;
}
```

#### 修改位置 2：确保 textarea 默认透明边框

`styles.css:4914-4916` 保持不变（已经是透明的）：
```css
.composer textarea {
  border: 1px solid transparent;
}
```

#### 修改位置 3：确保外层容器聚焦态正常

`styles.css:5411-5414` 保持不变（已经正确）：
```css
.composer-box:focus-within {
  border-color: var(--color-brand-text);
  box-shadow: none;
}
```

### 验收标准

- ✅ 点击输入框，只有外层容器边框变绿
- ✅ 不会出现双重边框
- ✅ 聚焦态清晰可见
- ✅ 符合 WCAG 可访问性标准

---

## 问题 2：主题色彩系统不完整（高优先级）

### 问题描述

用户反馈：切换主题后，只有部分元素变色，其他元素保持灰色，整体不协调，很难看。

### 问题原因

当前实现只改了 6 个 CSS 变量：
- `--color-background`
- `--color-background-subtle`
- `--color-text-primary`
- `--color-brand-bg`
- `--color-brand-text`
- `--color-brand-hover`

但 Notion 风格迁移后，CSS 里定义了更多变量：
- `--color-text-secondary`（次要文字）
- `--color-text-tertiary`（辅助文字）
- `--color-border`（边框）
- `--color-border-strong`（强调边框）
- `--color-surface`（悬浮元素背景）
- `--color-surface-hover`（悬停态背景）
- `--shadow-sm`、`--shadow-md`（阴影）

如果切换主题只改 6 个变量，其他变量保持默认值，会导致：
- 温暖橙主题：背景是米色，但按钮边框还是灰色 ❌
- 宁静蓝主题：背景是蓝色，但输入框边框还是灰色 ❌
- 视觉不协调，很难看

### 解决方案：完整的色彩系统

每个主题需要定义**完整的 13 个色彩变量 + 2 个阴影变量**，确保所有元素协调统一。

#### 色彩变量清单

```typescript
interface PresetPalette {
  // 背景层次（3 层）
  background: string;           // 主背景（纯白或接近白）
  backgroundSubtle: string;     // 次级背景（侧边栏、卡片）
  surface: string;              // 悬浮元素背景（按钮默认态）
  surfaceHover: string;         // 悬停态背景
  
  // 文字层次（3 层）
  textPrimary: string;          // 主要文字（标题、正文）
  textSecondary: string;        // 次要文字（描述、标签）
  textTertiary: string;         // 辅助文字（占位符、禁用）
  
  // 品牌色（3 层）
  brandBg: string;              // 品牌背景（标签、徽章）
  brandText: string;            // 品牌文字（链接、按钮）
  brandHover: string;           // 品牌悬停态
  
  // 边框（2 层）
  border: string;               // 默认边框
  borderStrong: string;         // 强调边框
  
  // 阴影（2 层，字符串格式）
  shadowSm: string;             // 小阴影
  shadowMd: string;             // 中阴影
}
```

#### 完整的主题配置

**位置**：`apps/excel-addin/src/hooks/useTheme.ts:16-58`

**替换为**：

```typescript
interface PresetPalette {
  background: string;
  backgroundSubtle: string;
  surface: string;
  surfaceHover: string;
  textPrimary: string;
  textSecondary: string;
  textTertiary: string;
  brandBg: string;
  brandText: string;
  brandHover: string;
  border: string;
  borderStrong: string;
  shadowSm: string;
  shadowMd: string;
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
    shadowMd: "0 2px 4px rgba(0, 0, 0, 0.06)"
  },
  "warm-orange": {
    // 温暖米色主题（对眼友好）
    background: "#faf9f7",
    backgroundSubtle: "#f5f3f0",
    surface: "#f0ede9",
    surfaceHover: "#ebe7e2",
    textPrimary: "#3d3730",
    textSecondary: "#6b6358",
    textTertiary: "#9a9186",
    brandBg: "#f0ebe5",
    brandText: "#7d6854",
    brandHover: "#e8e2db",
    border: "#e5e0db",
    borderStrong: "#d6cfc8",
    shadowSm: "0 1px 2px rgba(61, 55, 48, 0.04)",
    shadowMd: "0 2px 4px rgba(61, 55, 48, 0.06)"
  },
  "calm-blue": {
    // 石板灰蓝主题（对眼友好）
    background: "#f8f9fb",
    backgroundSubtle: "#f1f3f6",
    surface: "#ebeef2",
    surfaceHover: "#e3e7ec",
    textPrimary: "#2f3d4f",
    textSecondary: "#5a6678",
    textTertiary: "#8a95a5",
    brandBg: "#e8ecf1",
    brandText: "#5a6b7d",
    brandHover: "#dfe4ea",
    border: "#e3e8ed",
    borderStrong: "#d1d9e2",
    shadowSm: "0 1px 2px rgba(47, 61, 79, 0.04)",
    shadowMd: "0 2px 4px rgba(47, 61, 79, 0.06)"
  },
  "vivid-green": {
    // 鼠尾草绿主题（对眼友好）
    background: "#f8faf9",
    backgroundSubtle: "#f2f5f3",
    surface: "#ecf0ed",
    surfaceHover: "#e4e9e6",
    textPrimary: "#2f3d35",
    textSecondary: "#5a6a60",
    textTertiary: "#8a9890",
    brandBg: "#e9efec",
    brandText: "#5a7065",
    brandHover: "#dfe6e2",
    border: "#e4ebe7",
    borderStrong: "#d3ddd7",
    shadowSm: "0 1px 2px rgba(47, 61, 53, 0.04)",
    shadowMd: "0 2px 4px rgba(47, 61, 53, 0.06)"
  }
};
```

#### 应用完整色彩系统

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
}
```

### 色彩设计理念

**所有主题都采用低饱和度、灰调色彩，长时间看不累眼：**

1. **默认 Notion**：暖灰调，克制、专业
2. **温暖米色**：像纸张、羊皮纸，温暖但不刺眼
3. **石板灰蓝**：像阴天的天空，平静柔和
4. **鼠尾草绿**：像植物叶子的灰绿色，自然舒适

### 验收标准

- ✅ 切换主题后，所有元素（背景、文字、按钮、边框、阴影）都协调统一
- ✅ 温暖米色主题：整体米色调，没有灰色元素突兀
- ✅ 石板灰蓝主题：整体蓝灰调，没有米色元素突兀
- ✅ 鼠尾草绿主题：整体绿灰调，没有蓝色元素突兀
- ✅ 色彩柔和不刺眼，长时间使用不累眼
- ✅ 文字对比度符合 WCAG AA 标准（≥ 4.5:1）

---

## 问题 3：背景图片层级错误，遮挡文字（高优先级）

用户反馈：上传背景图片后，图片在文字上方，遮挡了内容。

### 问题原因

背景图层（`::before`）和遮罩层（`::after`）的层级设置正确，但**内容层的子元素没有设置 `position: relative` 和 `z-index`**，导致内容没有被提升到最上层。

### 正确的层级结构

```
z-index: 0  → 背景图层（::before）
z-index: 1  → 遮罩层（::after）
z-index: 2  → 所有内容（.chat-shell > *）
```

### 解决方案

#### 修改位置 1：确保容器使用相对定位

搜索 `styles.css` 中的 `.chat-shell` 或主容器，确保有：
```css
.chat-shell {
  position: relative;
}
```

#### 修改位置 2：确保背景层和遮罩层样式存在

在 `styles.css` 的主题彩蛋部分（约 4946 行附近）添加：

```css
/* 背景图层 */
.chat-shell::before {
  content: '';
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background-image: var(--custom-background-image);
  background-size: cover;
  background-position: center;
  background-repeat: no-repeat;
  opacity: var(--custom-background-opacity, 0);
  pointer-events: none;
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

#### 修改位置 3：确认容器选择器

当前代码使用 `.chat-shell` 作为背景容器：
```typescript
// useTheme.ts:84
const container = document.querySelector<HTMLElement>(".chat-shell");
```

**需要确认**：实际的主容器类名是 `.chat-shell` 还是其他名称（如 `.app-container`、`.conversation-container`）？

**操作**：
- 在浏览器中打开插件
- F12 打开开发者工具
- 查看 DOM 结构，找到最外层容器的 class 名称
- 如果不是 `.chat-shell`，需要同步修改：
  - `useTheme.ts:84, 88, 98` 中的选择器
  - `styles.css` 中的背景层样式选择器

### 验收标准

- ✅ 上传背景图片后，文字清晰可见，不被遮挡
- ✅ 背景图在最底层，遮罩层在中间，内容在最上层
- ✅ 所有交互元素（按钮、输入框、列表）都可正常点击

---

## 问题 4：工具抽屉配色不配套（高优先级）

用户反馈：切换主题后，工具抽屉（Tool Drawer）的颜色没有跟着变，和整体风格不匹配。

### 问题原因

工具抽屉是独立的浮层组件，有自己的背景色、边框色、按钮样式等。当前实现只更新了主容器的 CSS 变量，但工具抽屉内部的样式没有使用这些变量，或者使用了硬编码的颜色值。

### 涉及的组件

需要确保以下组件都使用主题变量：

1. **工具抽屉本身**
   - 背景色：`var(--color-background)`
   - 边框色：`var(--color-border)`
   - 阴影：`var(--shadow-md)`

2. **工具卡片**
   - 背景色：`var(--color-surface)`
   - 悬停态：`var(--color-surface-hover)`
   - 边框色：`var(--color-border)`

3. **工具抽屉内的按钮**
   - 默认背景：`var(--color-surface)`
   - 悬停背景：`var(--color-surface-hover)`
   - 主按钮背景：`var(--color-brand-text)`
   - 主按钮文字：`#ffffff`

4. **工具抽屉内的输入框**
   - 背景色：`var(--color-surface)`
   - 边框色：`var(--color-border)`
   - 聚焦边框：`var(--color-brand-text)`

5. **工具抽屉内的文字**
   - 标题：`var(--color-text-primary)`
   - 描述：`var(--color-text-secondary)`
   - 占位符：`var(--color-text-tertiary)`

### 解决方案

#### 排查步骤

1. **搜索工具抽屉相关样式**

   在 `styles.css` 中搜索：
   - `.tool-drawer`
   - `.tool-card`
   - `.tool-field`
   - `.tool-parameters`
   - `.tool-guide`

2. **检查是否使用硬编码颜色**

   查找这些样式中是否有：
   - `background: #fff` 或 `background: #f5f5f5`
   - `color: #333` 或 `color: #666`
   - `border: 1px solid #ddd`

3. **替换为主题变量**

   将硬编码颜色替换为：
   - `background: var(--color-surface)`
   - `color: var(--color-text-primary)`
   - `border: 1px solid var(--color-border)`

#### 示例修改

**修改前**（硬编码）：
```css
.tool-drawer {
  background: #ffffff;
  border: 1px solid #e0e0e0;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

.tool-card {
  background: #f9f9f9;
  border: 1px solid #ddd;
}

.tool-card:hover {
  background: #f0f0f0;
}
```

**修改后**（使用主题变量）：
```css
.tool-drawer {
  background: var(--color-background);
  border: 1px solid var(--color-border);
  box-shadow: var(--shadow-md);
}

.tool-card {
  background: var(--color-surface);
  border: 1px solid var(--color-border);
}

.tool-card:hover {
  background: var(--color-surface-hover);
}
```

#### 需要检查的具体位置

搜索 `styles.css` 中以下选择器，确保使用主题变量：

1. `.tool-drawer` - 抽屉容器
2. `.tool-card` - 工具卡片
3. `.tool-field` - 工具字段
4. `.tool-parameters` - 参数面板
5. `.tool-guide` - 工具指引
6. `.tool-dialog` - 工具对话框
7. `.secondary-button` - 次要按钮（在工具抽屉中）
8. `.primary-button` - 主要按钮（在工具抽屉中）

### 批量替换策略

可以使用正则表达式批量替换：

**替换背景色**：
- 查找：`background:\s*#(ffffff|fff|f[0-9a-f]{5}|fafafa|f5f5f5)`
- 根据上下文替换为：
  - 主背景 → `var(--color-background)`
  - 次级背景 → `var(--color-surface)`

**替换文字色**：
- 查找：`color:\s*#(000000|333|666|242424)`
- 根据上下文替换为：
  - 主文字 → `var(--color-text-primary)`
  - 次文字 → `var(--color-text-secondary)`

**替换边框色**：
- 查找：`border.*#(e0e0e0|ddd|ccc|d3d2cf)`
- 替换为：`border: ... var(--color-border)`

### 验收标准

- ✅ 切换到温暖米色主题，工具抽屉整体呈米色调
- ✅ 切换到石板灰蓝主题，工具抽屉整体呈蓝灰调
- ✅ 切换到鼠尾草绿主题，工具抽屉整体呈绿灰调
- ✅ 工具抽屉内的按钮、输入框、卡片都和主界面风格一致
- ✅ 没有灰色、白色等不协调的元素突兀地出现

---

## 问题 5：图片上传功能检查（中优先级）

用户反馈：有主题切换，但没有图片上传，应用不了。

### 排查步骤

按照问题 3 的修复完成后，测试图片上传：

1. **添加调试日志（临时）**

   在 `useTheme.ts:166` 的 `uploadBackground` 函数开头添加：
   ```typescript
   console.log('[Theme] 上传开始:', file.name, file.size, file.type);
   ```

   在每个关键步骤后添加日志，追踪上传流程。

2. **测试上传**
   - 长按 Logo 打开主题面板
   - 点击"上传图片"
   - 选择一张图片（< 10MB）
   - 查看浏览器控制台日志
   - 查看是否有错误信息

3. **如果上传成功但背景不显示**
   - 按照问题 3 的方案，检查层级设置
   - 检查容器选择器是否正确
   - 检查 CSS 背景样式是否存在

---

## 实施顺序

### 第一步：修复输入框双边框（最高优先级）

修改 `apps/excel-addin/src/styles.css:4907-4912`，分离 `.composer textarea:focus` 的样式，设置 `border-color: transparent`。

### 第二步：完善主题色彩系统（高优先级）

1. 修改 `apps/excel-addin/src/hooks/useTheme.ts:16-58`，添加完整的 14 个色彩变量
2. 修改 `apps/excel-addin/src/hooks/useTheme.ts:71-80`，应用完整的色彩变量

### 第三步：检查图片上传功能（中优先级）

1. 确认容器选择器（`.chat-shell` 或 `.app-container`）
2. 确认 CSS 背景样式存在
3. 添加调试日志
4. 测试上传流程

### 第四步：测试

```bash
npm run build:addin
npm run lint
npm run test:addin
```

在 Excel 中测试：
1. 点击输入框 → 只有外层边框变绿，无双重边框
2. 长按 Logo → 打开主题面板
3. 切换预设主题 → 所有元素协调统一，色彩柔和
4. 上传图片 → 背景显示，透明度可调

---

## Git 提交信息

```
fix: 修复输入框双边框问题并完善主题色彩系统为完整的 14 变量配置
```

---

## 注意事项

1. **输入框双边框**：只保留外层容器聚焦态，移除内层 textarea 的边框变色
2. **色彩协调**：每个主题必须定义完整的 14 个变量，确保所有元素统一
3. **对眼友好**：所有主题采用低饱和度、灰调色彩
4. **容器选择器**：确认实际容器类名，确保 JS 和 CSS 使用相同的选择器
