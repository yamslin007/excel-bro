# Codex 调试任务：格仔点击 + 背景图片问题

## 用户反馈

1. **背景图片看不到了**：上传图片后，背景不显示
2. **格仔点击没反应**：点击格仔后，没有转身/眨眼/点头动画

---

## 问题 1：背景图片不显示

### 代码检查结果

✅ CSS 样式存在（styles.css:4999-5032）  
✅ CSS 变量定义存在（styles.css:4953-4954）  
✅ useTheme hook 存在

### 可能原因

#### 原因 1：层级冲突（最可能）

Codex 添加的主题图案层可能和背景图层冲突了。

**当前层级**：
```css
.chat-shell::before {
  z-index: -1;  /* 图案层 */
  background-image: var(--theme-pattern, none);
}

.chat-shell[data-has-background="true"]::before {
  background-image: var(--custom-background-image);  /* 覆盖图案 */
  z-index: 0;  /* 提升到 0 */
}
```

**问题**：`::before` 同一个伪元素，后面的样式会覆盖前面的，但是 `background-image` 和 `z-index` 同时存在时可能冲突。

**解决方案**：分离图案层和背景图层，使用不同的伪元素。

**修改位置**：`styles.css:4986-5032`

**改为**：

```css
.chat-shell {
  position: relative;
  overflow: hidden;
}

/* 图案层（::before，最底层） */
.chat-shell::before {
  content: "";
  position: absolute;
  inset: 0;
  z-index: -1;
  pointer-events: none;
  background-color: var(--color-background);
  background-image: var(--theme-pattern, none);
  background-position: center;
  background-repeat: repeat;
  background-size: 60px 60px;
}

/* 隐藏图案层当有自定义背景时 */
.chat-shell[data-has-background="true"]::before {
  display: none;
}

/* 自定义背景图层（新增一个专门的 div） */
.chat-shell::after {
  content: "";
  position: absolute;
  inset: 0;
  z-index: 0;
  pointer-events: none;
  background-image: var(--custom-background-image);
  background-position: center;
  background-repeat: no-repeat;
  background-size: cover;
  opacity: 0;
  transition: opacity 0.3s ease;
}

.chat-shell[data-has-background="true"]::after {
  opacity: var(--custom-background-opacity);
}

/* 遮罩层（需要新增一个 div，因为伪元素只有两个） */
/* 这个需要在 JSX 中添加 <div className="chat-shell-mask"></div> */
.chat-shell-mask {
  position: absolute;
  inset: 0;
  z-index: 1;
  pointer-events: none;
  background: var(--color-background);
  opacity: 0;
  transition: opacity 0.3s ease;
}

.chat-shell[data-has-background="true"][data-auto-mask="true"] .chat-shell-mask {
  opacity: 0.6;
}

/* 内容层提升到最上层 */
.chat-shell > *:not(.chat-shell-mask) {
  position: relative;
  z-index: 2;
}
```

**但这需要修改 JSX！** 因为伪元素只有 `::before` 和 `::after` 两个，而我们需要三层（图案、背景、遮罩）。

#### 更简单的解决方案：合并图案和背景

既然图案是主题装饰，背景是用户自定义，可以让背景覆盖图案。

**修改位置**：`styles.css:4986-5032`

**改为**：

```css
.chat-shell {
  position: relative;
  overflow: hidden;
}

/* 背景层（::before） - 默认显示图案，上传图片后显示图片 */
.chat-shell::before {
  content: "";
  position: absolute;
  inset: 0;
  z-index: 0;
  pointer-events: none;
  background-color: var(--color-background);
  background-image: var(--theme-pattern, none);
  background-position: center;
  background-repeat: repeat;
  background-size: 60px 60px;
  transition: opacity 0.3s ease;
}

/* 有自定义背景时，切换为背景图片 */
.chat-shell[data-has-background="true"]::before {
  background-image: var(--custom-background-image);
  background-position: center;
  background-repeat: no-repeat;
  background-size: cover;
  opacity: var(--custom-background-opacity);
}

/* 遮罩层（::after） */
.chat-shell::after {
  content: "";
  position: absolute;
  inset: 0;
  z-index: 1;
  pointer-events: none;
  background: var(--color-background);
  opacity: 0;
  transition: opacity 0.3s ease;
}

.chat-shell[data-has-background="true"][data-auto-mask="true"]::after {
  opacity: 0.6;
}

/* 内容层提升到最上层 */
.chat-shell > * {
  position: relative;
  z-index: 2;
}
```

这样：
- **无自定义背景**：`::before` 显示图案，`::after` 透明
- **有自定义背景**：`::before` 显示背景图（覆盖图案），`::after` 显示遮罩

#### 原因 2：useTheme hook 没有正确设置 data 属性

检查 `useTheme.ts:84-98` 是否正确设置了 `data-has-background` 和 `data-auto-mask`。

**当前代码**：
```typescript
const container = document.querySelector<HTMLElement>(".chat-shell");
if (container) {
  container.setAttribute("data-has-background", "true");
  if (autoMask) {
    container.setAttribute("data-auto-mask", "true");
  }
}
```

**检查**：
1. 容器类名是否正确（`.chat-shell`）
2. 是否在正确的时机调用
3. 浏览器中检查元素是否有这些 data 属性

#### 原因 3：CSS 变量没有正确设置

检查 `useTheme.ts:166-193` 中的 `uploadBackground` 函数，确认是否正确设置了 `--custom-background-image`。

**当前代码**：
```typescript
const dataUrl = await blobToDataURL(blob);
const style = rootStyle();
style.setProperty("--custom-background-image", `url(${dataUrl})`);
style.setProperty("--custom-background-opacity", String(opacity / 100));
```

**检查**：
1. `dataUrl` 是否正确生成
2. CSS 变量是否正确设置到 `:root`
3. 浏览器开发者工具中检查 CSS 变量值

---

## 问题 2：格仔点击没反应

### 代码检查结果

✅ CSS 动画存在（styles.css:5306-5385）  
✅ 事件监听存在（App.tsx:1028-1056）  
✅ animatePet 函数存在（App.tsx:129-136）

### 可能原因

#### 原因 1：pet-avatar 元素不存在或被隐藏

**检查**：
1. 浏览器中检查 DOM，确认 `.pet-avatar` 元素存在
2. 确认 `petVisible` 状态为 `true`

#### 原因 2：z-index 层级问题

格仔可能被其他元素遮挡，导致点击事件无法触发。

**检查**：
1. 浏览器中检查 `.pet-avatar` 的 z-index
2. 确认没有其他元素覆盖在格仔上方

**修复**：
在 `styles.css` 中搜索 `.pet-avatar`，确保有足够高的 z-index：

```css
.pet-avatar {
  z-index: 100;  /* 确保在最上层 */
  cursor: pointer;  /* 显示可点击 */
}
```

#### 原因 3：事件监听没有正确绑定

**检查**：
1. 在 `handlePetClick` 函数开头添加 `console.log('Pet clicked!', target);`
2. 点击格仔，查看控制台是否有日志
3. 如果没有日志，说明事件没有触发
4. 如果有日志但没有动画，说明 `animatePet` 函数有问题

**调试代码**（临时添加）：

**位置**：`App.tsx:1028`

**改为**：
```typescript
const handlePetClick = (event: MouseEvent) => {
  const target = event.target as HTMLElement | null;
  console.log('[Pet] Click detected:', target);
  
  if (!target?.closest(".pet-avatar")) {
    console.log('[Pet] Click not on pet avatar');
    return;
  }
  
  const animation =
    PET_ANIMATIONS[Math.floor(Math.random() * PET_ANIMATIONS.length)];
  console.log('[Pet] Animating with:', animation);
  animatePet(animation);
};
```

**位置**：`App.tsx:129`

**改为**：
```typescript
function animatePet(className: string, duration = 800): void {
  const pet = document.querySelector<HTMLElement>(".pet-avatar");
  console.log('[Pet] animatePet called:', className, pet);
  
  if (!pet) {
    console.error('[Pet] pet-avatar element not found!');
    return;
  }
  
  pet.classList.remove(...PET_ANIMATIONS, "turtle-sleepy", "turtle-encourage");
  pet.classList.add(className);
  console.log('[Pet] Animation class added:', pet.className);
  
  window.setTimeout(() => {
    pet.classList.remove(className);
    console.log('[Pet] Animation class removed');
  }, duration);
}
```

#### 原因 4：CSS 动画被禁用

如果用户系统设置了"减少动画"，动画会被禁用。

**检查**：
在浏览器中执行：
```javascript
window.matchMedia('(prefers-reduced-motion: reduce)').matches
```

如果返回 `true`，说明动画被禁用。

**解决方案**：
让用户临时关闭系统的"减少动画"设置，或者移除 `styles.css:5387-5399` 中的 `@media (prefers-reduced-motion: reduce)` 规则。

---

## 实施步骤

### 第一步：修复背景图层级冲突

修改 `styles.css:4986-5032`，使用更简单的方案（背景图覆盖图案）。

### 第二步：检查 pet-avatar 样式

在 `styles.css` 中搜索 `.pet-avatar`，确保有：
```css
.pet-avatar {
  z-index: 100;
  cursor: pointer;
  pointer-events: auto;
}
```

### 第三步：添加调试日志

在 `App.tsx` 中添加 `console.log`，追踪格仔点击事件。

### 第四步：测试

```bash
npm run build:addin
```

**Excel 中测试**：

1. **背景图片**：
   - 长按 Logo 打开主题面板
   - 上传一张图片
   - 查看控制台日志
   - 检查元素是否有 `data-has-background="true"`
   - 检查 CSS 变量 `--custom-background-image` 是否有值
   - 确认背景图显示

2. **格仔点击**：
   - 点击格仔
   - 查看控制台日志
   - 确认事件触发
   - 确认动画播放

### 第五步：移除调试日志

测试完成后，移除所有 `console.log`。

---

## 验收标准

✅ 上传背景图片后，背景显示，图案消失  
✅ 调节透明度滑块，背景透明度实时变化  
✅ 开启自动遮罩，遮罩层显示，文字清晰  
✅ 点击格仔，随机播放转身/眨眼/点头/挥手动画  
✅ 60 秒无操作，格仔进入睡眠状态  
✅ 任意交互后，格仔从睡眠状态恢复

---

## Git 提交信息

```
fix: 修复背景图片层级冲突和格仔点击动画
```
