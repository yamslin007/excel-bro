# 主题定制彩蛋功能设计方案

> **核心概念**：长按顶部 Logo 2 秒触发隐藏的主题定制面板，支持预设主题和自定义背景上传（10MB 以内），采用毛玻璃（glassmorphism）视觉风格。

---

## 📋 功能概述

### 触发方式

**长按 "Excel Bro" Logo 文字 2 秒 → 触发主题面板**

```
┌─────────────────────────┐
│ 🐢 Excel Bro  [≡]      │  ← 长按这里 2 秒
└─────────────────────────┘
```

**交互流程**：
1. 用户长按 Logo
2. 长按 1 秒时：Logo 开始轻微发光（提示用户继续按）
3. 长按 2 秒时：触发成功，Logo 闪烁一下，主题面板从下方滑入
4. 如果中途松开：发光消失，不触发

**为什么是 2 秒**：
- 避免误触（正常点击 < 0.5 秒）
- 足够短，不会让用户失去耐心
- 真正的"彩蛋"感（需要主动探索）

---

## 🎨 主题面板设计

### 面板结构（毛玻璃风格）

```
┌───────────────────────────────┐
│  🎨 主题定制                   │
│  ─────────────────────────    │
│                                │
│  【预设主题】                  │
│   ○ 默认 Notion               │
│   ○ 温暖橙                     │
│   ○ 宁静蓝                     │
│   ○ 活力绿                     │
│                                │
│  【自定义背景】                │
│   📁 上传图片                  │
│   （支持 jpg/png/webp，<10MB）│
│                                │
│   🎨 背景透明度                │
│   [▓▓▓▓░░░░░░] 40%            │
│                                │
│   ✓ 自动遮罩（确保可读性）     │
│                                │
│  【预览效果】                  │
│   [消息气泡示例]               │
│   [按钮示例]                   │
│                                │
│   [应用主题]  [重置]  [删除背景] │
└───────────────────────────────┘
```

### 毛玻璃效果规范

```css
.theme-panel {
  position: fixed;
  bottom: 0;
  left: 0;
  right: 0;
  height: 70vh;
  max-height: 600px;
  
  /* 毛玻璃核心 */
  background: rgba(255, 255, 255, 0.85);
  backdrop-filter: blur(20px) saturate(180%);
  -webkit-backdrop-filter: blur(20px) saturate(180%);
  
  /* 边框和阴影 */
  border-top: 1px solid rgba(255, 255, 255, 0.3);
  box-shadow: 0 -8px 32px rgba(0, 0, 0, 0.1);
  
  /* 圆角（顶部） */
  border-radius: var(--radius-lg) var(--radius-lg) 0 0;
  
  /* 动画 */
  transform: translateY(100%);
  transition: transform 300ms cubic-bezier(0.4, 0, 0.2, 1);
}

.theme-panel.open {
  transform: translateY(0);
}

/* 不支持 backdrop-filter 时的降级 */
@supports not (backdrop-filter: blur(20px)) {
  .theme-panel {
    background: rgba(255, 255, 255, 0.95);
  }
}
```

---

## 🎨 预设主题规范

### 主题 1：默认 Notion（当前样式）

```css
--color-background: #FFFFFF;
--color-background-subtle: #F7F6F3;
--color-text-primary: #37352F;
--custom-background-image: none;
--custom-background-opacity: 0;
```

### 主题 2：温暖橙

```css
--color-background: #FFFBF5;
--color-background-subtle: #FFF4E6;
--color-text-primary: #4A3520;
--color-brand-bg: #FFE8CC;
--color-brand-text: #C87A3A;
/* 可选：背景渐变 */
--custom-background-image: linear-gradient(135deg, #FFF5E6 0%, #FFE8CC 100%);
--custom-background-opacity: 0.3;
```

### 主题 3：宁静蓝

```css
--color-background: #F8FBFF;
--color-background-subtle: #EFF6FF;
--color-text-primary: #1E3A5F;
--color-brand-bg: #DBEAFE;
--color-brand-text: #3B82F6;
/* 可选：背景渐变 */
--custom-background-image: linear-gradient(135deg, #EFF6FF 0%, #DBEAFE 100%);
--custom-background-opacity: 0.3;
```

### 主题 4：活力绿（当前品牌色增强）

```css
--color-background: #F7FFF7;
--color-background-subtle: #ECFDF5;
--color-text-primary: #1A4D2E;
--color-brand-bg: #D1FAE5;
--color-brand-text: #059669;
/* 可选：背景渐变 */
--custom-background-image: linear-gradient(135deg, #ECFDF5 0%, #D1FAE5 100%);
--custom-background-opacity: 0.3;
```

---

## 📁 自定义背景上传（重点）

### 存储方案：IndexedDB（支持 10MB 图片）

**为什么不用 localStorage**：
- localStorage 限制 5-10MB
- 图片转 base64 后体积增加 33%
- 10MB 图片 → 13.3MB base64，会超出限制

**IndexedDB 优势**：
- 存储容量大（Chrome 通常几百 MB）
- 直接存 Blob 二进制，不需要 base64 转换
- 性能更好

**存储架构**：
```
IndexedDB: excel_bro_theme (版本 1)
├─ objectStore: backgrounds
│  └─ key: 'current'
│     value: Blob (JPEG 二进制数据)
└─ objectStore: settings
   └─ key: 'theme_config'
      value: {
        preset: string,
        opacity: number,
        autoMask: boolean
      }
```

### 图片处理策略（保持高质量）

**原则**：
- ✅ 接受 < 10MB 的 jpg/png/webp
- ✅ 保持原始分辨率（不缩放）
- ✅ 转换为 JPEG 质量 90%（平衡质量和体积）
- ✅ 高 DPI 屏幕（Retina）也清晰

**处理流程**：
```typescript
async function processAndSaveImage(file: File): Promise<void> {
  // 1. 校验文件大小
  if (file.size > 10 * 1024 * 1024) {
    throw new Error('图片文件过大，请选择小于 10MB 的图片');
  }

  // 2. 读取图片
  const blob = await convertToJPEG(file, 0.90);
  
  // 3. 存储到 IndexedDB
  await saveBackgroundImage(blob);
  
  // 4. 应用到界面
  const url = URL.createObjectURL(blob);
  applyBackgroundImage(url);
}

async function convertToJPEG(file: File, quality: number): Promise<Blob> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    const reader = new FileReader();
    
    reader.onload = (e) => {
      img.onload = () => {
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d')!;
        
        // 保持原始尺寸
        canvas.width = img.width;
        canvas.height = img.height;
        
        // 绘制图片
        ctx.drawImage(img, 0, 0);
        
        // 转为 JPEG Blob（质量 90%）
        canvas.toBlob(
          (blob) => {
            if (blob) {
              resolve(blob);
            } else {
              reject(new Error('图片处理失败'));
            }
          },
          'image/jpeg',
          quality
        );
      };
      
      img.onerror = () => reject(new Error('图片加载失败'));
      img.src = e.target!.result as string;
    };
    
    reader.onerror = () => reject(new Error('文件读取失败'));
    reader.readAsDataURL(file);
  });
}
```

### IndexedDB 工具封装

```typescript
// utils/imageStorage.ts

const DB_NAME = 'excel_bro_theme';
const DB_VERSION = 1;
const STORE_BACKGROUNDS = 'backgrounds';
const STORE_SETTINGS = 'settings';

// 打开数据库
async function openDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    
    request.onerror = () => reject(request.error);
    request.onsuccess = () => resolve(request.result);
    
    request.onupgradeneeded = (event) => {
      const db = (event.target as IDBOpenDBRequest).result;
      
      // 创建背景图片存储
      if (!db.objectStoreNames.contains(STORE_BACKGROUNDS)) {
        db.createObjectStore(STORE_BACKGROUNDS);
      }
      
      // 创建设置存储
      if (!db.objectStoreNames.contains(STORE_SETTINGS)) {
        db.createObjectStore(STORE_SETTINGS);
      }
    };
  });
}

// 保存背景图片
export async function saveBackgroundImage(blob: Blob): Promise<void> {
  const db = await openDB();
  const tx = db.transaction(STORE_BACKGROUNDS, 'readwrite');
  const store = tx.objectStore(STORE_BACKGROUNDS);
  
  return new Promise((resolve, reject) => {
    const request = store.put(blob, 'current');
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

// 加载背景图片
export async function loadBackgroundImage(): Promise<string | null> {
  const db = await openDB();
  const tx = db.transaction(STORE_BACKGROUNDS, 'readonly');
  const store = tx.objectStore(STORE_BACKGROUNDS);
  
  return new Promise((resolve, reject) => {
    const request = store.get('current');
    
    request.onsuccess = () => {
      const blob = request.result as Blob | undefined;
      if (blob) {
        // 创建临时 URL（记得在不用时释放）
        const url = URL.createObjectURL(blob);
        resolve(url);
      } else {
        resolve(null);
      }
    };
    
    request.onerror = () => reject(request.error);
  });
}

// 删除背景图片
export async function clearBackgroundImage(): Promise<void> {
  const db = await openDB();
  const tx = db.transaction(STORE_BACKGROUNDS, 'readwrite');
  const store = tx.objectStore(STORE_BACKGROUNDS);
  
  return new Promise((resolve, reject) => {
    const request = store.delete('current');
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

// 保存主题设置
export async function saveThemeSettings(settings: ThemeSettings): Promise<void> {
  const db = await openDB();
  const tx = db.transaction(STORE_SETTINGS, 'readwrite');
  const store = tx.objectStore(STORE_SETTINGS);
  
  return new Promise((resolve, reject) => {
    const request = store.put(settings, 'theme_config');
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

// 加载主题设置
export async function loadThemeSettings(): Promise<ThemeSettings | null> {
  const db = await openDB();
  const tx = db.transaction(STORE_SETTINGS, 'readonly');
  const store = tx.objectStore(STORE_SETTINGS);
  
  return new Promise((resolve, reject) => {
    const request = store.get('theme_config');
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
}

// 类型定义
export interface ThemeSettings {
  preset: 'default' | 'warm-orange' | 'calm-blue' | 'vivid-green';
  opacity: number;  // 0-100
  autoMask: boolean;
  hasCustomBackground: boolean;
}
```

### 背景应用方式

**CSS 实现**（在 `<body>` 或 `.app-container` 上应用）：

```css
.app-container {
  position: relative;
  background: var(--color-background);
}

/* 自定义背景层 */
.app-container::before {
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

/* 遮罩层（确保可读性） */
.app-container::after {
  content: '';
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: var(--color-background);
  opacity: 0.6;  /* 60% 白色遮罩 */
  pointer-events: none;
  z-index: 1;
  display: var(--custom-background-mask-display, none);
}

/* 当有自定义背景且启用自动遮罩时显示 */
.app-container[data-has-background="true"][data-auto-mask="true"]::after {
  display: block;
}

/* 内容层 */
.app-container > * {
  position: relative;
  z-index: 2;
}
```

**JavaScript 应用**：

```typescript
async function applyBackgroundImage(url: string, opacity: number) {
  // 设置 CSS 变量
  document.documentElement.style.setProperty(
    '--custom-background-image',
    `url(${url})`
  );
  document.documentElement.style.setProperty(
    '--custom-background-opacity',
    (opacity / 100).toString()
  );
  
  // 设置容器属性
  const container = document.querySelector('.app-container');
  if (container) {
    container.setAttribute('data-has-background', 'true');
    container.setAttribute('data-auto-mask', 'true');
  }
}

async function clearBackground() {
  document.documentElement.style.setProperty('--custom-background-image', 'none');
  document.documentElement.style.setProperty('--custom-background-opacity', '0');
  
  const container = document.querySelector('.app-container');
  if (container) {
    container.removeAttribute('data-has-background');
    container.removeAttribute('data-auto-mask');
  }
  
  await clearBackgroundImage();
}
```

### 透明度调整

```tsx
<div className="opacity-control">
  <label>背景透明度</label>
  <input
    type="range"
    min="0"
    max="100"
    value={backgroundOpacity}
    onChange={(e) => {
      const opacity = Number(e.target.value);
      setBackgroundOpacity(opacity);
      document.documentElement.style.setProperty(
        '--custom-background-opacity',
        (opacity / 100).toString()
      );
    }}
  />
  <span>{backgroundOpacity}%</span>
</div>
```

---

## 🐢 格仔互动保持独立

**格仔的互动功能**（与主题功能完全独立）：

### 互动 1：点击格仔 → 随机动画

```typescript
const animations = [
  'turtle-spin',      // 转身
  'turtle-blink',     // 眨眼
  'turtle-nod',       // 点头
  'turtle-wave',      // 招手
];

function handlePetClick() {
  const randomAnim = animations[Math.floor(Math.random() * animations.length)];
  const petElement = document.querySelector('.pet-avatar');
  
  petElement?.classList.add(randomAnim);
  
  setTimeout(() => {
    petElement?.classList.remove(randomAnim);
  }, 800);
}
```

### 互动 2：空闲状态 → 打瞌睡

```typescript
let idleTimer: NodeJS.Timeout;

function resetIdleTimer() {
  clearTimeout(idleTimer);
  const petElement = document.querySelector('.pet-avatar');
  petElement?.classList.remove('turtle-sleepy');
  
  idleTimer = setTimeout(() => {
    petElement?.classList.add('turtle-sleepy');
  }, 60000); // 1 分钟无操作
}

// 监听用户活动
['click', 'keydown', 'scroll'].forEach(event => {
  document.addEventListener(event, resetIdleTimer);
});

// 初始化
resetIdleTimer();
```

### 互动 3：发送消息后 → 点头鼓励

```typescript
function onMessageSent() {
  const petElement = document.querySelector('.pet-avatar');
  petElement?.classList.add('turtle-encourage');
  
  setTimeout(() => {
    petElement?.classList.remove('turtle-encourage');
  }, 1000);
}
```

**CSS 动画示例**：

```css
/* 转身 */
@keyframes turtle-spin {
  0%, 100% { transform: rotateY(0deg); }
  50% { transform: rotateY(180deg); }
}

/* 眨眼 */
@keyframes turtle-blink {
  0%, 90%, 100% { opacity: 1; }
  95% { opacity: 0.3; }
}

/* 点头 */
@keyframes turtle-nod {
  0%, 100% { transform: translateY(0); }
  25%, 75% { transform: translateY(-4px); }
  50% { transform: translateY(0); }
}

/* 招手 */
@keyframes turtle-wave {
  0%, 100% { transform: rotate(0deg); }
  25% { transform: rotate(-15deg); }
  75% { transform: rotate(15deg); }
}

/* 打瞌睡 */
@keyframes turtle-sleepy {
  0%, 100% { transform: scale(1); opacity: 1; }
  50% { transform: scale(0.95); opacity: 0.7; }
}

.turtle-spin { animation: turtle-spin 0.8s ease; }
.turtle-blink { animation: turtle-blink 0.4s ease; }
.turtle-nod { animation: turtle-nod 0.6s ease; }
.turtle-wave { animation: turtle-wave 0.6s ease; }
.turtle-sleepy { animation: turtle-sleepy 2s ease infinite; }
.turtle-encourage { animation: turtle-nod 0.6s ease; }
```

---

## 🛠️ 技术实现清单

### 组件结构

```
apps/excel-addin/src/
├── components/
│   ├── ThemePanel.tsx              # 主题面板主组件
│   ├── PresetThemeSelector.tsx     # 预设主题选择器
│   ├── CustomBackgroundUpload.tsx  # 自定义背景上传
│   └── ThemePreview.tsx            # 主题预览区域
├── hooks/
│   ├── useTheme.ts                 # 主题管理 Hook
│   └── useLongPress.ts             # 长按检测 Hook
├── utils/
│   ├── imageStorage.ts             # IndexedDB 图片存储
│   └── imageProcessing.ts          # 图片处理（转 JPEG）
└── styles/
    └── theme-panel.css             # 主题面板样式（毛玻璃）
```

### 状态管理

```typescript
// 主题状态
interface ThemeState {
  preset: 'default' | 'warm-orange' | 'calm-blue' | 'vivid-green';
  customBackgroundUrl: string | null;  // ObjectURL
  backgroundOpacity: number;           // 0-100
  autoMask: boolean;                   // 是否启用自动遮罩
}

// 默认状态
const defaultThemeState: ThemeState = {
  preset: 'default',
  customBackgroundUrl: null,
  backgroundOpacity: 40,
  autoMask: true,
};
```

### 长按检测 Hook

```typescript
// hooks/useLongPress.ts
import { useRef } from 'react';

export function useLongPress(
  callback: () => void,
  duration: number = 2000
) {
  const timerRef = useRef<NodeJS.Timeout | null>(null);
  const startTimeRef = useRef<number>(0);

  const start = () => {
    startTimeRef.current = Date.now();
    const logoElement = document.querySelector('.app-logo');
    
    timerRef.current = setTimeout(() => {
      // 触发成功
      callback();
      logoElement?.classList.add('trigger-success');
      setTimeout(() => {
        logoElement?.classList.remove('trigger-success');
      }, 400);
    }, duration);
    
    // 1 秒后开始发光
    setTimeout(() => {
      if (timerRef.current) {
        logoElement?.classList.add('glow');
      }
    }, 1000);
  };

  const clear = () => {
    if (timerRef.current) {
      clearTimeout(timerRef.current);
      timerRef.current = null;
    }
    startTimeRef.current = 0;
    const logoElement = document.querySelector('.app-logo');
    logoElement?.classList.remove('glow');
  };

  return {
    onMouseDown: start,
    onMouseUp: clear,
    onMouseLeave: clear,
    onTouchStart: start,
    onTouchEnd: clear,
  };
}
```

### Logo 发光效果

```css
.app-logo {
  transition: all 200ms ease;
  cursor: pointer;
  user-select: none;
  -webkit-user-select: none;
}

.app-logo.glow {
  filter: drop-shadow(0 0 8px rgba(46, 128, 72, 0.5));
  transform: scale(1.05);
}

.app-logo.trigger-success {
  animation: logo-flash 0.4s ease;
}

@keyframes logo-flash {
  0%, 100% { opacity: 1; }
  50% { opacity: 0.3; }
}
```

---

## 📂 涉及文件

| 文件 | 改动内容 | 优先级 |
|------|---------|--------|
| `apps/excel-addin/src/components/ThemePanel.tsx` | 新增：主题面板组件 | 🔥 高 |
| `apps/excel-addin/src/hooks/useTheme.ts` | 新增：主题管理 Hook（使用 IndexedDB） | 🔥 高 |
| `apps/excel-addin/src/hooks/useLongPress.ts` | 新增：长按检测 Hook | 🔥 高 |
| `apps/excel-addin/src/utils/imageStorage.ts` | 新增：IndexedDB 封装 | 🔥 高 |
| `apps/excel-addin/src/utils/imageProcessing.ts` | 新增：图片处理工具（转 JPEG 90% 质量） | 🔥 高 |
| `apps/excel-addin/src/styles.css` | 添加：毛玻璃样式、主题变量、格仔动画 | 🔥 高 |
| `apps/excel-addin/src/App.tsx` | 修改：添加 Logo 长按监听、挂载主题面板、初始化主题 | 🔥 高 |
| `apps/excel-addin/src/Header.tsx`（如果有独立头部组件） | 修改：Logo 长按触发逻辑 | 🔥 高 |

---

## 🧪 验收标准

### 功能验收

1. **长按触发**
   - ✅ 长按 Logo 1 秒时，Logo 开始发光
   - ✅ 长按满 2 秒时，Logo 闪烁一下，主题面板从下方滑入
   - ✅ 中途松开鼠标，发光消失，不触发面板

2. **预设主题切换**
   - ✅ 点击预设主题，立即应用新色彩
   - ✅ 色彩符合设计规范（温暖橙、宁静蓝、活力绿、默认）
   - ✅ 切换主题后，IndexedDB 持久化

3. **自定义背景上传（10MB 支持）**
   - ✅ 点击"上传图片"，打开文件选择器
   - ✅ 只能选择 jpg/png/webp 格式
   - ✅ 文件大于 10MB 时，提示错误
   - ✅ 上传成功后，图片显示在背景（保持原始分辨率）
   - ✅ 图片转为 JPEG 质量 90%，存储到 IndexedDB
   - ✅ 高 DPI 屏幕（Retina）显示清晰
   - ✅ 透明度滑块可调节背景强度（0-100%）
   - ✅ 自动遮罩确保文字可读

4. **背景管理**
   - ✅ "删除背景"按钮可清除自定义背景
   - ✅ 关闭 Excel 重新打开，背景图片保持
   - ✅ 刷新页面，背景图片从 IndexedDB 加载

5. **格仔互动独立**
   - ✅ 点击格仔，随机播放动画（转身/眨眼/点头/招手）
   - ✅ 1 分钟无操作，格仔进入打瞌睡状态
   - ✅ 发送消息后，格仔点头鼓励
   - ✅ 格仔互动不会触发主题面板

### 视觉验收

1. **毛玻璃效果**
   - ✅ 主题面板背景半透明，有模糊效果
   - ✅ `backdrop-filter` 在支持的浏览器生效
   - ✅ 不支持的浏览器降级为半透明背景（无模糊）

2. **图片质量**
   - ✅ 上传的图片清晰（质量 90%）
   - ✅ 在高分辨率屏幕上不模糊
   - ✅ `background-size: cover` 自动适配容器

3. **动画流畅**
   - ✅ 面板滑入/滑出动画流畅（300ms）
   - ✅ Logo 发光效果平滑
   - ✅ 格仔动画可爱自然

4. **可读性保证**
   - ✅ 任何背景下，文字都清晰可读
   - ✅ 自动遮罩确保对比度符合 WCAG AA 标准
   - ✅ 按钮、输入框等交互元素清晰可见

### 性能验收

1. **加载速度**
   - ✅ 首次加载不受影响（主题面板懒加载）
   - ✅ 从 IndexedDB 加载背景图片 < 500ms

2. **运行流畅**
   - ✅ 毛玻璃效果不卡顿（60fps）
   - ✅ 在低端设备上，`backdrop-filter` 自动降级
   - ✅ IndexedDB 读写不阻塞 UI

3. **内存管理**
   - ✅ 使用 `URL.createObjectURL()` 创建的临时 URL
   - ✅ 在不需要时调用 `URL.revokeObjectURL()` 释放内存

---

## 🚀 实施步骤

### 阶段 1：基础架构（2 小时）

1. **创建 IndexedDB 工具**
   - `imageStorage.ts`：封装 IndexedDB 操作
   - 实现：`saveBackgroundImage`、`loadBackgroundImage`、`clearBackgroundImage`

2. **创建图片处理工具**
   - `imageProcessing.ts`：图片转 JPEG 质量 90%
   - 保持原始分辨率

3. **创建 Hooks**
   - `useTheme.ts`：主题状态管理（使用 IndexedDB）
   - `useLongPress.ts`：长按检测

4. **测试 IndexedDB**
   - 上传图片 → 存储 → 刷新页面 → 成功加载

---

### 阶段 2：长按触发和面板 UI（2 小时）

1. **添加 Logo 长按触发**
   - 在 `App.tsx` 或 `Header.tsx` 的 Logo 上绑定 `useLongPress`
   - 实现发光效果和触发逻辑

2. **创建主题面板组件**
   - `ThemePanel.tsx`：毛玻璃面板
   - 滑入/滑出动画
   - 预设主题选择器布局

3. **测试**：长按 Logo → 面板滑入 → 点击遮罩层关闭

---

### 阶段 3：预设主题（1 小时）

1. **定义预设主题配置**
   - 在 `useTheme.ts` 中定义 4 个预设主题
   - 色彩变量映射

2. **实现主题切换逻辑**
   - 点击预设主题 → 更新 CSS 变量
   - 持久化到 IndexedDB

3. **测试**：切换主题 → 色彩立即生效 → 刷新页面 → 主题保持

---

### 阶段 4：自定义背景（3 小时）

1. **实现图片上传**
   - 文件选择器
   - 文件大小校验（< 10MB）
   - FileReader 读取

2. **实现图片处理**
   - 转换为 JPEG 质量 90%
   - 保持原始尺寸

3. **应用背景图片**
   - CSS `::before` 伪元素
   - 透明度调节
   - 自动遮罩层

4. **实现删除功能**
   - "删除背景"按钮
   - 清除 IndexedDB 和 CSS 变量

5. **测试**：
   - 上传 10MB 图片 → 成功处理并显示
   - 调节透明度 → 实时更新
   - 刷新页面 → 背景保持
   - 删除背景 → 恢复默认

---

### 阶段 5：格仔互动（1.5 小时）

1. **实现格仔点击动画**
   - 随机选择动画（转身/眨眼/点头/招手）
   - CSS 动画定义

2. **实现空闲检测**
   - 1 分钟无操作 → 打瞌睡
   - 用户活动时重置计时器

3. **实现消息发送鼓励**
   - 监听消息发送事件
   - 触发点头动画

4. **测试**：
   - 点击格仔 → 随机动画播放
   - 空闲 1 分钟 → 打瞌睡
   - 发消息 → 点头鼓励

---

### 阶段 6：细节打磨和测试（1.5 小时）

1. **兼容性处理**
   - 检测 `backdrop-filter` 支持
   - 不支持时降级处理
   - 检测 IndexedDB 支持

2. **性能优化**
   - 图片懒加载
   - 防抖处理
   - 内存管理（释放 ObjectURL）

3. **全量测试**
   - 功能测试（长按、切换、上传、删除）
   - 视觉测试（毛玻璃、动画、可读性）
   - 性能测试（帧率、加载速度）
   - 大文件测试（9MB 图片）

4. **测试通过**
   ```bash
   npm run build:addin
   npm run lint
   npm run test:addin
   ```

---

## 📝 Git 提交规范

每个阶段完成后创建提交：

1. `feat: 添加 IndexedDB 图片存储和主题管理基础架构`
2. `feat: 实现 Logo 长按触发和主题面板 UI（毛玻璃）`
3. `feat: 实现预设主题切换功能`
4. `feat: 支持自定义背景图片上传（10MB，质量 90%）`
5. `feat: 为格仔添加互动动画（点击/空闲/鼓励）`
6. `polish: 优化主题功能性能和浏览器兼容性`

---

## 💡 后续增强方向（可选）

1. **更多预设主题**
   - 深色模式
   - 节日主题（春节、圣诞）
   - 季节主题（春夏秋冬）

2. **格仔更多互动**
   - 语音提示（文字转语音）
   - 表情系统（开心/困惑/惊讶）
   - 根据工作时长变化状态

3. **主题市场**
   - 云端主题库
   - 用户分享主题
   - 主题点赞和下载

4. **高级定制**
   - 字体选择
   - 圆角大小调节
   - 动画速度调节

---

## ⚠️ 注意事项

1. **格仔保持可见**
   - 任何主题下，格仔必须清晰可见
   - 不要被背景图遮挡

2. **可读性优先**
   - 自定义背景必须有遮罩层
   - 任何情况下文字对比度 ≥ 4.5:1（WCAG AA）

3. **性能监控**
   - `backdrop-filter` 在低端设备可能卡顿
   - 必须提供降级方案

4. **存储管理**
   - IndexedDB 容量虽大，但也要合理使用
   - 只保留一张背景图片（覆盖旧的）

5. **浏览器兼容性**
   - `backdrop-filter` 需要 `-webkit-` 前缀
   - Safari 需要特殊处理
   - 不支持 IndexedDB 的浏览器降级为预设主题（无自定义上传）

6. **内存泄漏防护**
   - `URL.createObjectURL()` 创建的 URL 必须手动释放
   - 在组件卸载时调用 `URL.revokeObjectURL()`

---

## 🔧 关键代码示例

### 初始化主题（App.tsx）

```typescript
useEffect(() => {
  async function initTheme() {
    // 1. 加载主题设置
    const settings = await loadThemeSettings();
    if (settings) {
      applyPresetTheme(settings.preset);
    }
    
    // 2. 加载自定义背景
    const backgroundUrl = await loadBackgroundImage();
    if (backgroundUrl) {
      applyBackgroundImage(backgroundUrl, settings?.opacity || 40);
    }
  }
  
  initTheme();
  
  // 清理函数：释放 ObjectURL
  return () => {
    const bgUrl = document.documentElement.style.getPropertyValue('--custom-background-image');
    if (bgUrl && bgUrl.startsWith('blob:')) {
      URL.revokeObjectURL(bgUrl);
    }
  };
}, []);
```

### 完整上传流程

```typescript
async function handleImageUpload(event: React.ChangeEvent<HTMLInputElement>) {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    // 1. 校验
    if (file.size > 10 * 1024 * 1024) {
      alert('图片文件过大，请选择小于 10MB 的图片');
      return;
    }

    // 2. 显示加载状态
    setUploading(true);

    // 3. 处理图片
    const blob = await convertToJPEG(file, 0.90);

    // 4. 存储到 IndexedDB
    await saveBackgroundImage(blob);

    // 5. 应用到界面
    const url = URL.createObjectURL(blob);
    applyBackgroundImage(url, backgroundOpacity);

    // 6. 保存设置
    await saveThemeSettings({
      ...currentSettings,
      hasCustomBackground: true,
    });

    setUploading(false);
  } catch (error) {
    console.error('图片上传失败:', error);
    alert('图片上传失败，请重试');
    setUploading(false);
  }
}
```

---

**文档版本**：v1.0  
**创建日期**：2026-08-18  
**更新日期**：2026-08-18（支持 10MB 图片，使用 IndexedDB）  
**预计实现时间**：11 小时（分 6 个阶段）  
**优先级**：🟡 中（彩蛋功能，不影响核心体验）  
**状态**：已实现 ✅
