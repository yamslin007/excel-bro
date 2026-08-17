# 贡献指南

本文档面向第一次接触仓库的开发者，说明如何准备环境、运行常用命令、
提交变更和做代码检查。所有命令都在仓库根目录的 PowerShell 中执行。

## 环境准备

- Node.js：实测为 v24.12.0，建议使用当前 LTS 或与 `package.json` 的
  engines 保持兼容
- npm：实测为 11.8.0
- Python：3.9+
- PowerShell 5.1 或 PowerShell 7

首次安装：

```powershell
npm install
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r server\requirements.txt
Copy-Item server\.env.example server\.env
```

如果 PowerShell 阻止虚拟环境脚本，可仅为当前终端放开：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
.\.venv\Scripts\Activate.ps1
```

## 常用命令

### 后端开发

```powershell
npm run dev:server
```

服务默认监听 `http://127.0.0.1:8765`，源码修改后会自动重载。

健康检查：

```powershell
Invoke-RestMethod http://127.0.0.1:8765/health
Invoke-RestMethod http://127.0.0.1:8765/api/models
Invoke-RestMethod http://127.0.0.1:8765/api/settings/model
```

### 前端开发

```powershell
npm run dev:addin
```

任务窗格地址是 `https://localhost:3000`。开发证书首次使用时会请求 Windows
信任证书。

浏览器模式可以直接打开 `https://localhost:3000` 调试布局，但真实 Excel
行为仍需在 Excel 中验证。

### Excel 旁加载

另开一个终端执行：

```powershell
npm run start:excel
```

该命令会创建并打开一个临时测试工作簿，加载项只旁加载到这个工作簿。停止旁加载：

```powershell
npm run stop:excel
```

修改 `apps/excel-addin/manifest.xml` 后，需要先停止再重新旁加载，仅刷新任务
窗格不会更新功能区按钮。

## 测试

### 前端

```powershell
npm run test:addin
```

当前基线为 30 个测试文件、258 项测试。

### 后端

Windows 下 `pytest` 的 `tmp_path` 需要项目内 basetemp，否则可能因临时目录权限
报错。推荐直接运行：

```powershell
python -m pytest server/tests -q --basetemp=.pytest-tmp
```

不要省略 `--basetemp=.pytest-tmp`。

同时建议运行 Python 编译检查：

```powershell
python -m compileall -q server/app
```

## 构建

```powershell
npm run build:addin
```

该命令先执行 `tsc --noEmit` 做类型检查，再执行 Vite 生产构建。类型错误会在这
一步暴露。

## 代码检查

```powershell
npm run lint
```

当前 ESLint 覆盖 `apps/excel-addin/src`，只启用 React Hooks 的两条规则：

- `react-hooks/rules-of-hooks`: error
- `react-hooks/exhaustive-deps`: warn

检查基线见 `docs/LINT_BASELINE.md`。lint 目前只提示 Hooks 依赖问题，不应顺手
批量修改业务逻辑。

## 提交规范

参照现有 git log，提交信息使用简洁中文，格式为：

```text
type: 说明
```

常见 type：

- `feat`：新功能
- `fix`：缺陷修复
- `refactor`：重构或类型收敛
- `docs`：文档
- `test`：测试
- `chore`：工程配置

示例：

```text
fix: 修复 useWorkbookContext 依赖
refactor: 收敛公共类型定义
chore: 接入 react-hooks ESLint 规则
```

每个提交只处理一个主题；涉及架构、协议、启动方式或环境变量变化时，同步更新
`docs/` 和根目录 `README.md`。
