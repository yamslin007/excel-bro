# 调试指南

## Excel 加载项调试

### 旁加载

先启动后端和前端：

```powershell
npm run dev:server
```

```powershell
npm run dev:addin
```

另开一个终端旁加载：

```powershell
npm run start:excel
```

`start:excel` 使用 `--no-debug`，只注册并旁加载清单，不开启依赖 Visual Studio
的 Office 直接调试器；React/CSS 仍由 Vite 热更新。停止旁加载：

```powershell
npm run stop:excel
```

修改 `apps/excel-addin/manifest.xml` 后必须先停止再重新旁加载，否则功能区按钮
和清单更新不会生效。

### 任务窗格 DevTools

Excel 中的前端错误优先打开 WebView 开发者工具查看，而不是只看系统终端。可从
Excel 加载项调试入口打开 DevTools；如果入口不可用，可先确认旁加载是否成功。

常见检查点：

- 任务窗格请求的是 `https://localhost:3000/index.html`
- 本地服务请求是 `http://127.0.0.1:8765`
- 网络失败先确认两个端口都有进程监听

## 前后端联调

前端开发服务器由 Vite 监听：

- 地址：`https://localhost:3000`
- 配置：`apps/excel-addin/vite.config.ts`
- `server.port: 3000`，`strictPort: true`

后端 FastAPI 服务监听：

- 地址：`http://127.0.0.1:8765`
- 启动命令：`npm run dev:server`

前端 API 地址默认从 `apps/excel-addin/src/api.ts` 读取：

```ts
const API_BASE_URL =
  import.meta.env.VITE_API_BASE_URL ?? "http://127.0.0.1:8765";
```

当前 `vite.config.ts` 没有配置反向代理，任务窗格直接访问
`VITE_API_BASE_URL`。后端 CORS 只允许 `localhost:3000` / `127.0.0.1:3000`
来源，对应配置在 `server/app/main.py`。

后端健康检查：

```powershell
Invoke-RestMethod http://127.0.0.1:8765/health
Invoke-RestMethod http://127.0.0.1:8765/api/models
Invoke-RestMethod http://127.0.0.1:8765/api/settings/model
```

## 常见问题排查

### 前端 3000 端口或后端 8765 端口被占用

查看监听状态：

```powershell
Get-NetTCPConnection -LocalPort 3000,8765 -State Listen
```

如果端口被旧进程占用，先停止旧终端或对应进程，再重新启动。

### 首次启动前端提示证书问题

`npm run dev:addin` 首次会使用本地 HTTPS 证书。Windows 弹出信任提示时确认信任；
如果证书过期或失效，可重新运行开发证书工具后再启动。

### Excel 中功能区没有 Excel Bro 标签

确认已经执行 `npm run start:excel`，且 `manifest.xml` 修改后已经停止并重新旁加载。
普通新建工作簿不会自动继承开发旁加载，`start:excel` 只把加载项注册到它创建的
临时测试工作簿。

### 前端请求后端失败

1. 确认后端终端正在运行 `npm run dev:server`。
2. 确认后端健康检查返回正常。
3. 确认前端没有通过错误的 `VITE_API_BASE_URL` 指向其他地址。
4. 检查 WebView DevTools 的 Network 面板，查看失败的是 CORS、连接拒绝还是
   4xx/5xx。

### pytest 在 Windows 下因 tmp_path 报权限错误

后端测试必须指定项目内 basetemp：

```powershell
python -m pytest server/tests -q --basetemp=.pytest-tmp
```

不要省略 `--basetemp=.pytest-tmp`。

### 浏览器模式无法验证 Office.js 行为

直接打开 `https://localhost:3000` 和 `https://localhost:3000/focus.html` 只能
检查布局、历史对话、模型选择和普通 API。当前工作簿读取、多工作表顺序、Office.js
动作、写入验收、专注窗口同步等仍需在 Excel 中验证。

### 修改类型或协议

前端类型定义已收敛到 `apps/excel-addin/src/types/`，领域协议仍以
`apps/excel-addin/src/contracts.ts` 与 `server/app/models.py` 为准。改动后至少
运行：

```powershell
npm run build:addin
npm run test:addin
python -m pytest server/tests -q --basetemp=.pytest-tmp
python -m compileall -q server/app
```
