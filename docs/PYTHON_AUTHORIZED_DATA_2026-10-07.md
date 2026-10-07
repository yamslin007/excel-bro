# 隔离 Python：内核门禁与授权数据接线（2026-10-07）

本批完成内部执行安全门禁、易失数据登记、本机数据 API、共享客户端及 Excel/WPS 只读适配。**不等于产品 Python 已可用**：`pythonExecution.enabled=false`；两端按钮/生命周期接线、模型工具成功通道、结果库、结果转 `AnalysisPlan`、文件夹适配与安装包未实现。未操作真实工作簿、未提交、未打包；已有工作区改动保留。

## 1. 真实内核门禁

旧镜像三个真实用例先失败：`os.execv` 替换 worker 并运行 `/bin/sh`；有剩余 PID 配额时 `fork` 成功；原生 `execveat` 未被拒绝。`pids-limit` 不是禁止执行外部程序的证明。

`docker_worker.py` 在可信 Python 启动、输入解析之后，任何生成代码之前安装附加 seccomp BPF。限定 Linux/amd64、容器标记、UID 65534；设置 no_new_privs，以 TSYNC 同步线程。架构不匹配终止；拒绝 x32 与旧 x32 syscall 编号、execve/execveat、clone/clone3/fork/vfork、socket/socketpair。安装失败返回固定 `isolation_not_applied`，不执行代码，不提供测试/环境变量绕过。保留 Docker 默认 seccomp 与原有断网、只读根、非 root、资源限额；命名空间/AST/monkeypatch 不充当安全边界。

过滤器安装时序与不可放宽语义依据 [Linux seccomp 文档](https://docs.kernel.org/userspace-api/seccomp_filter.html) 和 [seccomp 系统调用](https://man7.org/linux/man-pages/man2/seccomp.2.html)。这是一组经真实用例验证的附加拒绝策略，**不是对所有容器/内核逃逸攻击面的完整审计**；其它架构未支持。

本批重新构建，实际 RepoDigest：

`excel-bro-python-sandbox@sha256:812d1ccea0e507e8fe0990c874d3be043f0bc52dfd4d60932338d555d955e630`

标签 `excel-bro-python-sandbox:build-c6df1e645cc6`，相同本机端点/引擎；基础镜像仍固定 `python@sha256:bf44cdfcb76cd3b41e879bc058fc37ec5872002ccfde7fcb765e218cde0cd79c`。worker 构建输入 SHA256 `aa64d4c0ce64dea44ef3a90c1ce01c23af6c3b00dfacefc99d702ce9d0bbf1dc`。回填真实 RepoDigest，不以 image ID 推导；开关不变。旧镜像/旧记录仅表示上一阶段。

5 个新增真实 Docker 用例通过：替换进程拒绝、剩余 PID 下 fork 拒绝、原生 syscall 拒绝、IPv4/IPv6/Unix socket 拒绝、登记数据实际在容器中计算 `{sum:3,count:2}`。原有 15 个真实用例也通过，numpy/pandas 可用。网络正对照与 PID 定量对照用**固定可信测试代码**绕过 worker 门禁，以分别证明 Docker 的网络/PID 机制；不作为产品备用通道，不挂载数据。宿主 worker 测试改为失败关闭/可信 stub，不再在宿主执行分析代码。

## 2. 独立本机数据通道

协议同步 core/contracts、Excel contracts 转导出和 models.py，**没有给 PlanRequest 新增数据或秘密字段**：

- `POST /api/python/data/scope`：只登记元数据，返回随机 scopeToken。
- `POST /api/python/data/snapshot`：精确区域的已授权值，只进本机服务内存；返回随机 datasetHandle 和内容指纹。
- `POST /api/python/data/revoke`：撤销该范围及所有句柄。

三接口均先检查实际连接为回环、请求主机为 localhost/回环、Origin 是允许的 Excel/WPS 页面且与 hostId 对应、JSON Content-Type、总开关。忽略代理自报身份；禁止无 Origin CLI 和远端请求。请求流逐块限字节；解析/校验失败只返回固定文案，不使用含原始输入的 422 回显；响应 no-store。后端应绑定回环、不得将该接口公开代理到网络。

授权绑定版本 1 包含 taskId、hostId、工作簿身份 hash、authorizationVersion、明确 sheetId/name/range。**登记绑定不证明客户端自报的文件身份真实**；实际宿主证据由两端适配器提供，服务器不能独立感知 Excel/WPS 当前文件。随机 token 只留受控内存；不进入聊天、模型、检查点、URL 或 localStorage。

`PythonDatasetRegistry` 仅内存：精确绑定/Origin 检查、矩阵形状与有限标量检查、不可变编码快照、服务器生成句柄与指纹。新授权替换同任务旧授权；同源新快照替换旧句柄；拒绝跨任务/文件/授权版本/范围/Origin、伪造句柄或指纹、重复句柄、调用预算超限。撤销即时失效；TTL 在访问时清理。单服务进程使用，重启全部失效，未实现跨进程/持久续跑。

限额集中 `capabilities.pythonData`：32 scope/易失任务预算、每 scope 8 数据集、每集 20,000 格/1 MiB、总内存编码数据 16 MiB、单次分析数据 1 MiB、HTTP 请求约 1.06 MiB、单请求超时 10 秒、TTL 900 秒；另沿用每任务最多 12 次分析。重授权/撤销不重置同一易失任务预算或延长最初截止时间；易失任务过期后不支持当作持久任务续跑。无磁盘溢写、账本记录、日志数据或宿主路径解析。

## 3. 两端对齐与使用前复核

共享 `collectPythonData` 在任何宿主检查/读取前检查关闭状态和本机 API URL。拒绝远端、用户信息、非根路径、查询片段；fetch 禁重定向、no-store、无凭据，回执流有界。先登记元数据、再在读取前后复核授权；连续两读不一致则不上传并撤销。使用 `revalidate()` 前再次读取源值比对本地指纹；变化则请求撤销；`dispose()`/AbortSignal 撤销。离线撤销可能无法送达，服务器 TTL 是兜底，不能声称离线即时清空。

Office 使用当前 document 的强身份证据与授权 sheet ID 解析读取，不重新按同名替代表读取。ID 解析依据 [WorksheetCollection.getItem 官方协议](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetcollection?view=excel-js-preview)。WPS 使用既有授权读取和对象身份验证；初次不采用与授权引用不同的工作簿作新基线，后续 unknown/new wrapper 不冒充授权。两端均克隆授权元数据防止原地修改，读取前后守卫，只有 readValues、无写入方法调用。

**适配函数已实现，面板尚未调用它们**。未接入按钮、目标变化监听器与任务结束钩子；未来接线必须在用户确认后创建会话，目标/范围变化时调用 dispose，每次执行前后 revalidate。双读不是宿主原子事务，不替代后续预览/写前复核。文件夹模式本批未接线，不读取当前工作簿磁盘文件冒充未保存值。

内部 `run_registered_analysis_internal` 解析登记句柄，将值仅经 stdin 送入本机容器；运行中登记失效触发取消，结束后再核实，丢弃迟到成功结果。它没有 HTTP 执行路由或模型工具成功通道，不返回公开 resultHandle。普通/SDK 工具仍失败关闭；健康但未接线时用准确固定 `product_not_ready` 回执，不误报 runner_missing。历史参数投影仍不保留自报句柄，SDK 内部转录限制未解决。

## 4. 验证与剩余项

新增数据登记/API 29 测试、共享客户端 14 测试、Office/WPS 适配器各 5 测试、真实门禁/登记计算 5 测试。三个独立内存变异分别移除绑定、迟到结果校验、撤销：5/1/2 条变红；每次新进程还原套件全绿。脚本 `scripts/validate_python_data_mutations.py` 不修改生产/测试文件。额外用例确认重授权/撤销不重置调用预算或 TTL。

最终回归：server **1087 passed / 1 skipped**（既有真实 symlink 创建权限限制，不是 Docker skip）；20 项真实 Docker 测试通过；Excel **781**、WPS **536**、core **268**；双端 build 和 compileall 通过。Excel 首轮并发运行有 1 个授权 UI 用例超时，单套件复验及第二轮全量通过，未修改原守卫/断言；不据此宣称该时序用例永不波动。

未实现：产品模型调用接线、结果句柄库/结果出站策略、结果转白名单 AnalysisPlan、双端授权/分析/预览 UI 及任务生命周期、文件夹适配、SDK 转录隐私闭环/检查点恢复、安装包与朋友电脑验证。**下一批应先实现结果登记与白名单预览编译，然后再接两端 UI；不能仅把 enabled 改成 true。**

可重复验证（新的 basetemp 目录，勿重用/删除既有用户目录）：

```powershell
.venv\Scripts\python.exe -B -m pytest server/tests/test_python_datasets.py server/tests/test_python_process_gate.py -q -p no:cacheprovider
.venv\Scripts\python.exe -B -m scripts.validate_python_data_mutations
npm run test:core
npm run test:addin
npm run test:wps
npm run build:addin
npm run build:wps
```
