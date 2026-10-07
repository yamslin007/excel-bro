# Docker Python 本机真实验收（2026-10-07）

> 本文是第一轮运行时验收快照。随后内核门禁/授权数据批次已重建镜像，当前 RepoDigest 为 `sha256:812d1ccea0e507e8fe0990c874d3be043f0bc52dfd4d60932338d555d955e630`，详见 [最新实施记录](./PYTHON_AUTHORIZED_DATA_2026-10-07.md)。以下旧 digest、ENETUNREACH 探针与未实现项不再作为最新状态。

## 本轮范围与状态

用户启动 Docker Desktop 后，明确批准构建专用镜像、下载公开分析依赖并用合成数据验收。运行时已真实执行 Python，不再只是 mock 骨架；**产品功能仍关闭**，不是完整 Python 产品交付。未使用真实工作簿/客户数据/密钥，未修改已有 Redis/MySQL 容器，不提交 Git 或发布安装包。

## 构建证据

- 本机 Linux 引擎：`npipe:////./pipe/docker_engine`，ID `c08e8ab2-2cbb-4c24-ad6e-e4b8efcf1b85`。
- 构建器：显式 `--builder default`，单节点、driver `docker`、endpoint `default`、status `running`，由已核实的本机 `-H` 绑定 daemon；不依赖用户 current builder 或远端环境变量。
- 临时 CLI 配置仅有已安装的 Docker Desktop buildx 插件目录和空 `auths`，不读取用户凭据配置；构建前后引擎 ID 相同。
- 归档只有 `.dockerignore`、`docker/Dockerfile`、`app/docker_worker.py`。不目录遍历、不发送其它 server 源码/环境配置。
- 基础镜像：`python@sha256:bf44cdfcb76cd3b41e879bc058fc37ec5872002ccfde7fcb765e218cde0cd79c`（本次由公开 `python:3.13-slim` 拉取后解析真实 RepoDigest）。
- 专用标签：`excel-bro-python-sandbox:build-9cef581abb9a`。
- 本次实际 RepoDigest：`excel-bro-python-sandbox@sha256:a3fa19d9815ad4b2eb974e096de9d00e229608ab6534c566ac28ffcfd282c4f6`，已回填 `config/capabilities.json.pythonDocker.imageDigest`。
- 本次实际 image ID 也是 `sha256:a3fa19d9815ad4b2eb974e096de9d00e229608ab6534c566ac28ffcfd282c4f6`；二者恰好一致不代表可通用互换。只用 `image inspect` 的实际 RepoDigests 配置运行引用。
- 镜像 Linux/amd64，用户 `65534:65534`；预装 numpy 2.1.3、pandas 2.2.3。传递依赖本次为 python-dateutil 2.9.0.post0、pytz/tzdata 2026.5、six 1.17.0，尚非完整依赖锁文件，不宣称跨日期重建得到相同镜像。

构建时三个文件 SHA256：

| 文件 | SHA256 |
| --- | --- |
| `.dockerignore` | `dd76609483beeb4da83736e8df259fd76d054614c5d1156a3c806fb6919c5184` |
| `docker/Dockerfile` | `8ada7eb7ac9f934090a0fcef9a4be6f24077a9f9cd76d5aa93e925c352c411e0` |
| `app/docker_worker.py` | `5bcaf0a571aecfe712376308825dbdb959ec863b8a76b3be576cc7aa7bcef573` |

## 本次修复与真实测试

1. 新增显式三文件 tar 的可信构建脚本，默认全排除 `.dockerignore`，移除 `.env.example` 重新纳入规则；独立核实本机 builder，不把运行期端点验证当构建器验证。
2. 库测试初次失败：OpenBLAS 尝试按硬件数量创建线程，触发既有 `pids-limit=1`。显式容器 `--env` 固定 OpenBLAS/OpenMP/MKL/NumExpr 各为单线程，宿主值不可覆盖；不扩大 PID 限额。修复后 numpy/pandas 实际返回 `{sum:6,rows:3}`。
3. 短输出直到 EOF 才读到：真实管道测试先红，`BufferedReader.read(4096)` 换为 `read1(4096)`；修复后短探针可运行中观察，取消测试不再靠启动后固定一秒猜测。
4. 收紧资源探针：内存读取实际 cgroup 128 MiB 上限后超限退出 137；CPU 读取 RLIMIT_CPU `(2,2)` 后内核终止；线程实际 cap=8、已创建线程后触限；scratch 先成功小写，再因 ENOSPC=28 超出 1 MiB。普通启动失败/协议失败/超时不能冒充这些成功证据。
5. 网络对照纠正：Docker Desktop 的 bridge 网关属于 Linux VM，不能把它当 Windows 主机。用同镜像 bridge 对照实际连接受控宿主监听器，再把同一个可达数值 IP/端口交给 network=none 探针，要求 ENETUNREACH=101；不接受 DNS 失败或超时。测试专用正对照不修改生产 runner 的断网配置，无挂载/密钥/端口发布，并按确切容器名清理。

真实套件共 **15 项**：结构化分析、库导入/计算、受控网络/外网断路证据、宿主服务地址区分、辅助宿主文件探针、合成环境凭据未继承、内存、CPU、进程/线程、scratch、stdout、stderr、只读根/fresh scratch、墙钟超时、运行中取消。

文件探针确认宿主哨兵存在和容器读取代码正常，但 Windows 路径在 Linux 不存在本身**不单独证明**文件隔离。容器启动链不包含宿主挂载；此测试不能替代对所有 Docker/内核攻击面的审计。只读根实际写入得到 EROFS=30，scratch 两次运行均无前次文件，非 root UID 回读为 65534。超时/取消/限额探针按本次精确容器身份确认不存在，查询失败不能冒充清理成功。

## 可重复命令

在仓库根运行，Docker Desktop 保持运行并切到 Linux 容器模式：

```powershell
.venv\Scripts\python.exe -m scripts.build_python_sandbox
# 构建成功输出 BUILD_RECEIPT；核实实际 RepoDigest 后更新 pythonDocker.imageDigest。
# 没有真实 RepoDigest 就拒绝，不拿 image ID 填这个字段。
.venv\Scripts\python.exe -B -m pytest server/tests/test_python_docker_real.py -q -p no:cacheprovider
```

构建入口目前支持已安装 Docker Desktop 的 bundled buildx 目录布局；其它布局失败关闭，不宣称全平台构建支持。Dockerfile 的无默认 `PYTHON_BASE` 产生 InvalidDefaultArgInFrom 检查警告，脚本提供不可变基础引用，实际构建成功；不允许跳过脚本用未核实 builder 或漏传该参数。

## 不能开放的剩余边界

- `pythonExecution.enabled=false`，模型产品工具尚不能调用真实执行器；本轮只用了内部入口和合成数据。
- `pids-limit` 不等于完全禁 exec；容器内替换当前进程等完整禁止机制未实现。断网/只读/非 root 不能冒充该机制。
- datasetHandle 授权登记/核对、resultHandle 结果库、结果编译为白名单 `AnalysisPlan`、Excel/WPS/文件夹点击链路与安装包未实现。
- SDK CLI 内部转录不可投影的限制保持不变；不因 Docker 通过而宣称两个后端隐私闭环等同。
- 本机镜像不是朋友电脑已安装：新的机器仍要准备获准的本机引擎、同一镜像及身份验证。不能直接把这次验收称为“一键安装包已完成”。

## 最终回归与反向变异

| 检查 | 结果 |
| --- | --- |
| 服务端全量 `pytest server/tests -q -p no:cacheprovider --basetemp .tmp-server-docker-final-20261007` | 1053 passed，1 skipped，54.32 秒 |
| 真实 Docker 套件 + runner + build 定向 | 41 passed，1 skipped，30.20 秒（其中真实套件 15 项均通过） |
| 变异进程结束后 runner/build 复绿 | 26 passed，1 skipped |
| Excel `test:addin` | 776 passed |
| WPS `test:wps` | 531 passed |
| `build:addin` / `build:wps` | 均 exit 0 |
| `compileall -q server/app scripts/build_python_sandbox.py` | exit 0 |

唯一跳过的是构建上下文的真实符号链接创建测试：当前用户/策略不允许该测试创建 symlink；**不是 Docker 测试跳过，不记为通过**。构建器拒绝、三文件归档哨兵、固定线程参数和真实管道均执行通过。WPS 的既有 React act 提醒、Vite 大 chunk 提醒仍存在，不影响测试/构建退出，未纳入本批 UI 优化。

四项独立进程内反向变异（不写生产源码，不在宿主执行模型分析代码）：短输出退回 `read` → 1 红；OpenBLAS 固定值改为 999 → 1 红；builder 校验 no-op → 3 红；归档纳入合成 `.env` → 1 红。各项变异进程结束后原代码复跑全绿，不把变异红灯替代真实隔离验证。

最终 `docker ps -a` 仅保留原先 `flashflick-redis`（ID `83ca85b9a868`）和 `flashflick-mysql`（ID `2d314d78e5f2`），均仍为 8 个月前停止的状态；无本次测试容器残留。镜像 RepoDigest 与三个构建输入的 SHA256 再次回读一致。生成的专用镜像/基础镜像保留供后续验证；只清理本轮自建容器及构建用临时 CLI 配置，不全局 prune 或删除用户已有镜像/目录。

官方机制参考：[默认 builder 的 daemon 绑定](https://docs.docker.com/build/builders/)、[tar 构建上下文](https://docs.docker.com/build/concepts/context/)、[容器运行与资源限制](https://docs.docker.com/engine/containers/run/)、[OpenBLAS 单线程设置](https://www.openmathlib.org/OpenBLAS/docs/faq/)。本机测试结论以上述实际运行证据为准，不从文档直接推断通过。
