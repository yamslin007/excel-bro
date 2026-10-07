# 隔离 Python：易失结果与双端本地预览准备

本批接通内部合成分析→易失结果句柄→本机读取→白名单值矩阵编译→Office/WPS 只读预览准备。**不是产品 Python 可用**：总开关仍 false，模型调用/面板按钮/原生副本/确认写回/任务生命周期未接线。未调用付费模型、未碰真实工作簿、未提交/打包/重建镜像/安装依赖，原有改动保留。

## 实现和隐私边界

`python_results.py` 独立校验结果：只能为 `{kind:"matrix",values:[...]}`，没有动作、公式协议、坐标或自由元数据；规则非空矩阵、有限安全标量、UTF-16 字符数、UTF-8 总字节。整数和整数形式的浮点不能超 JS 安全整数，拒绝 NaN/Inf、对象、孤立 surrogate。普通负数可以写，可能被宿主按公式/转义解释的字符串前缀（空白后 =/+/-/@/单引号）及 NUL 不能编译，拒绝而非悄悄改值。

`PythonDatasetRegistry` 同锁保存不可变编码结果，绑定 scopeToken、完整 task/host/file-hash/授权版本/明确来源、Origin、有序 datasetHandle/内容指纹及结果元数据。随机 resultHandle、SHA256、行列数由服务端产生。只返回新的解码副本，不返回原内部对象。替换源快照删除依赖结果；重授权、撤销、原任务 900 秒截止、服务重启失效，取回不延长 TTL。限额 `pythonResult`：每 scope 12 结果、结果编码总内存 16 MiB、每项 1 MiB/500 行/50 列/20,000 格/每格 32767 UTF-16 单位。不是磁盘结果库，不支持跨进程或跨关闭续跑。

内部 `run_registered_result_internal` 复用既有授权数据→真实隔离计算链，严格验证后发布结果，仅返回 `{status:"completed",receipt:{resultHandle,contentFingerprint,rows,columns}}`。发布后再次检查取消/授权，迟到结果精确丢弃。`execute_python_analysis` 仍无产品成功通道，completed/receipt 不被模型出站闸接受；不把本地计算结果或代码放入模型、事件/检查点。

新增本机 POST `/api/python/data/result`、`/api/python/data/result/discard`：请求为 scope+秘密 token+有序 datasets 回执+结果回执，复用数据 API 的回环/Host/Origin/hostId/JSON/开关/有界请求检查。错误固定 409 文案、不回显输入，成功 no-store。结果 HTTP 不返回路径或代码。读取结果只进本机页面内存；**没有 HTTP 执行接口**。契约同步 core/contracts.ts、Excel contracts.ts 导出、models.py，PlanRequest/AnalysisPlan 动作协议不扩大。

core 的 getResult 读取前后重读真实源值/授权，响应逐块限大小、禁重定向、实际 JSON UTF-8 字节 SHA256；不能改为 JSON.stringify(parsed)，否则 Python 的 1.0/指数格式与 JS 不一致。结果失效只提示重新读取分析，不冒充检测到文件变化。dispose/取消撤销；离线撤销最佳努力，TTL 兜底，不声称立刻清空服务端。

共享编译器目标由可信用户流程明确提供，必须是本次已登记授权表的 sheet ID/name 与合法 A1 区域、尺寸精确匹配；Python 输出无权指定。只生成现有 `writeValues` 与 `rangeEquals`，值和验收矩阵分别克隆。stage 校验捕获的真实来源描述与原 scope 一致、计划和 scope 不可变，准备后再次取回验证结果活性；结果、计划和源值只活在内存。

## 双端实现与未接线部分

- Office `prepareOfficePythonResult`：核对原数据会话绑定的宿主，使用实际本地 preflight/OfficeApi 能力检查，按已授权 sheet ID 读取目标 values/formulas；目标快照、源、文件/授权、计划、结果活性守卫。没有写入/副本/持久账本。**完整可写性检查、原生副本和确认写回 UI 未接线**。
- WPS `prepareWpsPythonResult`：相同编译器/会话活性，实际 prepareWpsPlan 核对源、目标指纹、保护/合并、宿主版本、本地 xlsx 与自身写入上限；目标/源/计划变化拒绝。**原生副本和确认写回 UI 未接线**。结果全局上限不覆盖 WPS 较小单次写入上限；超限明确拒绝，不降级截断。
- 当前仅开发内部模块，两端均没有用户可点击的 Python 流程，不能标“实机待验”代替未实现。以后接线需把 assertCurrent 放在首次实际写入前，同时沿用所有既有宿主写入守卫，不以元数据检查代替完整可写性证明。

尤其不能直接将带结果矩阵的计划交给旧 plan_validate/规划请求、历史或会保存值矩阵的执行/撤销账本。两端需要非持久执行与会话内恢复路径，或者先明确申请改变隐私边界；本批不改旧执行器、不暗中放宽。文件夹 Python、SDK 内部转录投影/恢复、安装包和朋友电脑安装验收均未实现。

## 验证证据

server 新结果套件 42 条：完整绑定、精确输入顺序、不可变取回、源替换/重授权/撤销/过期/重启/丢弃、验证期间过期、所有配额、严格矩阵与安全整数、HTTP 实际字节指纹（1.0/小数/中文 emoji）、双宿主本机读取/丢弃、未知字段无回显、默认关闭/模型拒绝、发布后取消精确清理。

共享层：编译器标量/Unicode/A1 边界；两宿主各跑真实 fetch 夹具，检查本机 URL/请求体不含结果值、无 storage、源变化前后检测、篡改/过期/超大响应、取消/关闭、公开 scope/host 不能扩大授权、计划不可变和准备后的结果活性。Office 8 条、WPS 11 条使用实际 preflight/prepareWpsPlan，不 mock 核心准备器，零宿主 setter/持久 storage 调用。真实宿主/真实模型尚未验收。

新增真实 Docker 用例：已登记数据在当前固定镜像内运行 `sum/count` 矩阵分析，返回结果句柄，从登记器取回 `[[3,2]]`；不是宿主裸计算或 skip。当前 process_gate 六项全绿，原有 15 项真实 Docker 套件在全量回归也通过，合计 21 项，无镜像/运行环境变更。

独立进程内反向变异脚本 `scripts/validate_python_result_mutations.py`，不编辑生产/测试文件：拆回执/输入绑定→3 红，拆矩阵校验→9 红，拆发布后取消守卫→1 红；每项还原的新进程结果套件 42 全绿。这证明这些指定测试有判别力，不等于全部隔离威胁已审计。

收尾发现并修正 BOM 校验遗漏：先解码再重编码计算 SHA256 会被 TextDecoder 自动移除 BOM。新增双宿主 BOM 篡改响应用例先 2 红，改为对收到的 Uint8Array 直接 hash 后通过；不把“解码后相同”冒充“原始响应字节未变化”。

本批全量：server **1130 passed / 1 skipped**（既有 symlink 创建权限限制，非 Docker）；Excel **789**、WPS **547**、core **306**。双端 tsc/vite build 和 compileall 通过。WPS 保留已有 React act 警告，构建保留 chunk 大小警告。

最终定向复验曾遗漏独立 basetemp，12 个依赖 tmp_path 的旧隐私测试因默认 `C:/Users/jia/AppData/Local/Temp/pytest-of-jia` 无访问权限而在 setup 报错；改为新的明确工作区 basetemp 后 116 条全绿，不改变系统权限、不删除旧目录。这是测试临时目录环境问题，不用修改隐私守卫解决。

单列限制：`npm run build --workspace @excel-bro/core` 有 15 处既有 `ColumnMappingPlanAction` 联合类型未收窄的测试类型错误（blankPolicy.test.ts、columnMapping.test.ts、columnMappingSemantics.test.ts），相关源码/测试均未被本批修改。core 执行测试通过，双端构建通过，不能因此把 core 独立 tsc 报成成功；本批不改无关测试断言来凑绿。

可重复运行：

```powershell
.venv\Scripts\python.exe -B -m pytest server/tests/test_python_results.py server/tests/test_python_process_gate.py -q -p no:cacheprovider
.venv\Scripts\python.exe -B -m scripts.validate_python_result_mutations
npm run test:core
npm run test:addin
npm run test:wps
npm run build:addin
npm run build:wps
.venv\Scripts\python.exe -m compileall -q server/app
```

## 下一阶段

先设计并同时实现两端“确认目标/来源→隔离分析→选择明确目标区域→本地原生副本审阅→确认写回”的产品接口、状态和生命周期，明确取消/授权变化处置与非持久执行恢复；普通后端先用可信任务绑定的新执行接口，不复用自报句柄作为授权。SDK 不能绕过其未解决的内部转录边界。完成必要写前守卫、双端测试和独立真实测试工作簿验收后再决定开关；目前仍不能仅改 enabled=true。
