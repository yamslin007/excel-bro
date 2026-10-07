# ADR 0001：任务调度账本与文件夹写入排他

状态：已决定，阶段 A 已实施  
日期：2026-09-08  
依据：[任务断点恢复指导 v0.3](../TASK_CHECKPOINT_RECOVERY_GUIDE.md)

## 决定

1. 使用 Python 标准库 SQLite 保存任务身份、状态、事件、结果和后续 checkpoint。数据库位于统一应用数据目录 `tasks/state.sqlite3`。
2. 任务账本是唯一能够批准 Excel Bro 再次调度写入的记录；备份 manifest 继续保存文件恢复证据，实际文件可以否决账本中的完成判断。
3. 首期以工作目录为保守写入边界，使用进程内非重入锁和跨进程文件锁串行化执行与撤销。锁文件位于应用数据目录，不在用户工作目录创建空备份目录。
4. 流式执行第一次收到 operationId 时原子登记；内存状态被清理后，从账本重放结果和事件，不重新执行。同一身份内容不同返回冲突；记录仍是 running 而内存现场不存在时进入核对，不自动接管。
5. SQLite 使用 WAL、FULL synchronous、busy timeout 和显式事务。工作簿替换不属于 SQLite 事务，后续仍须按 manifest 和实际指纹核对。
6. 首建并发必须先在**独立 bootstrap 连接**上完成 WAL 切换与建表并关闭，之后才允许其它连接打开同一文件；bootstrap 自身与常规连接在 Windows/WAL 下被 SQLite 短暂持成只读句柄时**重开**（不降级为只读、不跳过写入）。原因见下方"首次并发落库的可写性"。
7. **读路径不得取写锁**（第十批 F4）：纯查询走 `mode=ro` 只读连接（`_connect_readonly`），绝不复用可写连接——可写连接会先 `BEGIN IMMEDIATE` 探测，使纯读被串行到写事务之后（实测：另一连接持写事务时 `ledger.get()` 阻塞 10.745s 后报 `database is locked`；改只读连接后 0.003s 返回）。可写探测只用于**可写连接**，不得下放到读路径。

## 原因

版本化 JSON 可以实现单文件原子替换，但任务状态与事件需一起提交、跨线程竞争需条件更新，SQLite 的事务和唯一键更直接。它是 Python 自带组件，不引入模型或任意执行能力。

目录级锁牺牲同一目录内不同文件的并行度，但当前计划可以动态创建或改名输出，执行前无法始终准确列出全部目标。先保证不出现同目录交叉发布，待目标冻结协议通过故障测试后再评估按文件细化。

## 首次并发落库的可写性（2026-10-07 定位）

现象：`test_register_is_atomic_and_reuses_existing_identity`（8 线程 / 16 并发首次登记）在 Windows 上稳定报 `sqlite3.OperationalError: attempt to write a readonly database`（约 3–6/16 次）。**不是**文件只读属性、ACL、也无关于 `tmp_path`/`\\?\` 路径解析、`TMP` 环境变量或 `EXCEL_BRO_CONFIG_DIR`——同目录下的 DB 文件属性为普通可写，根因是**并发首建 + WAL 切换**。

隔离实验（均在本机复现并重复验证）：① 预创建 DB 文件后并发 0/16 失败；② 16 连接在**同一新库**上并发首建 → 稳定失败；③ 仅去掉 `PRAGMA journal_mode = WAL`（用 rollback journal）→ 0/640 失败，命中 WAL；④ 高并发（32 线程/64 任务）下 WAL 仍有失败。结论：SQLite 在 Windows 上把"WAL 切换进行中"被并发打开的连接持成只读句柄；`busy_timeout` 不适用（不是锁竞争）。

修复（保留 WAL 设计，双管齐下）：① **bootstrap 串行化 + 当前版本短路**——`_ensure_schema` 在独立连接上完成 `journal_mode=WAL`、建表、`user_version` 后关闭；若库的 `user_version` 已达当前版本（`_schema_is_current`）则**直接返回**，不再重跑 journal_mode 切换与迁移（新实例重跑会与开启中的写事务冲突）；② **可写探测重开**——bootstrap 连接与常规可写连接都先 `BEGIN IMMEDIATE`/`ROLLBACK` 探测可写，遇 `readonly` 短暂状态按退避重开（`_READONLY_RETRY_ATTEMPTS`），绝不降级为只读或跳过写入。组合后在本机 0 失败（真实并发 640/640、高压 640/640、极压 768/768）。红绿与独立变异：`server/tests/test_task_ledger.py::test_first_initialization_under_concurrency_never_reports_readonly` + `scripts/mutate_task_ledger_wal.py`（M13/M14）。

**根因更正（第十批 F4，2026-10-07）**：上文早期解释里"提交一次真实写使 `-wal`/`-shm` 实体化稳定"经独立复现**不成立**——bootstrap 连接关闭后 `-wal`/`-shm` 并不存在（`outputs/repro_f4_lock_and_wal.py`），种子写对该现象无影响。因果实验（`outputs/repro_f4_causality.py` / `_causality2.py`，40 轮 × 8 线程）显示：**只连一次**得 24 次 readonly 失败，**失败重开**得 0 次——真正生效的机制是**按准确错误类别重开连接**，与种子写无关。据此已删除种子写（`_initialize_schema` 不再插入占位行），并新增 `test_new_instance_on_existing_db_does_not_collide_with_open_writer`、`test_read_is_not_serialized_behind_concurrent_write_lock` 两条判别测试与 M27/M28 独立变异。

## 不提供的保证

- 这项决定本身不提供节点续跑、断电耐久性或 Office.js 跨重启恢复。
- 锁不能阻止用户或其他应用修改文件，写前条件和写后验证仍必须保留。
- 旧 manifest 没有任务现场，只能导入为恢复证据，不能补造可续跑任务。

## 验收重点

- 同 operationId 重放不重复执行；内容变化返回冲突。
- 内存记录被清理后，完成事件和结果仍可补收。
- 服务重启留下的 running 记录不能自动重写文件。
- 不同 operationId 对同一工作目录的执行、撤销或恢复互斥。
- 测试通过环境变量使用隔离应用数据目录，不写入真实用户目录。
