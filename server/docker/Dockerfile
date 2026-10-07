# 隔离 Python 分析镜像（固定、可追溯）。
# 构建后分别记录本机 image ID 与 RepoDigests，不把两者当成同一种摘要。
# 模型代码只能通过 stdin 传入，由 /app/worker.py 固定入口执行。
#
# 构建上下文（重要）：本 Dockerfile 位于 server/docker/，但工作程序源文件在
# server/app/docker_worker.py。因此构建时**上下文必须是 server/**，命令为：
#     python -m scripts.build_python_sandbox
# 该脚本从 server/ 只选三份明确文件，核实本机构建器并锁定基础镜像 digest。
# 若误用 `docker build server/docker` 作上下文，下方的 COPY 会找不到源文件而失败。
# 这与 test_dockerfile_copy_source_exists_in_build_context 的校验口径一致
# （该测试以 server/docker 为上下文，会要求源文件在可寻址位置）。

# 必须传入经过核实的不可变基础镜像引用；没有 tag 回退。
ARG PYTHON_BASE
FROM ${PYTHON_BASE}

# 预装分析依赖（固定版本，运行时不联网安装）
RUN pip install --no-cache-dir numpy==2.1.3 pandas==2.2.3

# 非 root 用户（容器内 nobody:nogroup = 65534:65534）
# 工作程序与可写 scratch 目录
RUN mkdir -p /app /scratch && chown -R 65534:65534 /app /scratch && chmod 1777 /scratch

# 固定入口程序：源文件随构建上下文拷贝（上下文 = server/）
COPY app/docker_worker.py /app/worker.py

# 不设 CMD/ENTRYPOINT 到镜像层——由可信侧 docker run 显式指定固定入口
# （协调器决定 python3 /app/worker.py，模型不能改）

USER 65534:65534
