"""隔离 Python 分析的 Docker 运行时适配器（三层分离：协调器 → 适配器 → 工作程序）。
设计依据：``docs/SDK_ISOLATED_PYTHON_DESIGN_2026-10-06.md`` §4。
本模块是**可替换的运行时适配器**：Docker 是第一种实现，不是唯一前提。
模型只能提供待分析代码；镜像名、启动参数、挂载、网络、用户、限额、环境变量
与清理目标全部由本模块的可信代码决定。
安全约束（全部由环境强制 + 可信侧核对，不靠提示词/monkeypatch）：
- ``--network none``：无互联网、局域网、宿主回环；
- ``--read-only`` + ``--tmpfs /scratch``：只读根、仅限大小的临时可写目录；
- ``--user``（非 root，拒绝一切 root UID 形式）；
- ``--memory``/``--memory-swap``/``--cpus``/``--pids-limit``：内存/swap/CPU 速率/进程线程限额；
- ``--cap-drop ALL`` + ``--security-opt no-new-privileges``：去 capabilities、禁提权；
- ``--log-driver none``：容器输出不落进宿主 Docker 日志（模型输出可能是隐私）；
- 不挂载用户目录、仓库、工作簿、凭据、Docker socket 或控制接口；
- 不传宿主环境变量，只显式传工作程序需要的最小变量；
- **绑定并锁定本机端点**：先用 ``docker context inspect`` + ``docker info`` 核实
  实际端点是本机 npipe/unix socket（拒绝 tcp/ssh/未知），再把整个生命周期
  （health/image/create/start/stop/rm/verify）锁到该端点；不修改用户全局 context。
IPC：固定入口程序从 stdin 读有大小上限的 JSON，把有类型结果写到 stdout。
不使用 pickle，不拼 shell，不把模型代码放进命令行参数。
有界与清理（第八批修复）：
- 输出**运行期有界**：reader 线程并发读取 stdout/stderr，各自独立字节上限；
  任一通道触顶立即 kill 容器并进入终止核对——不再把内存风险转移到宿主磁盘，
  也不等进程自然结束。原始输出不落宿主临时文件/日志/历史/账本；
- 每次运行绑定**可信侧运行身份**（唯一容器名 + docker create 得到的容器 ID），
  成功/失败/超时/超限/取消/CLI 异常均有明确生命周期；
- **清理可核实**：``rm -f`` 后必须查询确认容器消失，区分「确认已消失
  （gone）」与「无法查询（unknown）」；清理未确认时不发布结果、不复用沙箱；
- 取消支持 ``cancel_check`` 回调，清理后按原契约传播（不吞成成功/普通不可用），
  取消后的迟到结果一律拒绝；
- 回包**严格校验**：status 枚举、``completed`` 必须有 ``result`` 键
  （显式 null 合法，缺键不合法）、拒绝 NaN/Infinity/过深/元素爆炸；
  ``status=error``（即便退出码 0）一律映射 ``error``。
本模块**不做**：datasetHandle 授权登记、resultHandle 持久结果库、
结果转 AnalysisPlan、Excel/WPS UI——这些仍在协调器之外，未实现。
"""
from __future__ import annotations
import json
import math
import os
import re
import subprocess
import threading
import uuid
from dataclasses import dataclass, field
# ---------------------------------------------------------------------------
# 固定安全类别文案（供协调器映射为有类型回执，不携带任何自由文本）。
# ---------------------------------------------------------------------------
RUNNER_UNAVAILABLE = "隔离 Python 运行时不可用：本机未检测到可用的 Docker Linux 容器引擎。"
ENGINE_NOT_VERIFIED = "隔离 Python 运行时不可用：本机 Docker 引擎未通过身份验证，拒绝执行。"
ENDPOINT_NOT_LOCAL = "隔离 Python 运行时不可用：Docker 端点不是经核实的本机端点，拒绝执行。"
IMAGE_MISSING = "隔离 Python 运行时不可用：所需的隔离镜像尚未就绪。"
IMAGE_NOT_IMMUTABLE = "隔离 Python 运行时不可用：隔离镜像缺少不可变身份（摘要），拒绝执行。"
ISOLATION_NOT_APPLIED = "隔离 Python 运行时不可用：隔离参数未落实，拒绝执行。"
EXECUTION_ERROR = "隔离 Python 分析在执行环节发生内部错误，本轮未运行任何代码。"
OUTPUT_LIMIT_EXCEEDED = "隔离 Python 分析输出超过上限，本轮结果未发布。"
CLEANUP_UNVERIFIED = "隔离 Python 运行时未能在执行后核实容器已清理，本轮结果未发布。"
# 工作程序回包允许的状态枚举；``completed`` 必须带 ``result``。
_WORKER_OK_STATUS = "completed"
_WORKER_ERROR_STATUS = "error"
# 容器名前缀（可信侧生成，模型不能影响）。
_CONTAINER_NAME_PREFIX = "excel-bro-py-"
# 结果结构边界（深度/元素数量）默认值——真值由 capabilities 提供。
_DEFAULT_MAX_RESULT_DEPTH = 32
_DEFAULT_MAX_RESULT_ELEMENTS = 200000
# 最小、**显式**的 CLI 环境：剔除所有可能与远端端点/凭据相关的宿主变量。
_CLI_ENV_ALLOWED_KEYS = ("PATH", "SystemRoot", "TEMP", "TMP", "COMSPEC")
_CLI_ENV_BLOCKED = ("DOCKER_HOST", "DOCKER_CONTEXT", "DOCKER_TLS_VERIFY", "DOCKER_CERT_PATH",
                    "DOCKER_CONFIG", "DOCKER_API_VERSION")
# 本机端点 scheme：仅接受 Unix domain socket / Windows named pipe。
# 注意：**不能只判前缀**——形似本机实为远端的串（如 npipe:////remote-host/...）
# 前缀同样合法。真实判定统一走 :func:`_classify_local_endpoint` 的结构校验；
# 该常量保留为变异脚本（反向变异 M15）的对照种子。
_LOCAL_ENDPOINT_SCHEMES = ("npipe://", "unix://")
def _cli_env() -> dict[str, str]:
    """构造 docker CLI 的最小环境：只保留必要键，**显式剔除**远端端点变量。
    Docker CLI 会读取 ``DOCKER_HOST``/``DOCKER_CONTEXT`` 指向远程守护进程。
    从零构造环境、不继承宿主，并剔除被屏蔽键。
    注意：这只屏蔽**环境变量**通道；持久化的 current context 仍需
    :func:`verify_local_endpoint` 用实际端点核实（见第八批缺陷 B3）。
    """
    env: dict[str, str] = {}
    source = os.environ
    for key in _CLI_ENV_ALLOWED_KEYS:
        if key in source:
            env[key] = source[key]
    if "PATH" not in env:
        env["PATH"] = "/usr/bin:/bin"
    for blocked in _CLI_ENV_BLOCKED:
        env.pop(blocked, None)
    return env
# ---------------------------------------------------------------------------
# 端点核实与锁定
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class VerifiedEndpoint:
    """经核实的**本机** Docker 端点。整个生命周期锁到这里，不再重新解析。"""
    scheme: str            # "npipe" | "unix"
    host: str              # 原始 Host 串，仅用于展示/审计
    cli_context: str | None = None  # 若有具名本机 context，可显式锁定
    def cli_prefix(self) -> list[str]:
        """返回锁定端点的 CLI 参数前缀（显式 -H，避免运行期重新解析 context）。"""
        return ["-H", self.host]
def _run_docker_cli(args: list[str], *, timeout: float = 30.0,
                    endpoint: VerifiedEndpoint | None = None) -> subprocess.CompletedProcess[str]:
    """用固定 docker CLI 执行，只接受可信侧构造的参数列表。
    不使用 shell=True；不继承宿主环境（见 :func:`_cli_env`）。
    ``endpoint`` 非空时把 ``-H <host>`` 插到参数最前，把本次调用锁到经验证的
    本机端点（不在运行期重新解析 current context）。
    """
    full = ["docker"]
    if endpoint is not None:
        # 最内层防线：任何带端点的 CLI 调用都必须先通过本机结构校验，
        # 非法端点一律拒绝（不冒充本机、不锁定到远端）。
        if not endpoint_is_locally_verified(endpoint):
            raise DockerRuntimeError("endpoint_not_local", ENDPOINT_NOT_LOCAL)
        full += endpoint.cli_prefix()
    full += args
    return subprocess.run(
        full,
        capture_output=True,
        text=True,
        timeout=timeout,
        shell=False,
        env=_cli_env(),
    )
def _parse_endpoint_host(raw: str) -> str:
    """从 ``docker context inspect``/``docker info`` 输出里提取 **docker 端点** 的 Host。

    优先解析结构化字段 ``Endpoints.docker.Host``（context inspect 的权威来源），
    避免被同一输出里**先出现的无关 URL**（如 Metadata.otel 的 ``unix://...``）
    误导——那些不是 docker 守护进程端点，锁进生命周期会打到错误目标。
    结构化字段缺失时才退回正则（用于 ``docker info`` 的自由文本形态）。
    """
    # 1) 权威路径：结构化 JSON 的 Endpoints.docker.Host
    try:
        parsed = json.loads(raw)
    except (ValueError, TypeError):
        parsed = None
    candidates = parsed if isinstance(parsed, list) else [parsed] if parsed is not None else []
    for item in candidates:
        if not isinstance(item, dict):
            continue
        endpoints = item.get("Endpoints")
        if isinstance(endpoints, dict):
            docker = endpoints.get("docker")
            if isinstance(docker, dict) and isinstance(docker.get("Host"), str):
                host = docker["Host"].strip()
                if host:
                    return host
    # 2) 退回正则（docker info 的自由文本；仍是 ``docker info`` 的端点形态）
    m = re.search(r"((?:npipe|unix|tcp|ssh|fd)://[^\s\"']+)", raw)
    return m.group(1) if m else ""
def _classify_local_endpoint(host: str) -> tuple[bool, str]:
    """**结构校验**端点是否为本机 docker 端点。

    不能只用 ``startswith("npipe://")``/``"unix://"`` 判前缀：形如
    ``npipe:////remote-host/pipe/docker_engine`` 的串前缀合法、实质指向远端主机，
    会被前缀判定误接受。此函数按各形态的**结构**逐段校验，只接受显式支持的本机
    形态，并拒绝远端 authority、空路径、畸形地址与无法确认的写法。

    返回 ``(True, "")`` 或 ``(False, 原因文案)``。不把任意主机别名改写为 local。
    """
    if not isinstance(host, str):
        return False, "端点 Host 不是字符串"
    raw = host.strip()
    if not raw:
        return False, "端点 Host 为空"
    # 只按 "scheme://" 切开，保留 scheme 之后剩余部分（可能自 `/` 或空开始）。
    if "://" not in raw:
        return False, f"端点缺少 scheme 分隔符：{raw.split(':', 1)[0]}"
    scheme, remainder = raw.split("://", 1)
    scheme = scheme.strip().lower()
    if scheme == "npipe":
        return _classify_local_named_pipe(remainder)
    if scheme == "unix":
        return _classify_local_unix_socket(remainder)
    return False, f"Docker 端点不是本机 npipe/unix socket：{scheme or '(空)'}"
def _classify_local_named_pipe(remainder: str) -> tuple[bool, str]:
    r"""校验 Windows 命名管道形态：只接受本机 ``\\.\pipe\<name>``。

    合法写法（舍去 ``npipe://`` 前缀后的剩余部分）：
    - ``//./pipe/<name>``    → 本机命名空间 ``\\.\pipe\<name>``
    - ``//?/pipe/<name>``    → 扩展长度前缀 ``\\?\pipe\<name>``
    远端主机名（``//remote-host/pipe/...``）或路径不完整一律拒绝。
    """
    rest = remainder.strip()
    if not rest.startswith("//"):
        return False, f"本机命名管道必须以 // 开头：{rest or '(空)'}"
    # 去前导 "//"，得到 "<authority>/pipe/<name>"
    body = rest[2:]
    if not body:
        return False, "本机命名管道缺少主机段"
    # 扩展长度前缀写法 "\\?\pipe\..." 在剥去 "npipe://" 后是 "?/pipe/..."，
    # 与本机命名空间 ".\pipe\..." 等价，统一归一为 "./..." 再校验。
    if body.startswith("?/"):
        body = "./" + body[2:]
    if not body.startswith("./"):
        authority = body.split("/", 1)[0]
        return False, f"命名管道不是本机形态（主机段 {authority!r}）：{rest}"
    tail = body[2:]  # 去掉 "./"
    if not tail.startswith("pipe/"):
        return False, f"本机命名管道缺少 pipe/ 段：{rest}"
    pipe_name = tail[len("pipe/"):].strip()
    if not pipe_name or "/" in pipe_name or pipe_name in (".", ".."):
        return False, f"本机命名管道名为空或非法：{rest}"
    return True, ""
def _classify_local_unix_socket(remainder: str) -> tuple[bool, str]:
    """校验 Unix domain socket 形态：必须是**本机绝对路径**，不得带远端 authority。

    - ``/var/run/docker.sock``         → 绝对路径，接受
    - ``/run/user/1000/docker.sock``   → 绝对路径，接受
    - ``//remote-host/...``            → 以 // 开头 = 远端 authority 或 UNC，拒绝
    - ``remote-host/...``              → 相对路径，拒绝
    """
    rest = remainder.strip()
    if not rest:
        return False, "Unix socket 路径为空"
    if rest.startswith("//"):
        return False, f"Unix socket 不得带主机名（远端 authority）：{rest}"
    if not rest.startswith("/"):
        return False, f"Unix socket 必须是本机绝对路径：{rest}"
    if rest == "/" or rest.rstrip("/") == "":
        return False, "Unix socket 路径为根目录，非法"
    return True, ""
def verify_local_endpoint() -> tuple[bool, str | None]:
    """核实 docker 实际端点是否是**经确认的本机**端点。
    仅剔除 ``DOCKER_HOST``/``DOCKER_CONTEXT`` 环境变量不等于绑定本机：Docker 仍会
    读取**持久化的 current context**（可在磁盘上指向远端）。因此必须查询实际端点，
    拒绝 tcp/ssh/fd/未知端点，只接受本机 npipe/unix socket。
    返回 ``(True, host)`` 或 ``(False, 原因文案)``。
    """
    # 1) 查当前 context 的端点 Host（context inspect 自身不带 -H，先取默认解析）
    try:
        ctx = _run_docker_cli(["context", "inspect"], timeout=15.0)
    except (subprocess.TimeoutExpired, OSError):
        return False, "无法查询 Docker context"
    if ctx.returncode == 0:
        host = _parse_endpoint_host(ctx.stdout)
    else:
        # 无 context（旧版/异常）→ 退回 docker info 的端点字段
        try:
            info = _run_docker_cli(["info", "--format", "{{.DockerRootDir}}"], timeout=15.0)
        except (subprocess.TimeoutExpired, OSError):
            return False, "无法查询 Docker 引擎"
        host = ""
    # 2) 若 context 未给出 Host（如默认 context），用 info 的 ClientInfo/Swarm 段兜底
    if not host:
        try:
            raw = _run_docker_cli(["info", "--format", "{{json .}}"], timeout=15.0)
        except (subprocess.TimeoutExpired, OSError):
            return False, "无法查询 Docker 端点"
        if raw.returncode != 0:
            return False, "Docker 引擎不可用"
        host = _parse_endpoint_host(raw.stdout)
    if not host:
        # 端点不可确认 → 失败关闭（不假定本机）
        return False, "无法确认 Docker 端点是否为本机"
    ok, reason = _classify_local_endpoint(host)
    if not ok:
        return False, reason
    return True, host
def verified_endpoint() -> VerifiedEndpoint:
    """核实并构造锁定的本机端点。失败抛 :class:`DockerRuntimeError`。"""
    ok, host = verify_local_endpoint()
    if not ok or not host:
        raise DockerRuntimeError("endpoint_not_local", ENDPOINT_NOT_LOCAL)
    scheme = host.split("://", 1)[0]
    return VerifiedEndpoint(scheme=scheme, host=host)
def endpoint_is_locally_verified(endpoint: "VerifiedEndpoint") -> bool:
    """再次校验端点实例本身是否为本机形态。

    执行路径**不得**仅凭对象名 ``VerifiedEndpoint`` 就信任——尤其是外部/测试构造的
    实例。此函数对该实例的 host 重新做结构校验，只有通过才允许进入执行路径。
    """
    if not isinstance(endpoint, VerifiedEndpoint):
        return False
    ok, _reason = _classify_local_endpoint(endpoint.host)
    return ok
def endpoint_cli_prefix(endpoint: VerifiedEndpoint) -> list[str]:
    """返回把一次 docker CLI 调用锁到经验证端点的参数前缀。
    公开别名（供协调器与测试引用）：一旦端点被核实，整个运行生命周期
    （健康检查/镜像检查/运行/终止/删除/核查）都必须带同一前缀，不得在运行期
    重新解析 current context（否则前后使用不同端点 → 清理可能打到别处）。
    """
    if not endpoint_is_locally_verified(endpoint):
        raise DockerRuntimeError("endpoint_not_local", ENDPOINT_NOT_LOCAL)
    return endpoint.cli_prefix()
# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class DockerRuntimeConfig:
    """可信侧固定的容器运行配置。模型不能改动任何字段。"""
    image: str
    image_digest: str | None = None
    non_root_user: str = "65534:65534"  # nobody:nogroup
    network: str = "none"
    read_only: bool = True
    scratch_dir: str = "/scratch"
    scratch_size: str = "256m"
    cap_drop: str = "ALL"
    no_new_privileges: bool = True
    worker_entrypoint: str = "python3"
    worker_script: str = "/app/worker.py"
    log_driver: str = "none"
    def image_ref(self) -> str:
        """返回**不可变**镜像引用。
        必须带 digest（``name@sha256:...``）。缺少 digest 时**不静默降级**为可变 tag，
        而是返回一个**不可能被 Docker 解析**的哨兵（带 ``@sha256:`` 前缀但没有有效
        摘要），确保任何下游调用都会失败而不是悄悄用可变 tag 运行。
        调用方应先经 :func:`validate_isolation_config` 拒绝该配置。
        """
        if self.image_digest:
            return f"{self.image}@{self.image_digest}"
        # 失败关闭：返回带摘要语法但摘要缺失的引用，Docker 无法解析它。
        return f"{self.image}@sha256:"
    def is_immutable(self) -> bool:
        return bool(self.image_digest and self.image_digest.startswith("sha256:"))
@dataclass
class DockerRunResult:
    """有类型的执行结果。异常/超时/取消都映射为固定状态，不回显原始文本。"""
    status: str  # "completed" | "timeout" | "cancelled" | "limit_exceeded" | "error"
    exit_code: int | None
    result_payload: dict | None  # 工作程序返回的有类型结果
    reason_category: str | None  # 固定类别，供协调器映射回执
    def is_success(self) -> bool:
        return self.status == "completed" and self.result_payload is not None
class DockerRuntimeError(RuntimeError):
    """运行时不可用/配置损坏/隔离未落实时的失败关闭异常。
    消息只携带固定类别文案，绝不携带原始异常文本、容器 stdout/stderr 或模型代码。
    """
    def __init__(self, reason_category: str, message: str) -> None:
        super().__init__(message)
        self.reason_category = reason_category
def _docker_config() -> DockerRuntimeConfig:
    """从 capabilities.json 读取 Docker 运行配置（可信侧固定，模型不能改）。"""
    from .capabilities import capabilities
    section = "pythonDocker"
    cfg = capabilities().get(section, {})
    image = cfg.get("image", "")
    if not image:
        raise DockerRuntimeError("isolation_not_applied", ISOLATION_NOT_APPLIED)
    return DockerRuntimeConfig(
        image=image,
        image_digest=cfg.get("imageDigest") or None,
        non_root_user=cfg.get("nonRootUser", "65534:65534"),
    )
# ---------------------------------------------------------------------------
# 配置校验（含 root UID 变体、镜像不可变身份）
# ---------------------------------------------------------------------------
def _is_root_user_spec(spec: str) -> bool:
    """判断 user 规格是否是 root（或包含 root UID）。
    覆盖：``0``、``0:0``、``root``、``root:root``、``0:1000``、``root:1000``、
    ``00:0``、``0:nogroup`` 等——只要 **uid 段**解析为 0 或 root 即判 root。
    """
    if not isinstance(spec, str):
        return True
    text = spec.strip()
    if not text:
        return True
    uid_part = text.split(":", 1)[0].strip()
    if uid_part.lower() in ("root", ""):
        return True
    # 纯数字（允许前导零/加号）→ 数值 0 即 root
    if re.fullmatch(r"[0-9]+", uid_part):
        try:
            return int(uid_part) == 0
        except ValueError:
            return True
    return False
def validate_isolation_config(config: DockerRuntimeConfig) -> tuple[bool, str | None]:
    """逐项校验隔离配置是否完整落实。缺任一项即返回 ``(False, 原因)``。
    只判断、不改值：这是「隔离未落实就不执行」的显式门。
    """
    problems: list[str] = []
    if config.network != "none":
        problems.append("network 必须为 none")
    if not config.read_only:
        problems.append("根文件系统必须只读")
    if config.cap_drop != "ALL":
        problems.append("必须 --cap-drop ALL")
    if not config.no_new_privileges:
        problems.append("必须启用 no-new-privileges")
    if _is_root_user_spec(config.non_root_user):
        problems.append("必须以非 root 运行（拒绝一切 root UID 形式）")
    if config.log_driver != "none":
        problems.append("必须禁用容器日志驱动")
    if not config.image:
        problems.append("必须指定固定镜像")
    elif not config.is_immutable():
        problems.append("必须使用带摘要的不可变镜像身份（imageDigest）")
    if not config.worker_script.startswith("/"):
        problems.append("工作程序入口必须是镜像内绝对路径")
    if problems:
        return False, "; ".join(problems)
    return True, None
# ---------------------------------------------------------------------------
# 引擎与镜像
# ---------------------------------------------------------------------------
def verify_engine(endpoint: VerifiedEndpoint | None = None) -> bool:
    """验证本机 Docker Linux 引擎可用（客户端 + 服务端通信 + 本机 Linux 容器）。
    ``.OSType == linux`` **不是**本机身份的证明，仅说明容器类型；本机身份由
    :func:`verify_local_endpoint` 单独核实。返回 False 即失败关闭。
    """
    try:
        proc = _run_docker_cli(["info", "--format", "{{.OSType}}"], timeout=15.0,
                               endpoint=endpoint)
    except (subprocess.TimeoutExpired, OSError):
        return False
    if proc.returncode != 0:
        return False
    return proc.stdout.strip() == "linux"
def image_present(config: DockerRuntimeConfig, endpoint: VerifiedEndpoint | None = None) -> bool:
    """检查所需镜像是否存在（按固定引用，不下载、不构建）。"""
    if not config.is_immutable():
        return False
    try:
        proc = _run_docker_cli(["image", "inspect", config.image_ref()], timeout=15.0,
                               endpoint=endpoint)
    except (subprocess.TimeoutExpired, OSError):
        return False
    return proc.returncode == 0
# ---------------------------------------------------------------------------
# 启动参数
# ---------------------------------------------------------------------------
def _require_positive_int(value: object, label: str) -> int:
    """资源**整数配额**严格校验：拒绝 bool/非整数/零/负数/缺失。

    bool 是 int 子类，`True` 会被 `isinstance(int)` 误接受，故先剔 bool。
    绝不把非法值静默钳位或降级为默认/无限。
    """
    if isinstance(value, bool) or not isinstance(value, int):
        raise DockerRuntimeError("invalid_resource_limit", "隔离资源限额必须是正整数")
    if value <= 0:
        raise DockerRuntimeError("invalid_resource_limit", "隔离资源限额必须是正整数")
    return value
def _require_positive_finite_float(value: object, label: str) -> float:
    """资源**浮点配额**严格校验：拒绝 bool/非数值/NaN/Infinity/零/负数。"""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise DockerRuntimeError("invalid_resource_limit", "隔离资源限额必须是有限正数")
    numeric = float(value)
    if not math.isfinite(numeric) or numeric <= 0:
        raise DockerRuntimeError("invalid_resource_limit", "隔离资源限额必须是有限正数")
    return numeric
def _validate_resource_limits(
    *,
    cpu_quota: object,
    cpu_seconds: object,
    memory_mb: object,
    pids_limit: object,
    scratch_mb: object,
) -> None:
    """在**构造任何 CLI / 起 Docker 之前**校验全部资源限额。

    不信任类型注解，也不信任已格式化好的 CLI 字符串——调用方可能传入
    bool/NaN/Infinity/零/负数/字符串。非法一律抛错（调用点必须接住返回值语义：
    这里抛异常即中断，不会继续启动容器）。
    """
    _require_positive_finite_float(cpu_quota, "cpu_quota")
    _require_positive_int(cpu_seconds, "cpu_seconds")
    _require_positive_int(memory_mb, "memory_mb")
    _require_positive_int(pids_limit, "pids_limit")
    _require_positive_int(scratch_mb, "scratch_mb")
def build_run_arguments(
    config: DockerRuntimeConfig,
    *,
    cpu_quota: float,
    cpu_seconds: int,
    memory_mb: int,
    pids_limit: int,
    scratch_mb: int,
    container_name: str,
) -> tuple[list[str], dict[str, str]]:
    """构造固定启动链参数（模型不能影响任何一项）。
    返回 (docker run 参数列表, 最小环境变量)。
    - ``cpu_quota``（``--cpus``）是 CPU **速率**上限，不等于累计 CPU 时间；
      累计 CPU 时间由 ``cpu_seconds`` 表达（经 ``--ulimit cpu=<s>`` 落实，
      RLIMIT_CPU 超限则内核终止进程）。两者语义不同，都接线、都校验。
    - ``pids_limit``（``--pids-limit``）限制进程/线程数，**不表示禁止 exec**；
      禁止外部程序执行由 ``--network none``、只读根、去 capabilities 与
      no-new-privileges 共同限制，其"完全禁止 exec"未落实（见文档记录）。
    - 输入 JSON 通过 stdin 传给工作程序，绝不放进命令行参数。
    """
    # 适配器入口：构造 CLI 之前先严格校验资源限额（拒绝 bool/NaN/Inf/零/负/非整）。
    _validate_resource_limits(
        cpu_quota=cpu_quota,
        cpu_seconds=cpu_seconds,
        memory_mb=memory_mb,
        pids_limit=pids_limit,
        scratch_mb=scratch_mb,
    )
    args: list[str] = [
        "docker", "run", "--rm",
        "--name", container_name,
        "--network", config.network,
        "--user", config.non_root_user,
        "--cap-drop", config.cap_drop,
        "--security-opt", "no-new-privileges",
        "--log-driver", config.log_driver,  # none：输出不落宿主 Docker 日志
        "--memory", f"{memory_mb}m",
        "--memory-swap", f"{memory_mb}m",  # 禁用 swap：swap == memory
        "--cpus", f"{cpu_quota}",
        "--ulimit", f"cpu={int(cpu_seconds)}:{int(cpu_seconds)}",
        "--pids-limit", str(pids_limit),
        "--read-only",
        "--tmpfs", f"{config.scratch_dir}:size={scratch_mb}m,mode=1777",
        # Fixed container values, never inherited host values. Numerical libraries
        # must not try to spawn hardware-count threads under the one-process limit.
        "--env", "OPENBLAS_NUM_THREADS=1",
        "--env", "OMP_NUM_THREADS=1",
        "--env", "MKL_NUM_THREADS=1",
        "--env", "NUMEXPR_MAX_THREADS=1",
        "--env", "NUMEXPR_NUM_THREADS=1",
        "--env", "PYTHONUNBUFFERED=1",
        "--env", "PYTHONDONTWRITEBYTECODE=1",
        "-i",
        config.image_ref(),
        config.worker_entrypoint,
        config.worker_script,
    ]
    env: dict[str, str] = {
        "PYTHONUNBUFFERED": "1",
        "PYTHONDONTWRITEBYTECODE": "1",
    }
    return args, env
def _new_container_name() -> str:
    """可信侧生成唯一容器名（含随机 UUID），并发/重放不会误停他人容器。"""
    return f"{_CONTAINER_NAME_PREFIX}{uuid.uuid4().hex[:16]}"
# ---------------------------------------------------------------------------
# 有界输出收集器（并发、双通道独立上限、提前终止）
# ---------------------------------------------------------------------------
class _OutputCollector:
    """并发读取容器 stdout/stderr 的**有界**收集器。
    - 每个通道独立字节上限；一旦触顶立即置 ``over_limit``，不再累积
      （内存封顶；也不落宿主磁盘——只在内存里保留有界片段）。
    - 段落只保留前 ``cap`` 字节（多 1 字节哨兵用于判定"确实超了"）。
    - 不把原始输出写入宿主临时文件/日志；仅在有界内存中暂存，用于协议解析。
    """
    def __init__(self, stdout_cap: int, stderr_cap: int) -> None:
        self._stdout_cap = max(1, int(stdout_cap))
        self._stderr_cap = max(1, int(stderr_cap))
        self._stdout = bytearray()
        self._stderr = bytearray()
        self._over = False
        self._lock = threading.Lock()
    def feed_stdout(self, chunk: bytes) -> None:
        with self._lock:
            if self._over:
                return
            room = self._stdout_cap + 1 - len(self._stdout)
            if room > 0:
                self._stdout.extend(chunk[:room])
            if len(self._stdout) > self._stdout_cap:
                self._over = True
    def feed_stderr(self, chunk: bytes) -> None:
        with self._lock:
            if self._over:
                return
            room = self._stderr_cap + 1 - len(self._stderr)
            if room > 0:
                self._stderr.extend(chunk[:room])
            if len(self._stderr) > self._stderr_cap:
                self._over = True
    def over_limit(self) -> bool:
        with self._lock:
            return self._over
    def stdout_bytes(self) -> bytes:
        with self._lock:
            return bytes(self._stdout)
    def stderr_bytes(self) -> bytes:
        with self._lock:
            return bytes(self._stderr)
    def stdout_size(self) -> int:
        with self._lock:
            return len(self._stdout)
    def stderr_size(self) -> int:
        with self._lock:
            return len(self._stderr)
def _pump(stream, collector: _OutputCollector, is_stdout: bool) -> None:
    """后台线程：并发把容器流的有界片段灌进收集器。"""
    feed = collector.feed_stdout if is_stdout else collector.feed_stderr
    # BufferedReader.read(n) waits for n bytes or EOF; read1 preserves short,
    # flushed output while the worker is still running. Raw streams retain read.
    read_available = getattr(stream, "read1", stream.read)
    try:
        while True:
            chunk = read_available(4096)
            if not chunk:
                break
            feed(chunk)
            if collector.over_limit():
                # 触顶后继续读走（防止子进程管道写阻塞），但不再累积
                continue
    except (ValueError, OSError):
        return
def _feed_stdin(stream, data: bytes, state: dict) -> None:
    """后台线程：把输入 JSON 写入容器 stdin 后关闭。

    **必须与读取线程并发**——若在主线程同步写完 stdin，容器不消费输入（或管道
    背压）会阻塞在这里，wall timeout 与取消都来不及生效（第十批 F1）。

    结果写入 ``state``：
    - ``state["fed"]=True``：输入已完整写入并关闭；
    - ``state["error"]=str``：写入失败（BrokenPipe/OSError）——**不得吞掉后
      当作成功继续**，由生命周期判定映射为固定失败类别；
    - ``state["done"]=True``：线程结束（无论成败）。
    """
    try:
        stream.write(data)
        stream.close()
        state["fed"] = True
    except (BrokenPipeError, OSError, AttributeError, ValueError) as exc:
        # 容器可能在读完输入前就退出/被杀（如立即崩溃），写入失败要如实记录。
        state["error"] = type(exc).__name__
    finally:
        state["done"] = True
# ---------------------------------------------------------------------------
# 清理（可核实）
# ---------------------------------------------------------------------------
CLEANUP_GONE = "gone"
CLEANUP_UNKNOWN = "unknown"
def _container_exists(container_name: str, endpoint: VerifiedEndpoint | None = None) -> bool | None:
    """查询容器是否存在。返回 True/False；查询失败返回 None（未知）。"""
    try:
        proc = _run_docker_cli(["container", "inspect", container_name], timeout=15.0,
                               endpoint=endpoint)
    except (subprocess.TimeoutExpired, OSError):
        return None
    if proc.returncode == 0:
        return True
    # 明确的 "no such container" → 不存在
    if "no such" in (proc.stderr or "").lower():
        return False
    return None
def _terminate_process(proc) -> None:
    """终止 CLI 进程并收束管道/进程句柄（不落宿主文件）。

    用于所有提前终止（超时/取消/超限/输入失败）与 finally 收尾：kill + wait，
    再尽力关闭 stdin/stdout/stderr，避免管道句柄泄漏。
    """
    if proc is None:
        return
    try:
        if proc.poll() is None:
            proc.kill()
    except (OSError, ValueError):
        pass
    try:
        proc.wait(timeout=10)
    except Exception:  # noqa: BLE001 - 收尾路径不抛
        pass
    for name in ("stdin", "stdout", "stderr"):
        stream = getattr(proc, name, None)
        close = getattr(stream, "close", None)
        if callable(close):
            try:
                close()
            except (OSError, ValueError):
                pass


def _cleanup_container(container_name: str,
                       endpoint: VerifiedEndpoint | None = None) -> str:
    """强制清理具名容器并**核实结果**。只作用于本容器（不用 prune/宽泛前缀）。
    返回：
    - ``"gone"``：确认容器已消失；
    - ``"unknown"``：清理失败或核查无法确认（**不得**当作清理成功）。
    """
    # 1) 强制删除（幂等：容器不存在也算成功）
    try:
        _run_docker_cli(["rm", "-f", container_name], timeout=20.0, endpoint=endpoint)
    except (subprocess.TimeoutExpired, OSError):
        pass
    # 2) 必须核查：查询容器是否仍存在
    exists = _container_exists(container_name, endpoint=endpoint)
    if exists is False:
        return CLEANUP_GONE
    # 仍存在或查询失败 → 未知（不冒充成功）
    return CLEANUP_UNKNOWN


def _finish_with_cleanup(
    container_name: str,
    endpoint: VerifiedEndpoint | None,
    *,
    status: str,
    exit_code: int | None,
    reason_category: str | None,
    cleanup_state: str | None = None,
) -> "DockerRunResult":
    """统一收束：**任何**分支结束前都必须核实清理结果并消费它。

    「已尝试清理」≠「已确认清理」。若核查结果不是 ``gone``（即 ``unknown``），
    一律返回固定 ``cleanup_unverified`` 语义——**不冒充**普通 cancelled /
    limit_exceeded / completed，且**不发布结果**。原始中断原因保留在
    ``reason_category`` 之外不对外回显（不回显自由文本）。

    ``cleanup_state`` 可由调用方传入（避免重复核查）；为 ``None`` 时本函数自行核查。
    """
    state = cleanup_state if cleanup_state is not None else _cleanup_container(
        container_name, endpoint=endpoint
    )
    if state != CLEANUP_GONE:
        # 清理未确认：不发布结果、不复用沙箱，保留固定类别（不带自由文本）。
        return DockerRunResult("error", exit_code, None, "cleanup_unverified")
    return DockerRunResult(status, exit_code, None, reason_category)
# ---------------------------------------------------------------------------
# 回包严格校验（含结构边界、非有限数）
# ---------------------------------------------------------------------------
def _json_depth(value, _depth: int = 0, _limit: int = _DEFAULT_MAX_RESULT_DEPTH) -> int:
    """计算 JSON 结构最大深度（用于拒绝过深结果）。"""
    if _depth > _limit:
        return _depth
    if isinstance(value, dict):
        return max([_depth] + [_json_depth(v, _depth + 1, _limit) for v in value.values()])
    if isinstance(value, (list, tuple)):
        return max([_depth] + [_json_depth(v, _depth + 1, _limit) for v in value])
    return _depth
def _json_element_count(value) -> int:
    """计算 JSON 元素数量（容器 + 标量，用于拒绝元素爆炸）。"""
    if isinstance(value, dict):
        return 1 + sum(_json_element_count(v) for v in value.values())
    if isinstance(value, (list, tuple)):
        return 1 + sum(_json_element_count(v) for v in value)
    return 1
def _has_non_finite(value) -> bool:
    """递归检测 float 非有限值（NaN/Infinity 不允许出现在结果里）。"""
    if isinstance(value, float):
        return value != value or value in (float("inf"), float("-inf"))
    if isinstance(value, dict):
        return any(_has_non_finite(v) for v in value.values())
    if isinstance(value, (list, tuple)):
        return any(_has_non_finite(v) for v in value)
    return False
def _validate_result_structure(value, *, max_depth: int, max_elements: int) -> bool:
    """结果结构边界：深度、元素数量、非有限数。合法返回 True。"""
    if _has_non_finite(value):
        return False
    if _json_depth(value, _limit=max_depth) > max_depth:
        return False
    if _json_element_count(value) > max_elements:
        return False
    return True
def _validate_worker_payload(
    payload,
    *,
    max_depth: int = _DEFAULT_MAX_RESULT_DEPTH,
    max_elements: int = _DEFAULT_MAX_RESULT_ELEMENTS,
) -> tuple[bool, str | None, object]:
    """严格校验工作程序回包。
    返回 ``(ok, error_reason, result)``：
    - ``ok=True``：合法 ``completed``，``result`` 为结果值（可为 None）；
    - ``ok=False``：非法/错误回包，``error_reason`` 为固定类别。
    规则：
    - 必须是对象；
    - ``status`` 必须是字符串且为已知枚举；
    - ``completed`` 时**必须**存在 ``result`` 键（显式 ``null`` 合法，缺键不合法
      ——那是协议错误，不是"代码没返回结果"）；
    - 结果必须通过结构边界（深度/元素数量/非有限数）；
    - ``status == "error"`` 一律不成功（即使退出码 0）。
    """
    if not isinstance(payload, dict):
        return False, "execution_error", None
    status = payload.get("status")
    if not isinstance(status, str):
        return False, "execution_error", None
    if status == _WORKER_OK_STATUS:
        if "result" not in payload:
            return False, "execution_error", None
        value = payload.get("result")
        if not _validate_result_structure(value, max_depth=max_depth, max_elements=max_elements):
            return False, "execution_error", None
        return True, None, value
    if status == _WORKER_ERROR_STATUS:
        return False, "execution_error", None
    return False, "execution_error", None
# ---------------------------------------------------------------------------
# 执行
# ---------------------------------------------------------------------------
def _emit_protocol_payload(stdout_bytes: bytes) -> tuple[bool, object, str | None]:
    """解析 stdout 单包协议（严格单 JSON 对象，拒绝多包/杂乱）。
    返回 ``(ok, payload, reason)``。
    """
    try:
        text = stdout_bytes.decode("utf-8")
    except UnicodeDecodeError:
        return False, None, "invalid_encoding"
    text = text.strip()
    if not text:
        return False, None, "empty"
    # 严格单包：必须恰好一个 JSON 对象，前后无多余非空白内容
    decoder = json.JSONDecoder()
    try:
        obj, idx = decoder.raw_decode(text)
    except ValueError:
        return False, None, "invalid_json"
    if text[idx:].strip():
        return False, None, "trailing_data"
    if not isinstance(obj, dict):
        return False, None, "not_object"
    return True, obj, None
def run_python_isolated(
    config: DockerRuntimeConfig,
    *,
    code: str,
    input_payload: dict,
    cpu_quota: float,
    memory_mb: int,
    pids_limit: int,
    scratch_mb: int,
    max_output_bytes: int,
    wall_timeout_seconds: float,
    cpu_seconds: int = 60,
    max_stderr_bytes: int = 262144,
    max_result_depth: int = _DEFAULT_MAX_RESULT_DEPTH,
    max_result_elements: int = _DEFAULT_MAX_RESULT_ELEMENTS,
    cancel_check=None,
    endpoint: VerifiedEndpoint | None = None,
) -> DockerRunResult:
    """执行一次隔离 Python 分析。失败关闭：任何环节异常都有明确生命周期。
    生命周期：配置校验 → 端点核实（若未锁定）→ 引擎/镜像 → 运行（并发有界读取）
    → 输出超限/超时/取消/异常 → 终止容器 → 核查清理 → 返回有类型结果。
    """
    # 0. 隔离配置校验
    ok, _reason = validate_isolation_config(config)
    if not ok:
        raise DockerRuntimeError("isolation_not_applied", ISOLATION_NOT_APPLIED)
    # 0b. 资源限额严格校验：必须在任何 Docker 调用之前拒绝非法参数
    #     （bool/NaN/Infinity/零/负/非整数值），不得静默降级为默认或无限。
    _validate_resource_limits(
        cpu_quota=cpu_quota,
        cpu_seconds=cpu_seconds,
        memory_mb=memory_mb,
        pids_limit=pids_limit,
        scratch_mb=scratch_mb,
    )
    _require_positive_int(max_output_bytes, "max_output_bytes")
    _require_positive_int(max_stderr_bytes, "max_stderr_bytes")
    _require_positive_int(max_result_depth, "max_result_depth")
    _require_positive_int(max_result_elements, "max_result_elements")
    _require_positive_finite_float(wall_timeout_seconds, "wall_timeout_seconds")
    # 1. 锁定本机端点（调用方未提供时自行核实）
    if endpoint is None:
        endpoint = verified_endpoint()
    # 1b. 即便调用方传入实例，也必须**重新校验**其结构为本机形态；
    #     不得仅凭对象名 VerifiedEndpoint 就信任（避免非法端点进入执行路径）。
    if not endpoint_is_locally_verified(endpoint):
        raise DockerRuntimeError("endpoint_not_local", ENDPOINT_NOT_LOCAL)
    # 2. 引擎与镜像
    if not verify_engine(endpoint):
        raise DockerRuntimeError("runner_missing", RUNNER_UNAVAILABLE)
    if not image_present(config, endpoint):
        raise DockerRuntimeError("image_missing", IMAGE_MISSING)
    # 3. 取消检查（执行前）
    if cancel_check is not None and cancel_check():
        return DockerRunResult("cancelled", None, None, "cancelled")
    request_payload = {"code": code, "input": input_payload}
    input_bytes = json.dumps(request_payload, ensure_ascii=False).encode("utf-8")
    container_name = _new_container_name()
    args, env = build_run_arguments(
        config,
        cpu_quota=cpu_quota,
        cpu_seconds=cpu_seconds,
        memory_mb=memory_mb,
        pids_limit=pids_limit,
        scratch_mb=scratch_mb,
        container_name=container_name,
    )
    # 用锁定端点执行 docker run
    full_args = ["docker", *endpoint.cli_prefix(), *args[1:]]
    collector = _OutputCollector(stdout_cap=max_output_bytes, stderr_cap=max_stderr_bytes)
    proc = None
    terminally_cleaned = False
    import time as _time
    # 生命周期的统一截止时刻：从**启动之前**开始计时，覆盖
    # 「启动 + 输入传输 + 运行 + 结果收集」。不能从"输入写完"才开始计时（F1）。
    deadline = _time.monotonic() + float(wall_timeout_seconds)
    try:
        proc = subprocess.Popen(
            full_args,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=False,
            env=env,
        )
        # 输入传输、双通道读取**并发**——任一方都不阻塞生命周期监控。
        stdin_state: dict = {"fed": False, "error": None, "done": False}
        t_in = threading.Thread(
            target=_feed_stdin, args=(proc.stdin, input_bytes, stdin_state), daemon=True
        )
        t_out = threading.Thread(target=_pump, args=(proc.stdout, collector, True), daemon=True)
        t_err = threading.Thread(target=_pump, args=(proc.stderr, collector, False), daemon=True)
        t_in.start()
        t_out.start()
        t_err.start()
        # 等待：输入失败 / 输出超限 / 取消 / 超时 / 自然结束
        status_early: str | None = None
        while True:
            # 输入传输失败不得被吞掉后继续发布成功结果（F1）——直接进入失败分支。
            if stdin_state["error"] is not None:
                status_early = "input_error"
                break
            if proc.poll() is not None:
                break
            if collector.over_limit():
                status_early = "limit_exceeded"
                break
            if cancel_check is not None and cancel_check():
                status_early = "cancelled"
                break
            if _time.monotonic() > deadline:
                status_early = "timeout"
                break
            _time.sleep(0.05)
        # 任何提前终止都 kill 容器并核实清理
        if status_early is not None:
            _terminate_process(proc)
            state = _cleanup_container(container_name, endpoint=endpoint)
            terminally_cleaned = True
            if status_early == "input_error":
                # 输入未能送达 → 固定执行错误类别；清理未知时优先 cleanup_unverified。
                if state != CLEANUP_GONE:
                    return DockerRunResult("error", None, None, "cleanup_unverified")
                return DockerRunResult("error", None, None, "execution_error")
            return _finish_with_cleanup(
                container_name, endpoint,
                status=status_early, exit_code=None, reason_category=status_early,
                cleanup_state=state,
            )
        # 自然结束：收线程（读取线程收尾，输入线程也应收束）
        t_out.join(timeout=5)
        t_err.join(timeout=5)
        t_in.join(timeout=5)
        # 取消后迟到结果 → 拒绝（清理核查同样必须被消费，F2）
        if cancel_check is not None and cancel_check():
            state = _cleanup_container(container_name, endpoint=endpoint)
            terminally_cleaned = True
            return _finish_with_cleanup(
                container_name, endpoint,
                status="cancelled", exit_code=None, reason_category="cancelled",
                cleanup_state=state,
            )
        # 输入传输失败（进程已自然结束但输入没送成）→ 固定执行错误
        if stdin_state["error"] is not None:
            state = _cleanup_container(container_name, endpoint=endpoint)
            terminally_cleaned = True
            if state != CLEANUP_GONE:
                return DockerRunResult("error", proc.returncode, None, "cleanup_unverified")
            return DockerRunResult("error", proc.returncode, None, "execution_error")
        # 输出超限（与进程结束同时到达）→ 清理核查同样必须被消费（F2）
        if collector.over_limit():
            state = _cleanup_container(container_name, endpoint=endpoint)
            terminally_cleaned = True
            return _finish_with_cleanup(
                container_name, endpoint,
                status="limit_exceeded", exit_code=proc.returncode,
                reason_category="output_limit_exceeded", cleanup_state=state,
            )
        # 清理容器（docker run --rm 通常已自清，仍核实一遍）
        state = _cleanup_container(container_name, endpoint=endpoint)
        terminally_cleaned = True
        if state != CLEANUP_GONE:
            return DockerRunResult("error", proc.returncode, None, "cleanup_unverified")
        if proc.returncode != 0:
            return DockerRunResult("error", proc.returncode, None, "execution_error")
        ok_parse, payload, parse_reason = _emit_protocol_payload(collector.stdout_bytes())
        if not ok_parse:
            return DockerRunResult("error", proc.returncode, None, "execution_error")
        ok_payload, err_reason, result_value = _validate_worker_payload(
            payload, max_depth=max_result_depth, max_elements=max_result_elements,
        )
        if not ok_payload:
            return DockerRunResult("error", proc.returncode, None, err_reason or "execution_error")
        return DockerRunResult(
            "completed", 0,
            {"status": "completed", "result": result_value},
            None,
        )
    except OSError as exc:
        if not terminally_cleaned:
            if proc is not None:
                _terminate_process(proc)
            _cleanup_container(container_name, endpoint=endpoint)
            terminally_cleaned = True
        raise DockerRuntimeError("execution_error", EXECUTION_ERROR) from exc
    finally:
        # 关掉可能仍开的管道（不落宿主文件），并收束 CLI 进程。
        _terminate_process(proc)
