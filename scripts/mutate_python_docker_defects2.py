"""反向变异：Docker 隔离执行器第八批缺陷（B1–B6）修复的判别力（Request I）。

用法：python scripts/mutate_python_docker_defects2.py M1|...|M26|M23a|M23b
M23 是兼容入口，依次执行 M23a/M23b，绝不同时拆两道输入失败守卫。
M23a/M23b 在独立测试进程内存中加载变异，不写回生产源码。

口径与其余变异脚本一致（node 在本沙箱无法 spawn 子进程，故用 Python 实现）。
铁律：① 基线先 GREEN；② 信号钩子还原；③ touched 登记；④ SEED 唯一（从函数
声明行/唯一片段开始，避免留残体）。

每个变异拆掉**对应行为本身**（不是改注释/测试名/无关变量），且各变异有唯一判别点。
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
PY = REPO / ".venv" / "Scripts" / "python.exe"
RUNNER = REPO / "server" / "app" / "python_docker_runner.py"
IGNORE = REPO / "server" / ".dockerignore"
DEFECT_TESTS = "server/tests/test_python_docker_defects2.py"

NODE = "\n"


def _j(*lines: str) -> str:
    return NODE.join(lines)


# (文件, 期望变红的测试名)
TARGETS = {
    # B1 输出有界：拆掉 stdout 封顶（不再 over_limit）
    "M1": (RUNNER, ["test_output_collector_enforces_stdout_cap"]),
    # B1 stderr 独立上限：拆掉 stderr 封顶
    "M2": (RUNNER, ["test_output_collector_enforces_stderr_cap_independently"]),
    # B2 清理可核实：把"清理失败/未知"直接当作 gone
    "M3": (RUNNER, ["test_cleanup_failure_not_reported_as_success",
                     "test_cleanup_query_failure_is_unknown_not_gone"]),
    # F1/F2 取消守卫：执行前取消不再短路（会继续起 Docker）
    "M4": (RUNNER, ["test_cancel_before_execution_short_circuits"]),
    # B3 端点核实：把远端端点当本机接受（拆掉 scheme 分派的拒绝分支）
    "M5": (RUNNER, ["test_engine_endpoint_rejects_remote_tcp",
                     "test_persistent_remote_context_rejected_even_with_empty_env"]),
    # B4 root UID：只拒 "0"，漏过 0:1000/root:1000
    "M6": (RUNNER, ["test_non_root_validation_rejects_root_uid_variants"]),
    # B4 镜像不可变：digest 缺失时静默降级可变 tag
    "M7": (RUNNER, ["test_image_ref_requires_immutable_identity"]),
    # B4 非有限数：不再拒绝 NaN/Infinity
    "M8": (RUNNER, ["test_worker_payload_rejects_nan_and_infinity"]),
    # B4 结果深度：不再拒绝过深
    "M9": (RUNNER, ["test_worker_payload_rejects_excessive_depth"]),
    # B5 构建上下文：.dockerignore 不再排除 .env
    "M10": (IGNORE, ["test_build_context_excludes_sensitive_files",
                      "test_build_context_actual_file_set_excludes_secrets"]),
    # B6 真实验收：确切容器身份核对被移除（静态判别）
    "M11": (REPO / "server" / "tests" / "test_python_docker_real.py",
            ["test_real_test_container_identity_verified"]),
    # B3 端点解析：返回第一个 URL 匹配（会取到 Host 之前的无关 URL）
    "M12": (RUNNER, ["test_endpoint_parser_ignores_non_docker_urls_before_host"]),
    # B3b 本机端点结构校验：退回"只判 scheme 前缀"（形似本机实为远端的串被接受）
    "M15": (RUNNER, ["test_endpoint_rejects_remote_named_pipe_host",
                     "test_endpoint_rejects_remote_or_malformed_hosts",
                     "test_verified_endpoint_instance_is_validated_not_trusted_by_name"]),
    # B3b 生命周期端点锁定：执行路径不再重新校验端点实例（按名字信任）
    "M16": (RUNNER, ["test_unverified_endpoint_cannot_enter_execution_path"]),
    # B3b 最内层防线：CLI 执行器不再校验端点实例
    "M17": (RUNNER, ["test_cli_level_guard_rejects_unverified_endpoint"]),
    # B7 资源参数：capability_int 重新接受 bool（True 冒充 1）
    "M18": (REPO / "server" / "app" / "capabilities.py",
            ["test_capability_int_rejects_bool"]),
    # B7 资源参数：构建参数前不再严格校验（bool 配额进入 CLI）
    "M19": (RUNNER, ["test_build_run_arguments_rejects_invalid_int_quota",
                     "test_run_python_isolated_rejects_bool_quota_before_docker"]),
    # B7 资源参数：执行入口不再在起 Docker 前校验（非法配额进入执行）
    "M20": (RUNNER, ["test_run_python_isolated_rejects_bool_quota_before_docker",
                     "test_run_python_isolated_rejects_bad_cpu_quota_before_docker",
                     "test_run_python_isolated_rejects_negative_wall_timeout_before_docker"]),
    # B7 资源参数：capability_float 不再拒绝 NaN/Infinity
    "M21": (REPO / "server" / "app" / "capabilities.py",
            ["test_capability_float_rejects_non_finite"]),
    # F1 stdin 阻塞绕过超时：输入改回主线程同步写（deadline 从写完才起）
    "M22": (RUNNER, ["test_stdin_block_does_not_bypass_wall_timeout",
                     "test_input_transfer_runs_concurrently_with_monitoring"]),
    # F1 运行中输入失败：只拆循环内守卫
    "M23a": (RUNNER, ["test_input_transfer_failure_not_swallowed_as_success"]),
    # F1 自然结束时输入失败：只拆结束后的守卫
    "M23b": (RUNNER, ["test_input_transfer_failure_after_process_exit_not_published"]),
    # F2 迟到取消不消费清理核查：直接返回普通 cancelled
    "M24": (RUNNER, ["test_late_cancel_with_unknown_cleanup_returns_cleanup_unverified"]),
    # F2 自然结束超限不消费清理核查：直接返回普通 limit_exceeded
    "M25": (RUNNER, ["test_natural_end_limit_exceeded_with_unknown_cleanup"]),
    # F2 正常完成不消费清理核查：unknown 仍发布 completed
    "M26": (RUNNER, ["test_normal_completion_with_unknown_cleanup_not_published"]),
}

SEEDS = {
    # feed_stdout 的最后一行封顶判定
    "M1": _j(
        "            if len(self._stdout) > self._stdout_cap:",
        "                self._over = True",
    ),
    "M2": _j(
        "            if len(self._stderr) > self._stderr_cap:",
        "                self._over = True",
    ),
    # _cleanup_container 的核查分支
    "M3": _j(
        "    exists = _container_exists(container_name, endpoint=endpoint)",
        "    if exists is False:",
        "        return CLEANUP_GONE",
        "    # 仍存在或查询失败 → 未知（不冒充成功）",
        "    return CLEANUP_UNKNOWN",
    ),
    # 执行前取消短路守卫
    "M4": _j(
        "    # 3. 取消检查（执行前）",
        "    if cancel_check is not None and cancel_check():",
        "        return DockerRunResult(\"cancelled\", None, None, \"cancelled\")",
    ),
    # 端点 scheme 分派的拒绝分支（远端 tcp/ssh/fd/未知一律接受）
    "M5": _j(
        "    if scheme == \"npipe\":",
        "        return _classify_local_named_pipe(remainder)",
        "    if scheme == \"unix\":",
        "        return _classify_local_unix_socket(remainder)",
        "    return False, f\"Docker 端点不是本机 npipe/unix socket：{scheme or '(空)'}\"",
    ),
    # root uid 判定：数值 0 即 root
    "M6": _j(
        "    if re.fullmatch(r\"[0-9]+\", uid_part):",
        "        try:",
        "            return int(uid_part) == 0",
        "        except ValueError:",
        "            return True",
    ),
    # image_ref 的失败关闭
    "M7": _j(
        "        if self.image_digest:",
        "            return f\"{self.image}@{self.image_digest}\"",
        "        # 失败关闭：返回带摘要语法但摘要缺失的引用，Docker 无法解析它。",
        "        return f\"{self.image}@sha256:\"",
    ),
    # 非有限数检测
    "M8": _j(
        "    if isinstance(value, float):",
        "        return value != value or value in (float(\"inf\"), float(\"-inf\"))",
    ),
    # 结构校验里的深度拒绝
    "M9": _j(
        "    if _json_depth(value, _limit=max_depth) > max_depth:",
        "        return False",
    ),
    # .dockerignore 里的 .env 规则
    "M10": ".env\n.env.*\n!/.env.example",
    # 真实测试里的确切身份核对
    "M11": "        assert exists is False, f\"确切容器 {name} 未被清理（exists={exists}）\"",
    # 端点解析：结构化 Endpoints.docker.Host 优先分支
    "M12": _j(
        "    try:",
        "        parsed = json.loads(raw)",
        "    except (ValueError, TypeError):",
        "        parsed = None",
        "    candidates = parsed if isinstance(parsed, list) else [parsed] if parsed is not None else []",
        "    for item in candidates:",
        "        if not isinstance(item, dict):",
        "            continue",
        "        endpoints = item.get(\"Endpoints\")",
        "        if isinstance(endpoints, dict):",
        "            docker = endpoints.get(\"docker\")",
        "            if isinstance(docker, dict) and isinstance(docker.get(\"Host\"), str):",
        "                host = docker[\"Host\"].strip()",
        "                if host:",
        "                    return host",
    ),
    # _classify_local_endpoint 的结构校验主体（scheme 分派 + 形态校验入口）
    "M15": _j(
        "    if not isinstance(host, str):",
        "        return False, \"端点 Host 不是字符串\"",
        "    raw = host.strip()",
        "    if not raw:",
        "        return False, \"端点 Host 为空\"",
        "    # 只按 \"scheme://\" 切开，保留 scheme 之后剩余部分（可能自 `/` 或空开始）。",
        "    if \"://\" not in raw:",
        "        return False, f\"端点缺少 scheme 分隔符：{raw.split(':', 1)[0]}\"",
        "    scheme, remainder = raw.split(\"://\", 1)",
        "    scheme = scheme.strip().lower()",
        "    if scheme == \"npipe\":",
        "        return _classify_local_named_pipe(remainder)",
        "    if scheme == \"unix\":",
        "        return _classify_local_unix_socket(remainder)",
        "    return False, f\"Docker 端点不是本机 npipe/unix socket：{scheme or '(空)'}\"",
    ),
    # run_python_isolated 里对传入端点实例的再校验
    "M16": _j(
        "    # 1b. 即便调用方传入实例，也必须**重新校验**其结构为本机形态；",
        "    #     不得仅凭对象名 VerifiedEndpoint 就信任（避免非法端点进入执行路径）。",
        "    if not endpoint_is_locally_verified(endpoint):",
        "        raise DockerRuntimeError(\"endpoint_not_local\", ENDPOINT_NOT_LOCAL)",
    ),
    # _run_docker_cli 里对端点的最内层校验
    "M17": _j(
        "        # 最内层防线：任何带端点的 CLI 调用都必须先通过本机结构校验，",
        "        # 非法端点一律拒绝（不冒充本机、不锁定到远端）。",
        "        if not endpoint_is_locally_verified(endpoint):",
        "            raise DockerRuntimeError(\"endpoint_not_local\", ENDPOINT_NOT_LOCAL)",
    ),
    # capability_int 的 bool 剔除
    "M18": _j(
        "    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:",
        "        raise RuntimeError(f\"能力配置 {section}.{key} 必须是正整数\")",
    ),
    # build_run_arguments 入口的资源校验调用
    "M19": _j(
        "    # 适配器入口：构造 CLI 之前先严格校验资源限额（拒绝 bool/NaN/Inf/零/负/非整）。",
        "    _validate_resource_limits(",
        "        cpu_quota=cpu_quota,",
        "        cpu_seconds=cpu_seconds,",
        "        memory_mb=memory_mb,",
        "        pids_limit=pids_limit,",
        "        scratch_mb=scratch_mb,",
        "    )",
    ),
    # run_python_isolated 入口的资源校验调用
    "M20": _j(
        "    # 0b. 资源限额严格校验：必须在任何 Docker 调用之前拒绝非法参数",
        "    #     （bool/NaN/Infinity/零/负/非整数值），不得静默降级为默认或无限。",
        "    _validate_resource_limits(",
        "        cpu_quota=cpu_quota,",
        "        cpu_seconds=cpu_seconds,",
        "        memory_mb=memory_mb,",
        "        pids_limit=pids_limit,",
        "        scratch_mb=scratch_mb,",
        "    )",
        "    _require_positive_int(max_output_bytes, \"max_output_bytes\")",
        "    _require_positive_int(max_stderr_bytes, \"max_stderr_bytes\")",
        "    _require_positive_int(max_result_depth, \"max_result_depth\")",
        "    _require_positive_int(max_result_elements, \"max_result_elements\")",
        "    _require_positive_finite_float(wall_timeout_seconds, \"wall_timeout_seconds\")",
    ),
    # capability_float 的非有限数拒绝
    "M21": _j(
        "    if not math.isfinite(value):",
        "        raise RuntimeError(f\"能力配置 {section}.{key} 必须是有限数值\")",
    ),
    # F1 输入传输与监控并发（改回主线程同步写 stdin）
    "M22": _j(
        "        # 输入传输、双通道读取**并发**——任一方都不阻塞生命周期监控。",
        "        stdin_state: dict = {\"fed\": False, \"error\": None, \"done\": False}",
        "        t_in = threading.Thread(",
        "            target=_feed_stdin, args=(proc.stdin, input_bytes, stdin_state), daemon=True",
        "        )",
    ),
    "M23a": _j(
        "            # 输入传输失败不得被吞掉后继续发布成功结果（F1）——直接进入失败分支。",
        "            if stdin_state[\"error\"] is not None:",
        "                status_early = \"input_error\"",
        "                break",
    ),
    "M23b": _j(
        "        # 输入传输失败（进程已自然结束但输入没送成）→ 固定执行错误",
        "        if stdin_state[\"error\"] is not None:",
        "            state = _cleanup_container(container_name, endpoint=endpoint)",
        "            terminally_cleaned = True",
        "            if state != CLEANUP_GONE:",
        "                return DockerRunResult(\"error\", proc.returncode, None, \"cleanup_unverified\")",
        "            return DockerRunResult(\"error\", proc.returncode, None, \"execution_error\")",
    ),
    # F2 迟到取消消费清理核查
    "M24": _j(
        "        if cancel_check is not None and cancel_check():",
        "            state = _cleanup_container(container_name, endpoint=endpoint)",
        "            terminally_cleaned = True",
        "            return _finish_with_cleanup(",
        "                container_name, endpoint,",
        "                status=\"cancelled\", exit_code=None, reason_category=\"cancelled\",",
        "                cleanup_state=state,",
        "            )",
    ),
    # F2 自然结束输出超限消费清理核查
    "M25": _j(
        "        if collector.over_limit():",
        "            state = _cleanup_container(container_name, endpoint=endpoint)",
        "            terminally_cleaned = True",
        "            return _finish_with_cleanup(",
        "                container_name, endpoint,",
        "                status=\"limit_exceeded\", exit_code=proc.returncode,",
        "                reason_category=\"output_limit_exceeded\", cleanup_state=state,",
        "            )",
    ),
    # F2 正常完成前的清理核查消费
    "M26": _j(
        "        state = _cleanup_container(container_name, endpoint=endpoint)",
        "        terminally_cleaned = True",
        "        if state != CLEANUP_GONE:",
        "            return DockerRunResult(\"error\", proc.returncode, None, \"cleanup_unverified\")",
        "        if proc.returncode != 0:",
    ),
}

MUTANTS = {
    "M1": _j(
        "            if False:  # [变异M1] 拆掉 stdout 封顶",
        "                self._over = True",
    ),
    "M2": _j(
        "            if False:  # [变异M2] 拆掉 stderr 封顶",
        "                self._over = True",
    ),
    "M3": _j(
        "    # [变异M3] 忽略核查结果：一律冒充已清理",
        "    return CLEANUP_GONE",
    ),
    "M4": _j(
        "    # [变异M4] 拆掉执行前取消短路（继续起 Docker）",
    ),
    "M5": _j(
        "    # [变异M5] 拆掉 scheme 分派的拒绝分支（远端端点一律接受）",
        "    return True, \"\"",
    ),
    "M6": _j(
        "    if uid_part == \"0\":  # [变异M6] 只拒裸 0",
        "        return True",
    ),
    "M7": _j(
        "        if self.image_digest:",
        "            return f\"{self.image}@{self.image_digest}\"",
        "        return self.image  # [变异M7] 静默降级可变 tag",
    ),
    "M8": _j(
        "    if False:  # [变异M8] 不检测非有限数",
        "        return value != value",
    ),
    "M9": _j(
        "    if False:  # [变异M9] 不拒绝过深",
        "        return False",
    ),
    "M10": "# [变异M10] 删掉 .env 排除规则\nexample-keep",
    "M11": "        # [变异M11] 删掉确切容器身份核对（不再断言已不存在）",
    # 拆掉结构化 Host 优先解析（把结构化分支整体去掉 → 退回"第一个 URL 匹配"错误行为）
    "M12": _j(
        "    # [变异M12] 拆掉结构化 Host 优先解析",
        "    if False:",
        "        return \"\"",
    ),
    # 退回"只判 scheme 前缀"（不复核结构；形似本机实为远端的串被接受）
    "M15": _j(
        "    # [变异M15] 退回 scheme 前缀判定（不做结构校验）",
        "    if not isinstance(host, str) or not host.strip():",
        "        return False, \"端点 Host 为空\"",
        "    if not host.startswith(_LOCAL_ENDPOINT_SCHEMES):",
        "        return False, f\"Docker 端点不是本机 npipe/unix socket：{host.split('://')[0]}\"",
        "    return True, \"\"",
    ),
    # 拆掉执行路径的端点实例再校验（按名字信任 VerifiedEndpoint）
    "M16": _j(
        "    # [变异M16] 拆掉执行路径的端点实例再校验（按名字信任）",
    ),
    # 拆掉 CLI 执行器的最内层端点校验（按名字信任 VerifiedEndpoint）
    "M17": _j(
        "        # [变异M17] 拆掉 CLI 执行器最内层端点校验（按名字信任）",
    ),
    # 重新接受 bool（不再剔除，True 冒充 1）
    "M18": _j(
        "    # [变异M18] 重新接受 bool（True 冒充 1）",
        "    if not isinstance(value, int) or value <= 0:",
        "        raise RuntimeError(f\"能力配置 {section}.{key} 必须是正整数\")",
    ),
    # 拆掉适配器入口的资源校验（非法配额直接进入 CLI）
    "M19": _j(
        "    # [变异M19] 拆掉适配器入口的资源校验",
        "    pass",
    ),
    # 拆掉执行入口的资源校验（非法配额直接进入执行）
    "M20": _j(
        "    # [变异M20] 拆掉执行入口的资源校验",
        "    pass",
    ),
    # 不再拒绝非有限数（NaN/Infinity 可绕过上限）
    "M21": _j(
        "    # [变异M21] 不再拒绝非有限数",
        "    pass",
    ),
    # 输入改回主线程同步写（deadline 从写完才起 → stdin 阻塞绕过超时）
    "M22": _j(
        "        stdin_state: dict = {\"fed\": False, \"error\": None, \"done\": False}",
        "        # [变异M22] 主线程同步写 stdin（deadline 从写完才起）",
        "        proc.stdin.write(input_bytes)",
        "        proc.stdin.close()",
        "        stdin_state[\"fed\"] = True",
        "        t_in = threading.Thread(target=lambda: None, daemon=True)",
    ),
    "M23a": _j(
        "            # [变异M23a] 只拆循环内输入失败守卫",
        "            pass",
    ),
    "M23b": _j(
        "        # [变异M23b] 只拆自然结束后的输入失败守卫",
        "        pass",
    ),
    # 迟到取消直接返回普通 cancelled（不消费清理核查）
    "M24": _j(
        "        if cancel_check is not None and cancel_check():",
        "            _cleanup_container(container_name, endpoint=endpoint)",
        "            terminally_cleaned = True",
        "            # [变异M24] 忽略清理核查结果，冒充普通取消",
        "            return DockerRunResult(\"cancelled\", None, None, \"cancelled\")",
    ),
    # 自然结束超限直接返回普通 limit_exceeded（不消费清理核查）
    "M25": _j(
        "        if collector.over_limit():",
        "            _cleanup_container(container_name, endpoint=endpoint)",
        "            terminally_cleaned = True",
        "            # [变异M25] 忽略清理核查结果，冒充普通超限",
        "            return DockerRunResult(\"limit_exceeded\", proc.returncode, None, \"output_limit_exceeded\")",
    ),
    # 正常完成时清理 unknown 仍发布结果
    "M26": _j(
        "        state = _cleanup_container(container_name, endpoint=endpoint)",
        "        terminally_cleaned = True",
        "        # [变异M26] 忽略清理核查结果，冒充 completed",
        "        if proc.returncode != 0:",
    ),
}

_touched: dict[Path, str] = {}


def restore_all() -> None:
    for path, content in list(_touched.items()):
        path.write_text(content, encoding="utf-8", newline="")
    _touched.clear()


def _install_signal_hooks() -> None:
    import signal

    def handler(signum, frame):
        restore_all()
        sys.exit(130)

    for sig in (signal.SIGINT, signal.SIGTERM):
        try:
            signal.signal(sig, handler)
        except (ValueError, OSError):
            pass


def run_pytest(test_path: str, name: str, *, runner_source: str | None = None) -> str:
    base = tempfile.mkdtemp(prefix=f"mutdef2_{name}_")
    env = dict(os.environ)
    env["TMP"] = base
    env["TEMP"] = base
    env["PYTEST_ADDOPTS"] = ""
    pytest_args = [test_path, "-q", "-p", "no:cacheprovider",
                   "--basetemp", str(Path(base) / "bt")]
    if runner_source is None:
        command = [str(PY), "-B", "-m", "pytest", *pytest_args]
    else:
        # 这里只加载可信仓库源码的单守卫变异；绝不执行模型传入的 code。
        bootstrap = (
            "import sys, pytest; from server.app import python_docker_runner as runner; "
            "exec(compile(sys.stdin.read(), runner.__file__, 'exec'), runner.__dict__); "
            "raise SystemExit(pytest.main(sys.argv[1:]))"
        )
        command = [str(PY), "-B", "-c", bootstrap, *pytest_args]
    proc = subprocess.run(
        command, input=runner_source,
        cwd=str(REPO), capture_output=True, text=True, timeout=300, env=env,
    )
    return (proc.stdout or "") + "\n" + (proc.stderr or "")


def _counts(text: str) -> tuple[int, int]:
    p = re.search(r"(\d+) passed", text)
    f = re.search(r"(\d+) failed", text)
    return (int(p.group(1)) if p else 0, int(f.group(1)) if f else 0)


def main(which: str | None = None) -> int:
    if which is None:
        which = sys.argv[1] if len(sys.argv) > 1 else ""
    if which == "M23":
        for independent in ("M23a", "M23b"):
            print(f"[M23 兼容入口] 独立运行 {independent}")
            status = main(independent)
            if status:
                return status
        return 0
    if which not in TARGETS:
        print("用法: python scripts/mutate_python_docker_defects2.py "
              + "|".join([*TARGETS, "M23"]))
        return 2
    _install_signal_hooks()
    file, expect_fail = TARGETS[which]

    baseline = run_pytest(DEFECT_TESTS, "baseline")
    bp, bf = _counts(baseline)
    print(f"[基线] {bp} passed, {bf} failed")
    if bp == 0 or bf:
        restore_all()
        print("BASELINE-NOT-GREEN：基线未全绿，终止（源码未改动）。")
        print(baseline[-1500:])
        return 4

    # 换行符归一化：目标文件可能是 CRLF/LF，SEED 统一按 LF 比对
    raw = file.read_bytes().decode("utf-8")
    crlf = "\r\n" in raw
    original = raw
    norm = original.replace("\r\n", "\n")
    seed = SEEDS[which]

    occurrences = norm.count(seed)
    if occurrences != 1:
        print(f"SKIP: {which} SEED 出现 {occurrences} 次（期望 1）。")
        print("SEED 片段：\n" + seed)
        return 3

    in_memory = which in ("M23a", "M23b")
    if not in_memory:
        _touched[file] = original
    try:
        mutated_norm = norm.replace(seed, MUTANTS[which])
        if in_memory:
            other = "M23b" if which == "M23a" else "M23a"
            assert mutated_norm.count(SEEDS[other]) == 1, "另一道输入守卫必须完整保留"
            out = run_pytest(DEFECT_TESTS, which + "_mutated", runner_source=mutated_norm)
        else:
            mutated = mutated_norm.replace("\n", "\r\n") if crlf else mutated_norm
            file.write_text(mutated, encoding="utf-8", newline="")
            out = run_pytest(DEFECT_TESTS, "mutated")
        failed = re.findall(r"FAILED (server/tests/\w+\.py::[\w\[\]\-]+)", out)
        failed_names = sorted({n.split("::")[1].split("[")[0] for n in failed})
        hits = [n for n in expect_fail if n in failed_names]
        print(f"[变异后] 失败 {len(failed_names)} 条：{', '.join(failed_names) or '（无）'}")
        print(f"[命中判别点] {', '.join(hits) if hits else '无'}")
        red = bool(failed_names) and bool(hits)
        print("RED-OK：变异使判别测试独立变红。" if red else "RED-MISS：变异未按预期变红。")
        if not red:
            print(out[-2500:])
    finally:
        restore_all()

    restored = run_pytest(DEFECT_TESTS, "restored")
    rp, rf = _counts(restored)
    print(f"[还原后] {rp} passed, {rf} failed")
    green = rp > 0 and rf == 0
    print("RESTORED-GREEN：还原后复跑全绿。" if green else "RESTORED-NOT-GREEN：还原后仍有失败！")
    return 0 if (red and green) else 5


if __name__ == "__main__":
    sys.exit(main())
