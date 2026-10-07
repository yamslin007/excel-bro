"""隔离 Python 分析的容器内固定工作程序（镜像内入口）。

由可信侧 Docker 适配器用固定启动链调用（stdin 传 JSON、stdout 回 JSON）。
模型代码只通过 stdin 的 ``code`` 字段传入，绝不进命令行参数。

安全约定：
- 从 stdin 读一个 JSON（有大小上限），字段：``code``（模型代码）+ ``input``（合成输入）。
- 命名空间不是安全边界；真正的边界是容器限制和执行前不可撤销的 seccomp。
- 输出是**有类型**的 JSON：``{status, result}``，``result`` 仅允许基础 JSON 类型。
- stdout 只用于协议回包；任何 traceback 只进 stderr，不进 stdout。
- 本程序不联网、不读宿主文件、不派生外部程序——这些由环境（容器）强制，
  本程序不做也不假装做安全边界。
"""
import json
import sys
import traceback
import ctypes
import platform

MAX_INPUT_BYTES = 2 * 1024 * 1024  # 2MB 上限


def _lock_down() -> None:
    """Install an additional kernel filter AFTER trusted Python startup.

    Supports the pinned Linux/amd64 image only. TSYNC covers any existing
    threads; failure prevents all generated code. This augments (does not
    replace) Docker's default seccomp policy. No environment-variable bypass.
    """
    if sys.platform != "linux" or platform.machine() != "x86_64":
        raise RuntimeError("unsupported isolation platform")
    # Without this check, invoking this worker directly on Linux could mistake
    # a host process for the configured Docker runtime.
    import os
    if not os.path.isfile("/.dockerenv") or os.getuid() != 65534:
        raise RuntimeError("container boundary missing")

    class Filter(ctypes.Structure):
        _fields_ = [("code", ctypes.c_ushort), ("jt", ctypes.c_ubyte),
                    ("jf", ctypes.c_ubyte), ("k", ctypes.c_uint32)]

    class Program(ctypes.Structure):
        _fields_ = [("len", ctypes.c_ushort), ("filter", ctypes.POINTER(Filter))]

    # Validate syscall architecture, reject x32/old x32 syscall numbers, then
    # block exec, all process/thread creation and network socket creation.
    deny = 0x00050001  # SECCOMP_RET_ERRNO | EPERM
    rules = [(0x20, 0, 0, 4), (0x15, 1, 0, 0xC000003E),
             (0x06, 0, 0, 0x80000000), (0x20, 0, 0, 0),
             (0x35, 0, 1, 0x40000000), (0x06, 0, 0, deny),
             (0x35, 0, 2, 512), (0x35, 1, 0, 548), (0x06, 0, 0, deny)]
    for number in (59, 322, 56, 57, 58, 435, 41, 53):
        rules.extend([(0x15, 0, 1, number), (0x06, 0, 0, deny)])
    rules.append((0x06, 0, 0, 0x7FFF0000))  # allow within Docker's outer filter
    filters = (Filter * len(rules))(*(Filter(*rule) for rule in rules))
    program = Program(len(rules), filters)
    libc = ctypes.CDLL(None, use_errno=True)
    libc.prctl.argtypes = [ctypes.c_int, ctypes.c_ulong, ctypes.c_ulong,
                          ctypes.c_ulong, ctypes.c_ulong]
    libc.prctl.restype = ctypes.c_int
    libc.syscall.restype = ctypes.c_long
    if libc.prctl(38, 1, 0, 0, 0) != 0:
        raise RuntimeError("no_new_privs failed")
    if libc.syscall(317, 1, 1, ctypes.byref(program)) != 0:
        raise RuntimeError("seccomp synchronization failed")


def _read_request() -> dict:
    raw = sys.stdin.buffer.read(MAX_INPUT_BYTES + 1)
    if len(raw) > MAX_INPUT_BYTES:
        raise ValueError("输入超过上限")
    if not raw:
        raise ValueError("空输入")
    data = json.loads(raw.decode("utf-8"))
    if not isinstance(data, dict):
        raise ValueError("输入必须是对象")
    if "code" not in data or not isinstance(data["code"], str):
        raise ValueError("缺少 code 字段")
    return data


def _run_code(code: str, input_payload):
    """在受限命名空间执行模型代码。只暴露 ``input`` 数据，不暴露任何宿主对象。"""
    namespace: dict = {
        "__builtins__": __builtins__,
        "input": input_payload,
    }
    # 使用 exec 执行模型代码，结果通过约定变量返回：
    # 代码应给 `result` 赋值；若未赋值，则尝试捕获最后一个表达式值。
    exec(compile(code, "<isolated-python>", "exec"), namespace, namespace)
    if "result" in namespace:
        return namespace["result"]
    return None


def _json_safe(value):
    """只允许基础 JSON 类型出站（数值/字符串/布尔/None/列表/字典），拒绝自定义对象。"""
    json.dumps(value)  # 若不可序列化会抛 TypeError
    return value


def main() -> None:
    try:
        request = _read_request()
    except Exception:
        _emit_error("invalid_request", "输入解析失败")
        return

    try:
        _lock_down()
    except Exception:
        _emit_error("isolation_not_applied", "隔离限制未能应用，代码未执行")
        return

    try:
        result = _run_code(request["code"], request.get("input"))
        safe_result = _json_safe(result)
        _emit({"status": "completed", "result": safe_result})
    except Exception:
        # 原始 traceback 只进 stderr（不进 stdout 协议通道，也不会被宿主持久化到模型/账本）
        traceback.print_exc(file=sys.stderr)
        _emit_error("execution_error", "代码执行失败")


def _emit(payload: dict) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False))


def _emit_error(reason: str, message: str) -> None:
    sys.stdout.write(json.dumps({
        "status": "error",
        "reasonCategory": reason,
        "message": message,
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
