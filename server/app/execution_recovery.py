"""未完成任务恢复包：本机后端应用数据目录里的原子文件存储。

**为什么需要它（2026-10-27 真实 WPS 验收）**：宿主侧的 `PluginStorage` 在"完整退出
WPS 再打开"这条路径上实测读不到原执行记录，于是"已核对第二步、仍可继续第三步"在真实
机器上断链。恢复包因此改由**本机后端**持久保存；`PluginStorage` / `localStorage` 只
保留索引与缓存，不再是跨进程恢复的唯一材料来源。

边界（不得放宽）：

- **不上云、不进模型。** 只是本机磁盘上的一份受控材料，接口只监听 127.0.0.1。
- **不进普通账本、日志、遥测。** 材料含单元格原值，本模块不写任何日志；异常消息只
  描述状态（缺失 / 过期 / 损坏 / 配额），绝不拼接正文。
- **与账本分工。** `office_executions` 是可变执行进度（游标 / 在途 / 租约）的真相；
  恢复包只保存不可变的原始恢复材料。两者不复制对方的真相。
- **落盘原子。** 单文件写：临时文件 → fsync → 原子替换（只重试替换），失败抛
  `RecoveryPackagePersistenceError`，调用方必须停止写入，不得"只存内存也继续写"。
- **配额不静默删。** 容量不足时先清理**已过期**的包；仍不足就明确报错并拒绝新增，
  绝不为了腾地方淘汰尚未结束任务的恢复材料。

文件布局（`material.dat`）：第 1 行是元数据 JSON（摘要/期限/字节数），其后是**原样
字节**的恢复包正文。用"一行元数据 + 原样正文"而不是把正文塞进 JSON 字符串，是为了
让读回的内容与客户端写下的内容逐字节相同（客户端据此做落盘读回校验）。
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import threading
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from .app_data import application_data_root
from .capabilities import capability_int
from .models import ExecutionRecoveryPackage

_LOCK = threading.Lock()

_OPERATION_ID = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.:-]{7,99}$")
_HEADER_VERSION = 1
_FILE_NAME = "material.dat"


class RecoveryPackageError(ValueError):
    """恢复包不可用。消息只描述状态，绝不包含正文。"""

    state = "invalid"


class RecoveryPackageMissing(RecoveryPackageError):
    state = "missing"


class RecoveryPackageExpired(RecoveryPackageError):
    state = "expired"


class RecoveryPackageCorrupt(RecoveryPackageError):
    state = "corrupt"


class RecoveryPackageConflict(RecoveryPackageError):
    state = "conflict"


class RecoveryPackageQuota(RecoveryPackageError):
    state = "quota"


class RecoveryPackagePersistenceError(OSError):
    """落盘失败：调用方必须停止写入（不得只留内存材料就继续写原表）。"""


def _utc(moment: datetime) -> str:
    return moment.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")


def _instant(value: Any) -> datetime | None:
    """把登记里的时间解析成**带明确时区**的时刻；确定不了就返回 `None`。

    `datetime.fromisoformat("2035-01-01T00:00:00")` 不报错，返回的是 **naive** 时间，
    随后与带 UTC 的 `now` 比较直接抛
    `TypeError: can't compare offset-naive and offset-aware datetimes`。2026-10-27 的
    现场就是只把这一个字段改成无时区：读回、同正文重存、保留期清理三个入口全炸。

    这里把"没有时区"归为**无法确定**：既不按本机时区猜（猜错方向会让材料提前失效或
    永不失效，两种都比"明确说损坏"糟），也不自动补 UTC。`Z`、`+00:00` 与其它合法偏移
    都按同一时刻比较。

    **这是读取、保存、清理共用的唯一解析口**：三处各写一份 `fromisoformat` 正是这次
    缺陷的成因（只修一处，另外两处照旧会抛 TypeError）。
    """

    if not isinstance(value, str):
        return None
    text = value.strip()
    if not text:
        return None
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00").replace("z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is None or parsed.utcoffset() is None:
        return None
    return parsed


def _aware(moment: datetime | None, what: str) -> datetime:
    """内部时刻的调用约定：必须带时区；`None` 表示"此刻（UTC）"。

    naive 的 `now` 同样会在比较处抛 TypeError，所以不能只修文件读取那一侧。这里显式
    拒绝并给出与材料无关的固定措辞 —— 不猜本机时区。
    """

    if moment is None:
        return datetime.now(timezone.utc)
    if not isinstance(moment, datetime) or moment.tzinfo is None or moment.utcoffset() is None:
        raise RecoveryPackageError(f"{what}必须带明确时区。")
    return moment


def _root() -> Path:
    root = application_data_root().resolve() / "execution-recovery"
    root.mkdir(parents=True, exist_ok=True)
    if root.is_symlink():
        raise RecoveryPackageError("恢复包目录不可信。")
    return root


def _folder(operation_id: str) -> Path:
    if not _OPERATION_ID.match(operation_id):
        raise RecoveryPackageError("恢复包标识无效。")
    root = _root()
    folder = root / operation_id
    if folder.is_symlink():
        raise RecoveryPackageError("恢复包目录不可信。")
    return folder


def _atomic_write(path: Path, payload: bytes) -> None:
    temporary = path.with_name(f".{path.name}-{os.getpid()}-{time.time_ns()}.tmp")
    attempts = capability_int("taskRecovery", "recoveryPackageWriteAttempts")
    delay = capability_int("taskRecovery", "recoveryPackageWriteRetryDelayMs") / 1000
    try:
        with temporary.open("wb") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
    except OSError as error:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass
        raise RecoveryPackagePersistenceError("恢复包无法写入本机数据目录。") from error
    for attempt in range(attempts):
        try:
            temporary.replace(path)
            return
        except PermissionError as error:
            if attempt + 1 == attempts:
                try:
                    temporary.unlink(missing_ok=True)
                except OSError:
                    pass
                raise RecoveryPackagePersistenceError("恢复包无法写入本机数据目录。") from error
            time.sleep(delay)
        except OSError as error:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass
            raise RecoveryPackagePersistenceError("恢复包无法写入本机数据目录。") from error


def _decode(data: bytes) -> tuple[dict[str, Any], bytes]:
    newline = data.find(b"\n")
    if newline <= 0:
        raise RecoveryPackageCorrupt("恢复包文件结构损坏。")
    try:
        header = json.loads(data[:newline].decode("utf-8"))
        body = data[newline + 1:]
    except (UnicodeDecodeError, json.JSONDecodeError):
        raise RecoveryPackageCorrupt("恢复包文件结构损坏。") from None
    if not isinstance(header, dict) or header.get("headerVersion") != _HEADER_VERSION:
        raise RecoveryPackageCorrupt("恢复包文件版本不受支持。")
    if hashlib.sha256(body).hexdigest() != header.get("digest"):
        raise RecoveryPackageCorrupt("恢复包摘要与正文不一致。")
    if header.get("size") != len(body):
        raise RecoveryPackageCorrupt("恢复包字节数与登记不一致。")
    return header, body


def _read_header(path: Path) -> dict[str, Any] | None:
    """只读元数据；用于清理与配额统计，不加载正文。"""

    try:
        with path.open("rb") as stream:
            line = stream.readline(4096)
        header = json.loads(line.decode("utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError):
        return None
    return header if isinstance(header, dict) and header.get("headerVersion") == _HEADER_VERSION else None


def _expiry(moment: datetime) -> str:
    return _utc(_aware(moment, "保留期起算时间")
                + timedelta(days=capability_int("taskRecovery", "unfinishedRecoveryPackageRetentionDays")))


def default_expiry(moment: datetime | None = None) -> str:
    """首次保存时该给的到期时间（保留期从此刻起算）。

    只用于**第一次**保存。之后任何一次写入都必须沿用账本登记的这一个值 ——
    重复 PUT、读回水合、过期清理都不许重新算一遍。
    """

    return _expiry(moment)


def _deadline(header: dict[str, Any]) -> datetime | None:
    """元数据里的到期时刻；字段缺失/类型不对/非法文本/**无时区**一律 `None`。"""

    return _instant(header.get("expiresAt"))


def _validated(package_json: str) -> ExecutionRecoveryPackage:
    try:
        return ExecutionRecoveryPackage.model_validate_json(package_json)
    except Exception as error:
        # Pydantic 的错误文本会带被校验的输入片段，不能外抛（含单元格原值）。
        raise RecoveryPackageError("恢复包正文未通过校验。") from None


def prune_expired(now: datetime | None = None) -> dict[str, int]:
    """删除已过期的恢复包。返回计数；单个目录清理失败只计数，不影响其它包。"""

    with _LOCK:
        return _prune_locked(_aware(now, "清理基准时间"))


def _prune_locked(moment: datetime) -> dict[str, int]:
    """`prune_expired` 的加锁内实现。调用方必须已持有 `_LOCK`（不可重入锁）。

    单条材料出问题**只影响它自己**：不删、继续处理同目录其它包。三个计数各有含义，
    不许互相顶替：

    - `removed`：确认已过期并删除；
    - `failed`：清理动作本身失败（IO/权限）；
    - `unconfirmed`：材料还在，但**期限读不出来** ⇒ 不知道是否过期。既不能当过期删掉
      （删掉就没有排查依据），也不能替它算一个新期限。
    """

    removed = failed = unconfirmed = 0
    retention = capability_int("taskRecovery", "unfinishedRecoveryPackageRetentionDays") * 86400
    for folder in sorted(_root().iterdir()):
        if not folder.is_dir():
            continue
        try:
            path = folder / _FILE_NAME
            header = _read_header(path) if path.exists() else None
            if header is None:
                # 结构损坏的包不能当成"有效材料"一直占着配额；但它也可能只是写了一半，
                # 因此只清理超过保留期的孤儿目录，不删除新鲜目录。
                age = moment.timestamp() - folder.stat().st_mtime
                if age < retention:
                    continue
            else:
                deadline = _deadline(header)
                if deadline is None:
                    # 期限读不出来 ≠ 已经过期。留着，并如实记一次"未确认"。
                    unconfirmed += 1
                    continue
                if deadline > moment:
                    continue
            for entry in folder.iterdir():
                entry.unlink(missing_ok=True)
            folder.rmdir()
            removed += 1
        except OSError:
            failed += 1
        except Exception:
            # 兜底：一条材料的任何意外都不许中断整批清理，也不许顺手把它删掉。
            unconfirmed += 1
    return {"removed": removed, "failed": failed, "unconfirmed": unconfirmed}


def _usage() -> tuple[int, int]:
    count = total = 0
    for folder in _root().iterdir():
        path = folder / _FILE_NAME
        if not folder.is_dir() or not path.is_file():
            continue
        count += 1
        try:
            total += path.stat().st_size
        except OSError:
            continue
    return count, total


def _material(header: dict[str, Any], package_json: str) -> dict[str, Any]:
    return {
        "packageJson": package_json,
        "digest": header["digest"],
        "size": header["size"],
        "createdAt": header["createdAt"],
        "expiresAt": header["expiresAt"],
        "savedAt": header["savedAt"],
    }


def _quota_exceeded(count: int, used: int, incoming: int) -> bool:
    return (count + 1 > capability_int("taskRecovery", "recoveryPackageMaxOperations")
            or used + incoming > capability_int("taskRecovery", "recoveryPackageTotalMaxBytes"))


def save(operation_id: str, package_json: str, expires_at: str | None = None) -> dict[str, Any]:
    """保存（幂等）恢复包；**首次**有效期一旦定下就不可续期。

    三条不许越过的边界：

    1. 同 `operationId` + 同正文 ⇒ 幂等，原样返回**首次**的 `savedAt`/`expiresAt`，
       不重写文件、不改时间。
    2. 同 `operationId` + 同正文但**已经过期** ⇒ 抛 `RecoveryPackageExpired`。
       绝不"生成新期限再覆盖一次" —— 那等于让重复 PUT 变成免费的续期按钮。
    3. 同 `operationId` + 不同正文 ⇒ 冲突；磁盘上的材料已损坏 ⇒ 损坏。
       两者都不能被当成"有效旧记录"顺手覆盖掉。

    `expires_at` 由调用方（HTTP 层）从账本登记的**首次**到期时间传入。文件被保留期清理
    之后再次 PUT，仍然按这个值判 —— 所以"删掉文件重建"也绕不过首次有效期；
    没有登记值时才用 `default_expiry()` 从此刻起算。
    """

    if not _OPERATION_ID.match(operation_id):
        raise RecoveryPackageError("恢复包标识无效。")
    body = package_json.encode("utf-8")
    package = _validated(package_json)
    if package.operationId != operation_id:
        raise RecoveryPackageError("恢复包与请求路径的操作标识不一致。")
    if len(body) > capability_int("taskRecovery", "recoveryPackageMaxBytes"):
        raise RecoveryPackageQuota("恢复包超过单任务容量上限，已拒绝保存。")

    moment = datetime.now(timezone.utc)
    # 账本登记的首次期限走**同一个**解析口：无时区/非法值在这里就是"无效"，绝不进入比较。
    registered = _instant(expires_at) if expires_at is not None else None
    if expires_at is not None and registered is None:
        raise RecoveryPackageError("恢复包保留期限字段无效。")

    with _LOCK:
        folder = _folder(operation_id)
        path = folder / _FILE_NAME
        if path.exists():
            try:
                header, existing = _decode(path.read_bytes())
            except OSError:
                raise RecoveryPackageCorrupt("恢复包无法读取，已拒绝覆盖。") from None
            except RecoveryPackageCorrupt:
                # 磁盘上已有的东西坏了 ⇒ 既不能当"有效旧记录"返回，也不能悄悄覆盖重发一个新期限。
                raise RecoveryPackageCorrupt("本机恢复材料已损坏，拒绝覆盖。") from None
            if existing != body:
                # 同一个 operationId 只对应一次原始执行；材料被换掉说明不是同一件事。
                raise RecoveryPackageConflict("这次执行已经有另一份恢复材料，已拒绝覆盖。")
            deadline = _deadline(header)
            if deadline is None:
                raise RecoveryPackageCorrupt("恢复包期限字段损坏，拒绝覆盖。")
            if deadline <= moment:
                raise RecoveryPackageExpired("恢复材料已超过保留期，不能重新保存或续期。")
            # 幂等：不重写文件，首次 `savedAt`/`expiresAt` 原样返回。
            return _material(header, package_json)

        deadline = registered or _instant(default_expiry(moment))
        assert deadline is not None
        if deadline <= moment:
            # 文件已被清理，但这次执行的首次有效期还在账本里：仍然过期，不许重新生效。
            raise RecoveryPackageExpired("恢复材料已超过保留期，不能重新保存或续期。")

        def within_quota() -> bool:
            count, used = _usage()
            return not _quota_exceeded(count, used, len(body))

        if not within_quota():
            # 先给**已过期**的材料让路；绝不淘汰尚未结束任务的恢复包。
            _prune_locked(moment)
            if not within_quota():
                raise RecoveryPackageQuota("本机恢复材料已达配额上限；未清理尚未结束任务的恢复包。")

        header = {
            "headerVersion": _HEADER_VERSION,
            "operationId": operation_id,
            "digest": hashlib.sha256(body).hexdigest(),
            "size": len(body),
            "createdAt": package.createdAt,
            "savedAt": _utc(moment),
            "expiresAt": _utc(deadline),
        }
        folder.mkdir(parents=True, exist_ok=True)
        _atomic_write(path, json.dumps(header, ensure_ascii=False).encode("utf-8") + b"\n" + body)
        return _material(header, package_json)


def load(operation_id: str, now: datetime | None = None) -> dict[str, Any]:
    """读回恢复包。缺失 / 过期 / 损坏分别抛出，调用方必须区别对待。"""

    moment = _aware(now, "读回基准时间")
    folder = _folder(operation_id)
    path = folder / _FILE_NAME
    if not path.is_file():
        raise RecoveryPackageMissing("本机没有这次执行的恢复包。")
    try:
        data = path.read_bytes()
    except OSError:
        raise RecoveryPackageCorrupt("恢复包无法读取。") from None
    header, body = _decode(data)
    if header.get("operationId") != operation_id:
        raise RecoveryPackageCorrupt("恢复包与请求的操作标识不一致。")
    deadline = _deadline(header)
    if deadline is None:
        # 字段缺失 / null / 非法文本 / **无时区**：读得出文件，却没法确定这份材料
        # 什么时候失效。按损坏处理 —— 不能猜一个时区再拿去比较。
        raise RecoveryPackageCorrupt("恢复包期限字段损坏。")
    if deadline <= moment:
        raise RecoveryPackageExpired("恢复包已超过保留期，不能用于继续执行。")
    try:
        package_json = body.decode("utf-8")
    except UnicodeDecodeError:
        raise RecoveryPackageCorrupt("恢复包正文不是有效文本。") from None
    _validated(package_json)
    return {
        "packageJson": package_json,
        "digest": header["digest"],
        "size": header["size"],
        "createdAt": header["createdAt"],
        "expiresAt": header["expiresAt"],
        "savedAt": header["savedAt"],
    }


def raw_digest(operation_id: str) -> tuple[str, int] | None:
    """删除**之前**那堆字节的指纹（文件字节的 sha256 + 字节数）；不存在返回 `None`。

    为什么是"文件字节"而不是 header 里的 `digest`：终态回收时材料可能**已经过期或损坏**
    —— 那恰恰是最需要留痕的时刻，而 header 不可信、也不一定读得出。被删掉的对象本来就是
    这堆字节，所以它的身份就是这堆字节的摘要。

    这里**不**校验协议、不解析正文、不判期限、不碰 mtime：它只回答"是哪一份"，
    不回答"能不能用"。可用性是 `load` / `status` 的事，两者不许混。

    只读；不加锁也能接受（唯一调用点是删除之前的同一条请求路径），读失败按"拿不到身份"
    返回 `None` —— 调用方据此**拒绝删除**，绝不拿一个猜出来的摘要去留痕。
    """

    folder = _folder(operation_id)
    path = folder / _FILE_NAME
    if not path.is_file():
        return None
    try:
        data = path.read_bytes()
    except OSError:
        return None
    return hashlib.sha256(data).hexdigest(), len(data)


def remove(operation_id: str) -> bool:
    """终态回收：删除恢复包。返回是否真的删掉了（不存在即幂等成功，返回 False）。

    唯一的合法调用点是**任务真的结束**之后（例如原地撤销成功、原表已回到执行前）。
    执行器返回成功时**不许**调用它 —— 那时审阅还没做，材料仍是核对与放弃的唯一依据。
    删除失败抛 `RecoveryPackagePersistenceError`：调用方可以如实报"材料未能回收"，
    但不得因为回收失败而回退或重放任何宿主写入；材料自然会随保留期到期被清掉。
    """

    with _LOCK:
        folder = _folder(operation_id)
        path = folder / _FILE_NAME
        if not path.exists():
            return False
        try:
            path.unlink()
            folder.rmdir()
        except OSError as error:
            raise RecoveryPackagePersistenceError("恢复包无法从本机数据目录删除。") from error
        return True


def status(operation_id: str, now: datetime | None = None) -> str:
    """恢复包状态：`ok` / `missing` / `expired` / `corrupt`。只读，不抛业务异常。

    调用约定错误（naive 的 `now`）在进入判定前就抛 —— 它属于调用方写错，不该被
    折成"材料损坏"这种关于材料的结论。
    """

    moment = _aware(now, "状态查询基准时间")
    try:
        load(operation_id, moment)
    except RecoveryPackageMissing:
        return "missing"
    except RecoveryPackageExpired:
        return "expired"
    except RecoveryPackageError:
        return "corrupt"
    except OSError:
        return "corrupt"
    return "ok"
