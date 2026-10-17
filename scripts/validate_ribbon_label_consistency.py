"""反向变异驱动：证明「两端功能区文案一致」+「图标 URL 与文件对得上」这两族守卫
**真的有判别力**。

被测机制 = `server/tests/test_installer_packaging.py` 里的三条用例：
  - `test_ribbon_entry_labels_match_across_both_hosts`
  - `test_ribbon_icon_urls_resolve_to_existing_files`
  - `test_ribbon_icon_legacy_aliases_match_versioned_files`

四个变异，分别打**不同的判别点**（同一文件上也可能有多处，所以一律一次只
apply/restore 一个）：

  M1 文本：Excel manifest 按钮文案改回「打开 AI 对话」  → 打「两端逐字段相等」
  M2 文本：WPS ribbon.xml 分组文案改成别的            → 打「两端相等」（反方向）
  M3 文本：Excel manifest 的 Icon.16 URL 指到一个不存在的版本号 → 打「文件存在」
  M4 文件：把去版本号的**别名**图标删掉                → 打「别名存在且同字节」

M3/M4 必须分开证：M3 只动 URL 字符串（文件都在），M4 只动文件（URL 都不变）。
只做一个的话，另一条用例可能是恒真的。

四条判据缺一不算过：
  1. 未变异时基线 GREEN（退出码 0 + 有效摘要 + 用例数达标）；
  2. 每个变异 apply 后用例必须变红（退出码非 0）；
  3. restore 后**被改的那个文件**逐字节一致（直接比字节，不只看 sha256）；
  4. 还原后复跑全绿，且所有文件最终都与基线一致。

用法：python scripts/validate_ribbon_label_consistency.py
（用项目自己的 .venv 跑：./.venv/Scripts/python.exe）
"""
from __future__ import annotations

import hashlib
import re
import signal
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
PYTHON = REPO / ".venv" / "Scripts" / "python.exe"
if not PYTHON.exists():
    PYTHON = Path(sys.executable)

MANIFEST = REPO / "apps" / "excel-addin" / "manifest.xml"
RIBBON = REPO / "apps" / "wps-addin" / "public" / "ribbon.xml"
EXCEL_ASSETS = REPO / "apps" / "excel-addin" / "public" / "assets"
TEST_FILE = "server/tests/test_installer_packaging.py"
# 该文件单独跑就是 17 项（24 是再并上 test_mac_packaging.py 的合计）。
# 阈值要按**本文件**的实数给，否则基线会被误判成不绿。
MIN_TESTS = 17

TARGETS = {
    "M1": {
        "kind": "text",
        "file": MANIFEST,
        "seed": '<bt:String id="Button.Label" DefaultValue="打开 Excel Bro" />',
        "mutant": '<bt:String id="Button.Label" DefaultValue="打开 AI 对话" />',
        "note": "Excel 按钮文案改回旧名（两端不再相等）",
    },
    "M2": {
        "kind": "text",
        "file": RIBBON,
        "seed": '<group id="excelBroWpsGroup" label="Excel Bro">',
        "mutant": '<group id="excelBroWpsGroup" label="WPS 助手">',
        "note": "WPS 分组文案改名（两端不再相等）",
    },
    "M3": {
        "kind": "text",
        "file": MANIFEST,
        "seed": '<bt:Image id="Icon.16" DefaultValue="https://localhost:3000/assets/excel-bro-icon-16-v2.png" />',
        "mutant": '<bt:Image id="Icon.16" DefaultValue="https://localhost:3000/assets/excel-bro-icon-16-v99.png" />',
        "note": "Excel Icon.16 指向不存在的版本号（URL 与文件对不上）",
    },
    "M4": {
        "kind": "delete-file",
        "file": EXCEL_ASSETS / "excel-bro-icon-16.png",
        "note": "删掉去版本号的别名图标（老 URL 会 404）",
    },
}

# ── 备份 / 还原钩子 ────────────────────────────────────────────────────────────
# 脚本被 Ctrl-C / 被 kill 打断时，若停在「已 apply」那一刻，文件就留在变异态，
# 下一次运行的基线就会记到变异态上，从而把状态错位误报成「还原后不一致」。
# 这里在动文件之前先把原始字节存下来，收到信号一律还原。
_BACKUP: dict[Path, bytes] = {}


def _install_signal_hooks() -> None:
    def handler(signum, _frame):  # noqa: ANN001 - 信号回调签名固定
        _restore_all()
        print(f"\n[mutate] 收到信号 {signum}，已还原文件后退出。", file=sys.stderr)
        raise SystemExit(130)

    for name in ("SIGINT", "SIGTERM", "SIGHUP"):
        sig = getattr(signal, name, None)
        if sig is not None:
            try:
                signal.signal(sig, handler)
            except (ValueError, OSError):
                pass


def _restore_all() -> None:
    for path, original in list(_BACKUP.items()):
        try:
            path.write_bytes(original)
        except OSError:
            pass
        _BACKUP.pop(path, None)


def _snapshot(paths) -> None:  # noqa: ANN001
    for path in paths:
        path = Path(path)
        if path.is_file():
            _BACKUP.setdefault(path, path.read_bytes())


# ── 字节级 IO ──────────────────────────────────────────────────────────────────
# **不要**用 `Path.read_text()` / `write_text()`：文本模式默认 `newline=None`，
# 读会把 CRLF 折成 LF、写又会按平台再翻译一次。这两个文件恰好是纯 LF 才没暴雷，
# 但拿它当通用工具用，任何 CRLF 目标文件 apply 一次就永久变成 LF，「逐字节一致」
# 这条判据自己就会失败。这里全程 bytes，换行是数据的一部分，不做任何规范化。
def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def target_files() -> dict[str, Path]:
    """每个变异各自的目标文件（判据 3/4 讲的是"文件逐字节一致"，
    所以基线按**文件**记、报错印**文件名**，不按变异编号记 —— 按编号记会把
    M1 的文件说成 M2 的问题，曾经真的这么错过一次）。"""
    return {name: Path(t["file"]) for name, t in TARGETS.items()}


def apply(which: str, restore: bool = False) -> tuple[bool, str]:
    target = TARGETS[which]
    path = Path(target["file"])
    _snapshot([path])

    if target["kind"] == "delete-file":
        if restore:
            if path.exists():
                return False, "ALREADY-PRESENT"
            path.write_bytes(_BACKUP[path])
            return True, "restored (文件写回)"
        if not path.exists():
            return False, "文件本来就不存在"
        path.unlink()
        return True, f"applied → 删除 {path.relative_to(REPO)}"

    text = path.read_bytes().decode("utf-8")
    if restore:
        count = text.count(target["mutant"])
        if count == 0:
            return False, "ALREADY-CLEAN"
        out = text.replace(target["mutant"], target["seed"])
        message = f"restored ({count} 处)"
    else:
        count = text.count(target["seed"])
        if count != 1:
            return False, f"SEED 出现 {count} 次（期望 1）"
        out = text.replace(target["seed"], target["mutant"])
        message = f"applied → {path.relative_to(REPO)}"

    path.write_bytes(out.encode("utf-8"))
    return True, message


def run_tests() -> tuple[int, str, int]:
    result = subprocess.run(
        [
            str(PYTHON), "-m", "pytest", TEST_FILE,
            "--basetemp=.tmp-pytest/bt-ribbon-mutation", "-p", "no:cacheprovider", "-q",
        ],
        cwd=REPO, capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    raw = re.sub(r"\x1b\[[0-9;]*m", "", result.stdout + result.stderr)
    summary = ""
    passed = 0
    for line in raw.splitlines():
        if re.search(r"\d+ (passed|failed)", line):
            summary = line.strip()
        match = re.search(r"(\d+) passed", line)
        if match:
            passed = int(match.group(1))
    return result.returncode, summary, passed


def main() -> int:
    _install_signal_hooks()

    files = target_files()
    baseline = {name: path.read_bytes() for name, path in files.items()}
    _snapshot(sorted(set(files.values())))
    for name, path in files.items():
        print(
            f"[mutate] 基线 {name} {path.relative_to(REPO)} "
            f"sha256={sha256(baseline[name])[:16]}"
        )

    # 前置状态闸：文本目标 SEED 恰好一次、变异态零次；文件目标必须存在。
    # 上一轮若被中断留下变异态，或另有一个实例并发在改同一个文件，都会在这里
    # 当场暴露，而不是等跑完四轮 pytest 之后误报成"还原后不一致"。
    for name, target in TARGETS.items():
        path = Path(target["file"])
        if target["kind"] == "delete-file":
            if not path.is_file():
                print(f"[mutate] 起始状态不对：{path.relative_to(REPO)} 不存在。", file=sys.stderr)
                return 1
            continue
        text = path.read_bytes().decode("utf-8")
        seed_count = text.count(target["seed"])
        mutant_count = text.count(target["mutant"])
        if seed_count != 1 or mutant_count != 0:
            print(
                f"[mutate] 起始状态不对：{path.relative_to(REPO)} "
                f"SEED={seed_count}（期望 1）变异态={mutant_count}（期望 0）。"
                f"文件可能停在上一次被中断的变异态，或被另一个实例并发修改。",
                file=sys.stderr,
            )
            return 1

    print("[mutate] 基线（未变异）应为 GREEN…")
    code, summary, passed = run_tests()
    if code != 0 or passed < MIN_TESTS:
        print(f"基线就不绿（code={code}, 摘要={summary or '无'}, passed={passed}）", file=sys.stderr)
        return 1
    print(f"[mutate] 基线 GREEN：{summary}")

    failures = 0
    for which in ("M1", "M2", "M3", "M4"):
        touched = Path(TARGETS[which]["file"])
        # apply 前再核一次：跑了四轮 pytest 之后，文件不能被别人动过。
        if not touched.is_file() or touched.read_bytes() != baseline[which]:
            print(
                f"[mutate] {which} 的目标文件在本轮开始前已被改动："
                f"{touched.relative_to(REPO)}", file=sys.stderr,
            )
            _restore_all()
            return 1

        ok, message = apply(which)
        if not ok:
            print(f"[mutate] {which} apply 失败：{message}", file=sys.stderr)
            _restore_all()
            return 1
        code, summary, _ = run_tests()
        verdict = "RED(OK)" if code != 0 else "GREEN(NG!)"
        print(f"[mutate] {which} {TARGETS[which]['note']} → {verdict} {summary or '(无摘要)'}")
        if code == 0:
            print(f"  {which} 拆掉守卫后仍然全绿 —— 断言没有判别力。", file=sys.stderr)
            failures += 1

        ok, message = apply(which, restore=True)
        if not ok:
            print(f"[mutate] {which} restore 失败：{message}", file=sys.stderr)
            _restore_all()
            return 1
        print(f"[mutate] {which} {message}")

        # 判据 3：只核**被改的那个**文件，且消息里印文件名；直接比字节。
        if not touched.is_file() or touched.read_bytes() != baseline[which]:
            print(
                f"[mutate] {which} 还原后 {touched.relative_to(REPO)} 与基线不一致",
                file=sys.stderr,
            )
            failures += 1

    for name, path in files.items():
        if not path.is_file() or path.read_bytes() != baseline[name]:
            print(f"[mutate] 最终 {path.relative_to(REPO)} 与基线不一致", file=sys.stderr)
            failures += 1
    print(f"[mutate] 全部文件最终与基线逐字节一致：{len(set(files.values()))} 个")

    print("[mutate] 还原后复跑应为 GREEN…")
    code, summary, passed = run_tests()
    if code != 0 or passed < MIN_TESTS:
        print(f"[mutate] 还原后复跑未全绿：{summary}", file=sys.stderr)
        failures += 1
    else:
        print(f"[mutate] 还原后 GREEN：{summary}")

    if failures:
        print(f"反向变异未通过：{failures} 处问题。", file=sys.stderr)
        return 1
    _BACKUP.clear()
    print("反向变异通过：M1–M4 各自变红，还原逐字节一致，复跑全绿。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
