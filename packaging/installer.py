from __future__ import annotations

import base64
import ctypes
import os
import shutil
import subprocess
import sys
import time
import traceback
import urllib.request
import winreg
import zipfile
from pathlib import Path


MANIFEST_ID = "9c758d40-c2b8-42d8-a6bc-735bd5c4f34c"
CATALOG_ID = "{41f62f5c-cd95-44f2-a0c6-c8cb847fe4e0}"
CATALOG_SHARE_NAME = "ExcelBroAddins"
DEVELOPER_KEY = r"Software\Microsoft\Office\16.0\Wef\Developer"
TRUSTED_CATALOG_KEY = (
    rf"Software\Microsoft\Office\16.0\Wef\TrustedCatalogs\{CATALOG_ID}"
)
SEE_MASK_NOCLOSEPROCESS = 0x00000040
SW_HIDE = 0
INFINITE = 0xFFFFFFFF


class _ShellExecuteInfo(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_ulong),
        ("fMask", ctypes.c_ulong),
        ("hwnd", ctypes.c_void_p),
        ("lpVerb", ctypes.c_wchar_p),
        ("lpFile", ctypes.c_wchar_p),
        ("lpParameters", ctypes.c_wchar_p),
        ("lpDirectory", ctypes.c_wchar_p),
        ("nShow", ctypes.c_int),
        ("hInstApp", ctypes.c_void_p),
        ("lpIDList", ctypes.c_void_p),
        ("lpClass", ctypes.c_wchar_p),
        ("hkeyClass", ctypes.c_void_p),
        ("dwHotKey", ctypes.c_ulong),
        ("hIconOrMonitor", ctypes.c_void_p),
        ("hProcess", ctypes.c_void_p),
    ]


def _payload_root() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


def _set_registry_string(root, path: str, name: str, value: str) -> None:
    with winreg.CreateKeyEx(root, path, 0, winreg.KEY_SET_VALUE) as key:
        winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)


def _set_registry_dword(root, path: str, name: str, value: int) -> None:
    with winreg.CreateKeyEx(root, path, 0, winreg.KEY_SET_VALUE) as key:
        winreg.SetValueEx(key, name, 0, winreg.REG_DWORD, value)


def _delete_registry_value(root, path: str, name: str) -> None:
    try:
        with winreg.OpenKey(root, path, 0, winreg.KEY_SET_VALUE) as key:
            winreg.DeleteValue(key, name)
    except FileNotFoundError:
        return


def _message(text: str, title: str, error: bool = False) -> None:
    flags = 0x10 if error else 0x40
    ctypes.windll.user32.MessageBoxW(None, text, title, flags)


def _powershell_literal(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def _encode_powershell(script: str) -> str:
    return base64.b64encode(script.encode("utf-16-le")).decode("ascii")


def _run_elevated_powershell(script: str) -> None:
    powershell = (
        Path(os.environ["SystemRoot"])
        / "System32"
        / "WindowsPowerShell"
        / "v1.0"
        / "powershell.exe"
    )
    info = _ShellExecuteInfo()
    info.cbSize = ctypes.sizeof(_ShellExecuteInfo)
    info.fMask = SEE_MASK_NOCLOSEPROCESS
    info.lpVerb = "runas"
    info.lpFile = str(powershell)
    info.lpParameters = (
        f"-NoProfile -EncodedCommand {_encode_powershell(script)}"
    )
    info.nShow = SW_HIDE
    shell_execute = ctypes.windll.shell32.ShellExecuteExW
    shell_execute.argtypes = [ctypes.POINTER(_ShellExecuteInfo)]
    shell_execute.restype = ctypes.c_bool
    wait_for_process = ctypes.windll.kernel32.WaitForSingleObject
    wait_for_process.argtypes = [ctypes.c_void_p, ctypes.c_ulong]
    wait_for_process.restype = ctypes.c_ulong
    get_exit_code = ctypes.windll.kernel32.GetExitCodeProcess
    get_exit_code.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong)]
    get_exit_code.restype = ctypes.c_bool
    close_handle = ctypes.windll.kernel32.CloseHandle
    close_handle.argtypes = [ctypes.c_void_p]
    close_handle.restype = ctypes.c_bool
    if not shell_execute(ctypes.byref(info)):
        error_code = ctypes.windll.kernel32.GetLastError()
        if error_code == 1223:
            raise RuntimeError(
                "已取消管理员授权，无法创建 Excel 本机加载项目录。"
            )
        raise ctypes.WinError(error_code)
    try:
        wait_for_process(info.hProcess, INFINITE)
        exit_code = ctypes.c_ulong()
        if not get_exit_code(info.hProcess, ctypes.byref(exit_code)):
            raise ctypes.WinError(ctypes.windll.kernel32.GetLastError())
        if exit_code.value != 0:
            raise RuntimeError(
                "创建 Excel 本机加载项目录失败，"
                f"管理员进程退出代码为 {exit_code.value}。"
            )
    finally:
        close_handle(info.hProcess)


def _catalog_share_script(catalog_dir: Path, principal: str) -> str:
    return "\n".join(
        [
            "$ErrorActionPreference = 'Stop'",
            (
                f"$existing = Get-SmbShare -Name "
                f"{_powershell_literal(CATALOG_SHARE_NAME)} "
                "-ErrorAction SilentlyContinue"
            ),
            (
                "if ($null -ne $existing) { "
                f"Remove-SmbShare -Name "
                f"{_powershell_literal(CATALOG_SHARE_NAME)} "
                "-Force -Confirm:$false }"
            ),
            (
                "New-SmbShare "
                f"-Name {_powershell_literal(CATALOG_SHARE_NAME)} "
                f"-Path {_powershell_literal(str(catalog_dir))} "
                f"-ReadAccess {_powershell_literal(principal)} "
                "| Out-Null"
            ),
        ]
    )


def _current_windows_principal() -> str:
    username = os.environ.get("USERNAME", "").strip()
    domain = os.environ.get("USERDOMAIN", "").strip()
    if not username:
        raise RuntimeError("无法确定当前 Windows 用户，不能限制加载项目录权限。")
    return f"{domain}\\{username}" if domain else username


def _catalog_url() -> str:
    computer_name = os.environ.get("COMPUTERNAME", "").strip()
    if not computer_name:
        raise RuntimeError("无法确定计算机名，不能注册 Excel 加载项目录。")
    return rf"\\{computer_name}\{CATALOG_SHARE_NAME}"


def _create_catalog_share(catalog_dir: Path) -> None:
    principal = _current_windows_principal()
    script = _catalog_share_script(catalog_dir, principal)
    _run_elevated_powershell(script)
    shared_manifest = Path(_catalog_url()) / "manifest.xml"
    if not shared_manifest.is_file():
        raise RuntimeError(
            f"无法读取本机加载项目录：{shared_manifest}。"
        )


def main() -> None:
    payload = _payload_root()
    local_app_data = Path(os.environ["LOCALAPPDATA"])
    install_dir = local_app_data / "Programs" / "Excel Bro"
    data_dir = local_app_data / "Excel Bro"
    data_dir.mkdir(parents=True, exist_ok=True)

    subprocess.run(
        ["taskkill", "/F", "/IM", "ExcelBro.exe"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    if install_dir.exists():
        shutil.rmtree(install_dir)
    install_dir.mkdir(parents=True)
    with zipfile.ZipFile(payload / "runtime.zip") as archive:
        archive.extractall(install_dir)
    with zipfile.ZipFile(payload / "frontend.zip") as archive:
        archive.extractall(install_dir / "frontend")
    catalog_dir = install_dir / "catalog"
    catalog_dir.mkdir()
    for name in (
        "localhost.cer",
        "localhost.crt",
        "localhost.key",
        "uninstall.ps1",
    ):
        shutil.copy2(payload / name, install_dir / name)
    shutil.copy2(payload / "manifest.xml", install_dir / "manifest.xml")
    shutil.copy2(payload / "manifest.xml", catalog_dir / "manifest.xml")

    subprocess.run(
        [
            "certutil.exe",
            "-user",
            "-addstore",
            "Root",
            str(install_dir / "localhost.cer"),
        ],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    _create_catalog_share(catalog_dir)
    catalog_url = _catalog_url()
    _delete_registry_value(
        winreg.HKEY_CURRENT_USER,
        DEVELOPER_KEY,
        MANIFEST_ID,
    )
    _set_registry_string(
        winreg.HKEY_CURRENT_USER,
        TRUSTED_CATALOG_KEY,
        "Id",
        CATALOG_ID,
    )
    _set_registry_string(
        winreg.HKEY_CURRENT_USER,
        TRUSTED_CATALOG_KEY,
        "Url",
        catalog_url,
    )
    _set_registry_dword(
        winreg.HKEY_CURRENT_USER,
        TRUSTED_CATALOG_KEY,
        "Flags",
        1,
    )
    _set_registry_string(
        winreg.HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows\CurrentVersion\Run",
        "Excel Bro",
        f'"{install_dir / "ExcelBro.exe"}"',
    )
    uninstall_key = (
        r"Software\Microsoft\Windows\CurrentVersion\Uninstall\Excel Bro"
    )
    uninstall_command = (
        'powershell.exe -NoProfile -ExecutionPolicy Bypass -File '
        f'"{install_dir / "uninstall.ps1"}"'
    )
    for name, value in {
        "DisplayName": "Excel Bro",
        "DisplayVersion": "0.1.0",
        "Publisher": "Excel Bro",
        "InstallLocation": str(install_dir),
        "DisplayIcon": str(install_dir / "ExcelBro.exe"),
        "UninstallString": uninstall_command,
        "QuietUninstallString": f"{uninstall_command} -Silent",
    }.items():
        _set_registry_string(
            winreg.HKEY_CURRENT_USER, uninstall_key, name, value
        )
    _set_registry_dword(
        winreg.HKEY_CURRENT_USER, uninstall_key, "NoModify", 1
    )
    _set_registry_dword(
        winreg.HKEY_CURRENT_USER, uninstall_key, "NoRepair", 1
    )
    subprocess.Popen(
        [str(install_dir / "ExcelBro.exe")],
        cwd=install_dir,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NO_WINDOW
        | subprocess.DETACHED_PROCESS,
        close_fds=True,
    )
    for _ in range(40):
        time.sleep(0.5)
        try:
            with urllib.request.urlopen(
                "http://127.0.0.1:8765/health", timeout=1
            ) as response:
                if response.status == 200:
                    break
        except OSError:
            continue
    else:
        raise RuntimeError(
            f"本地服务未能启动，请查看 {data_dir / 'excel-bro.log'}"
        )
    _message(
        "Excel Bro 已安装。\n\n"
        "首次使用请完全关闭并重新打开 Excel，然后依次选择：\n"
        "开始 → 加载项 → 更多加载项 → 高级 → 共享文件夹\n\n"
        "选择“Excel Bro”并点击“添加”。完成一次后，"
        "新工作簿也可以使用 Excel Bro。",
        "Excel Bro 安装完成",
    )


if __name__ == "__main__":
    log_path = Path(os.environ.get("TEMP", ".")) / "ExcelBro-install.log"
    try:
        main()
    except Exception:
        log_path.write_text(traceback.format_exc(), encoding="utf-8")
        _message(
            f"Excel Bro 安装失败。\n\n详细信息：{log_path}",
            "Excel Bro",
            error=True,
        )
        raise
