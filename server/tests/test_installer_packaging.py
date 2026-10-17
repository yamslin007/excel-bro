from __future__ import annotations

import hashlib
import re
import xml.etree.ElementTree as ET
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]
BUILD_SCRIPT_PATH = PROJECT_ROOT / "packaging" / "build_installer.ps1"
INNO_SCRIPT_PATH = PROJECT_ROOT / "packaging" / "excel_bro.iss"
TASK_SCRIPT_PATH = PROJECT_ROOT / "packaging" / "install_tasks.ps1"
MANIFEST_PATH = PROJECT_ROOT / "apps" / "excel-addin" / "manifest.xml"
WPS_RIBBON_PATH = PROJECT_ROOT / "apps" / "wps-addin" / "public" / "ribbon.xml"
WPS_RIBBON_TS_PATH = PROJECT_ROOT / "apps" / "wps-addin" / "src" / "ribbon.ts"
EXCEL_ASSETS_PATH = PROJECT_ROOT / "apps" / "excel-addin" / "public" / "assets"
WPS_ASSETS_PATH = PROJECT_ROOT / "apps" / "wps-addin" / "public" / "assets"

# 本仓库的 manifest 是**开发用**清单，图标一律由本机 dev server 提供。
DEV_ASSET_URL_PREFIX = "https://localhost:3000/assets/"

_BT = "{http://schemas.microsoft.com/office/officeappbasictypes/1.0}"
_CUSTOM_UI = "{http://schemas.microsoft.com/office/2006/01/customui}"

# 两端功能区入口的**唯一权威文案**。改产品名就要改这里，改这里就必须两端一起改。
EXPECTED_RIBBON_LABELS = {
    "Tab.Label": "Excel Bro",
    "Group.Label": "Excel Bro",
    "Button.Label": "打开 Excel Bro",
}


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _excel_ribbon_labels() -> dict[str, str]:
    """从 Excel 加载项清单里取功能区入口文案（`bt:String` 资源）。"""
    root = ET.parse(MANIFEST_PATH).getroot()
    return {
        node.get("id"): node.get("DefaultValue")
        for node in root.iter(f"{_BT}String")
        if node.get("id")
    }


def _wps_ribbon_labels() -> dict[str, str]:
    """从 WPS customUI 的 ribbon.xml 里取同一组文案（标签/分组/按钮的 label 属性）。"""
    root = ET.parse(WPS_RIBBON_PATH).getroot()
    tab = root.find(f".//{_CUSTOM_UI}tab")
    assert tab is not None, "ribbon.xml 里找不到 tab"
    group = tab.find(f"{_CUSTOM_UI}group")
    assert group is not None, "ribbon.xml 里找不到 group"
    button = group.find(f"{_CUSTOM_UI}button")
    assert button is not None, "ribbon.xml 里找不到 button"
    return {
        "Tab.Label": tab.get("label"),
        "Group.Label": group.get("label"),
        "Button.Label": button.get("label"),
    }


def test_ribbon_entry_labels_match_across_both_hosts() -> None:
    """两端功能区入口的文案必须一致，且等于约定的那一份。

    Excel 的 `manifest.xml` 与 WPS 的 `ribbon.xml` 是两套**互相独立**的宿主配置，
    没有任何机制强制它们同步，所以很容易各改各的：本次统一之前 Excel 是
    「Excel Bro / AI 助手 / 打开 AI 对话」，WPS 是「Excel Bro / Excel Bro /
    打开 Excel Bro」—— 同一个产品在两个宿主里叫两个名字。

    这里做两重判断（**读的是两个不同的 XML 文件，不是同一份数据的复述**）：
    ① 两端逐字段相等；② 逐字段等于约定的产品文案。
    只留 ① 的话，两边同时被改成同一个错名字也照样通过。
    """
    excel = _excel_ribbon_labels()
    wps = _wps_ribbon_labels()

    for key, expected in EXPECTED_RIBBON_LABELS.items():
        assert excel.get(key), f"Excel manifest 缺少 {key}"
        assert wps.get(key), f"WPS ribbon.xml 缺少 {key}"
        assert excel[key] == wps[key], (
            f"{key} 两端不一致：Excel={excel[key]!r} WPS={wps[key]!r}"
        )
        assert excel[key] == expected, (
            f"{key} 与约定文案不一致：实际 {excel[key]!r}，期望 {expected!r}"
        )


# 图标文件名里的版本号。`excel-bro-icon-16-v2.png` → 去掉 `-v2` 得到别名
# `excel-bro-icon-16.png`。见 `_icon_alias_name` 的注释说明它为什么必须存在。
_VERSION_TOKEN = re.compile(r"-v(\d+)(?=\.[a-z0-9]+$)")


def _local_name(tag: str) -> str:
    """去掉 `{namespace}` 前缀，只留本地名。"""
    return tag.rsplit("}", 1)[-1]


def _excel_icon_urls() -> list[str]:
    """Excel manifest 里全部图标 URL：`IconUrl` / `HighResolutionIconUrl` /
    带 `DefaultValue` 的 `bt:Image` 资源。

    按**本地名**匹配而不是 `root.find("IconUrl")`：`OfficeApp` 根节点带**默认
    命名空间**（`.../appforoffice/1.1`），带命名空间的 `find("IconUrl")` 匹配不到，
    会把 `IconUrl` 与 `HighResolutionIconUrl` 静默漏掉 —— 本用例第一版就踩了这个
    坑，恰好被下面 `checked` 的计数断言抓住（只数到 4 个）。
    """
    root = ET.parse(MANIFEST_PATH).getroot()
    urls: list[str] = []
    for node in root.iter():
        if _local_name(node.tag) not in ("IconUrl", "HighResolutionIconUrl", "Image"):
            continue
        value = node.get("DefaultValue")
        if value:
            urls.append(value)
    return urls


def _wps_ribbon_icon_path() -> str:
    """WPS 功能区图标的返回路径（`ribbon.GetImage()` 里那个字符串）。

    读**源码**而不是读构建产物：产物是生成的，源码才是权威。
    """
    source = _read(WPS_RIBBON_TS_PATH)
    marker = source.index("GetImage()")
    match = re.search(r'return\s+"([^"]+)"', source[marker:])
    assert match, "ribbon.ts 的 GetImage() 里找不到 return \"...\""
    return match.group(1)


def _icon_alias_name(name: str) -> str | None:
    """带版本号的图标文件名 → 去版本号的别名；没有版本号则返回 None。"""
    return _VERSION_TOKEN.sub("", name) if _VERSION_TOKEN.search(name) else None


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_ribbon_icon_urls_resolve_to_existing_files() -> None:
    """两端功能区点名的图标文件必须真实存在。

    这条守卫防的是**改名漂移**：manifest 与 `ribbon.ts` 各自写了一份图标路径，
    改了一处没改另一处、或者改了名字没同步文件，都不会有任何编译或运行时报错 ——
    宿主只会安静地显示一个破图标或旧图标。2026-10-12 的实际故障就是这一族：
    图片文件换了、URL 没换，宿主按 URL 吃缓存，功能区一直显示退役的「EB」字标。
    """
    checked = 0

    for url in _excel_icon_urls():
        assert url.startswith(DEV_ASSET_URL_PREFIX), (
            f"manifest 图标 URL 不在开发资源前缀下：{url!r}。"
            f"若这是切换到正式 CDN，请同步更新本用例的 {DEV_ASSET_URL_PREFIX!r}。"
        )
        name = url[len(DEV_ASSET_URL_PREFIX):].split("?", 1)[0].split("#", 1)[0]
        assert "/" not in name, f"图标 URL 带了子目录，本用例不覆盖：{url!r}"
        assert (EXCEL_ASSETS_PATH / name).is_file(), (
            f"manifest 指向的图标不存在：{name}（{EXCEL_ASSETS_PATH}）"
        )
        checked += 1

    wps_name = _wps_ribbon_icon_path()
    assert wps_name.startswith("assets/"), (
        f"WPS GetImage 返回值应以 assets/ 开头（相对 add-in 根解析）：{wps_name!r}"
    )
    assert (WPS_ASSETS_PATH / wps_name[len("assets/"):]).is_file(), (
        f"ribbon.ts 指向的图标不存在：{wps_name}（{WPS_ASSETS_PATH}）"
    )
    checked += 1

    assert checked >= 5, f"只检查到 {checked} 个图标，manifest 结构可能变了"


def test_ribbon_icon_legacy_aliases_match_versioned_files() -> None:
    """带版本号的图标，其**去版本号别名**必须存在且逐字节相同。

    为什么别名必须存在且同字节：两个宿主都按 URL 缓存功能区图标（实测 Excel 宿主
    一个多月没重新拉过图）。换图时若不同时换 URL，宿主永远拿缓存；而老 URL
    一旦 404，宿主又会显示破图标。所以 `scripts/build_logo_assets.py` 一次运行
    同时产出「带版本号的名字」和「无版本号的别名」，两份不可能漂移 —— 这条用例
    把那个不变量钉住。

    同时断言**至少有一个**图标带版本号：否则本用例会因为"没有别名可比"而空过，
    变成一条恒真的假守卫。
    """
    names = [
        url[len(DEV_ASSET_URL_PREFIX):].split("?", 1)[0].split("#", 1)[0]
        for url in _excel_icon_urls()
    ]
    names.append(_wps_ribbon_icon_path()[len("assets/"):])

    pairs = [(name, _icon_alias_name(name)) for name in names]
    versioned = [(name, alias) for name, alias in pairs if alias is not None]
    assert versioned, (
        "manifest / ribbon.ts 里的图标都不带版本号 —— 换图时宿主会吃缓存，"
        "请按 scripts/build_logo_assets.py 的 ICON_URL_VERSION 加版本号后缀。"
    )

    excel_dirs = (EXCEL_ASSETS_PATH, WPS_ASSETS_PATH)
    for name, alias in versioned:
        existing = [d / name for d in excel_dirs if (d / name).is_file()]
        assert existing, f"带版本号的图标不存在：{name}"
        for path in existing:
            alias_path = path.parent / alias
            assert alias_path.is_file(), (
                f"别名缺失：{alias_path}（宿主老 URL 会 404 出破图标）"
            )
            assert _sha256(path) == _sha256(alias_path), (
                f"别名与带版本号的文件字节不一致：{alias_path} vs {path}"
            )



def test_build_uses_standard_inno_setup_compiler() -> None:
    script = _read(BUILD_SCRIPT_PATH)

    assert "Find-InnoSetupCompiler" in script
    assert "ISCC.exe" in script
    assert "excel_bro.iss" in script
    assert "/DSourceDir=" in script
    assert "--onefile" not in script
    assert "installer.py" not in script


def test_installer_shows_normal_directory_and_ready_pages() -> None:
    script = _read(INNO_SCRIPT_PATH)

    assert "DefaultDirName={localappdata}\\Programs\\{#AppName}" in script
    assert "DisableDirPage=no" in script
    assert "DisableReadyPage=no" in script
    assert "DisableFinishedPage=no" in script
    assert "WizardStyle=modern" in script
    assert "AllowNetworkDrive=no" in script
    assert "AllowRootDirectory=no" in script
    assert "AllowUNCPath=no" in script


def test_installer_force_closes_only_its_background_runtime_on_upgrade() -> None:
    script = _read(INNO_SCRIPT_PATH)

    assert "CloseApplications=force" in script
    assert "CloseApplicationsFilter={#AppExeName}" in script
    assert "CloseApplicationsFilter=EXCEL.EXE" not in script


def test_installer_uses_standard_uninstaller_and_start_menu_entry() -> None:
    script = _read(INNO_SCRIPT_PATH)

    assert "Uninstallable=yes" in script
    assert "CreateUninstallRegKey=yes" in script
    assert 'Filename: "{uninstallexe}"' in script
    assert "CurUninstallStepChanged" in script
    assert "IntegrationParameters('Uninstall')" in script


def test_installer_is_per_user_and_preserves_model_configuration() -> None:
    inno_script = _read(INNO_SCRIPT_PATH)
    task_script = _read(TASK_SCRIPT_PATH)

    assert "PrivilegesRequired=lowest" in inno_script
    assert "Root: HKCU" in inno_script
    assert "model-connections.json" not in inno_script
    assert 'Join-Path $env:LOCALAPPDATA "Excel Bro"' not in task_script


def test_packaged_manifest_meets_office_minimum_version_and_icon_requirements() -> None:
    manifest = _read(MANIFEST_PATH)
    build_script = _read(BUILD_SCRIPT_PATH)

    assert "<Version>1.0.0.0</Version>" in manifest
    assert '<IconUrl DefaultValue="https://localhost:3000/' in manifest
    assert '<HighResolutionIconUrl DefaultValue="https://localhost:3000/' in manifest
    assert '$manifestVersion -lt [Version]"1.0.0.0"' in build_script
    assert "IconUrl 和 HighResolutionIconUrl" in build_script


def test_catalog_share_is_scoped_and_requires_elevation() -> None:
    script = _read(TASK_SCRIPT_PATH)

    assert "New-SmbShare" in script
    assert "-Name $shareLiteral" in script
    assert "-Path $catalogLiteral" in script
    assert "-ReadAccess $principalLiteral" in script
    assert "-Verb RunAs" in script
    assert "Remove-SmbShare" in script


def test_custom_local_install_directory_is_allowed_but_root_is_rejected() -> None:
    script = _read(TASK_SCRIPT_PATH)

    assert 'GetFileName($resolved) -ne "Excel Bro"' not in script
    assert "不能把 Excel Bro 直接安装到磁盘根目录" in script


def test_uninstall_rejects_an_unexpected_share_path() -> None:
    script = _read(TASK_SCRIPT_PATH)

    assert "共享路径与当前安装目录不一致，拒绝删除" in script
    assert "$actualPath.Equals(" in script
    assert "[StringComparison]::OrdinalIgnoreCase" in script


def test_certificate_cleanup_covers_previous_installer_builds() -> None:
    script = _read(TASK_SCRIPT_PATH)

    assert '$certificateSubject = "CN=Excel Bro localhost"' in script
    assert "Install-RootCertificate" in script
    assert "Remove-RootCertificates" in script
    assert "StoreLocation]::CurrentUser" in script
    assert "$store.Remove(" in script


def test_runtime_is_health_checked_during_install() -> None:
    script = _read(TASK_SCRIPT_PATH)

    assert "Start-ExcelBroRuntime" in script
    assert "http://127.0.0.1:8765/health" in script
    assert "本地服务未能启动" in script
    assert "$runtimeProcess.HasExited" in script
    assert "$listener.OwningProcess -eq $runtimeProcess.Id" in script


def test_legacy_installer_is_migrated_without_touching_user_data() -> None:
    script = _read(TASK_SCRIPT_PATH)

    assert "Remove-LegacyInstall" in script
    assert 'Join-Path $legacyDir "uninstall.ps1"' in script
    assert "Remove-Item -LiteralPath $legacyDir -Recurse -Force" in script
    assert 'Join-Path $env:LOCALAPPDATA "Excel Bro"' not in script


def test_all_old_excel_bro_processes_are_stopped_before_health_check() -> None:
    script = _read(TASK_SCRIPT_PATH)

    assert 'Get-Process -Name "ExcelBro"' in script
    assert "Stop-Process -Force" in script


def test_catalog_url_uses_the_actual_computer_name() -> None:
    script = _read(INNO_SCRIPT_PATH)

    assert "GetEnv('COMPUTERNAME')" in script
    assert "\\ExcelBroAddins" in script
