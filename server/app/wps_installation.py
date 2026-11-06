"""Installer-only WPS registration; never reads or changes workbook data."""
from __future__ import annotations

import os
import shutil
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

PLUGIN_NAME = "excel-bro-wps"
MAC_PUBLISH_PATH = Path("Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml")


def update_registration(user_home: Path, port: int, *, remove: bool = False) -> Path:
    if not 1 <= port <= 65535:
        raise ValueError("WPS 端口无效。")
    home = user_home.resolve(strict=True)
    target = home / MAC_PUBLISH_PATH
    for candidate in (target, *target.parents):
        if candidate == home:
            break
        if candidate.is_symlink():
            raise ValueError("WPS 注册路径包含符号链接，未修改。")
    if not target.exists() and remove:
        return target
    original = target.read_bytes() if target.exists() else b""
    if b"<!DOCTYPE" in original.upper() or b"<!ENTITY" in original.upper():
        raise ValueError("WPS 配置包含不支持的 XML 声明，未修改。")
    parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True))
    root = ET.fromstring(original, parser=parser) if original else ET.Element("jsplugins")
    if root.tag != "jsplugins":
        raise ValueError("WPS 配置根节点不正确，未修改。")
    owned = [node for node in root if node.tag == "jspluginonline"
             and node.get("name") == PLUGIN_NAME and node.get("type") == "et"]
    attributes = dict(name=PLUGIN_NAME, type="et", url=f"http://127.0.0.1:{port}/",
                      debug="", enable="enable_dev", install="null")
    if (remove and not owned) or (not remove and len(owned) == 1 and owned[0].attrib == attributes):
        return target
    for node in owned:
        root.remove(node)
    if not remove:
        ET.SubElement(root, "jspluginonline", attributes)
    content = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    target.parent.mkdir(parents=True, exist_ok=True)
    if original:
        # Preserve the previous configuration, including unrelated plugins.
        with tempfile.NamedTemporaryFile(prefix="publish.xml.excel-bro-backup-", dir=target.parent, delete=False) as backup:
            backup.write(original)
    temporary: str | None = None
    try:
        with tempfile.NamedTemporaryFile(dir=target.parent, delete=False) as output:
            temporary = output.name
            output.write(content)
            output.flush()
            os.fsync(output.fileno())
        if target.exists():
            shutil.copymode(target, temporary)
        os.replace(temporary, target)
        temporary = None
    finally:
        if temporary:
            Path(temporary).unlink(missing_ok=True)
    return target
