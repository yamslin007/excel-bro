from __future__ import annotations

import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

from server.app.wps_installation import MAC_PUBLISH_PATH, update_registration


def test_registration_merges_preserves_other_plugins_and_is_idempotent(tmp_path: Path) -> None:
    path = tmp_path / MAC_PUBLISH_PATH
    path.parent.mkdir(parents=True)
    original = b'<jsplugins><!--keep--><jspluginonline name="other" type="et" url="https://example.com/"/><jspluginonline name="excel-bro-wps" type="wps" url="https://example.com/word"/></jsplugins>'
    path.write_bytes(original)
    update_registration(tmp_path, 3889)
    installed = path.read_bytes()
    root = ET.fromstring(installed)
    assert root.find("./jspluginonline[@name='other']").get("url") == "https://example.com/"
    assert root.find("./jspluginonline[@type='wps']").get("url") == "https://example.com/word"
    assert root.find("./jspluginonline[@name='excel-bro-wps'][@type='et']").get("url") == "http://127.0.0.1:3889/"
    assert b"<!--keep-->" in installed
    backups = list(path.parent.glob("publish.xml.excel-bro-backup-*"))
    assert len(backups) == 1 and backups[0].read_bytes() == original
    update_registration(tmp_path, 3889)
    assert path.read_bytes() == installed
    assert len(list(path.parent.glob("publish.xml.excel-bro-backup-*"))) == 1
    update_registration(tmp_path, 3889, remove=True)
    remaining = ET.fromstring(path.read_bytes())
    assert len(remaining) == 2
    assert remaining.find("./jspluginonline[@type='wps']") is not None


def test_updates_duplicate_own_entries_without_deleting_unrelated(tmp_path: Path) -> None:
    path = tmp_path / MAC_PUBLISH_PATH
    path.parent.mkdir(parents=True)
    path.write_text('<jsplugins><jspluginonline name="excel-bro-wps" type="et" url="http://127.0.0.1:3987/"/><jspluginonline name="excel-bro-wps" type="et"/><jsplugin name="other"/></jsplugins>')
    update_registration(tmp_path, 3889)
    root = ET.fromstring(path.read_bytes())
    assert len(root.findall("jspluginonline")) == 1
    assert root.find("./jsplugin[@name='other']") is not None


@pytest.mark.parametrize("content", [b"<broken", b"<other/>", b'<!DOCTYPE jsplugins [<!ENTITY x "data">]><jsplugins>&x;</jsplugins>'])
def test_invalid_existing_configuration_is_never_overwritten(tmp_path: Path, content: bytes) -> None:
    path = tmp_path / MAC_PUBLISH_PATH
    path.parent.mkdir(parents=True)
    path.write_bytes(content)
    with pytest.raises((ValueError, ET.ParseError)):
        update_registration(tmp_path, 3889)
    assert path.read_bytes() == content


def test_uninstall_without_registration_is_noop(tmp_path: Path) -> None:
    update_registration(tmp_path, 3889, remove=True)
    assert not (tmp_path / MAC_PUBLISH_PATH).exists()


def test_symlink_configuration_is_not_modified(tmp_path: Path) -> None:
    other = tmp_path / "other.xml"
    other.write_text("<jsplugins/>")
    path = tmp_path / MAC_PUBLISH_PATH
    path.parent.mkdir(parents=True)
    try:
        path.symlink_to(other)
    except OSError:
        pytest.skip("Host does not permit symlink creation")
    with pytest.raises(ValueError, match="符号链接"):
        update_registration(tmp_path, 3889)
    assert other.read_text() == "<jsplugins/>"
