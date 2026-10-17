"""Ephemeral selection checks over registered files. Never reads cells or grants consent."""
from __future__ import annotations

import hashlib
import json
import os
import posixpath
import stat
from pathlib import Path
from xml.parsers import expat
from zipfile import ZipFile

from . import folder_workbooks as folders
from .capabilities import capability_int
from .models import FolderAdmissionRequest, FolderAdmissionResponse, ScopeAdmissionItem

_READ_ONLY = {"read_only", "protected", "write_unverified", "host_write_blocked", "locked", "directory_not_writable"}


def _item(file_id: str, name: str, reason: str):
    return ScopeAdmissionItem(id=f"{file_id}\x00{name}", reason=reason,
        status="ready" if reason == "none" else "read_only" if reason in _READ_ONLY else "unavailable")


def _metadata(archive: ZipFile, entry: str):
    if archive.getinfo(entry).file_size > capability_int("scopeAdmission", "maxMetadataBytes"):
        raise ValueError("metadata budget")
    nodes = []
    parser = expat.ParserCreate(namespace_separator="}")
    parser.StartElementHandler = lambda name, attributes: nodes.append((name.rsplit("}", 1)[-1], attributes))
    def reject(*_args):
        raise ValueError("DTD and external entities forbidden")
    parser.StartDoctypeDeclHandler = reject
    parser.ExternalEntityRefHandler = reject
    parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)
    parser.Parse(archive.read(entry), True)
    return nodes


def _protection(archive: ZipFile, entry: str) -> bool | None:
    # Parse only element names/attributes: business text is never delivered to a handler.
    if archive.getinfo(entry).file_size > capability_int("scopeAdmission", "maxMetadataBytes"):
        return None
    protected = False
    parser = expat.ParserCreate(namespace_separator="}")
    def start(name, attributes):
        nonlocal protected
        if name.rsplit("}", 1)[-1] == "sheetProtection":
            protected = attributes.get("sheet", "1").casefold() not in {"0", "false"}
    def reject(*_args):
        raise ValueError("DTD and external entities forbidden")
    parser.StartElementHandler = start
    parser.StartDoctypeDeclHandler = reject
    parser.ExternalEntityRefHandler = reject
    parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)
    with archive.open(entry) as source:
        for chunk in iter(lambda: source.read(65536), b""):
            parser.Parse(chunk, False)
    parser.Parse(b"", True)
    return protected


def inspect_file_selection(path: Path, names: list[str]) -> dict[str, str]:
    """Metadata and access probes only. Opening r+b never writes bytes or acquires consent."""
    with ZipFile(path) as archive:
        book = _metadata(archive, "xl/workbook.xml")
        relations = _metadata(archive, "xl/_rels/workbook.xml.rels")
        targets = {attributes["Id"]: attributes.get("Target", "") for name, attributes in relations
            if name == "Relationship" and attributes.get("TargetMode") != "External" and attributes.get("Type", "").endswith("/worksheet")}
        sheet_paths = {}
        for name, attributes in book:
            if name != "sheet":
                continue
            relation = next((value for key, value in attributes.items() if key.endswith("}id")), None)
            target = targets.get(relation, "")
            entry = posixpath.normpath(target.lstrip("/") if target.startswith("/") else "xl/" + target)
            if entry.startswith("xl/") and "\\" not in entry:
                sheet_paths[attributes.get("name")] = entry
        mode = path.stat()
        readonly = bool(getattr(mode, "st_file_attributes", 0) & getattr(stat, "FILE_ATTRIBUTE_READONLY", 1)) or not bool(mode.st_mode & stat.S_IWUSR)
        base = "read_only" if readonly else "directory_not_writable" if not os.access(path.parent, os.W_OK) else "locked" if folders._is_locked_for_writing(path) else "none"
        result = {}
        for name in names:
            if name not in sheet_paths:
                result[name] = "missing"
                continue
            protection = _protection(archive, sheet_paths[name])
            result[name] = base if base != "none" else "protected" if protection else "write_unverified" if protection is None else "none"
        return result


def check_folder_admission(request: FolderAdmissionRequest) -> FolderAdmissionResponse:
    session = folders._session(request.sessionId)
    root = session.root
    items = []
    fingerprints = []
    for selection in request.selections:
        path = session.files.get(selection.fileId)
        try:
            if path is None or not path.is_file():
                reasons = {name: "missing" for name in selection.sheets}
            elif not folders._is_relative_to(path.resolve(), root.resolve()) or folders._has_reparse_point(path, root):
                reasons = {name: "unsupported" for name in selection.sheets}
            else:
                before = folders._file_content_fingerprint(path)
                reasons = inspect_file_selection(path, selection.sheets)
                after = folders._file_content_fingerprint(path)
                if before != after:
                    reasons = {name: "target_changed" for name in selection.sheets}
                fingerprints.append([selection.fileId, selection.sheets, after])
        except Exception:
            # Never return paths, archive content, XML diagnostics, or business values.
            reasons = {name: "unreadable" for name in selection.sheets}
        items.extend(_item(selection.fileId, name, reasons[name]) for name in selection.sheets)
    fingerprint = hashlib.sha256(json.dumps(fingerprints, ensure_ascii=True).encode()).hexdigest() if all(i.status != "unavailable" for i in items) else None
    return FolderAdmissionResponse(items=items, fingerprint=fingerprint)
