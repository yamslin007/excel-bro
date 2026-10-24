"""原地撤销（restore）的自动化验收。

替身不是"再造一套模型"，而是把**真实的** `before.xlsx` 读回来：`SaveCopyAs` 用
openpyxl 写盘，`restore_session` 再用 openpyxl 打开它。这样被测的是
"归档里到底有什么"与"放回去之后目录/内容/可见性对不对"，而不是我们对归档的假设。

替身只模拟 COM 的形状：`Worksheets.Count/Item`、`Worksheet.Copy(Before, After)`、
`Worksheet.Cells(r, c).Formula`、`Workbook.Names`。删除时把幸存表里引用被删表的
公式改写成 `=#REF!...`，模拟真实宿主的行为 —— 这样"放回表之后 #REF! 不会自动复原"
这件事是被真的演出来的。
"""
from __future__ import annotations

import json
import os
import re
import time
from contextlib import contextmanager
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from server.app.models import WorksheetDeleteCheckpointRequest
from server.app.worksheet_checkpoint import WorksheetCheckpointService


class CellModel:
    def __init__(self, formula=""):
        self._formula = formula

    @property
    def Formula(self):
        return self._formula

    @Formula.setter
    def Formula(self, value):
        self._formula = value


class Count:
    def __init__(self, value):
        self.Count = value


class UsedRange:
    def __init__(self, sheet):
        self._sheet = sheet
        self.Rows = Count(sheet.rows())
        self.Columns = Count(sheet.columns())
        self.Row = 1
        self.Column = 1

    @property
    def Formula(self):
        sheet = self._sheet
        return tuple(
            tuple(sheet.cell(row, column).Formula for column in range(1, sheet.columns() + 1))
            for row in range(1, sheet.rows() + 1)
        )


class SheetModel:
    def __init__(self, name, book, visible=-1, cells=None):
        self.Name = name
        self.Parent = book
        self.Visible = visible
        self._cells = dict(cells or {})
        self.fail_copy = False

    def rows(self):
        return max([row for row, _ in self._cells] + [1])

    def columns(self):
        return max([column for _, column in self._cells] + [1])

    def cell(self, row, column):
        key = (row, column)
        if key not in self._cells:
            self._cells[key] = CellModel("")
        return self._cells[key]

    @property
    def UsedRange(self):
        return UsedRange(self)

    def Cells(self, row, column):
        return self.cell(row, column)

    def Delete(self):
        self.Parent.sheets.remove(self)

    def Copy(self, Before=None, After=None):
        if self.fail_copy:
            raise RuntimeError("copy failed")
        anchor = After if After is not None else Before
        if anchor is None:
            raise RuntimeError("copy without anchor is not modelled")
        target = anchor.Parent
        index = target.sheets.index(anchor) + (1 if After is not None else 0)
        clone = SheetModel(self.Name, target, self.Visible,
                           {key: CellModel(cell.Formula) for key, cell in self._cells.items()})
        target.sheets.insert(index, clone)
        return clone


class Collection:
    def __init__(self, items):
        self._items = items

    @property
    def Count(self):
        return len(self._items)

    def Item(self, index):
        return self._items[index - 1]


class NamesModel:
    def __init__(self, rows):
        self._rows = rows

    @property
    def Count(self):
        return len(self._rows)

    def Item(self, key):
        if isinstance(key, int):
            return self._rows[key - 1]
        for row in self._rows:
            if row.Name == key:
                return row
        raise KeyError(key)


class NameModel:
    def __init__(self, name, refers_to):
        self.Name = name
        self.RefersTo = refers_to


class AppModel:
    def __init__(self):
        self.DisplayAlerts = True
        self.EnableEvents = True
        self.AutomationSecurity = 1


class BookModel:
    def __init__(self, full_name, names=None, fail_on_copy=False):
        self.FullName = full_name
        self.Path = str(Path(full_name).parent)
        self.ReadOnly = False
        self.ProtectStructure = False
        self.FileFormat = 51
        self.Application = AppModel()
        self.sheets: list[SheetModel] = []
        self.names = list(names or [])
        self.copies = 0
        self.fail_on_copy = fail_on_copy

    @property
    def Worksheets(self):
        return Collection(self.sheets)

    @property
    def Sheets(self):
        return Collection(self.sheets)

    @property
    def Names(self):
        return NamesModel(self.names)

    def sheet(self, name):
        return next((sheet for sheet in self.sheets if sheet.Name == name), None)

    def SaveCopyAs(self, path):
        workbook = Workbook()
        workbook.remove(workbook.active)
        for sheet in self.sheets:
            written = workbook.create_sheet(sheet.Name)
            written.sheet_state = ("veryHidden" if sheet.Visible == 2
                                   else "hidden" if sheet.Visible == 0 else "visible")
            for (row, column), cell in sheet._cells.items():
                written.cell(row=row, column=column, value=cell.Formula or None)
        for item in self.names:
            # openpyxl 的 defined name 值不带前导 '='。
            workbook.defined_names.add(_defined_name(item.Name, item.RefersTo))
        workbook.save(path)

    def DeleteSheetLikeHost(self, name):
        """真实宿主删表后：引用它的公式与名称定义立刻变成 #REF!。"""
        pattern = re.compile(r"(?:'" + re.escape(name) + r"'|" + re.escape(name) + r")!", re.IGNORECASE)
        for sheet in self.sheets:
            if sheet.Name == name:
                continue
            for cell in sheet._cells.values():
                if isinstance(cell.Formula, str) and pattern.search(cell.Formula):
                    cell.Formula = pattern.sub("#REF!", cell.Formula)
        for item in self.names:
            if pattern.search(item.RefersTo):
                item.RefersTo = pattern.sub("#REF!", item.RefersTo)
        target = self.sheet(name)
        if target is not None:
            self.sheets.remove(target)


def _defined_name(name, refers_to):
    from openpyxl.workbook.defined_name import DefinedName
    return DefinedName(name, attr_text=refers_to.lstrip("="))


def _book_from_archive(path):
    """把归档读成一个可选表的替身宿主工作簿（只读使用，不写回）。"""
    workbook = load_workbook(path)
    book = BookModel(str(path))
    for title in workbook.sheetnames:
        sheet = workbook[title]
        state = sheet.sheet_state
        visible = -1 if state == "visible" else 2 if state == "veryHidden" else 0
        cells = {}
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                cells[(cell.row, cell.column)] = CellModel(str(cell.value))
        book.sheets.append(SheetModel(title, book, visible, cells))
    for item in workbook.defined_names.values():
        ref = item.attr_text
        book.names.append(NameModel(item.name, ref if ref.startswith("=") else "=" + ref))
    workbook.close()
    return book


class Native:
    def __init__(self, book):
        self.book = book
        self.opened = []

    @contextmanager
    def workbook(self, request):
        yield self.book

    def open_copy(self, request, path):
        self.opened.append(path)

    @contextmanager
    def restore_session(self, request, archive):
        archive_book = _book_from_archive(archive)
        try:
            yield self.book, archive_book
        finally:
            archive_book.sheets.clear()


def make_book(book_path, *, hidden=(), extra_sheet=False):
    """执行前的现场：Sheet1 引用 Sheet2/Sheet3，还有一条工作簿级名称。"""
    book = BookModel(book_path, names=[NameModel("汇总", "=SUM(Sheet2!A1:A2)+SUM(Sheet3!A1)")])
    for name in ("Sheet1", "Sheet2", "Sheet3", "Sheet4"):
        book.sheets.append(SheetModel(name, book, 0 if name in hidden else -1))
    book.sheets[0].cell(1, 1).Formula = "=Sheet2!A1+Sheet3!A1"
    book.sheets[1].cell(1, 1).Formula = "1"
    book.sheets[1].cell(2, 1).Formula = "2"
    book.sheets[2].cell(1, 1).Formula = "=Sheet2!A1*2"
    book.sheets[3].cell(1, 1).Formula = "=Sheet2!A1"
    return book


def request(book, **updates):
    payload = dict(
        operationId="b" * 32, host="wps-et", hostWindowId=11,
        workbookPath=book.FullName,
        authorizedSheets=[sheet.Name for sheet in book.sheets],
        plan={"id": "delete-plan", "title": "删除中间三张", "summary": "留 Sheet1",
              "actions": [{"type": "deleteWorksheet", "sheet": "Sheet2"},
                          {"type": "deleteWorksheet", "sheet": "Sheet3"},
                          {"type": "deleteWorksheet", "sheet": "Sheet4"}]})
    payload.update(updates)
    return WorksheetDeleteCheckpointRequest.model_validate(payload)


@pytest.fixture
def env(tmp_path, monkeypatch):
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path))
    book_path = str(tmp_path / "444.xlsx")
    book = make_book(book_path, hidden=("Sheet4",))
    native = Native(book)
    service = WorksheetCheckpointService(native)
    payload = request(book)
    # 先真实执行一次删除：归档、状态、回执全部走产品代码。
    receipt = service.execute(payload, "https://pane")
    assert receipt.status == "completed"
    for name in receipt.deletedSheets:
        book.DeleteSheetLikeHost(name)
    return service, native, book, payload


def test_restore_puts_sheets_back_in_place_with_objects_and_visibility(env):
    service, native, book, payload = env
    assert [sheet.Name for sheet in book.sheets] == ["Sheet1"]

    result = service.restore(payload.operationId, "https://pane")

    assert result.status == "completed", result.message
    assert result.restoredSheets == ["Sheet2", "Sheet3", "Sheet4"]
    assert [sheet.Name for sheet in book.sheets] == ["Sheet1", "Sheet2", "Sheet3", "Sheet4"]
    assert [sheet.Visible for sheet in book.sheets] == [-1, -1, -1, 0], "可见性必须与执行前一致"
    # 内容不只是"表在"：公式原文也回来了。
    assert book.sheet("Sheet2").cell(1, 1).Formula == "1"
    assert book.sheet("Sheet3").cell(1, 1).Formula == "=Sheet2!A1*2"
    # 原工作簿身份没变、没有打开任何新窗口。
    assert book.FullName.endswith("444.xlsx")
    assert native.opened == []


def test_restore_puts_sheets_back_before_a_later_survivor(tmp_path, monkeypatch):
    """删的是**中间一段**、后面还有幸存表：放回原位 ≠ 追加到末尾。

    只留 Sheet1 的现场分辨不出这两种实现（追加到末尾恰好等于原顺序），
    所以这里留一张 Sheet5 在后面 —— 位置判据必须挂在"它与末尾的差别"上。
    """
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path))
    book_path = str(tmp_path / "with-tail.xlsx")
    book = BookModel(book_path)
    for name in ("Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"):
        book.sheets.append(SheetModel(name, book))
    native = Native(book)
    service = WorksheetCheckpointService(native)
    payload = request(book)
    receipt = service.execute(payload, "https://pane")
    assert receipt.status == "completed", receipt.message
    for name in receipt.deletedSheets:
        book.DeleteSheetLikeHost(name)
    assert [sheet.Name for sheet in book.sheets] == ["Sheet1", "Sheet5"]

    result = service.restore(payload.operationId, "https://pane")

    assert result.status == "completed", result.message
    names = [sheet.Name for sheet in book.sheets]
    assert names == ["Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"], (
        f"放回顺序必须与执行前一致，实际是 {names}")


def test_restore_repairs_ref_formulas_and_workbook_names(env):
    service, native, book, payload = env
    assert "#REF!" in book.sheet("Sheet1").cell(1, 1).Formula
    assert "#REF!" in book.sheet("Sheet4").cell(1, 1).Formula if book.sheet("Sheet4") else True

    result = service.restore(payload.operationId, "https://pane")

    assert result.status == "completed", result.message
    # 只放回表并不会自动修好 #REF!：验收必须看到公式真的被改回来。
    assert book.sheet("Sheet1").cell(1, 1).Formula == "=Sheet2!A1+Sheet3!A1"
    assert book.sheet("Sheet4").cell(1, 1).Formula == "=Sheet2!A1"
    assert result.repairedFormulas >= 1
    name = book.Names.Item("汇总")
    assert name.RefersTo == "=SUM(Sheet2!A1:A2)+SUM(Sheet3!A1)"
    assert result.repairedNames == 1


def test_restore_keeps_later_edits_on_survivors(env):
    service, native, book, payload = env
    book.sheet("Sheet1").cell(5, 5).Formula = "用户后来写的"
    result = service.restore(payload.operationId, "https://pane")
    assert result.status == "completed", result.message
    assert book.sheet("Sheet1").cell(5, 5).Formula == "用户后来写的"


def test_restore_refuses_when_same_name_sheet_exists(env):
    service, native, book, payload = env
    book.sheets.append(SheetModel("Sheet2", book))
    result = service.restore(payload.operationId, "https://pane")
    assert result.status == "conflict"
    assert "Sheet2" in result.message
    # 不覆盖、不静默成功：目录原样。
    assert [sheet.Name for sheet in book.sheets] == ["Sheet1", "Sheet2"]
    assert book.sheet("Sheet2").cell(1, 1).Formula == ""


def test_partial_restore_is_reported_and_retry_does_not_duplicate(env):
    service, native, book, payload = env
    # 让第三张表的复制失败：前两张已经放回去了。
    original_copy = SheetModel.Copy
    calls = {"n": 0}

    def flaky(self, Before=None, After=None):
        calls["n"] += 1
        if calls["n"] == 3:
            raise RuntimeError("copy failed")
        return original_copy(self, Before, After)

    SheetModel.Copy = flaky
    try:
        first = service.restore(payload.operationId, "https://pane")
    finally:
        SheetModel.Copy = original_copy
    assert first.status == "partial"
    assert first.restoredSheets == ["Sheet2", "Sheet3"]
    # A raised Copy is not proof that the host made no change. The retry
    # re-reads the directory and only proceeds once absence is established.
    assert first.unverifiedSheets == ["Sheet4"]
    assert [sheet.Name for sheet in book.sheets] == ["Sheet1", "Sheet2", "Sheet3"]

    retry = service.restore(payload.operationId, "https://pane")
    assert retry.status == "completed", retry.message
    assert [sheet.Name for sheet in book.sheets] == ["Sheet1", "Sheet2", "Sheet3", "Sheet4"], "重试不得重复造表"
    assert retry.restoredSheets.count("Sheet2") == 1


def test_second_restore_is_idempotent(env):
    service, native, book, payload = env
    first = service.restore(payload.operationId, "https://pane")
    again = service.restore(payload.operationId, "https://pane")
    assert again == first
    assert [sheet.Name for sheet in book.sheets] == ["Sheet1", "Sheet2", "Sheet3", "Sheet4"]


def test_restore_keeps_backup_copy_available(env):
    service, native, book, payload = env
    service.restore(payload.operationId, "https://pane")
    opened = service.open(payload.operationId, "https://pane")
    assert opened["opened"]
    assert native.opened and native.opened[0].name.startswith("recovered-")


def test_newer_deletion_blocks_older_restore(env):
    service, native, book, payload = env
    # 再执行一次删除（更晚），且尚未撤销。
    if book.sheet("Sheet4") is None:
        book.sheets.append(SheetModel("Sheet4", book))
    time.sleep(0.01)
    second = request(book, operationId="c" * 32,
                     authorizedSheets=["Sheet1", "Sheet4"],
                     plan={"id": "delete-later", "title": "再删一张", "summary": "删 Sheet4",
                           "actions": [{"type": "deleteWorksheet", "sheet": "Sheet4"}]})
    later = service.execute(second, "https://pane")
    assert later.status == "completed"

    with pytest.raises(ValueError) as blocked:
        service.restore(payload.operationId, "https://pane")
    assert "更晚" in str(blocked.value)


def test_damaged_checkpoint_never_writes(env):
    service, native, book, payload = env
    # 直接按配置根定位这一次的检查点：rglob 会把同一次运行里别的用例的记录也算进来。
    folder = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "worksheet-checkpoints" / payload.operationId
    archive = folder / "before.xlsx"
    assert archive.is_file()
    archive.write_bytes(b"not a zip")
    with pytest.raises(ValueError) as blocked:
        service.restore(payload.operationId, "https://pane")
    assert "校验失败" in str(blocked.value)
    assert [sheet.Name for sheet in book.sheets] == ["Sheet1"]


def test_incomplete_deletion_is_not_restorable(tmp_path, monkeypatch):
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path))
    book_path = str(tmp_path / "incomplete.xlsx")
    book = make_book(book_path, extra_sheet=True)
    book.fail_on_copy = True
    service = WorksheetCheckpointService(Native(book))
    # 让归档校验失败 ⇒ 删除回执是 failed。
    original = BookModel.SaveCopyAs
    BookModel.SaveCopyAs = lambda self, path: Path(path).write_bytes(b"broken")
    try:
        receipt = service.execute(request(book), "https://pane")
    finally:
        BookModel.SaveCopyAs = original
    assert receipt.status == "failed"
    with pytest.raises(ValueError) as blocked:
        service.restore(receipt.operationId, "https://pane")
    assert "没有完整完成" in str(blocked.value)
