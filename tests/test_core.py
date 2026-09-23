import csv
from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_diff import ExcelDiffError, compare


def workbook(path: Path, sheets: dict[str, list[list[object]]]):
    wb = Workbook()
    first = True
    for name, rows in sheets.items():
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = name
        for row in rows:
            ws.append(row)
    wb.save(path)


def test_identical_workbooks(tmp_path):
    a, b = tmp_path / "a.xlsx", tmp_path / "b.xlsx"
    data = {"Sales": [["Item", "Qty"], ["Book", 2]]}
    workbook(a, data); workbook(b, data)
    assert compare(a, b).identical


def test_detects_cells_and_sheets(tmp_path):
    a, b = tmp_path / "a.xlsx", tmp_path / "b.xlsx"
    workbook(a, {"Sales": [["Item", "Qty"], ["Book", 2]], "Old": [[1]]})
    workbook(b, {"Sales": [["Item", "Qty"], ["Book", 3]], "New": [[1]]})
    report = compare(a, b)
    assert report.added_sheets == ["New"]
    assert report.removed_sheets == ["Old"]
    assert [(c.sheet, c.cell, c.before, c.after) for c in report.changes] == [("Sales", "B2", 2, 3)]


def test_missing_cell_is_change(tmp_path):
    a, b = tmp_path / "a.xlsx", tmp_path / "b.xlsx"
    workbook(a, {"Sheet": [["x", "gone"]]}); workbook(b, {"Sheet": [["x"]]})
    change = compare(a, b).changes[0]
    assert change.cell == "B1" and change.before == "gone" and change.after is None


def test_ignore_case(tmp_path):
    a, b = tmp_path / "a.xlsx", tmp_path / "b.xlsx"
    workbook(a, {"Sheet": [["Mosul"]]}); workbook(b, {"Sheet": [["MOSUL"]]})
    assert compare(a, b, ignore_case=True).identical


def test_csv_utf8_and_difference(tmp_path):
    a, b = tmp_path / "a.csv", tmp_path / "b.csv"
    a.write_text("name,value\nموصل,1\n", encoding="utf-8")
    b.write_text("name,value\nموصل,2\n", encoding="utf-8")
    report = compare(a, b)
    assert report.changes[0].cell == "B2"


def test_unsupported_format(tmp_path):
    p = tmp_path / "x.txt"; p.write_text("x")
    with pytest.raises(ExcelDiffError):
        compare(p, p)
