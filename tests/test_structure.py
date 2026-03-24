"""Tests for Stage 2 — Structure Detection."""
import pytest

from table2db.models import SheetData, WorkbookData
from table2db.pipeline.structure import detect_structure


def _make_wb(*sheets: SheetData, source: str = "test.xlsx") -> WorkbookData:
    return WorkbookData(source_file=source, sheets=list(sheets))


def test_detect_simple_header():
    sheet = SheetData(
        name="Test",
        rows=[
            ["id", "name", "val"],
            [1, "a", 10],
            [2, "b", 20],
        ],
    )
    wb, warnings = detect_structure(_make_wb(sheet))
    assert len(wb.sheets) == 1
    s = wb.sheets[0]
    assert s.header_row_start == 0
    assert s.headers == ["id", "name", "val"]
    assert len(s.rows) == 2
    assert s.rows[0] == [1, "a", 10]


def test_detect_offset_header():
    sheet = SheetData(
        name="Test",
        rows=[
            ["Report Title", None, None],
            [None, None, None],
            ["id", "name", "val"],
            [1, "a", 10],
            [2, "b", 20],
            [3, "c", 30],
        ],
    )
    wb, warnings = detect_structure(_make_wb(sheet))
    s = wb.sheets[0]
    assert s.header_row_start == 2
    assert s.headers == ["id", "name", "val"]
    assert len(s.rows) == 3


def test_normalize_duplicate_names():
    sheet = SheetData(
        name="Test",
        rows=[
            ["Name", "Name", "Name"],
            [1, 2, 3],
            [4, 5, 6],
            [7, 8, 9],
        ],
    )
    wb, _ = detect_structure(_make_wb(sheet))
    assert wb.sheets[0].headers == ["Name", "Name_1", "Name_2"]


def test_normalize_empty_names():
    sheet = SheetData(
        name="Test",
        rows=[
            [None, "a", None, "b", "c"],
            [1, 2, 3, 4, 5],
            [4, 5, 6, 7, 8],
            [7, 8, 9, 10, 11],
        ],
    )
    wb, _ = detect_structure(_make_wb(sheet))
    headers = wb.sheets[0].headers
    assert headers[0] == "column_1"
    assert headers[1] == "a"
    assert headers[2] == "column_3"


def test_normalize_whitespace():
    sheet = SheetData(
        name="Test",
        rows=[
            [" Name \n", "  Age  ", "Score"],
            [1, 2, 3],
            [4, 5, 6],
            [7, 8, 9],
        ],
    )
    wb, _ = detect_structure(_make_wb(sheet))
    assert wb.sheets[0].headers[0] == "Name"
    assert wb.sheets[0].headers[1] == "Age"


def test_multi_level_header():
    sheet = SheetData(
        name="Test",
        rows=[
            ["ID", "Personal Info", None, "Financial", None],
            [None, "Name", "Age", "Salary", "Bonus"],
            [1, "Alice", 30, 50000, 5000],
            [2, "Bob", 25, 45000, 4500],
            [3, "Charlie", 35, 60000, 6000],
        ],
        merge_map={
            # A1:A2 merged (col 0, rows 0-1)
            (0, 0): "ID",
            (1, 0): "ID",
            # B1:C1 merged (row 0, cols 1-2)
            (0, 1): "Personal Info",
            (0, 2): "Personal Info",
            # D1:E1 merged (row 0, cols 3-4)
            (0, 3): "Financial",
            (0, 4): "Financial",
        },
    )
    wb, _ = detect_structure(_make_wb(sheet))
    s = wb.sheets[0]
    assert s.headers == [
        "ID",
        "Personal Info_Name",
        "Personal Info_Age",
        "Financial_Salary",
        "Financial_Bonus",
    ]
    assert len(s.rows) == 3


def test_skip_title_row():
    sheet = SheetData(
        name="Test",
        rows=[
            ["Big Title", None, None, None, None],
            ["id", "name", "val", "extra", "more"],
            [1, "a", 10, "x", "y"],
            [2, "b", 20, "x", "y"],
            [3, "c", 30, "x", "y"],
        ],
    )
    wb, _ = detect_structure(_make_wb(sheet))
    s = wb.sheets[0]
    assert s.header_row_start == 1
    assert s.headers[0] == "id"


def test_empty_sheet_removed():
    sheet = SheetData(name="Empty", rows=[])
    wb, warnings = detect_structure(_make_wb(sheet))
    assert len(wb.sheets) == 0
    assert any("removed" in w.lower() for w in warnings)


def test_trailing_empty_rows_trimmed():
    sheet = SheetData(
        name="Test",
        rows=[
            ["id", "name", "val"],
            [1, "a", 10],
            [2, "b", 20],
            [3, "c", 30],
            [None, None, None],
            [None, None, None],
        ],
    )
    wb, _ = detect_structure(_make_wb(sheet))
    assert len(wb.sheets[0].rows) == 3


def test_multi_table_warning():
    rows = [
        ["id", "name", "val"],
        [1, "a", 10],
        [2, "b", 20],
        [None, None, None],
        [None, None, None],
        [None, None, None],
        ["code", "desc", None],
        ["X", "xray", None],
        ["Y", "yank", None],
    ]
    sheet = SheetData(name="Test", rows=rows)
    wb, warnings = detect_structure(_make_wb(sheet))
    # Island detector should split into two tables
    assert any("splitting" in w or "detected" in w for w in warnings)
    assert len(wb.sheets) == 2


def test_multi_table_detection():
    """Two tables separated by empty rows → split into two SheetData."""
    rows = []
    # Table 1 header + 3 data rows
    rows.append(["id", "name", "value"])
    rows.append([1, "Alice", 100])
    rows.append([2, "Bob", 200])
    rows.append([3, "Charlie", 300])
    # Gap
    rows.append([None, None, None])
    rows.append([None, None, None])
    # Table 2 header + 3 data rows
    rows.append(["code", "desc", None])
    rows.append(["A", "Alpha", None])
    rows.append(["B", "Beta", None])
    rows.append(["C", "Gamma", None])

    sheet = SheetData(name="Test", rows=rows)
    wb = WorkbookData(source_file="test.xlsx", sheets=[sheet])
    wb, warnings = detect_structure(wb)

    assert len(wb.sheets) == 2
    assert any("splitting" in w or "detected" in w for w in warnings)
    # First table should have 3 data rows
    assert len(wb.sheets[0].rows) == 3
    # Second table should have 3 data rows
    assert len(wb.sheets[1].rows) == 3


def test_single_table_empty_left_right_columns_trimmed():
    """A single table with empty left/right columns should have those columns trimmed."""
    sheet = SheetData(
        name="Test",
        rows=[
            [None, None, "id", "name", "val", None],
            [None, None, 1, "a", 10, None],
            [None, None, 2, "b", 20, None],
            [None, None, 3, "c", 30, None],
        ],
    )
    wb, _ = detect_structure(_make_wb(sheet))
    s = wb.sheets[0]
    assert s.headers == ["id", "name", "val"]
    assert len(s.rows) == 3
    assert s.rows[0] == [1, "a", 10]


def test_single_table_empty_top_bottom_rows_trimmed():
    """A single table with empty top/bottom rows should have those rows trimmed."""
    sheet = SheetData(
        name="Test",
        rows=[
            [None, None, None],
            [None, None, None],
            ["id", "name", "val"],
            [1, "a", 10],
            [2, "b", 20],
            [3, "c", 30],
            [None, None, None],
        ],
    )
    wb, _ = detect_structure(_make_wb(sheet))
    s = wb.sheets[0]
    assert s.headers == ["id", "name", "val"]
    assert len(s.rows) == 3
