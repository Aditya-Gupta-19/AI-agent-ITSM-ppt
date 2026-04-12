import os
import tempfile

import openpyxl
import pandas as pd
import pytest

from tools.t2_excel_reader import ExcelReader


def _write_workbook(wb, file_path: str) -> None:
    wb.save(file_path)


def _build_basic_workbook(sheet_names=("A", "B")) -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    ws0 = wb.active
    ws0.title = sheet_names[0]
    ws0.append(["Col1", "Ratio (%)"])
    ws0.append([1, 0.5])

    for sn in sheet_names[1:]:
        ws = wb.create_sheet(sn)
        ws.append(["Col1", "Ratio (%)"])
        ws.append([2, 0.7])

    return wb


def test_reads_all_sheets():
    wb = _build_basic_workbook(("SheetA", "SheetB", "SheetC"))
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "test.xlsx")
        _write_workbook(wb, path)
        reader = ExcelReader()
        sheets = reader.read_all_sheets(path)
        assert set(sheets.keys()) == {"SheetA", "SheetB", "SheetC"}


def test_detects_empty_sheet():
    wb = _build_basic_workbook(("HasData",))
    wb.create_sheet("Empty")
    ws_empty = wb["Empty"]
    # Leave it truly empty (no rows).
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "test.xlsx")
        _write_workbook(wb, path)
        reader = ExcelReader()
        sheets = reader.read_all_sheets(path)
        assert sheets["Empty"]["is_empty"] is True


def test_column_type_detection():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "T"
    ws.append(["Ratio (%)", "Other"])
    ws.append([0.5, "x"])
    ws.append([0.7, "y"])

    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "test.xlsx")
        _write_workbook(wb, path)
        reader = ExcelReader()
        sheets = reader.read_all_sheets(path)
        column_types = sheets["T"]["column_types"]
        assert column_types["Ratio (%)"] == "percent"


def test_handles_merged_cells():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Merged"
    ws.append(["A", "B"])
    ws.append([None, None])

    ws["A2"].value = "MergedVal"
    ws.merge_cells("A2:B2")

    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "test.xlsx")
        _write_workbook(wb, path)
        reader = ExcelReader()
        sheets = reader.read_all_sheets(path)
        df = sheets["Merged"]["dataframe"]
        assert df.iloc[0]["A"] == "MergedVal"
        assert df.iloc[0]["B"] == "MergedVal"

