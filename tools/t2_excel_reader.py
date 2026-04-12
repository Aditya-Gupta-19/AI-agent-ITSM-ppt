import os
import tempfile
from dataclasses import dataclass
from typing import Any, Dict, Optional

import openpyxl
import pandas as pd


@dataclass
class SheetData:
    dataframe: pd.DataFrame
    headers: list[str]
    row_count: int
    is_empty: bool
    column_types: Dict[str, str]


class ExcelReader:
    def __init__(self, config: Optional[dict] = None):
        self.kpi_sheet_name = None
        if config:
            self.kpi_sheet_name = (
                config.get("excel", {}).get("kpi_sheet_name")
                or config.get("excel", {}).get("kpi_sheet")
            )

    def read_all_sheets(self, file_path: str) -> Dict[str, Dict[str, Any]]:
        """
        Returns:
        {
          "sheet_name": {
            "dataframe": pd.DataFrame,
            "headers": list[str],
            "row_count": int,
            "is_empty": bool,
            "column_types": dict  # {col_name: "numeric" | "percent" | "text"}
          }
        }
        """

        wb = openpyxl.load_workbook(file_path, data_only=True)

        # Unmerge merged cells and forward-fill from the top-left cell.
        for ws in wb.worksheets:
            merged_ranges = list(ws.merged_cells.ranges)
            for merged in merged_ranges:
                # merged can be cast to str like "A1:B2"
                min_row = merged.min_row
                min_col = merged.min_col
                value = ws.cell(row=min_row, column=min_col).value
                ws.unmerge_cells(str(merged))
                for r in range(merged.min_row, merged.max_row + 1):
                    for c in range(merged.min_col, merged.max_col + 1):
                        ws.cell(row=r, column=c).value = value

        with tempfile.TemporaryDirectory() as d:
            tmp_path = os.path.join(d, "merged_unpacked.xlsx")
            wb.save(tmp_path)

            sheets: Dict[str, Dict[str, Any]] = {}
            with pd.ExcelFile(tmp_path, engine="openpyxl") as xl:
                for sheet_name in xl.sheet_names:
                    header_mode = None
                    if self.kpi_sheet_name and sheet_name == self.kpi_sheet_name:
                        # KPI sheet is parsed positionally (A/B/C); treat first row as data.
                        header_mode = None  # header=None
                    else:
                        header_mode = 0  # header row is row 1

                    if header_mode is None:
                        df = pd.read_excel(tmp_path, sheet_name=sheet_name, header=None, engine="openpyxl")
                    else:
                        df = pd.read_excel(tmp_path, sheet_name=sheet_name, header=0, engine="openpyxl")

                    # Drop fully empty rows.
                    df = df.dropna(how="all")

                    # Normalize headers as strings for downstream consistency.
                    headers = [str(h) for h in df.columns.tolist()] if len(df.columns) else []

                    # Detect empty sheet.
                    is_empty = df.empty or df.shape[0] == 0
                    row_count = int(df.shape[0])

                    column_types: Dict[str, str] = {}
                    if not is_empty:
                        for col in df.columns:
                            col_name = str(col)
                            column_types[col_name] = self._detect_column_type(df[col])

                    sheets[sheet_name] = {
                        "dataframe": df,
                        "headers": headers,
                        "row_count": row_count,
                        "is_empty": bool(is_empty),
                        "column_types": column_types,
                    }

            return sheets

    def _detect_column_type(self, series: pd.Series) -> str:
        # If column name will include "%", orchestrator should label it by header;
        # here we infer from values only.
        non_null = series.dropna()
        if non_null.empty:
            return "text"

        numeric = pd.to_numeric(non_null, errors="coerce").dropna()
        if numeric.empty:
            return "text"

        # Percent if values look like ratios (0..1).
        try:
            min_v = float(numeric.min())
            max_v = float(numeric.max())
        except Exception:
            return "text"

        if 0.0 <= min_v <= 1.0 and 0.0 <= max_v <= 1.0:
            return "percent"

        return "numeric"

