import re
from typing import Any, Dict, List

import pandas as pd


class KPIParser:
    def __init__(self):
        # Matches things like: "Delivery Percent (>70%)" or "Failed Change Ratio (<1%)"
        self._threshold_regex = re.compile(r"\(([<>≥≤]=?)\s*(\d+\.?\d*)\s*(%?)\)")

    def parse_cover_page(self, df: pd.DataFrame) -> Dict[str, Dict[str, str]]:
        """
        Reads Cover Page sheet. Returns KPI reference dict.
        Structure: {group_name: {kpi_name: threshold_description}}
        Look for group name headers in column A or B.
        KPI names in column B, thresholds/descriptions in column C.
        """

        kpi_reference: Dict[str, Dict[str, str]] = {}
        current_group: str = "Unknown"

        if df is None or df.empty:
            return kpi_reference

        # Be robust to header rows: parse positionally (first three columns).
        if df.shape[1] < 3:
            return kpi_reference

        colA = df.iloc[:, 0]
        colB = df.iloc[:, 1]
        colC = df.iloc[:, 2]

        for i in range(len(df)):
            a = colA.iloc[i]
            b = colB.iloc[i]
            c = colC.iloc[i]

            a_val = None if pd.isna(a) else str(a).strip()
            b_val = None if pd.isna(b) else str(b).strip()
            c_val = None if pd.isna(c) else str(c).strip()

            if a_val:
                current_group = a_val

            # If column C has a value, treat it as threshold/description tied to KPI name in B.
            if b_val and c_val:
                if current_group not in kpi_reference:
                    kpi_reference[current_group] = {}
                kpi_reference[current_group][b_val] = c_val

        return kpi_reference

    def parse_header_thresholds(self, headers: List[str]) -> Dict[str, Dict[str, Any]]:
        """
        Extracts thresholds embedded in column header names.
        Returns:
          {col_name: {operator, value, unit, original_header}}
        """

        thresholds: Dict[str, Dict[str, Any]] = {}
        for h in headers:
            if h is None:
                continue
            header = str(h)
            match = self._threshold_regex.search(header)
            if not match:
                continue

            operator = match.group(1)
            value = float(match.group(2))
            unit = match.group(3) if match.group(3) else ""
            thresholds[header] = {
                "operator": operator,
                "value": value,
                "unit": unit,
                "original_header": header,
            }

        return thresholds

    def evaluate_kpis(self, row_data: Dict[str, Any], thresholds: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Compares actual values against thresholds.
        Returns list of:
          {kpi_name, actual_value, threshold, status: "PASS"|"FAIL"|"UNKNOWN"}
        """

        results: List[Dict[str, Any]] = []

        def operator_eval(actual: float, operator: str, target: float) -> bool:
            op = operator
            # Normalize Unicode operators.
            if op == "≥":
                op = ">="
            if op == "≤":
                op = "<="

            if op == ">":
                return actual > target
            if op == "<":
                return actual < target
            if op == ">=":
                return actual >= target
            if op == "<=":
                return actual <= target
            if op == "==":
                return actual == target
            return False

        for header, t in thresholds.items():
            operator = t["operator"]
            target = t["value"]
            unit = t["unit"]

            raw_actual = row_data.get(header)
            if raw_actual is None or (isinstance(raw_actual, float) and pd.isna(raw_actual)):
                results.append(
                    {
                        "kpi_name": header,
                        "actual_value": None,
                        "threshold": f"{operator}{target}{unit}",
                        "status": "UNKNOWN",
                    }
                )
                continue

            # Convert actual to numeric.
            if isinstance(raw_actual, str):
                # Handle values like "47.83%" or "0.4783"
                raw = raw_actual.strip().replace("%", "")
                try:
                    actual_num = float(raw)
                except Exception:
                    results.append(
                        {
                            "kpi_name": header,
                            "actual_value": raw_actual,
                            "threshold": f"{operator}{target}{unit}",
                            "status": "UNKNOWN",
                        }
                    )
                    continue
            else:
                try:
                    actual_num = float(raw_actual)
                except Exception:
                    results.append(
                        {
                            "kpi_name": header,
                            "actual_value": raw_actual,
                            "threshold": f"{operator}{target}{unit}",
                            "status": "UNKNOWN",
                        }
                    )
                    continue

            # Normalize percent scale to 0..100 when threshold has '%'.
            actual_cmp = actual_num
            target_cmp = float(target)
            if unit == "%":
                # Support both ratio (0..1) and percent (0..100) scales.
                target_is_ratio = target_cmp <= 1.0
                actual_is_ratio = actual_num <= 1.0
                if target_is_ratio and not actual_is_ratio:
                    # Actual is percent, convert to ratio.
                    actual_cmp = actual_num / 100.0
                elif (not target_is_ratio) and actual_is_ratio:
                    # Target is percent, convert actual ratio to percent.
                    actual_cmp = actual_num * 100.0
                else:
                    actual_cmp = actual_num

            status = "PASS" if operator_eval(actual_cmp, operator, target_cmp) else "FAIL"

            # Produce a cleaner KPI name (remove threshold parentheses when possible).
            kpi_name = header
            if "(" in header:
                kpi_name = header.split("(")[0].strip()

            results.append(
                {
                    "kpi_name": kpi_name,
                    "actual_value": actual_cmp if unit == "%" else actual_num,
                    "threshold": f"{operator}{target}{unit}",
                    "status": status,
                }
            )

        return results

