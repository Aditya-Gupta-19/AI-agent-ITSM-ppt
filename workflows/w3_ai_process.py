from typing import Any, Dict, List, Optional


def run(
    sheet_name: str,
    sheet_info: Dict[str, Any],
    ai_engine: Any,
    kpi_dict: Dict[str, Any],
    logger: Any,
    excel_summary: str = "",
    summary_mode: str = "ai_write",
    kpi_rows: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """
    Workflow 3: Rule-based analysis per sheet.
    """

    df = sheet_info["dataframe"]
    if df is None or df.empty:
        return {
            "summary": f"The {sheet_name} team has no data records for this period.",
            "kpi_evaluation": [],
            "key_achievements": [],
            "insights": ["No data available for analysis."],
        }

    headers = sheet_info.get("headers") or [str(c) for c in df.columns.tolist()]
    rows = df.to_dict(orient="records")

    logger.info(f"Generating analysis for sheet '{sheet_name}' ({len(rows)} rows).")
    return ai_engine.generate_analysis(
        sheet_name, headers, rows, kpi_dict,
        excel_summary=excel_summary,
        summary_mode=summary_mode,
        kpi_rows=kpi_rows,
    )
