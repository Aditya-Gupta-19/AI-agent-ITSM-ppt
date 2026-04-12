from typing import Any, Dict, Tuple


def run(
    sheets_data: Dict[str, Dict[str, Any]],
    kpi_sheet_name: str,
    kpi_parser: Any,
    logger: Any,
) -> Tuple[Dict[str, Any], Dict[str, Dict[str, Any]]]:
    """
    Workflow 2: parse all sheets + KPI dict.

    Returns:
      (kpi_dict, sheets_data)
    """

    if kpi_sheet_name not in sheets_data:
        logger.warning(f"KPI sheet '{kpi_sheet_name}' not found. KPI reference will be empty.")
        return {}, sheets_data

    cover_df = sheets_data[kpi_sheet_name]["dataframe"]
    kpi_dict = kpi_parser.parse_cover_page(cover_df)
    logger.info(f"Parsed KPI reference for {len(kpi_dict)} groups from '{kpi_sheet_name}'.")

    return kpi_dict, sheets_data

