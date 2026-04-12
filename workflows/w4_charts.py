from typing import Any, Dict, Optional


def run(
    sheet_name: str,
    sheet_info: Dict[str, Any],
    chart_generator: Any,
    logger: Any,
) -> Optional[bytes]:
    """
    Workflow 4: generate chart PNG per sheet (optional).
    """
    df = sheet_info["dataframe"]
    column_types = sheet_info.get("column_types") or {}

    try:
        chart_bytes = chart_generator.generate(df=df, sheet_name=sheet_name, column_types=column_types)
        if chart_bytes:
            logger.info(f"Generated chart for '{sheet_name}'.")
        else:
            logger.info(f"No suitable chart for '{sheet_name}'.")
        return chart_bytes
    except Exception as e:
        logger.warning(f"Chart generation failed for '{sheet_name}': {e}")
        return None

