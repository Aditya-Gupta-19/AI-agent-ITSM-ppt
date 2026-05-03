from typing import Any, Dict, List, Optional, Tuple


def run(
    sheet_name: str,
    sheet_info: Dict[str, Any],
    chart_generator: Any,
    logger: Any,
    report_config: Optional[dict] = None,
) -> Tuple[Optional[bytes], List[dict]]:
    """
    Workflow 4: generate chart PNG(s) per sheet.
    Returns (first_chart_bytes, charts_list).
    """
    df = sheet_info["dataframe"]
    column_types = sheet_info.get("column_types") or {}

    try:
        if report_config is not None and hasattr(chart_generator, "generate_for_sheet"):
            charts_list = chart_generator.generate_for_sheet(
                sheet_name, df, column_types, report_config
            )
            first_bytes = next(
                (c["png_bytes"] for c in charts_list if c.get("png_bytes")), None
            )
            if charts_list:
                logger.info(f"Generated {len(charts_list)} chart(s) for '{sheet_name}'.")
            else:
                logger.info(f"No suitable chart for '{sheet_name}'.")
            return first_bytes, charts_list

        chart_bytes = chart_generator.generate(df=df, sheet_name=sheet_name, column_types=column_types)
        if chart_bytes:
            logger.info(f"Generated chart for '{sheet_name}'.")
        else:
            logger.info(f"No suitable chart for '{sheet_name}'.")
        return chart_bytes, []
    except Exception as e:
        logger.warning(f"Chart generation failed for '{sheet_name}': {e}")
        return None, []
