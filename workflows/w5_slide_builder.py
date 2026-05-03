from typing import Any, Dict, List


def run(
    pptx_builder: Any,
    report_date: str,
    groups_queue: List[Dict[str, Any]],
    total_group_count: int,
    output_path: str,
) -> str:
    """
    Workflow 5: assemble final PPTX (one slide per team/group).
    """
    pptx_builder.add_title_slide(report_date=report_date, group_count=total_group_count)

    for item in groups_queue:
        pptx_builder.add_group_slide(
            sheet_name=item["sheet_name"],
            ai_output=item["ai_output"],
            kpi_rows=item.get("kpi_rows") or [],
            chart_bytes=item.get("chart_bytes"),
            charts=item.get("charts_list"),
            summary_mode=item.get("summary_mode", "ai_write"),
            df=item.get("df"),
            kpi_definitions=item.get("kpi_definitions", []),
            column_types=item.get("column_types", {}),
            layout=item.get("layout", "standard"),
            multigroup_data=item.get("multigroup_data"),
            rag_thresholds=item.get("rag_thresholds"),
            parsed_sections=item.get("parsed_sections"),
        )

    return pptx_builder.save(output_path)
