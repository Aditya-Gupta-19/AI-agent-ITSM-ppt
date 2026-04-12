from typing import Any, Dict


def run(excel_reader: Any, excel_path: str, logger: Any) -> Dict[str, Dict[str, Any]]:
    """
    Workflow 1: load Excel file into sheet -> parsed dict.
    """
    logger.info(f"Ingesting Excel file: {excel_path}")
    return excel_reader.read_all_sheets(excel_path)

