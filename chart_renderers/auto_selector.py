import pandas as pd


def auto_select(df: pd.DataFrame, column_types: dict) -> str:
    """Returns chart type string based on data shape and column types."""
    if df is None or df.empty:
        return "none"

    numeric_cols = [c for c, t in column_types.items()
                    if t in ("numeric", "percent", "percent_decimal")]
    percent_cols = [c for c, t in column_types.items()
                    if t in ("percent", "percent_decimal")]
    row_count = len(df.dropna(how="all"))

    if row_count == 0:
        return "none"
    if row_count == 1 and len(percent_cols) >= 3:
        return "rag_scorecard"
    if row_count == 1 and len(numeric_cols) >= 2:
        return "grouped_bar"
    if row_count >= 2 and len(numeric_cols) >= 2:
        return "grouped_bar"
    if row_count >= 2 and len(numeric_cols) == 1:
        return "simple_bar"
    if len(percent_cols) == 2:
        return "pie"
    return "simple_bar"
