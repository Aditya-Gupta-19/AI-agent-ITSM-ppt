import matplotlib

matplotlib.use("Agg")

import pandas as pd

from tools.t5_chart_generator import ChartGenerator


def test_returns_png_for_numeric_data():
    gen = ChartGenerator()
    df = pd.DataFrame({"Ratio (%)": [0.5, 0.7], "Other": [1, 2]})
    column_types = {"Ratio (%)": "percent", "Other": "numeric"}
    png = gen.generate(df=df, sheet_name="Test", column_types=column_types)
    assert png is not None
    assert isinstance(png, (bytes, bytearray))
    assert len(png) > 1000


def test_generate_for_sheet_auto_selects():
    gen = ChartGenerator()
    df = pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6]})
    col_types = {"A": "numeric", "B": "numeric"}
    results = gen.generate_for_sheet("Test", df, col_types, {})
    assert isinstance(results, list)
    assert len(results) >= 1


def test_generate_for_sheet_none_type():
    gen = ChartGenerator()
    df = pd.DataFrame({"A": [1, 2]})
    col_types = {"A": "numeric"}
    config = {"Test": {"charts": [{"type": "none", "columns": "auto"}]}}
    results = gen.generate_for_sheet("Test", df, col_types, config)
    assert len(results) == 1
    assert results[0]["png_bytes"] is None


def test_generate_original_interface_unchanged():
    gen = ChartGenerator()
    df = pd.DataFrame({"Ratio (%)": [0.5, 0.7], "Other": [1, 2]})
    column_types = {"Ratio (%)": "percent", "Other": "numeric"}
    png = gen.generate(df=df, sheet_name="Test", column_types=column_types)
    assert png is not None
    assert isinstance(png, (bytes, bytearray))


def test_returns_none_when_no_numeric_data():
    gen = ChartGenerator()
    df = pd.DataFrame({"Text": ["a", "b"]})
    column_types = {"Text": "text"}
    png = gen.generate(df=df, sheet_name="Test", column_types=column_types)
    assert png is None

