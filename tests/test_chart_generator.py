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


def test_returns_none_when_no_numeric_data():
    gen = ChartGenerator()
    df = pd.DataFrame({"Text": ["a", "b"]})
    column_types = {"Text": "text"}
    png = gen.generate(df=df, sheet_name="Test", column_types=column_types)
    assert png is None

