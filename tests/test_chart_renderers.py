import matplotlib
matplotlib.use("Agg")

import pandas as pd
import pytest

from chart_renderers import (
    RENDERER_MAP, auto_select, get_renderer,
    SimpleBarRenderer, GroupedBarRenderer,
)
from chart_renderers.base_renderer import BaseRenderer


class _FailRenderer(BaseRenderer):
    def _render_impl(self, df, chart_spec, history):
        raise RuntimeError("Intentional failure")


def test_base_renderer_never_raises():
    r = _FailRenderer()
    result = r.render(pd.DataFrame({"A": [1, 2]}), {})
    assert result is None


def test_auto_select_rag_scorecard():
    df = pd.DataFrame({"A": [0.9], "B": [0.8], "C": [0.7]})
    col_types = {"A": "percent", "B": "percent", "C": "percent"}
    assert auto_select(df, col_types) == "rag_scorecard"


def test_auto_select_grouped_bar():
    df = pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6]})
    col_types = {"A": "numeric", "B": "numeric"}
    assert auto_select(df, col_types) == "grouped_bar"


def test_auto_select_none_empty():
    assert auto_select(pd.DataFrame(), {}) == "none"


def test_simple_bar_returns_png():
    df = pd.DataFrame({"Score": [10, 20, 30]})
    col_types = {"Score": "numeric"}
    renderer = SimpleBarRenderer()
    png = renderer.render(df, {"column_types": col_types})
    assert png is not None
    assert isinstance(png, bytes)
    assert len(png) > 500


def test_grouped_bar_returns_png():
    df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
    col_types = {"A": "numeric", "B": "numeric"}
    renderer = GroupedBarRenderer()
    png = renderer.render(df, {"column_types": col_types})
    assert png is not None
    assert isinstance(png, bytes)


def test_renderer_map_complete():
    expected = {
        "rag_scorecard", "grouped_bar", "pie", "stacked_bar_line",
        "multi_panel", "bar_line_combo", "line", "bar_dotted_line",
        "simple_bar", "dashboard_grid",
    }
    assert expected == set(RENDERER_MAP.keys())


def test_get_renderer_unknown_falls_back_to_simple_bar():
    r = get_renderer("nonexistent_type")
    assert isinstance(r, SimpleBarRenderer)
