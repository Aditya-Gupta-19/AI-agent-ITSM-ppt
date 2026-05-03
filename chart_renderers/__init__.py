from .auto_selector import auto_select
from .base_renderer import BaseRenderer
from .c1_rag_scorecard import RAGScorecardRenderer
from .c2_grouped_bar import GroupedBarRenderer
from .c3_pie import PieRenderer
from .c4_stacked_bar_line import StackedBarLineRenderer
from .c5_multi_panel import MultiPanelRenderer
from .c6_bar_line_combo import BarLineComboRenderer
from .c7_line import LineRenderer
from .c8_bar_dotted_line import BarDottedLineRenderer
from .c9_simple_bar import SimpleBarRenderer
from .c10_dashboard_grid import DashboardGridRenderer

RENDERER_MAP: dict = {
    "rag_scorecard": RAGScorecardRenderer,
    "grouped_bar": GroupedBarRenderer,
    "pie": PieRenderer,
    "stacked_bar_line": StackedBarLineRenderer,
    "multi_panel": MultiPanelRenderer,
    "bar_line_combo": BarLineComboRenderer,
    "line": LineRenderer,
    "bar_dotted_line": BarDottedLineRenderer,
    "simple_bar": SimpleBarRenderer,
    "dashboard_grid": DashboardGridRenderer,
}


def get_renderer(chart_type: str) -> BaseRenderer:
    cls = RENDERER_MAP.get(chart_type, SimpleBarRenderer)
    return cls()


__all__ = [
    "auto_select", "get_renderer", "RENDERER_MAP", "BaseRenderer",
    "RAGScorecardRenderer", "GroupedBarRenderer", "PieRenderer",
    "StackedBarLineRenderer", "MultiPanelRenderer", "BarLineComboRenderer",
    "LineRenderer", "BarDottedLineRenderer", "SimpleBarRenderer",
    "DashboardGridRenderer",
]
