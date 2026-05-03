from typing import Optional

import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import pandas as pd

from .base_renderer import BaseRenderer


class DashboardGridRenderer(BaseRenderer):
    def _render_impl(self, df: pd.DataFrame, chart_spec: dict,
                     history: Optional[pd.DataFrame]) -> Optional[bytes]:
        sub_specs = chart_spec.get("charts", [])
        if not sub_specs:
            return None

        rows = chart_spec.get("rows", 2)
        cols = chart_spec.get("cols", 3)

        fig = plt.figure(figsize=(14, 8))
        gs = gridspec.GridSpec(rows, cols, figure=fig, hspace=0.45, wspace=0.35)

        # Import here to avoid circular imports
        from chart_renderers import get_renderer

        for i, sub_spec in enumerate(sub_specs[:rows * cols]):
            row_idx = i // cols
            col_idx = i % cols
            ax = fig.add_subplot(gs[row_idx, col_idx])
            chart_type = sub_spec.get("type", "simple_bar")
            renderer = get_renderer(chart_type)
            renderer.render_in_axis(ax, df, {**sub_spec, "column_types": chart_spec.get("column_types", {})})

        return self._fig_to_bytes(fig)

    def _render_in_axis_impl(self, ax, df, chart_spec, history) -> None:
        ax.text(0.5, 0.5, "Dashboard Grid", ha="center", va="center",
                transform=ax.transAxes, fontsize=9)
