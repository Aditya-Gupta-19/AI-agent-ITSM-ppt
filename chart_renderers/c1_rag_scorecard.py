from typing import Optional

import matplotlib.patches as mpatches
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from .base_renderer import BaseRenderer

_GREEN = "#00B050"
_AMBER = "#FFC000"
_RED = "#FF0000"


def _rag_color(actual_pct: float, green_threshold: float, amber_threshold: float) -> str:
    if actual_pct >= green_threshold:
        return _GREEN
    if actual_pct >= amber_threshold:
        return _AMBER
    return _RED


class RAGScorecardRenderer(BaseRenderer):
    def _render_impl(self, df: pd.DataFrame, chart_spec: dict,
                     history: Optional[pd.DataFrame]) -> Optional[bytes]:
        col_types = chart_spec.get("column_types", {})
        pct_cols = [c for c in df.columns if col_types.get(str(c), "text")
                    in ("percent", "percent_decimal")]
        if not pct_cols:
            pct_cols = list(df.columns)

        green_t = chart_spec.get("green_threshold", 95.0)
        amber_t = chart_spec.get("amber_threshold", 70.0)

        rows_data = []
        rag_colors = []
        for col in pct_cols:
            val = pd.to_numeric(df[col].dropna().iloc[-1] if not df[col].dropna().empty else None,
                                errors="coerce")
            if val is None or (val != val):
                continue
            pct = self._norm_pct(val)
            color = _rag_color(pct, green_t, amber_t)
            rows_data.append((str(col).split("(")[0].strip(), f"{pct:.1f}%", color))
            rag_colors.append(color)

        if not rows_data:
            return None

        fig_h = max(2.0, len(rows_data) * 0.55 + 0.8)
        fig, ax = plt.subplots(figsize=(9, fig_h))
        ax.axis("off")

        # Build table data
        table_data = [[r[0], r[1], ""] for r in rows_data]
        col_labels = ["KPI", "Actual", "RAG"]
        col_widths = [0.6, 0.2, 0.2]

        tbl = ax.table(
            cellText=table_data,
            colLabels=col_labels,
            colWidths=col_widths,
            loc="center",
            cellLoc="left",
        )
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(9)
        tbl.scale(1, 1.4)

        # Style header row
        for col_idx in range(3):
            cell = tbl[0, col_idx]
            cell.set_facecolor("#1F3864")
            cell.get_text().set_color("white")
            cell.get_text().set_fontweight("bold")

        # Style data rows + draw RAG dots
        for row_idx, (_, _, color) in enumerate(rows_data, start=1):
            for col_idx in range(3):
                tbl[row_idx, col_idx].set_facecolor("#F2F2F2" if row_idx % 2 else "white")

            # Get cell position to draw dot
            cell = tbl[row_idx, 2]
            bbox = cell.get_bbox()
            # We'll draw the circle after rendering — use annotation instead
            cell.get_text().set_text("")  # clear text; circle drawn below

        fig.canvas.draw()

        # Draw RAG circles over the RAG column
        for row_idx, (_, _, color) in enumerate(rows_data, start=1):
            cell = tbl[row_idx, 2]
            x = cell.get_xy()[0] + cell.get_width() / 2
            y = cell.get_xy()[1] + cell.get_height() / 2
            ax.add_patch(mpatches.Circle(
                (x, y), radius=min(cell.get_height(), cell.get_width()) * 0.35,
                color=color, transform=ax.transData, zorder=5,
            ))

        # Overall RAG dot (worst)
        if rag_colors:
            overall = _RED if _RED in rag_colors else (_AMBER if _AMBER in rag_colors else _GREEN)
            ax.text(1.05, 0.5, "Overall\nRAG", transform=ax.transAxes,
                    ha="center", va="center", fontsize=9, fontweight="bold")
            ax.add_patch(mpatches.Circle(
                (1.12, 0.5), radius=0.05, color=overall,
                transform=ax.transAxes, zorder=5,
            ))

        fig.tight_layout()
        return self._fig_to_bytes(fig)

    def _render_in_axis_impl(self, ax, df, chart_spec, history) -> None:
        ax.text(0.5, 0.5, "RAG Scorecard", ha="center", va="center",
                transform=ax.transAxes, fontsize=9)
