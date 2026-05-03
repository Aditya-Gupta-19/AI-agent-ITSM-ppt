from typing import Optional

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from .base_renderer import BaseRenderer


class StackedBarLineRenderer(BaseRenderer):
    def _render_impl(self, df: pd.DataFrame, chart_spec: dict,
                     history: Optional[pd.DataFrame]) -> Optional[bytes]:
        fig, ax = plt.subplots(figsize=(9, 4))
        self._draw(ax, df, chart_spec)
        fig.tight_layout()
        return self._fig_to_bytes(fig)

    def _render_in_axis_impl(self, ax, df, chart_spec, history) -> None:
        self._draw(ax, df, chart_spec)

    def _draw(self, ax, df: pd.DataFrame, chart_spec: dict) -> None:
        col_types = chart_spec.get("column_types", {})
        all_cols = self._resolve_columns(df, chart_spec, col_types)

        stack_cols = chart_spec.get("stack_columns") or all_cols[:3]
        line_cols = chart_spec.get("line_columns") or (all_cols[3:5] if len(all_cols) > 3 else [])

        if not stack_cols:
            ax.text(0.5, 0.5, "No data", ha="center", va="center", transform=ax.transAxes)
            return

        x = np.arange(len(df))
        bottom = np.zeros(len(df))
        totals = np.zeros(len(df))

        for i, col in enumerate(stack_cols):
            vals = pd.to_numeric(df[col], errors="coerce").fillna(0).values
            color = self.MUFG_COLORS[i % len(self.MUFG_COLORS)]
            bars = ax.bar(x, vals, bottom=bottom, color=color, label=str(col), width=0.6)
            for bar, v, b in zip(bars, vals, bottom):
                if v > 0:
                    ax.text(bar.get_x() + bar.get_width() / 2, b + v / 2,
                            str(int(v)), ha="center", va="center",
                            fontsize=7, color="white", fontweight="bold")
            bottom += vals
            totals += vals

        for i, total in enumerate(totals):
            ax.text(i, total + 0.5, str(int(total)), ha="center", va="bottom",
                    fontsize=8, fontweight="bold")

        if line_cols:
            ax2 = ax.twinx()
            for i, col in enumerate(line_cols):
                vals = pd.to_numeric(df[col], errors="coerce").tolist()
                color = "#F0C040" if i == 0 else "#333333"
                ax2.plot(x, vals, color=color, marker="o", linewidth=2,
                         markersize=5, label=str(col))
                for xi, v in zip(x, vals):
                    if v == v:
                        ax2.annotate(str(int(v)), (xi, v),
                                     textcoords="offset points", xytext=(0, 6),
                                     ha="center", fontsize=7, color=color)
            # Combine legends
            h1, l1 = ax.get_legend_handles_labels()
            h2, l2 = ax2.get_legend_handles_labels()
            ax.legend(h1 + h2, l1 + l2, loc="upper center",
                      bbox_to_anchor=(0.5, -0.12), ncol=4, fontsize=7, frameon=False)
        else:
            ax.legend(loc="upper right", fontsize=7, frameon=False)

        ax.set_xticks(x)
        ax.set_xticklabels([str(i + 1) for i in range(len(df))], fontsize=8)
        ax.spines["top"].set_visible(False)
        ax.set_title(chart_spec.get("title", ""), fontsize=10, fontweight="bold",
                     color=self.MUFG_COLORS[0])
