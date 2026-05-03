from math import cos, sin, radians
from typing import Optional

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from .base_renderer import BaseRenderer

_PIE_COLORS = ["#F0C040", "#2E75B6", "#404040", "#C00000", "#0F6E56", "#C55A11", "#1F3864"]


class PieRenderer(BaseRenderer):
    def _render_impl(self, df: pd.DataFrame, chart_spec: dict,
                     history: Optional[pd.DataFrame]) -> Optional[bytes]:
        fig, ax = plt.subplots(figsize=(6, 4))
        self._draw(ax, df, chart_spec)
        fig.tight_layout()
        return self._fig_to_bytes(fig)

    def _render_in_axis_impl(self, ax, df, chart_spec, history) -> None:
        self._draw(ax, df, chart_spec)

    def _draw(self, ax, df: pd.DataFrame, chart_spec: dict) -> None:
        col_types = chart_spec.get("column_types", {})
        spec_cols = chart_spec.get("columns", "auto")

        # Find a text column with few unique values for categorical pie.
        if spec_cols == "auto":
            text_cols = [c for c in df.columns if col_types.get(str(c), "text") == "text"]
            if text_cols:
                vc = df[text_cols[0]].value_counts()
                if len(vc) <= 8:
                    labels = vc.index.tolist()
                    counts = vc.values.tolist()
                else:
                    labels, counts = self._sum_numeric_cols(df, col_types)
            else:
                labels, counts = self._sum_numeric_cols(df, col_types)
        else:
            cols = [c.strip() for c in str(spec_cols).split(",") if c.strip() in df.columns]
            labels = cols
            counts = [pd.to_numeric(df[c], errors="coerce").sum() for c in cols]

        if not labels or sum(counts) == 0:
            ax.text(0.5, 0.5, "No data", ha="center", va="center", transform=ax.transAxes)
            return

        total = sum(counts)
        colors = [_PIE_COLORS[i % len(_PIE_COLORS)] for i in range(len(labels))]
        wedges, _, _ = ax.pie(counts, labels=None, autopct="%1.0f%%",
                               startangle=90, colors=colors,
                               pctdistance=0.75, wedgeprops={"linewidth": 0.5, "edgecolor": "white"})
        # External labels
        for wedge, label, count in zip(wedges, labels, counts):
            angle = (wedge.theta2 + wedge.theta1) / 2
            x = 1.3 * cos(radians(angle))
            y = 1.3 * sin(radians(angle))
            pct = count / total * 100
            ax.text(x, y, f"{label}\n{int(count)}, {pct:.0f}%",
                    ha="center", va="center", fontsize=8)
        ax.set_title(chart_spec.get("title", ""), fontsize=10, fontweight="bold",
                     color=self.MUFG_COLORS[0])

    @staticmethod
    def _sum_numeric_cols(df, col_types):
        num_cols = [c for c in df.columns if col_types.get(str(c), "text") in
                    ("numeric", "percent", "percent_decimal")]
        labels = num_cols
        counts = [pd.to_numeric(df[c], errors="coerce").sum() for c in num_cols]
        return labels, counts
