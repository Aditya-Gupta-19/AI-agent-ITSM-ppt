from typing import Optional

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from .base_renderer import BaseRenderer


class SimpleBarRenderer(BaseRenderer):
    def _render_impl(self, df: pd.DataFrame, chart_spec: dict,
                     history: Optional[pd.DataFrame]) -> Optional[bytes]:
        fw, fh = chart_spec.get("figsize_hint", (9, 3.8))
        fig, ax = plt.subplots(figsize=(fw, fh))
        self._draw(ax, df, chart_spec)
        compact = fh <= 2.5
        if compact:
            fig.tight_layout(pad=0.4)
        else:
            fig.subplots_adjust(left=0.08, right=0.97, top=0.92, bottom=0.14)
        return self._fig_to_bytes(fig)

    def _render_in_axis_impl(self, ax, df, chart_spec, history) -> None:
        self._draw(ax, df, chart_spec)

    def _draw(self, ax, df: pd.DataFrame, chart_spec: dict) -> None:
        col_types = chart_spec.get("column_types", {})
        cols = self._resolve_columns(df, chart_spec, col_types)
        if not cols:
            ax.text(0.5, 0.5, "No data", ha="center", va="center", transform=ax.transAxes)
            return

        col = cols[0]
        values = pd.to_numeric(df[col], errors="coerce").fillna(0).tolist()
        if col_types.get(str(col)) in ("percent", "percent_decimal"):
            values = [self._norm_pct(v) for v in values]

        x = np.arange(len(values))
        bars = ax.bar(x, values, color=self.MUFG_COLORS[0], width=0.6, zorder=3)
        is_pct = col_types.get(str(col)) in ("percent", "percent_decimal")
        labels_fmt = [f"{v:.0f}%" if is_pct else f"{v:.1f}".rstrip("0").rstrip(".") for v in values]
        ax.bar_label(bars, labels=labels_fmt, padding=4, fontsize=12,
                     fontweight="bold", color="#333333", zorder=5)
        if values:
            max_v = max(v for v in values if v == v)
            if is_pct:
                ax.set_ylim(0, min(108, max_v * 1.08))
            else:
                ax.set_ylim(0, max_v * 1.28)
        ax.yaxis.grid(True, linestyle="--", alpha=0.4, zorder=0)
        ax.set_axisbelow(True)

        text_cols = [c for c in df.columns if col_types.get(str(c), "text") == "text"]
        x_labels = df[text_cols[0]].fillna("").astype(str).tolist() if text_cols else [f"W{i+1}" for i in range(len(values))]

        wrapped = [self._wrap_label(l) for l in x_labels]
        ax.set_xticks(x)
        ax.set_xticklabels(wrapped, fontsize=10, rotation=0, ha="center", linespacing=0.9)
        ax.tick_params(axis="y", labelsize=11)
        ax.set_title(chart_spec.get("title", str(col)), fontsize=11, fontweight="bold",
                     color=self.MUFG_COLORS[0], pad=4)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
