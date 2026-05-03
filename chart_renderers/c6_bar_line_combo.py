from typing import Optional

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from .base_renderer import BaseRenderer


class BarLineComboRenderer(BaseRenderer):
    def _render_impl(self, df: pd.DataFrame, chart_spec: dict,
                     history: Optional[pd.DataFrame]) -> Optional[bytes]:
        fw, fh = chart_spec.get("figsize_hint", (9, 3.8))
        fig, ax = plt.subplots(figsize=(fw, fh))
        self._draw(ax, df, chart_spec)
        compact = fh <= 2.5
        if compact:
            fig.tight_layout(pad=0.4)
        else:
            fig.subplots_adjust(left=0.08, right=0.97, top=0.80, bottom=0.14)
        return self._fig_to_bytes(fig)

    def _render_in_axis_impl(self, ax, df, chart_spec, history) -> None:
        self._draw(ax, df, chart_spec)

    def _draw(self, ax, df: pd.DataFrame, chart_spec: dict) -> None:
        col_types = chart_spec.get("column_types", {})
        cols = self._resolve_columns(df, chart_spec, col_types)
        if len(cols) < 2:
            ax.text(0.5, 0.5, "Need 2+ columns", ha="center", va="center",
                    transform=ax.transAxes)
            return

        bar_col = cols[0]
        line_col = cols[1]
        x = np.arange(len(df))

        bar_is_pct = col_types.get(str(bar_col)) in ("percent", "percent_decimal")
        line_is_pct = col_types.get(str(line_col)) in ("percent", "percent_decimal")

        bar_vals = pd.to_numeric(df[bar_col], errors="coerce").fillna(0).tolist()
        line_vals = pd.to_numeric(df[line_col], errors="coerce").tolist()
        if bar_is_pct:
            bar_vals = [self._norm_pct(v) for v in bar_vals]
        if line_is_pct:
            line_vals = [self._norm_pct(v) if v == v else float("nan") for v in line_vals]

        bars = ax.bar(x, bar_vals, color="#404040", width=0.6, label=str(bar_col), zorder=3)
        bar_labels = [
            f"{v:.0f}%" if bar_is_pct else (f"{int(v)}" if v == int(v) else f"{v:.1f}")
            for v in bar_vals
        ]
        ax.bar_label(bars, labels=bar_labels, padding=4, fontsize=12,
                     fontweight="bold", color="#404040", zorder=5)

        ax.plot(x, line_vals, color="#CC2222", marker="o", linewidth=2.5,
                markersize=7, label=str(line_col), zorder=4)
        for xi, v in zip(x, line_vals):
            if v == v:
                lbl = f"{v:.0f}%" if line_is_pct else f"{v:.0f}"
                ax.annotate(lbl, (xi, v), textcoords="offset points",
                            xytext=(0, 11), ha="center", fontsize=12,
                            fontweight="bold", color="#CC2222")

        all_vals = [v for v in bar_vals + line_vals if v == v]
        if all_vals:
            max_v = max(all_vals)
            has_pct = bar_is_pct or line_is_pct
            ax.set_ylim(0, min(108, max_v * 1.08) if has_pct else max_v * 1.3)

        text_cols = [c for c in df.columns if col_types.get(str(c), "text") == "text"]
        x_labels = df[text_cols[0]].fillna("").astype(str).tolist() if text_cols else [f"W{i+1}" for i in range(len(df))]
        wrapped = [self._wrap_label(l) for l in x_labels]
        ax.set_xticks(x)
        ax.set_xticklabels(wrapped, fontsize=10, rotation=0, ha="center", linespacing=0.9)
        ax.tick_params(axis="y", labelsize=11)
        ax.yaxis.grid(True, linestyle="--", alpha=0.4, zorder=0)
        ax.set_axisbelow(True)

        compact = chart_spec.get("figsize_hint", (9, 3.8))[1] <= 2.5
        if compact:
            ax.legend(loc="upper right", fontsize=8, frameon=True,
                      fancybox=False, edgecolor="#CCCCCC", facecolor="white", framealpha=0.9)
        else:
            ax.legend(loc="upper center", bbox_to_anchor=(0.5, 1.18),
                      ncol=2, fontsize=11, frameon=False)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.set_title(chart_spec.get("title", ""), fontsize=11, fontweight="bold",
                     color=self.MUFG_COLORS[0], pad=4)
