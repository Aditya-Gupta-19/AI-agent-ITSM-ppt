from typing import Optional

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from .base_renderer import BaseRenderer


class LineRenderer(BaseRenderer):
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
        if not cols:
            ax.text(0.5, 0.5, "No data", ha="center", va="center", transform=ax.transAxes)
            return

        text_cols = [c for c in df.columns if col_types.get(str(c), "text") == "text"]
        x_labels = df[text_cols[0]].fillna("").astype(str).tolist() if text_cols else [f"W{i+1}" for i in range(len(df))]

        x = np.arange(len(df))
        all_vals = []

        for i, col in enumerate(cols[:4]):
            vals = pd.to_numeric(df[col], errors="coerce").tolist()
            if col_types.get(str(col)) in ("percent", "percent_decimal"):
                vals = [self._norm_pct(v) for v in vals]
            color = self.MUFG_COLORS[i % len(self.MUFG_COLORS)]
            ax.plot(x, vals, marker="o", linewidth=2.5, markersize=7,
                    color=color, label=str(col), zorder=3)
            is_pct = col_types.get(str(col)) in ("percent", "percent_decimal")
            for xi, v in zip(x, vals):
                if v == v:
                    lbl = f"{v:.0f}%" if is_pct else (f"{int(v)}" if v == int(v) else f"{v:.1f}")
                    ax.annotate(lbl, (xi, v), textcoords="offset points",
                                xytext=(0, 10), ha="center", fontsize=11,
                                fontweight="bold", color=color)
                    all_vals.append(v)

        if all_vals:
            max_v = max(all_vals)
            has_pct = any(col_types.get(str(c)) in ("percent", "percent_decimal") for c in cols)
            ax.set_ylim(0, min(108, max_v * 1.08) if has_pct else max_v * 1.35)

        wrapped = [self._wrap_label(l) for l in x_labels]
        ax.set_xticks(x)
        ax.set_xticklabels(wrapped, fontsize=10, rotation=0, ha="center", linespacing=0.9)
        ax.tick_params(axis="y", labelsize=11)
        ax.yaxis.grid(True, linestyle="--", alpha=0.4, zorder=0)
        ax.set_axisbelow(True)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

        if len(cols) > 1:
            compact = chart_spec.get("figsize_hint", (9, 3.8))[1] <= 2.5
            if compact:
                ax.legend(loc="upper right", fontsize=8, frameon=True,
                          fancybox=False, edgecolor="#CCCCCC", facecolor="white", framealpha=0.9)
            else:
                ax.legend(loc="upper center", bbox_to_anchor=(0.5, 1.18),
                          ncol=min(len(cols), 3), fontsize=11, frameon=False)

        ax.set_title(chart_spec.get("title", ""), fontsize=11, fontweight="bold",
                     color=self.MUFG_COLORS[0], pad=4)
