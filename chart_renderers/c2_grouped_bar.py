from typing import Optional

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from .base_renderer import BaseRenderer


class GroupedBarRenderer(BaseRenderer):
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

    def _pick_series_columns(self, df: pd.DataFrame, chart_spec: dict, col_types: dict) -> list:
        spec_cols = chart_spec.get("columns", "auto")
        if spec_cols and spec_cols != "auto":
            return [c.strip() for c in str(spec_cols).split(",") if c.strip() in df.columns]
        pct_cols = [c for c in df.columns
                    if col_types.get(str(c)) in ("percent", "percent_decimal")]
        if pct_cols:
            return pct_cols[:4]
        return [c for c in df.columns if col_types.get(str(c)) == "numeric"][:4]

    def _get_x_labels(self, df: pd.DataFrame, col_types: dict) -> list:
        text_cols = [c for c in df.columns if col_types.get(str(c), "text") == "text"]
        if text_cols:
            return df[text_cols[0]].fillna("").astype(str).tolist()
        return [f"W{i + 1}" for i in range(len(df))]

    def _draw(self, ax, df: pd.DataFrame, chart_spec: dict) -> None:
        col_types = chart_spec.get("column_types", {})
        cols = self._pick_series_columns(df, chart_spec, col_types)
        if not cols:
            ax.text(0.5, 0.5, "No data", ha="center", va="center", transform=ax.transAxes)
            return

        n_series = len(cols)
        n_groups = len(df)
        x = np.arange(n_groups)
        bar_width = 0.7 / max(n_series, 1)

        all_vals = []
        for i, col in enumerate(cols):
            vals = pd.to_numeric(df[col], errors="coerce").fillna(0).tolist()
            if col_types.get(str(col)) in ("percent", "percent_decimal"):
                vals = [self._norm_pct(v) for v in vals]
            all_vals.extend([v for v in vals if not np.isnan(v)])

            offset = (i - (n_series - 1) / 2) * bar_width
            color = self.MUFG_COLORS[i % len(self.MUFG_COLORS)]
            bars = ax.bar(x + offset, vals, width=bar_width * 0.9,
                          color=color, label=str(col), zorder=3)

            is_pct = col_types.get(str(col)) in ("percent", "percent_decimal")
            labels_fmt = [
                f"{v:.0f}%" if is_pct else (f"{int(v)}" if v == int(v) else f"{v:.1f}")
                for v in vals
            ]
            ax.bar_label(bars, labels=labels_fmt, padding=3,
                         fontsize=9, fontweight="bold", color="#333333", zorder=5)

        if all_vals:
            max_v = max(all_vals)
            has_pct = any(col_types.get(str(c)) in ("percent", "percent_decimal") for c in cols)
            ax.set_ylim(0, min(108, max_v * 1.08) if has_pct else max_v * 1.28)

        x_labels = self._get_x_labels(df, col_types)
        wrapped = [self._wrap_label(l) for l in x_labels]
        ax.set_xticks(x)
        ax.set_xticklabels(wrapped, fontsize=10, rotation=0, ha="center", linespacing=0.9)
        ax.tick_params(axis="y", labelsize=9)

        compact = chart_spec.get("figsize_hint", (9, 3.8))[1] <= 2.5
        if compact:
            ax.legend(loc="upper right", fontsize=8, frameon=True,
                      fancybox=False, edgecolor="#CCCCCC", facecolor="white", framealpha=0.9)
        else:
            ax.legend(loc="upper center", bbox_to_anchor=(0.5, 1.18),
                      ncol=min(n_series, 4), fontsize=10, frameon=False)

        ax.yaxis.grid(True, linestyle="--", alpha=0.4, zorder=0)
        ax.set_axisbelow(True)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.set_title(chart_spec.get("title", ""), fontsize=11, fontweight="bold",
                     color=self.MUFG_COLORS[0], pad=4)
