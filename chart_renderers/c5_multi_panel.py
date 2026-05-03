from typing import Optional

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from .base_renderer import BaseRenderer


class MultiPanelRenderer(BaseRenderer):
    def _render_impl(self, df: pd.DataFrame, chart_spec: dict,
                     history: Optional[pd.DataFrame]) -> Optional[bytes]:
        col_types = chart_spec.get("column_types", {})
        panel_col = chart_spec.get("panel_column")

        # Default: use first text column as panel grouper
        if not panel_col:
            text_cols = [c for c in df.columns if col_types.get(str(c), "text") == "text"]
            panel_col = text_cols[0] if text_cols else None

        if not panel_col or panel_col not in df.columns:
            # Fallback: split rows into equal panels
            return self._render_row_panels(df, chart_spec)

        panels = df[panel_col].dropna().unique()[:5]
        n = len(panels)
        if n == 0:
            return None

        value_cols = self._resolve_columns(df, chart_spec, col_types)

        fig, axes = plt.subplots(1, n, figsize=(5 * n, 3.5), sharey=False)
        if n == 1:
            axes = [axes]

        for ax, panel_val in zip(axes, panels):
            subset = df[df[panel_col] == panel_val]
            x = np.arange(len(subset))
            for i, col in enumerate(value_cols[:3]):
                vals = pd.to_numeric(subset[col], errors="coerce").fillna(0).tolist()
                if col_types.get(str(col)) in ("percent", "percent_decimal"):
                    vals = [self._norm_pct(v) for v in vals]
                color = self.MUFG_COLORS[i % len(self.MUFG_COLORS)]
                bars = ax.bar(x + i * 0.25, vals, width=0.25, color=color, label=str(col))
                ax.bar_label(bars, fmt="%.0f", padding=1, fontsize=7)
            ax.set_title(str(panel_val), fontsize=9, fontweight="bold")
            ax.set_xticks(x)
            ax.set_xticklabels([str(j + 1) for j in range(len(subset))], fontsize=7)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)

        plt.suptitle(chart_spec.get("title", ""), fontsize=11, fontweight="bold", y=1.02)
        plt.tight_layout()
        return self._fig_to_bytes(fig)

    def _render_row_panels(self, df: pd.DataFrame, chart_spec: dict) -> Optional[bytes]:
        """Fallback: one panel per row, max 3."""
        n = min(len(df), 3)
        if n == 0:
            return None
        col_types = chart_spec.get("column_types", {})
        value_cols = self._resolve_columns(df, chart_spec, col_types)
        fig, axes = plt.subplots(1, n, figsize=(5 * n, 3.5))
        if n == 1:
            axes = [axes]
        for i, ax in enumerate(axes):
            row = df.iloc[i]
            vals = [pd.to_numeric(row.get(c), errors="coerce") for c in value_cols]
            ax.bar(range(len(vals)), [v if v == v else 0 for v in vals],
                   color=self.MUFG_COLORS[i % len(self.MUFG_COLORS)])
            ax.set_title(f"Row {i + 1}", fontsize=9)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
        plt.tight_layout()
        return self._fig_to_bytes(fig)

    def _render_in_axis_impl(self, ax, df, chart_spec, history) -> None:
        ax.text(0.5, 0.5, "Multi-Panel", ha="center", va="center",
                transform=ax.transAxes, fontsize=9)
