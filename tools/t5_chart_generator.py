import io
from typing import Dict, List, Optional

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from chart_renderers import auto_select, get_renderer


class ChartGenerator:
    @staticmethod
    def _resolve_config(sheet_name: str, report_config: dict) -> dict:
        """Exact-match first; fallback to longest prefix match (≥10 chars, case-insensitive)."""
        if not sheet_name or not report_config:
            return {}
        exact = report_config.get(sheet_name)
        if exact is not None:
            return exact
        sn_lower = sheet_name.lower().strip()
        best_key, best_len = None, 0
        for key in report_config:
            kl = key.lower().strip()
            min_len = min(len(sn_lower), len(kl))
            if min_len < 10:
                continue
            if sn_lower.startswith(kl[:min_len]) or kl.startswith(sn_lower[:min_len]):
                if min_len > best_len:
                    best_key, best_len = key, min_len
        return report_config.get(best_key, {}) if best_key else {}

    def generate(
        self,
        df: pd.DataFrame,
        sheet_name: str,
        column_types: Dict[str, str],
    ) -> Optional[bytes]:
        """
        Auto-selects chart type and returns PNG as bytes (BytesIO).
        Returns None if no suitable numeric data.
        """

        if df is None or df.empty:
            return None

        numeric_or_percent_cols = [
            col for col in df.columns if column_types.get(str(col)) in ("numeric", "percent")
        ]
        if not numeric_or_percent_cols:
            return None

        numeric_cols = [col for col in df.columns if column_types.get(str(col)) == "numeric"]
        percent_cols = [col for col in df.columns if column_types.get(str(col)) == "percent"]

        # Use only rows that have at least one numeric/percent value.
        plot_df = df.copy()
        plot_df = plot_df[numeric_or_percent_cols] if numeric_or_percent_cols else plot_df

        # Drop rows fully empty for plotting columns.
        plot_df = plot_df.dropna(how="all")
        if plot_df.empty:
            return None

        n_rows = plot_df.shape[0]

        PRIMARY = "#2E75B6"
        SECONDARY = "#1F3864"
        ACCENT = "#0F6E56"

        def normalize_percent(v):
            try:
                fv = float(v)
            except Exception:
                return np.nan
            if np.isnan(fv):
                return np.nan
            if fv <= 1.0:
                return fv * 100.0
            return fv

        # Style baseline.
        plt.ioff()
        fig, ax = plt.subplots(figsize=(7, 3.5))

        title = f"{sheet_name} — Performance Overview"
        ax.set_title(title, fontsize=12, fontweight="bold", color=SECONDARY)

        if n_rows == 1:
            row = plot_df.iloc[0]

            # Multiple percent columns -> horizontal bar, target 100%.
            # (Do this even if other numeric identifier columns exist, e.g., "Week No".)
            if len(percent_cols) >= 2:
                labels = [str(c) for c in percent_cols]
                values = [normalize_percent(row.get(c)) for c in percent_cols]
                values = [v for v in values if not np.isnan(v)]
                # If after normalization we dropped values, bail.
                if len(values) < 2:
                    plt.close(fig)
                    return None

                y = np.arange(len(labels))
                colors = [ACCENT] * len(labels)
                ax.barh(y, values, color=colors)
                ax.set_yticks(y, labels=labels)
                ax.set_xlabel("Percent (%)")
                ax.set_xlim(0, 100)
                ax.axvline(100, color=PRIMARY, linewidth=1.5)

                for i, v in enumerate(values):
                    ax.text(v + 1, i, f"{v:.1f}%", va="center", fontsize=8, color=SECONDARY)

            # Multiple numeric columns -> horizontal bar comparison.
            elif len(numeric_cols) >= 2:
                labels = [str(c) for c in numeric_cols]
                values = []
                for c in numeric_cols:
                    try:
                        values.append(float(row.get(c)))
                    except Exception:
                        values.append(np.nan)
                if sum(~np.isnan(values)) < 2:
                    plt.close(fig)
                    return None

                y = np.arange(len(labels))
                colors = [PRIMARY] * len(labels)
                ax.barh(y, values, color=colors)
                ax.set_yticks(y, labels=labels)
                ax.set_xlabel("Value")
                for i, v in enumerate(values):
                    if np.isnan(v):
                        continue
                    ax.text(v, i, f" {v:.2f}".rstrip("0").rstrip("."), va="center", fontsize=8, color=SECONDARY)

            else:
                plt.close(fig)
                return None

        else:
            # Grouped bar chart: rows as groups, columns as series.
            series_cols = numeric_or_percent_cols
            series_count = len(series_cols)
            group_count = n_rows

            group_labels = [str(i) for i in range(group_count)]

            # Values matrix [group, series]
            values = []
            for c in series_cols:
                col_vals = plot_df[c].tolist()
                if column_types.get(str(c)) == "percent":
                    col_vals = [normalize_percent(v) for v in col_vals]
                values.append(col_vals)
            # values is [series, group] -> transpose to [group, series]
            values = np.array(values).T

            x = np.arange(group_count)
            total_width = 0.85
            bar_width = total_width / max(series_count, 1)

            for s_idx, c in enumerate(series_cols):
                offset = (s_idx - (series_count - 1) / 2) * bar_width
                bar_positions = x + offset
                color = PRIMARY if s_idx == 0 else (SECONDARY if s_idx == 1 else ACCENT)
                bars = ax.bar(bar_positions, values[:, s_idx], width=bar_width, color=color, label=str(c))
                for b in bars:
                    v = b.get_height()
                    if np.isnan(v):
                        continue
                    ax.text(
                        b.get_x() + b.get_width() / 2,
                        v,
                        f"{v:.1f}".rstrip("0").rstrip("."),
                        ha="center",
                        va="bottom",
                        fontsize=7,
                        color=SECONDARY,
                        rotation=0,
                    )

            ax.set_xticks(x, group_labels)
            ax.set_ylabel("Value")
            ax.legend(loc="upper right", fontsize=8, frameon=False)

        # Remove top and right spines.
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

        fig.tight_layout()

        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf.getvalue()

    def generate_for_sheet(
        self,
        sheet_name: str,
        df: pd.DataFrame,
        column_types: Dict[str, str],
        report_config: dict,
    ) -> List[dict]:
        """
        Returns list of chart dicts:
          [{"chart_id": str, "png_bytes": bytes|None, "position": str, "title": str}]
        Falls back to auto_select when no config entry for sheet.
        """
        if "asset and configuration" in sheet_name.lower():
            return []

        team_config = self._resolve_config(sheet_name, report_config)
        charts_spec = team_config.get("charts", [])

        if not charts_spec:
            auto_type = auto_select(df, column_types)
            charts_spec = [{"type": auto_type, "columns": "auto"}]

        charts_spec = charts_spec[:2]  # hard cap — max 2 charts per slide

        positions = ["right", "bottom_right"]
        results = []

        figsize_hint = (9, 3.4) if len(charts_spec) == 1 else (9, 2.8)

        for i, spec in enumerate(charts_spec):
            chart_type = spec.get("type", "auto")
            if chart_type == "auto":
                chart_type = auto_select(df, column_types)

            position = positions[i] if i < len(positions) else "right"
            title = spec.get("title", f"{sheet_name} — Chart {i + 1}")

            if chart_type == "none":
                results.append({
                    "chart_id": f"chart_{i}",
                    "png_bytes": None,
                    "position": position,
                    "title": title,
                    "chart_type": "none",
                })
                continue

            renderer = get_renderer(chart_type)
            png = renderer.render(df, {**spec, "column_types": column_types,
                                        "figsize_hint": figsize_hint})
            results.append({
                "chart_id": f"chart_{i}",
                "png_bytes": png,
                "position": position,
                "title": title,
                "chart_type": chart_type,
            })

        return results

