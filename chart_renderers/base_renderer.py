import io
import logging
from abc import ABC, abstractmethod
from typing import Optional

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

logger = logging.getLogger(__name__)

MUFG_COLORS = ["#CC0000", "#595757", "#404040", "#1F3864", "#E63333", "#808080", "#2E75B6"]
# MUFG Red (primary), MUFG Gray, Dark Gray, Dark Blue, Light Red, Med Gray, Mid Blue


class BaseRenderer(ABC):
    MUFG_COLORS = MUFG_COLORS

    def render(
        self,
        df: pd.DataFrame,
        chart_spec: dict,
        history: Optional[pd.DataFrame] = None,
    ) -> Optional[bytes]:
        """NEVER raises — catches all exceptions, returns None on failure."""
        try:
            return self._render_impl(df, chart_spec, history)
        except Exception as e:
            logger.warning(f"{self.__class__.__name__} render failed: {e}")
            return None

    def render_in_axis(
        self,
        ax,
        df: pd.DataFrame,
        chart_spec: dict,
        history: Optional[pd.DataFrame] = None,
    ) -> None:
        """Draw into a provided matplotlib Axes. NEVER raises."""
        try:
            self._render_in_axis_impl(ax, df, chart_spec, history)
        except Exception as e:
            logger.warning(f"{self.__class__.__name__} render_in_axis failed: {e}")

    @abstractmethod
    def _render_impl(
        self,
        df: pd.DataFrame,
        chart_spec: dict,
        history: Optional[pd.DataFrame],
    ) -> Optional[bytes]:
        raise NotImplementedError

    def _render_in_axis_impl(self, ax, df, chart_spec, history) -> None:
        ax.text(0.5, 0.5, self.__class__.__name__,
                ha="center", va="center", transform=ax.transAxes, fontsize=9)

    @staticmethod
    def _fig_to_bytes(fig) -> bytes:
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf.getvalue()

    def _resolve_columns(
        self,
        df: pd.DataFrame,
        chart_spec: dict,
        col_types: dict,
        desired_types=("numeric", "percent", "percent_decimal"),
    ) -> list:
        spec_cols = chart_spec.get("columns", "auto")
        if spec_cols and spec_cols != "auto":
            return [c.strip() for c in str(spec_cols).split(",") if c.strip() in df.columns]
        return [c for c in df.columns if col_types.get(str(c), "text") in desired_types]

    @staticmethod
    def _norm_pct(v) -> float:
        """Convert 0-1 ratio to 0-100 percentage for display."""
        try:
            fv = float(v)
        except Exception:
            return float("nan")
        if fv <= 1.0:
            return fv * 100.0
        return fv

    @staticmethod
    def _wrap_label(label: str, max_chars: int = 9) -> str:
        """Wrap label at first space if longer than max_chars, producing 2-line label."""
        label = str(label)
        if len(label) <= max_chars:
            return label
        idx = label.find(' ', 3)
        if idx != -1:
            return label[:idx] + '\n' + label[idx + 1:]
        return label[:max_chars] + '\n' + label[max_chars:]

    @staticmethod
    def _peek_x_labels(df: pd.DataFrame, chart_spec: dict) -> list:
        """Peek at what x-axis labels will be without drawing."""
        col_types = chart_spec.get("column_types", {})
        text_cols = [c for c in df.columns if col_types.get(str(c), "text") == "text"]
        if text_cols:
            return df[text_cols[0]].fillna("").astype(str).tolist()
        return [f"W{i + 1}" for i in range(len(df))]
