from typing import Any, Dict, List, Optional


class RuleEngine:
    """
    Deterministic replacement for the AI engine.
    Computes overall_rag from evaluated KPI rows and generates
    template-based text when Excel comments are empty.
    No external dependencies; runs in milliseconds.
    """

    def __init__(self, config: dict):
        pass  # no AI configuration required

    def generate_analysis(
        self,
        sheet_name: str,
        headers: List[str],
        rows: List[Dict[str, Any]],
        kpi_dict: Dict[str, Any],
        excel_summary: str = "",
        summary_mode: str = "ai_write",
        kpi_rows: Optional[List[Dict[str, Any]]] = None,
    ) -> Dict[str, Any]:
        """
        Returns a dict compatible with the previous AI output format:
          overall_rag      : GREEN | AMBER | RED   (computed from kpi_rows)
          key_achievements : list[str]             (template-based fallback)
          insights         : list[str]             (template-based fallback)
          summary          : str                   (template-based fallback)
          kpi_evaluation   : []                    (KPI table built separately)
        """
        evaluated = [
            r for r in (kpi_rows or [])
            if r.get("kpi_name") != "Overall Performance"
        ]

        overall_rag = self._compute_rag(evaluated)
        achievements, focus_items, concerns = self._generate_templates(sheet_name, evaluated)

        return {
            "summary": concerns[0] if concerns else "",
            "kpi_evaluation": [],
            "key_achievements": achievements,
            "insights": focus_items,
            "overall_rag": overall_rag,
        }

    def _compute_rag(self, kpi_rows: List[Dict]) -> str:
        """
        RED   : any KPI status is FAIL
        AMBER : no FAIL but at least one AMBER
        GREEN : all evaluated KPIs are PASS
        AMBER : default when no threshold KPIs exist
        """
        statuses = {r.get("status", "UNKNOWN") for r in kpi_rows}
        if "FAIL" in statuses:
            return "RED"
        if "AMBER" in statuses:
            return "AMBER"
        if "PASS" in statuses:
            return "GREEN"
        return "AMBER"

    @staticmethod
    def _fmt_value(value: Any, threshold: str) -> str:
        """Format a KPI value for display in template text."""
        if value is None:
            return "N/A"
        try:
            fv = float(value)
            if "%" in threshold:
                display = fv * 100 if fv <= 1.0 else fv
                return f"{display:.1f}%"
            return str(int(fv)) if fv == int(fv) else f"{fv:.1f}"
        except Exception:
            return str(value)

    def _generate_templates(
        self, sheet_name: str, kpi_rows: List[Dict]
    ) -> tuple:
        """
        Returns (achievements, focus, concerns) as lists of strings
        derived entirely from KPI evaluation results.
        """
        pass_rows  = [r for r in kpi_rows if r.get("status") == "PASS"]
        amber_rows = [r for r in kpi_rows if r.get("status") == "AMBER"]
        fail_rows  = [r for r in kpi_rows if r.get("status") == "FAIL"]

        # ── KEY ACHIEVEMENTS: passing KPIs ────────────────────────────────
        achievements: List[str] = []
        for r in pass_rows[:3]:
            name = r["kpi_name"]
            val  = self._fmt_value(r.get("actual_value"), r.get("threshold", ""))
            thr  = r.get("threshold", "target")
            achievements.append(f"{name} met target at {val} (threshold: {thr})")
        if not achievements:
            achievements = [
                f"All monitored metrics captured for {sheet_name} this reporting period.",
                "Data validated and ready for management review.",
            ]

        # ── NEXT WEEK FOCUS: amber and fail KPIs needing attention ────────
        focus: List[str] = []
        for r in (amber_rows + fail_rows)[:3]:
            name = r["kpi_name"]
            val  = self._fmt_value(r.get("actual_value"), r.get("threshold", ""))
            thr  = r.get("threshold", "target")
            focus.append(
                f"Monitor {name} — currently at {val} against {thr} target"
            )
        if not focus:
            focus = [
                "Sustain current performance levels across all KPIs.",
                "Continue weekly data validation and reporting cadence.",
            ]

        # ── CONCERNS: failing KPIs requiring action ────────────────────────
        concerns: List[str] = []
        for r in fail_rows[:3]:
            name = r["kpi_name"]
            thr  = r.get("threshold", "target threshold")
            val  = self._fmt_value(r.get("actual_value"), r.get("threshold", ""))
            concerns.append(
                f"Action required: {name} fell below the target threshold of {thr}. "
                f"Current value: {val}."
            )
        if not concerns:
            concerns = [
                "No current concerns. All metrics performing at or above target thresholds."
            ]

        return achievements, focus, concerns


# Backward-compatibility alias — existing imports of AIEngine still work.
AIEngine = RuleEngine
