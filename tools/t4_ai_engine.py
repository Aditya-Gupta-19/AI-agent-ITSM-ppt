import json
import math
from typing import Any, Dict, List, Optional

import ollama


class AIEngine:
    def __init__(self, config: dict):
        self.provider = config.get("ai", {}).get("provider") or config.get("provider") or "ollama"
        ai_cfg = config.get("ai", {}) if isinstance(config, dict) else {}
        self.model = ai_cfg.get("model") or config.get("model") or "phi3"
        self.temperature = ai_cfg.get("temperature", 0.3)
        self.max_tokens = ai_cfg.get("max_tokens", 1000)

    def generate_analysis(
        self,
        sheet_name: str,
        headers: List[str],
        rows: List[Dict[str, Any]],
        kpi_dict: Dict[str, Any],
    ) -> Dict[str, Any]:
        """
        Builds prompt and calls AI. Returns structured dict.

        If JSON parse fails: retry once with stricter prompt.
        If second attempt fails: return fallback dict with rule-based content.
        """

        base_system = (
            "You are a senior IT service management analyst. "
            "You analyze operational data for IT teams. "
            "IMPORTANT: Always respond with valid JSON only. "
            "No explanation, no markdown, no text outside the JSON object."
        )

        kpi_reference = kpi_dict.get(sheet_name, {}) if isinstance(kpi_dict, dict) else {}
        data_profile = self._build_data_profile(headers, rows)
        user_prompt = self._build_user_prompt(sheet_name, headers, rows, kpi_reference, data_profile)

        try:
            text = self._call_ollama(base_system, user_prompt)
            parsed = self._parse_json_or_none(text)
            if not self._validate_output(parsed):
                raise ValueError("Invalid JSON structure")
            return parsed
        except Exception:
            # Retry once with stricter prompt.
            try:
                stricter_system = base_system + " Return ONLY the JSON object, with double quotes for all keys."
                stricter_user = user_prompt + "\nReturn ONLY the JSON object. No additional text."
                text2 = self._call_ollama(stricter_system, stricter_user)
                parsed2 = self._parse_json_or_none(text2)
                if not self._validate_output(parsed2):
                    raise ValueError("Invalid JSON structure after retry")
                return parsed2
            except Exception:
                return self._fallback_analysis(sheet_name, rows, data_profile)

    def _build_user_prompt(
        self,
        sheet_name: str,
        headers: List[str],
        rows: List[Dict[str, Any]],
        kpi_reference: Dict[str, Any],
        data_profile: Dict[str, Any],
    ) -> str:
        # Keep prompts compact: cap row count if many rows.
        limited_rows = rows[:20]
        return (
            f"Group: {sheet_name}\n"
            f"Columns: {headers}\n"
            f"Data: {limited_rows}\n"
            f"KPI reference: {kpi_reference}\n"
            f"Data profile: {data_profile}\n"
            "Write a professional summary in 3-5 complete sentences. "
            "The summary must mention: performance level, strongest metric, weakest metric/risk, "
            "and one concrete recommendation.\n"
            "Return ONLY JSON with keys: summary, kpi_evaluation, key_achievements, insights.\n"
        )

    def _call_ollama(self, system_prompt: str, user_prompt: str) -> str:
        response = ollama.chat(
            model=self.model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            options={
                "temperature": self.temperature,
                # Some Ollama builds accept num_predict; others ignore. Keep harmless.
                "num_predict": self.max_tokens,
            },
        )
        content = response.get("message", {}).get("content")
        if not content:
            raise RuntimeError("Empty AI response")
        return content

    def _parse_json_or_none(self, text: str) -> Optional[Dict[str, Any]]:
        if not text:
            return None
        try:
            return json.loads(text)
        except Exception:
            # Try to extract JSON object from within additional text.
            start = text.find("{")
            end = text.rfind("}")
            if start == -1 or end == -1 or end <= start:
                return None
            candidate = text[start : end + 1]
            try:
                return json.loads(candidate)
            except Exception:
                return None

    def _validate_output(self, data: Any) -> bool:
        if not isinstance(data, dict):
            return False
        required_keys = {"summary", "kpi_evaluation", "key_achievements", "insights"}
        if not required_keys.issubset(set(data.keys())):
            return False
        if not isinstance(data["kpi_evaluation"], list):
            return False
        if not isinstance(data["key_achievements"], list):
            return False
        if not isinstance(data["insights"], list):
            return False
        # Minimal type checks for robustness.
        if not isinstance(data["summary"], str):
            return False
        return True

    def _to_float(self, value: Any) -> Optional[float]:
        if value is None:
            return None
        try:
            if isinstance(value, str):
                value = value.strip().replace("%", "")
                if value == "":
                    return None
            num = float(value)
            if math.isnan(num):
                return None
            return num
        except Exception:
            return None

    def _build_data_profile(self, headers: List[str], rows: List[Dict[str, Any]]) -> Dict[str, Any]:
        row_count = len(rows)
        if row_count == 0:
            return {"row_count": 0, "metric_count": 0, "top_metrics": [], "risk_metrics": []}

        # Use latest row for snapshot (for ITSM weekly data this is typically the most relevant).
        latest = rows[-1]
        numeric_values: List[tuple[str, float]] = []
        percent_like_values: List[tuple[str, float]] = []
        for h in headers:
            v = self._to_float(latest.get(h))
            if v is None:
                continue
            numeric_values.append((h, v))
            if "%" in str(h) or 0.0 <= v <= 1.0:
                percent = v * 100.0 if v <= 1.0 else v
                percent_like_values.append((h, percent))

        top_metrics = sorted(percent_like_values, key=lambda x: x[1], reverse=True)[:3]
        risk_metrics = sorted(percent_like_values, key=lambda x: x[1])[:3]
        return {
            "row_count": row_count,
            "metric_count": len(numeric_values),
            "top_metrics": top_metrics,
            "risk_metrics": risk_metrics,
        }

    def _fallback_analysis(self, sheet_name: str, rows: List[Dict[str, Any]], data_profile: Dict[str, Any]) -> Dict[str, Any]:
        """Rule-based fallback if AI fails."""
        row_count = len(rows)
        top = data_profile.get("top_metrics") or []
        risk = data_profile.get("risk_metrics") or []

        top_text = ", ".join([f"{k}: {v:.1f}%" for k, v in top[:2]]) if top else "no standout KPI values were detected"
        risk_text = ", ".join([f"{k}: {v:.1f}%" for k, v in risk[:2]]) if risk else "no clear risk KPI was detected"

        summary = (
            f"The {sheet_name} team has {row_count} data record(s) in the current reporting window. "
            f"Performance highlights indicate {top_text}. "
            f"Potential risk areas include {risk_text}. "
            f"Recommended action is to sustain high-performing KPIs while prioritizing targeted corrective actions "
            f"for underperforming indicators in the next cycle."
        )
        return {
            "summary": summary,
            "kpi_evaluation": [],
            "key_achievements": [
                f"{sheet_name} data was processed successfully.",
                "Automated KPI snapshot generated for this period.",
            ],
            "insights": [
                "AI analysis unavailable. Rule-based summary was applied.",
                "Review low-performing KPIs and assign action owners.",
            ],
        }

