import json
from unittest.mock import Mock, patch

import pytest

from tools.t4_ai_engine import AIEngine


def _valid_ai_json():
    return json.dumps(
        {
            "summary": "Team summary",
            "kpi_evaluation": [],
            "key_achievements": ["Ach 1"],
            "insights": ["Insight 1"],
        }
    )


def test_fallback_on_connection_error():
    engine = AIEngine({"ai": {"provider": "ollama", "model": "phi3"}})
    with patch("tools.t4_ai_engine.ollama.chat", side_effect=Exception("connection error")):
        result = engine.generate_analysis(
            sheet_name="MIM",
            headers=["A"],
            rows=[{"A": 1}],
            kpi_dict={},
        )
        assert "AI analysis unavailable" in result["insights"][0]


def test_json_parse_success():
    engine = AIEngine({"ai": {"provider": "ollama", "model": "phi3"}})
    content = _valid_ai_json()
    mock_chat = Mock(return_value={"message": {"content": content}})
    with patch("tools.t4_ai_engine.ollama.chat", mock_chat):
        result = engine.generate_analysis(
            sheet_name="Change",
            headers=["B"],
            rows=[{"B": 2}],
            kpi_dict={},
        )
        assert result["summary"] == "Team summary"
        assert result["key_achievements"] == ["Ach 1"]
        assert result["insights"] == ["Insight 1"]


def test_use_excel_mode_forces_excel_summary():
    engine = AIEngine({"ai": {"provider": "ollama", "model": "phi3"}})
    content = json.dumps({
        "summary": "AI wrote this",
        "kpi_evaluation": [],
        "key_achievements": [],
        "insights": [],
    })
    with patch("tools.t4_ai_engine.ollama.chat", Mock(return_value={"message": {"content": content}})):
        result = engine.generate_analysis(
            "MIM", ["A"], [{"A": 1}], {},
            excel_summary="Team wrote this",
            summary_mode="use_excel",
        )
        assert result["summary"] == "Team wrote this"


def test_ai_write_context_in_prompt():
    engine = AIEngine({"ai": {"provider": "ollama", "model": "phi3"}})
    captured = []
    def fake_chat(model, messages, options):
        captured.extend(messages)
        return {"message": {"content": _valid_ai_json()}}
    with patch("tools.t4_ai_engine.ollama.chat", fake_chat):
        engine.generate_analysis(
            "NOC", ["X"], [{"X": 1}], {},
            excel_summary="Some context text",
            summary_mode="ai_write",
        )
    user_msg = next(m["content"] for m in captured if m["role"] == "user")
    assert "Team context" in user_msg


def test_overall_rag_in_fallback():
    engine = AIEngine({"ai": {"provider": "ollama", "model": "phi3"}})
    result = engine._fallback_analysis("MIM", [], {"row_count": 0, "top_metrics": [], "risk_metrics": []})
    assert result.get("overall_rag") == "AMBER"


def test_json_parse_retry():
    engine = AIEngine({"ai": {"provider": "ollama", "model": "phi3"}})

    bad = "not json"
    good = _valid_ai_json()
    mock_chat = Mock(side_effect=[{"message": {"content": bad}}, {"message": {"content": good}}])

    with patch("tools.t4_ai_engine.ollama.chat", mock_chat):
        result = engine.generate_analysis(
            sheet_name="NOC",
            headers=["X"],
            rows=[{"X": 3}],
            kpi_dict={},
        )
        assert mock_chat.call_count == 2
        assert result["summary"] == "Team summary"

