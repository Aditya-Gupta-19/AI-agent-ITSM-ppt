import pytest

from tools.t4_ai_engine import AIEngine, RuleEngine


def _make_kpi(name, status, actual=None, threshold=">=95%"):
    return {"kpi_name": name, "status": status, "actual_value": actual, "threshold": threshold}


# ── RAG computation ───────────────────────────────────────────────────────────

def test_rag_all_pass():
    engine = RuleEngine({})
    rows = [_make_kpi("KPI A", "PASS"), _make_kpi("KPI B", "PASS")]
    assert engine._compute_rag(rows) == "GREEN"


def test_rag_any_fail_is_red():
    engine = RuleEngine({})
    rows = [_make_kpi("KPI A", "PASS"), _make_kpi("KPI B", "FAIL")]
    assert engine._compute_rag(rows) == "RED"


def test_rag_amber_no_fail():
    engine = RuleEngine({})
    rows = [_make_kpi("KPI A", "PASS"), _make_kpi("KPI B", "AMBER")]
    assert engine._compute_rag(rows) == "AMBER"


def test_rag_fail_overrides_amber():
    engine = RuleEngine({})
    rows = [_make_kpi("KPI A", "AMBER"), _make_kpi("KPI B", "FAIL")]
    assert engine._compute_rag(rows) == "RED"


def test_rag_empty_rows_defaults_amber():
    engine = RuleEngine({})
    assert engine._compute_rag([]) == "AMBER"


# ── Template generation ───────────────────────────────────────────────────────

def test_achievements_from_pass_rows():
    engine = RuleEngine({})
    rows = [_make_kpi("SLA Compliance", "PASS", actual=0.98, threshold=">=95%")]
    achievements, _, _ = engine._generate_templates("MIM", rows)
    assert any("SLA Compliance" in a for a in achievements)


def test_focus_from_amber_and_fail():
    engine = RuleEngine({})
    rows = [
        _make_kpi("FCR Rate", "AMBER", actual=0.82, threshold=">=90%"),
        _make_kpi("Resolution Time", "FAIL", actual=0.65, threshold=">=95%"),
    ]
    _, focus, _ = engine._generate_templates("NOC", rows)
    assert len(focus) >= 2
    assert any("FCR Rate" in f for f in focus)
    assert any("Resolution Time" in f for f in focus)


def test_concerns_from_fail_rows():
    engine = RuleEngine({})
    rows = [_make_kpi("Change Success Ratio", "FAIL", actual=0.72, threshold=">=99%")]
    _, _, concerns = engine._generate_templates("Change", rows)
    assert any("Change Success Ratio" in c for c in concerns)
    assert any("Action required" in c for c in concerns)


def test_no_concern_fallback_when_all_pass():
    engine = RuleEngine({})
    rows = [_make_kpi("Uptime", "PASS", actual=1.0, threshold=">=99%")]
    _, _, concerns = engine._generate_templates("NOC", rows)
    assert any("No current concerns" in c for c in concerns)


def test_default_fallbacks_when_no_rows():
    engine = RuleEngine({})
    achievements, focus, concerns = engine._generate_templates("ServiceNow", [])
    assert len(achievements) > 0
    assert len(focus) > 0
    assert len(concerns) > 0


# ── generate_analysis output shape ───────────────────────────────────────────

def test_generate_analysis_returns_required_keys():
    engine = RuleEngine({})
    result = engine.generate_analysis(
        sheet_name="MIM",
        headers=["KPI", "Value"],
        rows=[{"KPI": "SLA", "Value": 0.95}],
        kpi_dict={},
        kpi_rows=[_make_kpi("SLA", "PASS", actual=0.95)],
    )
    assert "overall_rag" in result
    assert "key_achievements" in result
    assert "insights" in result
    assert "summary" in result
    assert "kpi_evaluation" in result


def test_generate_analysis_overall_performance_row_excluded():
    engine = RuleEngine({})
    rows = [
        _make_kpi("SLA", "PASS"),
        {"kpi_name": "Overall Performance", "status": "GREEN"},
    ]
    result = engine.generate_analysis("MIM", [], [], {}, kpi_rows=rows)
    assert result["overall_rag"] == "GREEN"


# ── Value formatting ──────────────────────────────────────────────────────────

def test_fmt_value_percent_threshold():
    assert RuleEngine._fmt_value(0.966, ">=95%") == "96.6%"


def test_fmt_value_integer():
    assert RuleEngine._fmt_value(42.0, ">=30") == "42"


def test_fmt_value_none():
    assert RuleEngine._fmt_value(None, ">=95%") == "N/A"


def test_fmt_value_non_numeric():
    assert RuleEngine._fmt_value("N/A", "") == "N/A"


# ── Backward-compatibility alias ──────────────────────────────────────────────

def test_aiengine_alias_is_rule_engine():
    assert AIEngine is RuleEngine


def test_aiengine_alias_instantiates():
    engine = AIEngine({})
    result = engine.generate_analysis("Change", [], [], {})
    assert "overall_rag" in result
