from tools.t3_kpi_parser import KPIParser


def test_parses_greater_than_threshold():
    parser = KPIParser()
    thresholds = parser.parse_header_thresholds(["Change Success Ratio (>99%)"])
    assert thresholds["Change Success Ratio (>99%)"]["operator"] == ">"
    assert thresholds["Change Success Ratio (>99%)"]["value"] == 99
    assert thresholds["Change Success Ratio (>99%)"]["unit"] == "%"


def test_parses_less_than_threshold():
    parser = KPIParser()
    thresholds = parser.parse_header_thresholds(["Failed Change Ratio (<1%)"])
    assert thresholds["Failed Change Ratio (<1%)"]["operator"] == "<"
    assert thresholds["Failed Change Ratio (<1%)"]["value"] == 1
    assert thresholds["Failed Change Ratio (<1%)"]["unit"] == "%"


def test_evaluates_pass():
    parser = KPIParser()
    header = "Some KPI (>% value%)"
    thresholds = {
        "X": {"operator": ">", "value": 99, "unit": "%", "original_header": "X"},
    }
    row_data = {"X": 100.0}
    results = parser.evaluate_kpis(row_data=row_data, thresholds=thresholds)
    assert results[0]["status"] == "PASS"


def test_kpi_evaluate_percent_decimal_display():
    parser = KPIParser()
    thresholds = {
        "SLA Compliance (>95%)": {
            "operator": ">", "value": 95, "unit": "%",
            "original_header": "SLA Compliance (>95%)",
        }
    }
    row_data = {"SLA Compliance (>95%)": 0.97}
    column_types = {"SLA Compliance (>95%)": "percent_decimal"}
    results = parser.evaluate_kpis(row_data=row_data, thresholds=thresholds,
                                   column_types=column_types)
    assert results[0]["status"] == "PASS"
    assert abs(results[0]["actual_value"] - 97.0) < 0.01


def test_evaluates_fail():
    parser = KPIParser()
    thresholds = {
        "Y": {"operator": ">", "value": 0.70, "unit": "%", "original_header": "Y"},
    }
    row_data = {"Y": 0.47}
    results = parser.evaluate_kpis(row_data=row_data, thresholds=thresholds)
    assert results[0]["status"] == "FAIL"

