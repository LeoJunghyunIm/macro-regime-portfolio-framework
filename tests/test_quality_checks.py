import pandas as pd

from quality_checks import check_latest_signal_validity


def test_latest_signal_validity_passes_for_aligned_valid_outputs():
    index = pd.DatetimeIndex(["2026-06-19"], name="AsOfDate")
    risk = pd.DataFrame({"Risk_Appetite_Valid": [True]}, index=index)
    rates = pd.DataFrame({"Rates_Inflation_Valid": [True]}, index=index)
    portfolio = pd.DataFrame({"Portfolio_Router_Valid": [True]}, index=index)

    result = check_latest_signal_validity(risk, rates, portfolio)

    assert result["status"] == "PASS"
