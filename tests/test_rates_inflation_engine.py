import pandas as pd

from rates_inflation_engine import run_rates_inflation


def test_rates_inflation_classifies_regime_one(config):
    index = pd.DatetimeIndex(["2026-06-19"], name="AsOfDate")
    data = pd.DataFrame(
        {
            "dgs10": [4.8],
            "dgs10_p80": [4.4],
            "dgs10_d4w": [0.2],
            "t10yie": [2.8],
            "t10yie_p80": [2.6],
            "t10yie_d4w": [0.1],
            "dfii10": [2.2],
            "dfii10_p80": [2.0],
            "dfii10_d4w": [0.1],
            "dgs2": [4.5],
        },
        index=index,
    )

    output = run_rates_inflation(data, config)
    latest = output.iloc[-1]

    assert latest["Rates_Inflation_Valid"]
    assert latest["Rates_Inflation_Regime"] == "Regime 1"
    assert bool(latest["Real_Yield_Tightening"])
