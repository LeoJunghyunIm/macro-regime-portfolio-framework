import pandas as pd

from portfolio_router import run_portfolio_router


def test_real_yield_constraint_caps_aggressive_overweights(config):
    index = pd.DatetimeIndex(["2026-06-19"], name="AsOfDate")
    risk = pd.DataFrame(
        {
            "Risk_Appetite_Valid": [True],
            "Risk_Appetite_Regime": ["Risk-On"],
            "Risk_Appetite_Score": [0],
        },
        index=index,
    )
    rates = pd.DataFrame(
        {
            "Rates_Inflation_Valid": [True],
            "Rates_Inflation_Regime": ["Regime 4"],
            "Real_Yield_Tightening": [True],
        },
        index=index,
    )

    output = run_portfolio_router(risk, rates, config)
    latest = output.iloc[-1]

    assert latest["Equity_Stance"] == "Neutral"
    assert latest["Duration_Stance"] == "Neutral"
    assert "capped" in latest["Portfolio_Adjustments"]
