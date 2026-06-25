import pandas as pd

from risk_appetite_engine import run_risk_appetite


def test_risk_appetite_classifies_risk_off(config):
    index = pd.DatetimeIndex(["2026-06-19"], name="AsOfDate")
    data = pd.DataFrame(
        {
            "hy_oas": [5.0],
            "hy_oas_p80": [4.0],
            "hy_oas_d4w": [0.4],
            "nfci": [0.5],
            "nfci_p80": [0.2],
            "vix": [32.0],
            "vix_p80": [25.0],
            "claims": [220000.0],
            "claims_p80": [260000.0],
        },
        index=index,
    )

    output = run_risk_appetite(data, config)
    latest = output.iloc[-1]

    assert latest["Risk_Appetite_Valid"]
    assert latest["HY_Reason"] == "Both"
    assert latest["Risk_Appetite_Score"] == 3
    assert latest["Risk_Appetite_Regime"] == "Risk-Off"
