"""
risk_appetite_engine.py

Risk Appetite Engine for the Macro Regime-to-Portfolio Decision Framework.

Purpose:
- Determine whether the broad market environment supports risk-taking.
- Convert credit, financial conditions, volatility, and labor signals into
  a simple Risk-On / Neutral / Risk-Off regime.

Rules:
- Credit Stress = 1 if HY OAS > rolling P80 OR HY OAS d4w > 0.
- NFCI Stress = 1 if NFCI > rolling P80.
- VIX Stress = 1 if VIX > rolling P80.
- Claims Stress = 1 if Initial Claims > rolling P80.

Important:
- Low VIX does NOT create a risk-on score.
- NFCI is level-only.
- Claims is level-only.
- No 0.5 weights.
- No negative scores.

Definition:
- d4w refers to the 4-week change after all series are resampled to
  week-ending Friday.
- d4w = current weekly value - value 4 weekly observations ago.
"""


def calculate_hy_reason(value, p80, d4w) -> str:
    """
    Calculate HY OAS reason code.

    Reason codes:
    - "Level": HY OAS is above rolling P80.
    - "Momentum": HY OAS has widened over the past 4 weeks.
    - "Both": HY OAS is above rolling P80 and widening.
    - "None": No credit stress.
    """
    level = value > p80
    momentum = d4w > 0

    if level and momentum:
        return "Both"
    if level:
        return "Level"
    if momentum:
        return "Momentum"
    return "None"


def calculate_risk_appetite_score(row) -> int:
    """
    Calculate Risk Appetite score.

    Risk Appetite Score =
        Credit_Flag + NFCI_Flag + VIX_Flag + Claims_Flag
    """
    required_flags = [
        "Credit_Flag",
        "NFCI_Flag",
        "VIX_Flag",
        "Claims_Flag",
    ]

    missing_flags = [flag for flag in required_flags if flag not in row]

    if missing_flags:
        raise KeyError(f"Missing Risk Appetite flag columns: {missing_flags}")

    return int(sum(row[flag] for flag in required_flags))


def classify_risk_appetite_regime(score) -> str:
    """
    Classify Risk Appetite regime based on Risk Appetite score.

    Regime map:
    - Score 0: Risk-On
    - Score 1-2: Neutral
    - Score 3-4: Risk-Off
    - Missing score: Insufficient Data
    """
    if score is None:
        return "Insufficient Data"

    try:
        if score != score:  # Handles NaN without requiring pandas import.
            return "Insufficient Data"
    except TypeError:
        return "Insufficient Data"

    score = int(score)

    if score == 0:
        return "Risk-On"
    if score in [1, 2]:
        return "Neutral"
    if score in [3, 4]:
        return "Risk-Off"

    raise ValueError(f"Invalid Risk Appetite score: {score}")


def calculate_risk_appetite_trend(current_row, prior_row) -> str:
    """
    Calculate the memo-level Risk Appetite trend.

    Current priority:
    1. If prior week is unavailable, return Stable with insufficient history note.
    2. If Risk_Appetite_Score increased vs prior week, return Deteriorating.
    3. If Risk_Appetite_Score decreased vs prior week, return Improving.
    4. If Risk_Appetite_Score is unchanged but HY_Reason includes Momentum, return Deteriorating.
    5. Otherwise, return Stable.

    Risk Appetite Trend is memo-level only.
    It is NOT a scored input.
    """
    if prior_row is None:
        return "Stable - insufficient prior history"

    current_score = int(current_row["Risk_Appetite_Score"])
    prior_score = int(prior_row["Risk_Appetite_Score"])

    if current_score > prior_score:
        return "Deteriorating"

    if current_score < prior_score:
        return "Improving"

    hy_reason = str(current_row["HY_Reason"])

    if hy_reason in ["Momentum", "Both"]:
        return "Deteriorating"

    return "Stable"


def generate_next_risk_appetite_trigger(current_row) -> str:
    """
    Generate rule-derived next Risk Appetite trigger.

    This is intentionally simple in the current framework.
    """
    score = int(current_row["Risk_Appetite_Score"])

    if score == 0:
        return (
            "If any stress flag turns ON and Risk Appetite score moves from 0 to 1, "
            "Risk Appetite shifts from Risk-On to Neutral."
        )

    if score == 1:
        return (
            "If one additional stress flag turns ON and Risk Appetite score moves from 1 to 2, "
            "Risk Appetite remains Neutral but risk appetite deteriorates."
        )

    if score == 2:
        return (
            "If one additional stress flag turns ON and Risk Appetite score moves from 2 to 3, "
            "Risk Appetite shifts from Neutral to Risk-Off."
        )

    if score in [3, 4]:
        return (
            "If stress flags decline and Risk Appetite score falls below 3, "
            "Risk Appetite shifts from Risk-Off to Neutral."
        )

    raise ValueError(f"Invalid Risk Appetite score for trigger generation: {score}")


def run_risk_appetite(weekly_data, config: dict):
    """
    Run Risk Appetite Engine.

    Expected Risk Appetite engine output columns:
    - AsOfDate
    - Risk_Appetite_Valid
    - Risk_Appetite_Regime
    - Risk_Appetite_Score
    - Credit_Flag
    - HY_Reason
    - NFCI_Flag
    - VIX_Flag
    - Claims_Flag
    - Risk_Appetite_Trend
    - Next_Risk_Appetite_Trigger

    Note:
    - Confidence is not created here.
    - Confidence is assigned later by quality_checks.py.
    """
    _ = config  # Reserved for future config-driven rule validation.

    required_columns = [
        "hy_oas",
        "hy_oas_p80",
        "hy_oas_d4w",
        "nfci",
        "nfci_p80",
        "vix",
        "vix_p80",
        "claims",
        "claims_p80",
    ]

    missing_columns = [
        column for column in required_columns if column not in weekly_data.columns
    ]

    if missing_columns:
        raise KeyError(f"Missing required Risk Appetite columns: {missing_columns}")

    risk_appetite = weekly_data.copy()

    risk_appetite["Risk_Appetite_Valid"] = ~risk_appetite[required_columns].isna().any(axis=1)

    risk_appetite["HY_Reason"] = "Insufficient Data"
    risk_appetite["Credit_Flag"] = None
    risk_appetite["NFCI_Flag"] = None
    risk_appetite["VIX_Flag"] = None
    risk_appetite["Claims_Flag"] = None
    risk_appetite["Risk_Appetite_Score"] = None
    risk_appetite["Risk_Appetite_Regime"] = "Insufficient Data"
    risk_appetite["Risk_Appetite_Trend"] = "Insufficient Data"
    risk_appetite["Next_Risk_Appetite_Trigger"] = "Insufficient data to generate Risk Appetite trigger."

    valid_mask = risk_appetite["Risk_Appetite_Valid"]

    risk_appetite.loc[valid_mask, "HY_Reason"] = risk_appetite.loc[valid_mask].apply(
        lambda row: calculate_hy_reason(
            value=row["hy_oas"],
            p80=row["hy_oas_p80"],
            d4w=row["hy_oas_d4w"],
        ),
        axis=1,
    )

    risk_appetite.loc[valid_mask, "Credit_Flag"] = (
        risk_appetite.loc[valid_mask, "HY_Reason"] != "None"
    ).astype(int)

    risk_appetite.loc[valid_mask, "NFCI_Flag"] = (
        risk_appetite.loc[valid_mask, "nfci"] > risk_appetite.loc[valid_mask, "nfci_p80"]
    ).astype(int)

    risk_appetite.loc[valid_mask, "VIX_Flag"] = (
        risk_appetite.loc[valid_mask, "vix"] > risk_appetite.loc[valid_mask, "vix_p80"]
    ).astype(int)

    risk_appetite.loc[valid_mask, "Claims_Flag"] = (
        risk_appetite.loc[valid_mask, "claims"] > risk_appetite.loc[valid_mask, "claims_p80"]
    ).astype(int)

    risk_appetite.loc[valid_mask, "Risk_Appetite_Score"] = risk_appetite.loc[valid_mask].apply(
        calculate_risk_appetite_score,
        axis=1,
    )

    risk_appetite.loc[valid_mask, "Risk_Appetite_Regime"] = risk_appetite.loc[
        valid_mask, "Risk_Appetite_Score"
    ].apply(classify_risk_appetite_regime)

    directions = []

    for i in range(len(risk_appetite)):
        current_row = risk_appetite.iloc[i]

        if not bool(current_row["Risk_Appetite_Valid"]):
            directions.append("Insufficient Data")
            continue

        prior_row = None

        for j in range(i - 1, -1, -1):
            candidate_prior = risk_appetite.iloc[j]
            if bool(candidate_prior["Risk_Appetite_Valid"]):
                prior_row = candidate_prior
                break

        directions.append(calculate_risk_appetite_trend(current_row, prior_row))

    risk_appetite["Risk_Appetite_Trend"] = directions

    risk_appetite.loc[valid_mask, "Next_Risk_Appetite_Trigger"] = risk_appetite.loc[valid_mask].apply(
        generate_next_risk_appetite_trigger,
        axis=1,
    )

    output_columns = [
        "Risk_Appetite_Valid",
        "Risk_Appetite_Regime",
        "Risk_Appetite_Score",
        "Credit_Flag",
        "HY_Reason",
        "NFCI_Flag",
        "VIX_Flag",
        "Claims_Flag",
        "Risk_Appetite_Trend",
        "Next_Risk_Appetite_Trigger",
    ]

    risk_appetite_output = risk_appetite[output_columns].copy()
    risk_appetite_output.index.name = "AsOfDate"

    return risk_appetite_output
