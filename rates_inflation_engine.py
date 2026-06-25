"""
rates_inflation_engine.py

Rates & Inflation Engine for the Macro Regime-to-Portfolio Decision Framework.

Purpose:
- Determine the rates/inflation regime.
- Translate rate pressure and inflation concern into a fixed-income regime.
- Calculate Real Yield Tightening as a Portfolio Router modifier.

Rules:
- Rate Pressure = ON if DGS10 > rolling P80 OR DGS10 d4w > 0.
- Inflation Concern = ON if T10YIE > rolling P80 OR T10YIE d4w > 0.
- Real Yield Tightening = ON if DFII10 > rolling P80 OR DFII10 d4w > 0.

Important:
- Rates & Inflation remains a 2x2 regime map.
- DFII10 is NOT a third Rates & Inflation axis.
- Real Yield Tightening is passed to Portfolio Router as a cap/constraint.
- DGS2 may be used for curve context in audit/memo, but it is not a core
  Rates & Inflation regime axis in the current framework.

Definition:
- d4w refers to the 4-week change after all series are resampled to
  week-ending Friday.
- d4w = current weekly value - value 4 weekly observations ago.
"""


def calculate_reason_code(value, p80, d4w) -> str:
    """
    Calculate reason code for level/momentum rules.

    Reason codes:
    - "Level": value is above rolling P80.
    - "Momentum": value increased over the past 4 weeks.
    - "Both": value is above rolling P80 and increased over the past 4 weeks.
    - "None": no pressure.
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


def classify_rates_inflation_regime(rate_pressure, inflation_concern) -> str:
    """
    Classify Rates & Inflation regime using the 2x2 map.

    Regime map:
    - Rate Pressure ON + Inflation Concern ON = Regime 1
    - Rate Pressure ON + Inflation Concern OFF = Regime 2
    - Rate Pressure OFF + Inflation Concern ON = Regime 3
    - Rate Pressure OFF + Inflation Concern OFF = Regime 4
    - Missing inputs = Insufficient Data
    """
    if rate_pressure is None or inflation_concern is None:
        return "Insufficient Data"

    try:
        if rate_pressure != rate_pressure or inflation_concern != inflation_concern:
            return "Insufficient Data"
    except TypeError:
        return "Insufficient Data"

    rate_pressure = bool(rate_pressure)
    inflation_concern = bool(inflation_concern)

    if rate_pressure and inflation_concern:
        return "Regime 1"

    if rate_pressure and not inflation_concern:
        return "Regime 2"

    if not rate_pressure and inflation_concern:
        return "Regime 3"

    return "Regime 4"


def generate_next_rates_inflation_trigger(current_row) -> str:
    """
    Generate rule-derived next Rates & Inflation trigger.

    This is intentionally simple in the current framework.
    """
    regime = str(current_row["Rates_Inflation_Regime"])

    if regime == "Insufficient Data":
        return "Insufficient data to generate Rates & Inflation trigger."

    rate_pressure = bool(current_row["Rate_Pressure"])
    inflation_concern = bool(current_row["Inflation_Concern"])

    if regime == "Regime 1":
        return (
            "If either Rate Pressure or Inflation Concern turns OFF, "
            "Rates & Inflation shifts out of Regime 1."
        )

    if regime == "Regime 2":
        return (
            "If T10YIE crosses P80 or d4w turns positive, Inflation Concern turns ON "
            "and Rates & Inflation shifts from Regime 2 to Regime 1."
        )

    if regime == "Regime 3":
        return (
            "If DGS10 crosses P80 or d4w turns positive, Rate Pressure turns ON "
            "and Rates & Inflation shifts from Regime 3 to Regime 1."
        )

    if regime == "Regime 4":
        return (
            "If DGS10 or T10YIE crosses P80 or d4w turns positive, "
            "Rates & Inflation shifts to a pressure regime."
        )

    raise ValueError(
        f"Invalid Rates & Inflation regime for trigger generation: {regime}. "
        f"Rate_Pressure={rate_pressure}, Inflation_Concern={inflation_concern}"
    )

def generate_real_yield_trigger(current_row) -> str:
    """
    Generate rule-derived Real Yield trigger.
    """
    if not bool(current_row.get("Rates_Inflation_Valid", False)):
        return "Insufficient data to generate Real Yield trigger."

    real_yield_tightening = bool(current_row["Real_Yield_Tightening"])

    if real_yield_tightening:
        return (
            "If Real Yield Tightening remains ON, aggressive Equity Overweight "
            "and Duration Overweight positions are capped."
        )

    return (
        "If DFII10 crosses P80 or d4w turns positive, Real Yield Tightening turns ON "
        "and aggressive Equity Overweight and Duration Overweight positions are capped."
    )


def run_rates_inflation(weekly_data, config: dict):
    """
    Run Rates & Inflation Engine.

    Expected Rates & Inflation engine output columns:
    - AsOfDate
    - Rates_Inflation_Valid
    - Rates_Inflation_Regime
    - Rate_Pressure
    - Rate_Reason
    - Inflation_Concern
    - Inflation_Reason
    - Real_Yield_Tightening
    - Real_Yield_Reason
    - Next_Rates_Inflation_Trigger
    - Real_Yield_Trigger

    Note:
    - DGS2 is available for curve context, audit, or memo support.
    - DGS2 does not define Rates & Inflation regime boundaries in the current framework.
    - Confidence is assigned later by quality_checks.py.
    """
    _ = config  # Reserved for future config-driven rule validation.

    required_columns = [
        "dgs10",
        "dgs10_p80",
        "dgs10_d4w",
        "t10yie",
        "t10yie_p80",
        "t10yie_d4w",
        "dfii10",
        "dfii10_p80",
        "dfii10_d4w",
        "dgs2",
    ]

    validity_columns = [
        "dgs10",
        "dgs10_p80",
        "dgs10_d4w",
        "t10yie",
        "t10yie_p80",
        "t10yie_d4w",
        "dfii10",
        "dfii10_p80",
        "dfii10_d4w",
    ]

    missing_columns = [
        column for column in required_columns if column not in weekly_data.columns
    ]

    if missing_columns:
        raise KeyError(f"Missing required Rates & Inflation columns: {missing_columns}")

    rates_inflation = weekly_data.copy()

    rates_inflation["Rates_Inflation_Valid"] = ~rates_inflation[validity_columns].isna().any(axis=1)

    rates_inflation["Rate_Reason"] = "Insufficient Data"
    rates_inflation["Inflation_Reason"] = "Insufficient Data"
    rates_inflation["Real_Yield_Reason"] = "Insufficient Data"
    rates_inflation["Rate_Pressure"] = None
    rates_inflation["Inflation_Concern"] = None
    rates_inflation["Real_Yield_Tightening"] = None
    rates_inflation["Rates_Inflation_Regime"] = "Insufficient Data"
    rates_inflation["Next_Rates_Inflation_Trigger"] = "Insufficient data to generate Rates & Inflation trigger."
    rates_inflation["Real_Yield_Trigger"] = "Insufficient data to generate Real Yield trigger."

    valid_mask = rates_inflation["Rates_Inflation_Valid"]

    rates_inflation.loc[valid_mask, "Rate_Reason"] = rates_inflation.loc[valid_mask].apply(
        lambda row: calculate_reason_code(
            value=row["dgs10"],
            p80=row["dgs10_p80"],
            d4w=row["dgs10_d4w"],
        ),
        axis=1,
    )

    rates_inflation.loc[valid_mask, "Inflation_Reason"] = rates_inflation.loc[valid_mask].apply(
        lambda row: calculate_reason_code(
            value=row["t10yie"],
            p80=row["t10yie_p80"],
            d4w=row["t10yie_d4w"],
        ),
        axis=1,
    )

    rates_inflation.loc[valid_mask, "Real_Yield_Reason"] = rates_inflation.loc[valid_mask].apply(
        lambda row: calculate_reason_code(
            value=row["dfii10"],
            p80=row["dfii10_p80"],
            d4w=row["dfii10_d4w"],
        ),
        axis=1,
    )

    rates_inflation.loc[valid_mask, "Rate_Pressure"] = (
        rates_inflation.loc[valid_mask, "Rate_Reason"] != "None"
    )

    rates_inflation.loc[valid_mask, "Inflation_Concern"] = (
        rates_inflation.loc[valid_mask, "Inflation_Reason"] != "None"
    )

    rates_inflation.loc[valid_mask, "Real_Yield_Tightening"] = (
        rates_inflation.loc[valid_mask, "Real_Yield_Reason"] != "None"
    )

    rates_inflation.loc[valid_mask, "Rates_Inflation_Regime"] = rates_inflation.loc[valid_mask].apply(
        lambda row: classify_rates_inflation_regime(
            rate_pressure=row["Rate_Pressure"],
            inflation_concern=row["Inflation_Concern"],
        ),
        axis=1,
    )

    rates_inflation.loc[valid_mask, "Next_Rates_Inflation_Trigger"] = rates_inflation.loc[valid_mask].apply(
        generate_next_rates_inflation_trigger,
        axis=1,
    )

    rates_inflation.loc[valid_mask, "Real_Yield_Trigger"] = rates_inflation.loc[valid_mask].apply(
        generate_real_yield_trigger,
        axis=1,
    )

    output_columns = [
        "Rates_Inflation_Valid",
        "Rates_Inflation_Regime",
        "Rate_Pressure",
        "Rate_Reason",
        "Inflation_Concern",
        "Inflation_Reason",
        "Real_Yield_Tightening",
        "Real_Yield_Reason",
        "Next_Rates_Inflation_Trigger",
        "Real_Yield_Trigger",
    ]

    rates_inflation_output = rates_inflation[output_columns].copy()
    rates_inflation_output.index.name = "AsOfDate"

    return rates_inflation_output
