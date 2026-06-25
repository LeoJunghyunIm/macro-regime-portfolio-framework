"""
Portfolio Router for the Macro Regime-to-Portfolio Decision Framework.

The router combines the Risk Appetite classification and Rates & Inflation
regime, then applies a real-yield constraint to produce transparent stances
for equity, duration, cash, and TIPS.

It is a rule-based decision framework, not an optimizer. Stance bands are
illustrative and are not optimized portfolio weights.
"""

from __future__ import annotations

import pandas as pd


def normalize_risk_appetite_regime_key(regime: str) -> str:
    """Convert a public Risk Appetite label into the configuration key."""
    mapping = {
        "Risk-On": "risk_on",
        "Neutral": "neutral",
        "Risk-Off": "risk_off",
    }
    if regime not in mapping:
        raise ValueError(f"Invalid Risk Appetite regime: {regime}")
    return mapping[regime]


def normalize_rates_inflation_regime_key(regime: str) -> str:
    """Convert a public Rates & Inflation label into the configuration key."""
    mapping = {
        "Regime 1": "regime_1",
        "Regime 2": "regime_2",
        "Regime 3": "regime_3",
        "Regime 4": "regime_4",
    }
    if regime not in mapping:
        raise ValueError(f"Invalid Rates & Inflation regime: {regime}")
    return mapping[regime]


def get_base_stance(
    risk_appetite_regime: str,
    rates_inflation_regime: str,
    config: dict,
) -> dict:
    """Read the base asset-class stances from the configured 3x4 matrix."""
    risk_key = normalize_risk_appetite_regime_key(risk_appetite_regime)
    rates_key = normalize_rates_inflation_regime_key(rates_inflation_regime)

    try:
        base_stance = config["portfolio_router"]["base_matrix"][risk_key][rates_key]
    except KeyError as error:
        raise KeyError(
            "Could not find a Portfolio Router stance for "
            f"Risk Appetite={risk_appetite_regime}, "
            f"Rates & Inflation={rates_inflation_regime}."
        ) from error

    allowed_stances = set(config["portfolio_router"]["allowed_stances"])
    invalid_stances = {
        asset: stance
        for asset, stance in base_stance.items()
        if stance not in allowed_stances
    }
    if invalid_stances:
        raise ValueError(
            f"Invalid stances in the Portfolio Router matrix: {invalid_stances}"
        )

    return dict(base_stance)


def apply_real_yield_modifier(
    base_stance: dict,
    real_yield_tightening: bool,
) -> tuple[dict, list[str]]:
    """Cap equity and duration overweight stances when real yields tighten."""
    final_stance = dict(base_stance)
    adjustments: list[str] = []

    if not real_yield_tightening:
        return final_stance, adjustments

    if final_stance.get("equity") == "Overweight":
        final_stance["equity"] = "Neutral"
        adjustments.append(
            "Equity Overweight capped to Neutral due to Real Yield Tightening."
        )

    if final_stance.get("duration") == "Overweight":
        final_stance["duration"] = "Neutral"
        adjustments.append(
            "Duration Overweight capped to Neutral due to Real Yield Tightening."
        )

    return final_stance, adjustments


def attach_stance_bands(final_stance: dict, config: dict) -> dict:
    """Attach the illustrative allocation band for each stance."""
    stance_bands = config["portfolio_router"]["stance_bands"]
    output = dict(final_stance)

    for asset, stance in final_stance.items():
        asset_key = asset.lower()
        stance_key = stance.lower()
        try:
            lower_bound, upper_bound = stance_bands[asset_key][stance_key]
        except KeyError as error:
            raise KeyError(
                f"Missing stance band for asset={asset_key}, stance={stance_key}."
            ) from error

        output[f"{asset_key}_band"] = (
            f"{int(lower_bound * 100)}-{int(upper_bound * 100)}%"
        )

    return output


def generate_portfolio_notes(
    base_stance: dict,
    final_stance: dict,
    real_yield_tightening: bool,
    adjustments: list[str],
) -> str:
    """Generate a concise explanation of the router result."""
    notes: list[str] = []

    if adjustments:
        notes.extend(adjustments)

    if real_yield_tightening and not adjustments:
        notes.append(
            "Real Yield Tightening is ON, but no equity or duration "
            "overweight stance required a cap."
        )

    if not real_yield_tightening:
        notes.append(
            "Real Yield Tightening is OFF; the base portfolio stance is unchanged."
        )

    if base_stance == final_stance:
        notes.append("Final stance matches the base portfolio stance.")
    else:
        notes.append("Final stance reflects the Real Yield Tightening constraint.")

    return " ".join(notes)


def run_portfolio_router(
    risk_appetite_output: pd.DataFrame,
    rates_inflation_output: pd.DataFrame,
    config: dict,
    quality_output: dict | None = None,
) -> pd.DataFrame:
    """Run the Portfolio Router across every aligned weekly observation."""
    required_risk_columns = [
        "Risk_Appetite_Valid",
        "Risk_Appetite_Regime",
        "Risk_Appetite_Score",
    ]
    required_rates_columns = [
        "Rates_Inflation_Valid",
        "Rates_Inflation_Regime",
        "Real_Yield_Tightening",
    ]

    missing_risk_columns = [
        column
        for column in required_risk_columns
        if column not in risk_appetite_output.columns
    ]
    missing_rates_columns = [
        column
        for column in required_rates_columns
        if column not in rates_inflation_output.columns
    ]

    if missing_risk_columns:
        raise KeyError(
            "Missing required Risk Appetite columns: "
            f"{missing_risk_columns}"
        )
    if missing_rates_columns:
        raise KeyError(
            "Missing required Rates & Inflation columns: "
            f"{missing_rates_columns}"
        )
    if not risk_appetite_output.index.equals(rates_inflation_output.index):
        raise ValueError(
            "Risk Appetite and Rates & Inflation outputs must share the same index."
        )

    output_rows: list[dict] = []

    for as_of_date in risk_appetite_output.index:
        risk_row = risk_appetite_output.loc[as_of_date]
        rates_row = rates_inflation_output.loc[as_of_date]

        risk_valid = bool(risk_row["Risk_Appetite_Valid"])
        rates_valid = bool(rates_row["Rates_Inflation_Valid"])
        portfolio_valid = risk_valid and rates_valid

        risk_regime = risk_row["Risk_Appetite_Regime"]
        rates_regime = rates_row["Rates_Inflation_Regime"]

        if not portfolio_valid:
            output_row = {
                "AsOfDate": as_of_date,
                "Portfolio_Router_Valid": False,
                "Risk_Appetite_Regime": risk_regime,
                "Risk_Appetite_Score": risk_row["Risk_Appetite_Score"],
                "Rates_Inflation_Regime": rates_regime,
                "Real_Yield_Tightening": None,
                "Equity_Stance": "Insufficient Data",
                "Duration_Stance": "Insufficient Data",
                "Cash_Stance": "Insufficient Data",
                "TIPS_Stance": "Insufficient Data",
                "Equity_Band": "N/A",
                "Duration_Band": "N/A",
                "Cash_Band": "N/A",
                "TIPS_Band": "N/A",
                "Portfolio_Adjustments": "Insufficient Data",
                "Portfolio_Notes": "Insufficient data to run the Portfolio Router.",
            }
        else:
            real_yield_tightening = bool(rates_row["Real_Yield_Tightening"])
            base_stance = get_base_stance(
                risk_appetite_regime=risk_regime,
                rates_inflation_regime=rates_regime,
                config=config,
            )
            final_stance, adjustments = apply_real_yield_modifier(
                base_stance=base_stance,
                real_yield_tightening=real_yield_tightening,
            )
            stance_with_bands = attach_stance_bands(final_stance, config)
            notes = generate_portfolio_notes(
                base_stance=base_stance,
                final_stance=final_stance,
                real_yield_tightening=real_yield_tightening,
                adjustments=adjustments,
            )

            output_row = {
                "AsOfDate": as_of_date,
                "Portfolio_Router_Valid": True,
                "Risk_Appetite_Regime": risk_regime,
                "Risk_Appetite_Score": int(risk_row["Risk_Appetite_Score"]),
                "Rates_Inflation_Regime": rates_regime,
                "Real_Yield_Tightening": real_yield_tightening,
                "Equity_Stance": stance_with_bands["equity"],
                "Duration_Stance": stance_with_bands["duration"],
                "Cash_Stance": stance_with_bands["cash"],
                "TIPS_Stance": stance_with_bands["tips"],
                "Equity_Band": stance_with_bands["equity_band"],
                "Duration_Band": stance_with_bands["duration_band"],
                "Cash_Band": stance_with_bands["cash_band"],
                "TIPS_Band": stance_with_bands["tips_band"],
                "Portfolio_Adjustments": (
                    " | ".join(adjustments) if adjustments else "None"
                ),
                "Portfolio_Notes": notes,
            }

        if quality_output is not None and "confidence" in quality_output:
            output_row["Confidence"] = quality_output["confidence"]

        output_rows.append(output_row)

    output = pd.DataFrame(output_rows).set_index("AsOfDate")
    output.index.name = "AsOfDate"
    return output
