"""
diagnostics.py

Diagnostic audit module for the Macro Regime-to-Portfolio Decision Framework.

Purpose:
- Evaluate historical regime distribution.
- Build Risk Appetite x Rates & Inflation co-occurrence matrix.
- Build stress-event false-negative / regime-response diagnostics.
- Support framework audit without turning the project into a return backtest.

Important:
- Diagnostics are not trading signals.
- Diagnostics are not return backtests.
- Diagnostics are used to understand how the framework behaved across
  historical regime combinations and major stress-event windows.
"""

import pandas as pd


def normalize_bool(value) -> bool:
    """
    Convert common boolean-like values to bool.

    Handles:
    - True / False
    - 1 / 0
    - "TRUE" / "FALSE"
    - "ON" / "OFF"
    - "VALID" / "INVALID"
    - "YES" / "NO"
    """
    if pd.isna(value):
        return False

    if isinstance(value, str):
        return value.strip().lower() in {
            "true",
            "valid",
            "on",
            "yes",
            "1",
        }

    try:
        return bool(value == 1 or value is True)
    except TypeError:
        return False


def build_valid_mask(validity_series: pd.Series) -> pd.Series:
    """
    Build a robust boolean mask for validity fields.
    """
    return validity_series.apply(normalize_bool)


def classify_pre_existing_stress(pre_row, focus: str) -> str:
    """
    Classify whether the framework was already stressed before the event window.

    Returns:
    - Yes: clearly defensive/stressed before event start
    - Partial: some deterioration/pressure was already present
    - No: benign before event start
    - N/A: no valid pre-event observation
    """
    if pre_row is None:
        return "N/A"

    if focus in ["Risk Appetite_RISK_OFF", "Risk Appetite_DETERIORATION"]:
        risk_appetite_regime = pre_row.get("Risk_Appetite_Regime", "N/A")
        risk_appetite_score = pd.to_numeric(pre_row.get("Risk_Appetite_Score"), errors="coerce")

        if risk_appetite_regime == "Risk-Off" or risk_appetite_score >= 3:
            return "Yes"

        if risk_appetite_regime == "Neutral" or risk_appetite_score >= 1:
            return "Partial"

        return "No"

    if focus == "Rates & Inflation_PRESSURE":
        pressure_count = int(normalize_bool(pre_row.get("Rate_Pressure")))
        pressure_count += int(normalize_bool(pre_row.get("Inflation_Concern")))
        pressure_count += int(normalize_bool(pre_row.get("Real_Yield_Tightening")))

        if pressure_count >= 2:
            return "Yes"

        if pressure_count == 1:
            return "Partial"

        return "No"

    return "N/A"


def build_cooccurrence_matrix(portfolio_output: pd.DataFrame) -> dict:
    """
    Build Risk Appetite x Rates & Inflation co-occurrence count and percentage matrices.

    Uses only valid Portfolio Router rows.

    Expected input columns:
    - Portfolio_Router_Valid
    - Risk_Appetite_Regime
    - Rates_Inflation_Regime

    Returns:
    {
        "count_matrix": DataFrame,
        "pct_matrix": DataFrame,
        "diagnostic_summary": DataFrame
    }
    """
    required_columns = [
        "Portfolio_Router_Valid",
        "Risk_Appetite_Regime",
        "Rates_Inflation_Regime",
    ]

    missing_columns = [
        column for column in required_columns
        if column not in portfolio_output.columns
    ]

    if missing_columns:
        raise KeyError(
            f"Missing required columns for co-occurrence diagnostic: {missing_columns}"
        )

    valid_mask = build_valid_mask(portfolio_output["Portfolio_Router_Valid"])
    valid_data = portfolio_output[valid_mask].copy()

    if valid_data.empty:
        empty_summary = pd.DataFrame(
            {
                "Metric": [
                    "Valid observations",
                    "Start date",
                    "End date",
                    "Most common Risk Appetite regime",
                    "Most common Rates & Inflation regime",
                ],
                "Value": [
                    0,
                    "N/A",
                    "N/A",
                    "N/A",
                    "N/A",
                ],
            }
        )

        return {
            "count_matrix": pd.DataFrame(),
            "pct_matrix": pd.DataFrame(),
            "diagnostic_summary": empty_summary,
        }

    risk_appetite_order = ["Risk-On", "Neutral", "Risk-Off"]
    rates_inflation_order = ["Regime 1", "Regime 2", "Regime 3", "Regime 4"]

    count_matrix = pd.crosstab(
        valid_data["Risk_Appetite_Regime"],
        valid_data["Rates_Inflation_Regime"],
    )

    count_matrix = count_matrix.reindex(
        index=risk_appetite_order,
        columns=rates_inflation_order,
        fill_value=0,
    )

    count_matrix["Total"] = count_matrix.sum(axis=1)

    total_row = pd.DataFrame(count_matrix.sum(axis=0)).T
    total_row.index = ["Total"]

    count_matrix = pd.concat([count_matrix, total_row])

    total_observations = int(count_matrix.loc["Total", "Total"])

    if total_observations > 0:
        pct_matrix = count_matrix / total_observations
    else:
        pct_matrix = count_matrix * 0

    dates = pd.to_datetime(valid_data.index)
    start_date = dates.min().date()
    end_date = dates.max().date()

    diagnostic_summary = pd.DataFrame(
        {
            "Metric": [
                "Valid observations",
                "Start date",
                "End date",
                "Most common Risk Appetite regime",
                "Most common Rates & Inflation regime",
            ],
            "Value": [
                total_observations,
                start_date,
                end_date,
                valid_data["Risk_Appetite_Regime"].mode().iloc[0],
                valid_data["Rates_Inflation_Regime"].mode().iloc[0],
            ],
        }
    )

    return {
        "count_matrix": count_matrix,
        "pct_matrix": pct_matrix,
        "diagnostic_summary": diagnostic_summary,
    }


def get_default_stress_events() -> list[dict]:
    """
    Define stress-event windows and expected framework responses.

    These windows are fixed ex ante for diagnostic purposes.
    This is not a return backtest.
    """
    return [
        {
            "event": "Global Financial Crisis",
            "window_start": "2007-07-01",
            "window_end": "2009-06-30",
            "stress_type": "Credit / Financial Conditions / Labor",
            "focus": "Risk Appetite_RISK_OFF",
            "expected_response": (
                "Risk Appetite should deteriorate materially, ideally reaching Risk-Off "
                "during peak stress. Portfolio Router should reduce equity risk and increase "
                "defensive positioning."
            ),
            "first_warning_rule": "Risk_Appetite_Score >= 1",
        },
        {
            "event": "2011 Eurozone / US Debt Ceiling Stress",
            "window_start": "2011-07-01",
            "window_end": "2011-10-31",
            "stress_type": "Volatility / Credit / Policy",
            "focus": "Risk Appetite_DETERIORATION",
            "expected_response": (
                "Risk Appetite should show at least Neutral or elevated volatility/credit stress."
            ),
            "first_warning_rule": "Risk_Appetite_Score >= 1",
        },
        {
            "event": "2015-2016 China / Oil / HY Stress",
            "window_start": "2015-08-01",
            "window_end": "2016-02-29",
            "stress_type": "Credit / Growth / Commodity",
            "focus": "Risk Appetite_DETERIORATION",
            "expected_response": (
                "HY OAS and broader risk appetite indicators should help Risk Appetite "
                "deteriorate or at least move to Neutral."
            ),
            "first_warning_rule": "Risk_Appetite_Score >= 1",
        },
        {
            "event": "2018 Volmageddon / Q4 Selloff",
            "window_start": "2018-02-01",
            "window_end": "2018-12-31",
            "stress_type": "Volatility / Rates / Risk Appetite",
            "focus": "Risk Appetite_DETERIORATION",
            "expected_response": (
                "Risk Appetite should show volatility or risk-appetite deterioration around "
                "the February volatility spike and/or Q4 selloff."
            ),
            "first_warning_rule": "Risk_Appetite_Score >= 1",
        },
        {
            "event": "COVID Shock",
            "window_start": "2020-02-01",
            "window_end": "2020-04-30",
            "stress_type": "Fast Volatility / Credit / Labor Shock",
            "focus": "Risk Appetite_RISK_OFF",
            "expected_response": (
                "Risk Appetite should deteriorate rapidly due to volatility, credit, financial "
                "conditions, and labor stress. Portfolio Router should move defensively."
            ),
            "first_warning_rule": "Risk_Appetite_Score >= 1",
        },
        {
            "event": "2022 Inflation / Rates Shock",
            "window_start": "2022-01-01",
            "window_end": "2022-12-31",
            "stress_type": "Rates / Inflation / Real Yield",
            "focus": "Rates & Inflation_PRESSURE",
            "expected_response": (
                "Rates & Inflation should show Rate Pressure and/or Inflation Concern. Real Yield "
                "Tightening should cap aggressive Equity OW and Duration OW."
            ),
            "first_warning_rule": (
                "Rate_Pressure ON or Inflation_Concern ON or Real_Yield_Tightening ON"
            ),
        },
        {
            "event": "2023 Regional Bank Stress",
            "window_start": "2023-03-01",
            "window_end": "2023-05-31",
            "stress_type": "Banking / Financial Conditions / Rates",
            "focus": "Risk Appetite_DETERIORATION",
            "expected_response": (
                "Risk Appetite should show some deterioration through financial conditions, "
                "volatility, or credit, but full Risk-Off may not be necessary if "
                "the shock is contained."
            ),
            "first_warning_rule": "Risk_Appetite_Score >= 1",
        },
    ]


def copy_optional_driver_column(
    combined: pd.DataFrame,
    source: pd.DataFrame,
    standard_column: str,
    aliases: list[str],
) -> None:
    """
    Copy optional driver column into a standard column name.

    If no alias exists, fill the standard column with False.
    """
    for alias in aliases:
        if alias in source.columns:
            combined[standard_column] = source[alias]
            return

    combined[standard_column] = False


def prepare_stress_diagnostic_data(
    risk_appetite_output: pd.DataFrame,
    rates_inflation_output: pd.DataFrame,
    portfolio_output: pd.DataFrame,
) -> pd.DataFrame:
    """
    Combine Risk Appetite, Rates & Inflation, and Portfolio Router outputs for stress-event diagnostics.
    """
    required_risk_appetite = [
        "Risk_Appetite_Valid",
        "Risk_Appetite_Regime",
        "Risk_Appetite_Score",
    ]

    required_rates_inflation = [
        "Rates_Inflation_Valid",
        "Rates_Inflation_Regime",
        "Rate_Pressure",
        "Inflation_Concern",
        "Real_Yield_Tightening",
    ]

    required_portfolio = [
        "Portfolio_Router_Valid",
        "Equity_Stance",
        "Cash_Stance",
    ]

    for column in required_risk_appetite:
        if column not in risk_appetite_output.columns:
            raise KeyError(f"Missing Risk Appetite column for stress diagnostic: {column}")

    for column in required_rates_inflation:
        if column not in rates_inflation_output.columns:
            raise KeyError(f"Missing Rates & Inflation column for stress diagnostic: {column}")

    for column in required_portfolio:
        if column not in portfolio_output.columns:
            raise KeyError(f"Missing Portfolio Router column for stress diagnostic: {column}")

    combined = pd.DataFrame(index=portfolio_output.index.copy())

    combined["Risk_Appetite_Valid"] = risk_appetite_output["Risk_Appetite_Valid"]
    combined["Risk_Appetite_Regime"] = risk_appetite_output["Risk_Appetite_Regime"]
    combined["Risk_Appetite_Score"] = pd.to_numeric(
        risk_appetite_output["Risk_Appetite_Score"],
        errors="coerce",
    )

    copy_optional_driver_column(
        combined=combined,
        source=risk_appetite_output,
        standard_column="Credit_Flag",
        aliases=[
            "Credit_Flag",
            "HY_Flag",
            "HY_OAS_Flag",
            "HY_Stress_Flag",
            "HY_OAS_Stress_Flag",
        ],
    )

    copy_optional_driver_column(
        combined=combined,
        source=risk_appetite_output,
        standard_column="NFCI_Flag",
        aliases=[
            "NFCI_Flag",
            "Financial_Conditions_Flag",
            "FinancialConditions_Flag",
        ],
    )

    copy_optional_driver_column(
        combined=combined,
        source=risk_appetite_output,
        standard_column="VIX_Flag",
        aliases=[
            "VIX_Flag",
            "Volatility_Flag",
            "High_Vol_Flag",
        ],
    )

    copy_optional_driver_column(
        combined=combined,
        source=risk_appetite_output,
        standard_column="Claims_Flag",
        aliases=[
            "Claims_Flag",
            "Initial_Claims_Flag",
            "Labor_Flag",
        ],
    )

    combined["Rates_Inflation_Valid"] = rates_inflation_output["Rates_Inflation_Valid"]
    combined["Rates_Inflation_Regime"] = rates_inflation_output["Rates_Inflation_Regime"]
    combined["Rate_Pressure"] = rates_inflation_output["Rate_Pressure"]
    combined["Inflation_Concern"] = rates_inflation_output["Inflation_Concern"]
    combined["Real_Yield_Tightening"] = rates_inflation_output["Real_Yield_Tightening"]

    combined["Portfolio_Router_Valid"] = portfolio_output["Portfolio_Router_Valid"]
    combined["Equity_Stance"] = portfolio_output["Equity_Stance"]
    combined["Cash_Stance"] = portfolio_output["Cash_Stance"]

    combined = combined.sort_index()

    return combined


def build_warning_mask(event_data: pd.DataFrame, focus: str) -> pd.Series:
    """
    Build event-specific first-warning mask.

    Risk Appetite-focused events:
    - first warning = Risk_Appetite_Score >= 1

    Rates & Inflation-focused events:
    - first warning = Rate_Pressure, Inflation_Concern, or Real_Yield_Tightening ON
    """
    if focus in ["Risk Appetite_RISK_OFF", "Risk Appetite_DETERIORATION"]:
        return event_data["Risk_Appetite_Score"].fillna(0) >= 1

    if focus == "Rates & Inflation_PRESSURE":
        return (
            event_data["Rate_Pressure"].apply(normalize_bool)
            | event_data["Inflation_Concern"].apply(normalize_bool)
            | event_data["Real_Yield_Tightening"].apply(normalize_bool)
        )

    raise ValueError(f"Invalid stress-event focus: {focus}")


def classify_miss_status(event_data: pd.DataFrame, focus: str) -> str:
    """
    Classify whether the framework missed the event.

    This is a false-negative / regime-response diagnostic, not a performance test.
    """
    if event_data.empty:
        return "Not Applicable"

    max_score = event_data["Risk_Appetite_Score"].max(skipna=True)
    any_risk_appetite_warning = bool((event_data["Risk_Appetite_Score"].fillna(0) >= 1).any())
    any_risk_off = bool((event_data["Risk_Appetite_Regime"] == "Risk-Off").any())

    any_rates_inflation_pressure = bool(
        (
            event_data["Rate_Pressure"].apply(normalize_bool)
            | event_data["Inflation_Concern"].apply(normalize_bool)
            | event_data["Real_Yield_Tightening"].apply(normalize_bool)
        ).any()
    )

    if focus == "Risk Appetite_RISK_OFF":
        if any_risk_off:
            return "No Miss"
        if any_risk_appetite_warning:
            return "Partial Miss"
        return "Miss"

    if focus == "Risk Appetite_DETERIORATION":
        if any_risk_appetite_warning:
            return "No Miss"
        return "Miss"

    if focus == "Rates & Inflation_PRESSURE":
        if any_rates_inflation_pressure:
            return "No Miss"
        if max_score >= 1:
            return "Partial Miss"
        return "Miss"

    return "Not Applicable"


def select_peak_stress_row(
    event_data: pd.DataFrame,
    focus: str,
) -> tuple[pd.Timestamp, pd.Series]:
    """
    Select the peak stress row for each event.

    Risk Appetite-focused events:
    - peak = highest Risk Appetite score.

    Rates & Inflation-focused events:
    - peak = highest Rates & Inflation pressure score:
      Rate_Pressure + Inflation_Concern + Real_Yield_Tightening.
    """
    if event_data.empty:
        raise ValueError("event_data is empty. Cannot select peak stress row.")

    if focus in ["Risk Appetite_RISK_OFF", "Risk Appetite_DETERIORATION"]:
        score_series = event_data["Risk_Appetite_Score"].fillna(-1)
        peak_date = score_series.idxmax()
        return peak_date, event_data.loc[peak_date]

    if focus == "Rates & Inflation_PRESSURE":
        pressure_score = (
            event_data["Rate_Pressure"].apply(normalize_bool).astype(int)
            + event_data["Inflation_Concern"].apply(normalize_bool).astype(int)
            + event_data["Real_Yield_Tightening"].apply(normalize_bool).astype(int)
        )

        peak_date = pressure_score.idxmax()
        return peak_date, event_data.loc[peak_date]

    raise ValueError(f"Invalid stress-event focus: {focus}")


def build_peak_drivers(peak_row: pd.Series) -> str:
    """
    Build a concise list of active drivers at the peak stress date.

    This explains why the framework detected stress instead of only showing
    the final regime label.
    """
    drivers = []

    risk_appetite_driver_map = {
        "Credit_Flag": "Credit",
        "NFCI_Flag": "Financial Conditions",
        "VIX_Flag": "Volatility",
        "Claims_Flag": "Labor",
    }

    for column, driver_name in risk_appetite_driver_map.items():
        if column in peak_row and normalize_bool(peak_row.get(column)):
            drivers.append(driver_name)

    if "Rate_Pressure" in peak_row and normalize_bool(peak_row.get("Rate_Pressure")):
        drivers.append("Rate Pressure")

    if (
        "Inflation_Concern" in peak_row
        and normalize_bool(peak_row.get("Inflation_Concern"))
    ):
        drivers.append("Inflation Concern")

    if (
        "Real_Yield_Tightening" in peak_row
        and normalize_bool(peak_row.get("Real_Yield_Tightening"))
    ):
        drivers.append("Real Yield Tightening")

    if not drivers:
        return "None"

    return ", ".join(dict.fromkeys(drivers))


def build_diagnostic_note(
    miss_status: str,
    already_stressed_at_start: str,
    first_warning_date,
    response_lag_weeks,
    peak_row: pd.Series | None,
    peak_drivers: str,
    real_yield_seen_in_window: bool,
    real_yield_at_peak: bool,
) -> str:
    """
    Build concise diagnostic note for the stress-event sheet.
    """
    if miss_status == "Not Applicable":
        return "No valid framework observations were available in this event window."

    if already_stressed_at_start in ["Yes", "Partial"]:
        warning_text = (
            f"Framework was already stressed before the event window "
            f"({already_stressed_at_start})."
        )
    elif first_warning_date == "N/A":
        warning_text = "No framework warning was observed during the event window."
    else:
        warning_text = (
            f"First warning appeared on {first_warning_date} "
            f"after {response_lag_weeks} weeks."
        )

    if peak_row is None:
        peak_text = "Peak regime information unavailable."
    else:
        peak_score = (
            int(peak_row["Risk_Appetite_Score"])
            if pd.notna(peak_row["Risk_Appetite_Score"])
            else "N/A"
        )

        peak_text = (
            f"Peak Risk Appetite={peak_row['Risk_Appetite_Regime']} "
            f"(score {peak_score}), "
            f"Rates & Inflation={peak_row['Rates_Inflation_Regime']}, "
            f"Equity={peak_row['Equity_Stance']}, "
            f"Cash={peak_row['Cash_Stance']}."
        )

    ry_window_text = (
        "Real Yield Tightening appeared during the window."
        if real_yield_seen_in_window
        else "Real Yield Tightening did not appear during the window."
    )

    ry_peak_text = (
        "Real Yield Tightening was ON at peak."
        if real_yield_at_peak
        else "Real Yield Tightening was OFF at peak."
    )

    driver_text = f"Peak drivers: {peak_drivers}."

    return (
        f"{warning_text} {peak_text} {driver_text} "
        f"{ry_window_text} {ry_peak_text} Miss status: {miss_status}."
    )


def build_stress_event_diagnostics(
    risk_appetite_output: pd.DataFrame,
    rates_inflation_output: pd.DataFrame,
    portfolio_output: pd.DataFrame,
    events: list[dict] | None = None,
) -> pd.DataFrame:
    """
    Build historical stress-event false-negative / regime-response diagnostic.

    This diagnostic asks:
    - Did the framework produce the expected risk/rates/inflation warning?
    - How long did it take?
    - What was the peak regime response?
    - Which drivers were active at peak?

    This is not a return backtest.
    """
    if events is None:
        events = get_default_stress_events()

    diagnostic_data = prepare_stress_diagnostic_data(
        risk_appetite_output=risk_appetite_output,
        rates_inflation_output=rates_inflation_output,
        portfolio_output=portfolio_output,
    )

    rows = []

    for event in events:
        event_name = event["event"]
        window_start = pd.Timestamp(event["window_start"])
        window_end = pd.Timestamp(event["window_end"])
        focus = event["focus"]

        window_mask = (
            (diagnostic_data.index >= window_start)
            & (diagnostic_data.index <= window_end)
        )

        event_data = diagnostic_data[window_mask].copy()

        if not event_data.empty:
            valid_mask = build_valid_mask(event_data["Portfolio_Router_Valid"])
            valid_event_data = event_data[valid_mask].copy()
        else:
            valid_event_data = event_data

        pre_event_data = diagnostic_data[
            (diagnostic_data.index < window_start)
            & build_valid_mask(diagnostic_data["Portfolio_Router_Valid"])
        ]

        if not pre_event_data.empty:
            pre_row = pre_event_data.iloc[-1]
            pre_risk_appetite_regime = pre_row["Risk_Appetite_Regime"]
            pre_rates_inflation_regime = pre_row["Rates_Inflation_Regime"]
            pre_equity_stance = pre_row["Equity_Stance"]
            pre_cash_stance = pre_row["Cash_Stance"]
            already_stressed_at_start = classify_pre_existing_stress(
                pre_row=pre_row,
                focus=focus,
            )
        else:
            pre_row = None
            pre_risk_appetite_regime = "N/A"
            pre_rates_inflation_regime = "N/A"
            pre_equity_stance = "N/A"
            pre_cash_stance = "N/A"
            already_stressed_at_start = "N/A"

        if valid_event_data.empty:
            rows.append(
                {
                    "Event": event_name,
                    "Window_Start": window_start.date(),
                    "Window_End": window_end.date(),
                    "Stress_Type": event["stress_type"],
                    "Expected_Response": event["expected_response"],
                    "First_Warning_Rule": event["first_warning_rule"],
                    "Framework_Valid": "No",
                    "Pre_Risk_Appetite_Regime": pre_risk_appetite_regime,
                    "Pre_Rates_Inflation_Regime": pre_rates_inflation_regime,
                    "Pre_Equity_Stance": pre_equity_stance,
                    "Pre_Cash_Stance": pre_cash_stance,
                    "Already_Stressed_At_Window_Start": already_stressed_at_start,
                    "First_Warning_Date": "N/A",
                    "Response_Lag_Weeks": "N/A",
                    "Peak_Stress_Date": "N/A",
                    "Peak_Risk_Appetite_Regime": "N/A",
                    "Peak_Risk_Appetite_Score": "N/A",
                    "Peak_Rates_Inflation_Regime": "N/A",
                    "Peak_Equity_Stance": "N/A",
                    "Peak_Cash_Stance": "N/A",
                    "Peak_Drivers": "N/A",
                    "Real_Yield_Tightening_Seen_In_Window": "N/A",
                    "Real_Yield_Tightening_At_Peak": "N/A",
                    "Miss_Status": "Not Applicable",
                    "Diagnostic_Note": (
                        "No valid Portfolio Router observations were available during this event window."
                    ),
                }
            )
            continue

        warning_mask = build_warning_mask(valid_event_data, focus)
        warning_data = valid_event_data[warning_mask]

        if already_stressed_at_start in ["Yes", "Partial"]:
            first_warning_date = "Pre-existing"
            response_lag_weeks = 0.0
        elif warning_data.empty:
            first_warning_date = "N/A"
            response_lag_weeks = "N/A"
        else:
            first_warning_timestamp = warning_data.index.min()
            first_warning_date = first_warning_timestamp.date()
            response_lag_weeks = round(
                (first_warning_timestamp - window_start).days / 7,
                1,
            )

        peak_date, peak_row = select_peak_stress_row(
            valid_event_data,
            focus,
        )

        peak_drivers = build_peak_drivers(peak_row)

        real_yield_seen_in_window = bool(
            valid_event_data["Real_Yield_Tightening"]
            .apply(normalize_bool)
            .any()
        )

        real_yield_at_peak = normalize_bool(
            peak_row.get("Real_Yield_Tightening")
        )

        miss_status = classify_miss_status(valid_event_data, focus)

        diagnostic_note = build_diagnostic_note(
            miss_status=miss_status,
            already_stressed_at_start=already_stressed_at_start,
            first_warning_date=first_warning_date,
            response_lag_weeks=response_lag_weeks,
            peak_row=peak_row,
            peak_drivers=peak_drivers,
            real_yield_seen_in_window=real_yield_seen_in_window,
            real_yield_at_peak=real_yield_at_peak,
        )

        rows.append(
            {
                "Event": event_name,
                "Window_Start": window_start.date(),
                "Window_End": window_end.date(),
                "Stress_Type": event["stress_type"],
                "Expected_Response": event["expected_response"],
                "First_Warning_Rule": event["first_warning_rule"],
                "Framework_Valid": "Yes",
                "Pre_Risk_Appetite_Regime": pre_risk_appetite_regime,
                "Pre_Rates_Inflation_Regime": pre_rates_inflation_regime,
                "Pre_Equity_Stance": pre_equity_stance,
                "Pre_Cash_Stance": pre_cash_stance,
                "Already_Stressed_At_Window_Start": already_stressed_at_start,
                "First_Warning_Date": first_warning_date,
                "Response_Lag_Weeks": response_lag_weeks,
                "Peak_Stress_Date": peak_date.date(),
                "Peak_Risk_Appetite_Regime": peak_row["Risk_Appetite_Regime"],
                "Peak_Risk_Appetite_Score": (
                    int(peak_row["Risk_Appetite_Score"])
                    if pd.notna(peak_row["Risk_Appetite_Score"])
                    else "N/A"
                ),
                "Peak_Rates_Inflation_Regime": peak_row["Rates_Inflation_Regime"],
                "Peak_Equity_Stance": peak_row["Equity_Stance"],
                "Peak_Cash_Stance": peak_row["Cash_Stance"],
                "Peak_Drivers": peak_drivers,
                "Real_Yield_Tightening_Seen_In_Window": (
                    "Yes" if real_yield_seen_in_window else "No"
                ),
                "Real_Yield_Tightening_At_Peak": (
                    "Yes" if real_yield_at_peak else "No"
                ),
                "Miss_Status": miss_status,
                "Diagnostic_Note": diagnostic_note,
            }
        )

    return pd.DataFrame(rows)