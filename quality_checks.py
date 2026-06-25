"""
quality_checks.py

QA and Confidence module for the Macro Regime-to-Portfolio Decision Framework.

Purpose:
- Check whether the weekly output is operationally reliable.
- Identify missing, stale, insufficient, or invalid signal issues.
- Assign High / Medium / Low confidence.

QA checks:
- Missing latest value
- Series freshness / stale data
- 52-week rolling window sufficiency
- As-of date alignment
- Portfolio Router output availability check
- Latest Risk Appetite / Rates & Inflation / Portfolio Router signal validity check
- Post-export success check

Important:
- Confidence is not a Risk Appetite or Rates & Inflation signal.
- Confidence is assigned after Risk Appetite/Rates & Inflation/Portfolio Router calculations.
- Confidence measures data/output reliability, not investment conviction.
- Export success is a post-export check, because it cannot be confirmed before files are created.
"""

from pathlib import Path

import pandas as pd


def check_missing_latest_values(weekly_data) -> dict:
    """
    Check whether latest as-of row has missing values.

    Expected output:
    {
        "check": "missing_latest_value",
        "status": "PASS/WARN/FAIL",
        "details": "..."
    }
    """
    if weekly_data is None or weekly_data.empty:
        return {
            "check": "missing_latest_value",
            "status": "FAIL",
            "details": "weekly_data is empty.",
        }

    latest_date = weekly_data.index.max()
    latest_row = weekly_data.loc[latest_date]

    missing_columns = latest_row[latest_row.isna()].index.tolist()

    if missing_columns:
        return {
            "check": "missing_latest_value",
            "status": "FAIL",
            "details": (
                f"Latest row contains missing values on {latest_date.date()}: "
                f"{missing_columns}"
            ),
        }

    return {
        "check": "missing_latest_value",
        "status": "PASS",
        "details": f"No missing values in latest row on {latest_date.date()}.",
    }


def check_series_freshness(raw_data, as_of_date, config: dict) -> dict:
    """
    Check whether each series is stale relative to config thresholds.

    the current framework staleness thresholds:
    - Daily market series: max stale days = 7
    - Weekly macro series: max stale days = 14

    Expected output:
    {
        "check": "series_freshness",
        "status": "PASS/WARN/FAIL",
        "details": "..."
    }
    """
    if raw_data is None or raw_data.empty:
        return {
            "check": "series_freshness",
            "status": "FAIL",
            "details": "raw_data is empty.",
        }

    if as_of_date is None:
        return {
            "check": "series_freshness",
            "status": "FAIL",
            "details": "as_of_date is missing.",
        }

    as_of_date = pd.Timestamp(as_of_date)

    max_stale_days = config["staleness"]["max_stale_days"]
    frequency_groups = config["staleness"]["series_frequency_groups"]

    issues = []

    for group_name, series_names in frequency_groups.items():
        allowed_stale_days = int(max_stale_days[group_name])

        for series_name in series_names:
            if series_name not in raw_data.columns:
                issues.append(f"{series_name}: missing from raw_data.")
                continue

            last_valid_date = raw_data[series_name].dropna().index.max()

            if pd.isna(last_valid_date):
                issues.append(f"{series_name}: no valid observations.")
                continue

            stale_days = (as_of_date - pd.Timestamp(last_valid_date)).days

            if stale_days > allowed_stale_days:
                issues.append(
                    f"{series_name}: last valid date {last_valid_date.date()}, "
                    f"{stale_days} days stale; max allowed {allowed_stale_days}."
                )

    if issues:
        return {
            "check": "series_freshness",
            "status": "WARN",
            "details": " | ".join(issues),
        }

    return {
        "check": "series_freshness",
        "status": "PASS",
        "details": f"All tracked series are fresh as of {as_of_date.date()}.",
    }


def check_rolling_window_sufficiency(weekly_data, config: dict) -> dict:
    """
    Check whether each required Risk Appetite/Rates & Inflation signal series has enough non-null
    weekly observations to support 52-week rolling percentile calculations.

    the current framework requirement:
    - Each required signal series should have at least 52 non-null weekly
      observations by the latest AsOfDate.
    """
    if weekly_data is None or weekly_data.empty:
        return {
            "check": "rolling_window_sufficiency",
            "status": "FAIL",
            "details": "weekly_data is empty.",
        }

    required_window = int(config["parameters"]["rolling_window_weeks"])
    latest_date = weekly_data.index.max()

    required_series = [
        "hy_oas",
        "nfci",
        "vix",
        "claims",
        "dgs10",
        "t10yie",
        "dfii10",
    ]

    missing_series = [
        series for series in required_series if series not in weekly_data.columns
    ]

    if missing_series:
        return {
            "check": "rolling_window_sufficiency",
            "status": "FAIL",
            "details": f"Missing required weekly series: {missing_series}",
        }

    insufficient_series = []
    series_status = []

    for series in required_series:
        non_null_count = weekly_data[series].dropna().shape[0]

        if non_null_count < required_window:
            insufficient_series.append(
                f"{series}: {non_null_count}/{required_window}"
            )
            series_status.append(f"{series}: FAIL")
        else:
            series_status.append(f"{series}: PASS")

    if insufficient_series:
        return {
            "check": "rolling_window_sufficiency",
            "status": "FAIL",
            "details": (
                f"Insufficient non-null weekly observations as of {latest_date.date()}: "
                f"{'; '.join(insufficient_series)}."
            ),
        }

    return {
        "check": "rolling_window_sufficiency",
        "status": "PASS",
        "details": (
            f"All required Risk Appetite/Rates & Inflation series have at least {required_window} "
            f"non-null weekly observations as of {latest_date.date()}. "
            f"{'; '.join(series_status)}."
        ),
    }


def check_as_of_date_alignment(risk_appetite_output, rates_inflation_output) -> dict:
    """
    Check whether Risk Appetite and Rates & Inflation use the same AsOfDate.

    Expected output:
    {
        "check": "as_of_date_alignment",
        "status": "PASS/WARN/FAIL",
        "details": "..."
    }
    """
    if risk_appetite_output is None or rates_inflation_output is None:
        return {
            "check": "as_of_date_alignment",
            "status": "FAIL",
            "details": "Risk Appetite output or Rates & Inflation output is missing.",
        }

    if risk_appetite_output.empty or rates_inflation_output.empty:
        return {
            "check": "as_of_date_alignment",
            "status": "FAIL",
            "details": "Risk Appetite output or Rates & Inflation output is empty.",
        }

    risk_appetite_latest = risk_appetite_output.index.max()
    rates_inflation_latest = rates_inflation_output.index.max()

    if risk_appetite_latest != rates_inflation_latest:
        return {
            "check": "as_of_date_alignment",
            "status": "FAIL",
            "details": (
                f"Latest AsOfDate mismatch: Risk Appetite={risk_appetite_latest.date()}, "
                f"Rates & Inflation={rates_inflation_latest.date()}."
            ),
        }

    if not risk_appetite_output.index.equals(rates_inflation_output.index):
        return {
            "check": "as_of_date_alignment",
            "status": "WARN",
            "details": (
                f"Latest AsOfDate matches at {risk_appetite_latest.date()}, "
                "but full Risk Appetite/Rates & Inflation index history differs."
            ),
        }

    return {
        "check": "as_of_date_alignment",
        "status": "PASS",
        "details": f"Risk Appetite and Rates & Inflation AsOfDate indexes match through {risk_appetite_latest.date()}.",
    }


def check_portfolio_output_available(portfolio_output) -> dict:
    """
    Check whether Portfolio Router output exists and contains final stance fields.

    The final memo depends on Portfolio Router.

    Expected output:
    {
        "check": "portfolio_output_available",
        "status": "PASS/WARN/FAIL",
        "details": "..."
    }
    """
    if portfolio_output is None:
        return {
            "check": "portfolio_output_available",
            "status": "WARN",
            "details": "Portfolio Router output was not provided.",
        }

    if portfolio_output.empty:
        return {
            "check": "portfolio_output_available",
            "status": "FAIL",
            "details": "Portfolio Router output is empty.",
        }

    required_columns = [
        "Equity_Stance",
        "Duration_Stance",
        "Cash_Stance",
        "TIPS_Stance",
    ]

    missing_columns = [
        column for column in required_columns if column not in portfolio_output.columns
    ]

    if missing_columns:
        return {
            "check": "portfolio_output_available",
            "status": "FAIL",
            "details": f"Portfolio Router output is missing required columns: {missing_columns}",
        }

    latest_date = portfolio_output.index.max()

    return {
        "check": "portfolio_output_available",
        "status": "PASS",
        "details": f"Portfolio Router output is available through {latest_date.date()}.",
    }


def check_latest_signal_validity(risk_appetite_output, rates_inflation_output, portfolio_output) -> dict:
    """
    Check whether the latest Risk Appetite, Rates & Inflation, and Portfolio Router outputs are valid.

    This prevents the framework from releasing a portfolio stance when
    one or more regime engines have insufficient data.
    """
    if risk_appetite_output is None or risk_appetite_output.empty:
        return {
            "check": "latest_signal_validity",
            "status": "FAIL",
            "details": "Risk Appetite output is missing or empty.",
        }

    if rates_inflation_output is None or rates_inflation_output.empty:
        return {
            "check": "latest_signal_validity",
            "status": "FAIL",
            "details": "Rates & Inflation output is missing or empty.",
        }

    if portfolio_output is None or portfolio_output.empty:
        return {
            "check": "latest_signal_validity",
            "status": "FAIL",
            "details": "Portfolio Router output is missing or empty.",
        }

    required_risk_appetite_columns = ["Risk_Appetite_Valid"]
    required_rates_inflation_columns = ["Rates_Inflation_Valid"]
    required_portfolio_columns = ["Portfolio_Router_Valid"]

    missing_risk_appetite_columns = [
        column for column in required_risk_appetite_columns if column not in risk_appetite_output.columns
    ]
    missing_rates_inflation_columns = [
        column for column in required_rates_inflation_columns if column not in rates_inflation_output.columns
    ]
    missing_portfolio_columns = [
        column for column in required_portfolio_columns if column not in portfolio_output.columns
    ]

    missing_columns = missing_risk_appetite_columns + missing_rates_inflation_columns + missing_portfolio_columns

    if missing_columns:
        return {
            "check": "latest_signal_validity",
            "status": "FAIL",
            "details": f"Missing validity columns: {missing_columns}",
        }

    latest_date = portfolio_output.index.max()

    if latest_date not in risk_appetite_output.index or latest_date not in rates_inflation_output.index:
        return {
            "check": "latest_signal_validity",
            "status": "FAIL",
            "details": "Latest Portfolio Router AsOfDate is not available in Risk Appetite or Rates & Inflation output.",
        }

    risk_appetite_valid = bool(risk_appetite_output.loc[latest_date, "Risk_Appetite_Valid"])
    rates_inflation_valid = bool(rates_inflation_output.loc[latest_date, "Rates_Inflation_Valid"])
    portfolio_valid = bool(portfolio_output.loc[latest_date, "Portfolio_Router_Valid"])

    if not risk_appetite_valid or not rates_inflation_valid or not portfolio_valid:
        return {
            "check": "latest_signal_validity",
            "status": "FAIL",
            "details": (
                f"Latest signal validity failed as of {latest_date.date()}. "
                f"Risk_Appetite_Valid={risk_appetite_valid}, "
                f"Rates_Inflation_Valid={rates_inflation_valid}, "
                f"Portfolio_Router_Valid={portfolio_valid}."
            ),
        }

    return {
        "check": "latest_signal_validity",
        "status": "PASS",
        "details": (
            f"Latest Risk Appetite, Rates & Inflation, and Portfolio Router outputs are valid as of {latest_date.date()}."
        ),
    }


def check_export_success(output_paths: dict) -> dict:
    """
    Post-export check: confirm whether expected output files were created.

    Expected files:
    - outputs/Macro_Regime_Framework_Output.xlsx
    - outputs/weekly_memo.txt

    Important:
    This check should run after export.py creates the output files.

    Expected output:
    {
        "check": "export_success",
        "status": "PASS/WARN/FAIL",
        "details": "..."
    }
    """
    if not output_paths:
        return {
            "check": "export_success",
            "status": "FAIL",
            "details": "output_paths is missing or empty.",
        }

    missing_files = []

    for label, path in output_paths.items():
        file_path = Path(path)

        if not file_path.exists():
            missing_files.append(f"{label}: {path}")

    if missing_files:
        return {
            "check": "export_success",
            "status": "FAIL",
            "details": f"Missing expected output files: {missing_files}",
        }

    return {
        "check": "export_success",
        "status": "PASS",
        "details": "All expected output files were created successfully.",
    }


def assign_confidence(qa_results: list[dict]) -> str:
    """
    Assign High / Medium / Low confidence based on pre-export QA results.

    the current framework logic:
    - High: all major checks PASS.
    - Medium: one or more WARN issues, no FAIL issues.
    - Low: one or more serious FAIL issues.

    Note:
    Export success is important operationally, but confidence should mainly
    reflect data reliability before the memo decision is generated.
    """
    if not qa_results:
        return "Low"

    statuses = [result.get("status") for result in qa_results]

    if "FAIL" in statuses:
        return "Low"

    if "WARN" in statuses:
        return "Medium"

    return "High"


def run_pre_export_qa_checks(
    raw_data,
    weekly_data,
    risk_appetite_output,
    rates_inflation_output,
    portfolio_output=None,
    config: dict | None = None,
):
    """
    Run pre-export QA checks and assign confidence.

    Pre-export QA checks:
    - Missing latest value
    - Series freshness
    - Rolling window sufficiency
    - As-of date alignment
    - Portfolio Router output availability
    - Latest Risk Appetite / Rates & Inflation / Portfolio Router signal validity

    Expected output:
    {
        "qa_results": [...],
        "confidence": "High/Medium/Low"
    }
    """
    if config is None:
        raise ValueError("config is required for pre-export QA checks.")

    if weekly_data is None or weekly_data.empty:
        qa_results = [
            {
                "check": "pre_export_qa",
                "status": "FAIL",
                "details": "weekly_data is empty.",
            }
        ]
        return {
            "qa_results": qa_results,
            "confidence": assign_confidence(qa_results),
        }

    as_of_date = weekly_data.index.max()

    qa_results = [
        check_missing_latest_values(weekly_data),
        check_series_freshness(raw_data, as_of_date, config),
        check_rolling_window_sufficiency(weekly_data, config),
        check_as_of_date_alignment(risk_appetite_output, rates_inflation_output),
    ]

    if portfolio_output is not None:
        qa_results.append(check_portfolio_output_available(portfolio_output))

        qa_results.append(
            check_latest_signal_validity(
                risk_appetite_output=risk_appetite_output,
                rates_inflation_output=rates_inflation_output,
                portfolio_output=portfolio_output,
            )
        )

    confidence = assign_confidence(qa_results)

    return {
        "qa_results": qa_results,
        "confidence": confidence,
    }


def run_post_export_checks(output_paths: dict) -> dict:
    """
    Run post-export checks after files are created.

    Post-export checks:
    - Export success

    Expected output:
    {
        "post_export_results": [...]
    }
    """
    post_export_results = [
        check_export_success(output_paths)
    ]

    return {
        "post_export_results": post_export_results
    }