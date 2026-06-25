"""
fetch_data.py

Data fetching module for the Macro Regime-to-Portfolio Decision Framework.

Purpose:
- Read FRED series IDs from config.yaml.
- Fetch required Risk Appetite and Rates & Inflation data from FRED.
- Optionally merge local historical HY OAS data.
- Prepare raw time-series data for weekly alignment.

Data series:

Risk Appetite:
- HY OAS: BAMLH0A0HYM2
- NFCI: NFCI
- VIX: VIXCLS
- Initial Claims: ICSA

Rates & Inflation:
- 10Y Treasury Yield: DGS10
- 10Y Breakeven Inflation: T10YIE
- 10Y Real Yield: DFII10
- 2Y Treasury Yield: DGS2
"""
import time
import os
from pathlib import Path

import pandas as pd
import yaml
from fredapi import Fred


def load_config(config_path: str = "config.yaml") -> dict:
    """
    Load project configuration from config.yaml.
    """
    path = Path(config_path)

    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {path.resolve()}")

    with path.open("r", encoding="utf-8") as file:
        config = yaml.safe_load(file)

    if not config:
        raise ValueError(f"Config file is empty or invalid: {path.resolve()}")

    return config


def fetch_fred_series(
    series_id: str,
    max_retries: int = 5,
    initial_wait_seconds: int = 10,
) -> pd.Series:
    """
    Fetch one FRED series with simple retry/backoff logic.

    Handles temporary FRED API rate-limit errors such as:
    - HTTP Error 429
    - "Too Many Requests"
    - "Exceeded Rate Limit"

    This protects the weekly workflow from failing when FRED temporarily
    rejects rapid repeated requests.
    """
    api_key = os.getenv("FRED_API_KEY")

    if api_key is None:
        raise ValueError(
            "FRED_API_KEY environment variable is not set. "
            "Please set it before running the weekly framework."
        )

    fred = Fred(api_key=api_key)

    wait_seconds = initial_wait_seconds

    for attempt in range(1, max_retries + 1):
        try:
            series = fred.get_series(series_id)
            series.name = series_id
            return series

        except ValueError as error:
            error_message = str(error)

            is_rate_limit_error = (
                "Too Many Requests" in error_message
                or "Exceeded Rate Limit" in error_message
                or "429" in error_message
            )

            if is_rate_limit_error and attempt < max_retries:
                print(
                    f"FRED rate limit hit for {series_id}. "
                    f"Retrying in {wait_seconds} seconds "
                    f"({attempt}/{max_retries})..."
                )

                time.sleep(wait_seconds)
                wait_seconds *= 2
                continue

            raise

        except Exception as error:
            if attempt < max_retries:
                print(
                    f"Temporary fetch error for {series_id}: {error}. "
                    f"Retrying in {wait_seconds} seconds "
                    f"({attempt}/{max_retries})..."
                )

                time.sleep(wait_seconds)
                wait_seconds *= 2
                continue

            raise


def convert_dates_to_previous_friday(dates: pd.Series) -> pd.Series:
    """
    Convert dates to the latest Friday on or before each date.

    Used for local historical HY OAS data that was downloaded from FRED
    using weekly frequency with a Saturday week-ending label.

    Example:
    - Saturday 2026-01-10 -> Friday 2026-01-09

    This is an anchor conversion, not a signal transformation.
    """
    dates = pd.to_datetime(dates)
    days_since_friday = (dates.dt.weekday - 4) % 7

    return dates - pd.to_timedelta(days_since_friday, unit="D")


def load_local_hy_oas_history(config: dict) -> pd.Series | None:
    """
    Optionally load local historical HY OAS data from a user-provided file.

    This supports longer-history diagnostics without including the raw
    historical HY OAS file in the public repository.

    Expected local file:
    - data/local/Risk_Appetite.xlsx

    Expected sheet:
    - RAW_HY

    Expected columns:
    - Date
    - Value

    If local history is disabled or the file is missing, return None.
    """
    local_config = config.get("local_data", {})

    use_local_history = bool(
        local_config.get("use_local_hy_oas_history", False)
    )

    if not use_local_history:
        return None

    file_path = Path(
        local_config.get(
            "hy_oas_history_file",
            "data/local/Risk_Appetite.xlsx",
        )
    )

    if not file_path.exists():
        print(
            f"Local HY OAS history file not found: {file_path}. "
            "Using FRED-only HY OAS."
        )
        return None

    sheet_name = local_config.get("hy_oas_history_sheet", "RAW_HY")
    date_column = local_config.get("hy_oas_date_column", "Date")
    value_column = local_config.get("hy_oas_value_column", "Value")
    anchor = local_config.get("hy_oas_anchor", "previous_friday")

    local_data = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        usecols=[date_column, value_column],
    )

    local_data = local_data.dropna(subset=[date_column, value_column]).copy()

    local_data[date_column] = pd.to_datetime(local_data[date_column])
    local_data[value_column] = pd.to_numeric(
        local_data[value_column],
        errors="coerce",
    )

    local_data = local_data.dropna(subset=[value_column]).copy()

    if anchor == "previous_friday":
        local_data[date_column] = convert_dates_to_previous_friday(
            local_data[date_column]
        )
    elif anchor == "as_reported":
        pass
    else:
        raise ValueError(
            f"Invalid hy_oas_anchor: {anchor}. "
            "Expected 'previous_friday' or 'as_reported'."
        )

    local_series = (
        local_data
        .drop_duplicates(subset=[date_column], keep="last")
        .set_index(date_column)[value_column]
        .sort_index()
    )

    local_series.name = "hy_oas"

    if local_series.empty:
        print(
            f"Local HY OAS history file was found at {file_path}, "
            "but no valid observations were loaded. Using FRED-only HY OAS."
        )
        return None

    print(
        "Loaded local HY OAS history: "
        f"{len(local_series)} observations, "
        f"{local_series.index.min().date()} to {local_series.index.max().date()}."
    )

    return local_series


def merge_local_hy_oas_history(raw_data: pd.DataFrame, config: dict) -> pd.DataFrame:
    """
    Merge optional local historical HY OAS data with FRED HY OAS data.

    Priority rule:
    - Local HY OAS extends the historical sample.
    - FRED HY OAS takes priority on overlapping exact dates because it is
      the latest directly fetched source.
    - If no local file exists, raw_data is returned unchanged.
    """
    local_hy_oas = load_local_hy_oas_history(config)

    if local_hy_oas is None:
        return raw_data

    raw_data = raw_data.copy()

    if "hy_oas" not in raw_data.columns:
        raw_data["hy_oas"] = pd.NA

    fred_hy_oas = raw_data["hy_oas"].dropna().copy()
    fred_hy_oas.name = "hy_oas"

    combined_hy_oas = pd.concat(
        [
            local_hy_oas,
            fred_hy_oas,
        ]
    )

    combined_hy_oas = combined_hy_oas.sort_index()
    combined_hy_oas = combined_hy_oas[
        ~combined_hy_oas.index.duplicated(keep="last")
    ]
    combined_hy_oas.name = "hy_oas"

    raw_without_hy = raw_data.drop(columns=["hy_oas"])

    merged_data = raw_without_hy.join(combined_hy_oas, how="outer")
    merged_data = merged_data.sort_index()

    ordered_columns = ["hy_oas"] + [
        column for column in raw_data.columns if column != "hy_oas"
    ]

    merged_data = merged_data[ordered_columns]

    print(
        "Merged HY OAS history: "
        f"{combined_hy_oas.index.min().date()} to "
        f"{combined_hy_oas.index.max().date()}."
    )

    return merged_data


def fetch_all_series(config: dict) -> pd.DataFrame:
    """
    Fetch all Risk Appetite and Rates & Inflation series defined in config.yaml.

    Returns:
        DataFrame containing all raw FRED series.
        Index is DatetimeIndex.
        Columns use internal project names:
        - hy_oas
        - nfci
        - vix
        - claims
        - dgs10
        - t10yie
        - dfii10
        - dgs2
    """
    series_groups = config["fred_series"]

    fetched_series = {}

    for group_name, group_series in series_groups.items():
        _ = group_name

        for internal_name, metadata in group_series.items():
            series_id = metadata["id"]

            print(f"Fetching {internal_name}: {series_id}")
            fetched_series[internal_name] = fetch_fred_series(series_id)

            # Small pause to reduce the chance of hitting FRED rate limits.
            time.sleep(1)

    raw_data = pd.concat(fetched_series.values(), axis=1)
    raw_data.columns = fetched_series.keys()
    raw_data = raw_data.sort_index()

    raw_data = merge_local_hy_oas_history(raw_data, config)

    return raw_data


def align_to_weekly_friday(raw_data: pd.DataFrame) -> pd.DataFrame:
    """
    Align raw series to week-ending Friday.

    As-of convention:
    AsOfDate = latest completed Friday after resampling all series to
    week-ending Friday.

    Method:
    - Sort by date.
    - Resample all series to weekly Friday frequency.
    - Use the last available observation within each week.
    - Forward-fill missing values after weekly resampling.
    - Drop incomplete future Friday buckets.
    """
    if raw_data is None or raw_data.empty:
        raise ValueError("raw_data is empty. Cannot align to weekly Friday.")

    raw_data = raw_data.sort_index()

    latest_raw_date = pd.Timestamp(raw_data.index.max()).normalize()
    today = pd.Timestamp.today().normalize()

    reference_date = min(latest_raw_date, today)

    # Python weekday: Monday=0, Tuesday=1, ..., Friday=4, Saturday=5, Sunday=6.
    days_since_friday = (reference_date.weekday() - 4) % 7
    latest_completed_friday = reference_date - pd.Timedelta(days=days_since_friday)

    weekly_data = raw_data.resample("W-FRI").last()
    weekly_data = weekly_data[weekly_data.index <= latest_completed_friday]
    weekly_data = weekly_data.ffill()
    weekly_data.index.name = "AsOfDate"

    if weekly_data.empty:
        raise ValueError("weekly_data is empty after completed-Friday filtering.")

    return weekly_data


def add_signal_features(weekly_data: pd.DataFrame, config: dict) -> pd.DataFrame:
    """
    Add rolling percentile and 4-week change features to weekly data.

    Definitions:
    - rolling P80 = 52-week rolling 80th percentile.
    - d4w = current weekly value minus value 4 weekly observations ago.
    - All calculations are performed after resampling to week-ending Friday.
    """
    if weekly_data is None or weekly_data.empty:
        raise ValueError("weekly_data is empty. Cannot add signal features.")

    window = config["parameters"]["rolling_window_weeks"]
    threshold = config["parameters"]["percentile_threshold"]

    if threshold != 0.80:
        raise ValueError(
            f"Expected percentile_threshold = 0.80, but got {threshold}."
        )

    feature_data = weekly_data.copy()

    for column in weekly_data.columns:
        feature_data[f"{column}_p80"] = (
            weekly_data[column]
            .rolling(window=window, min_periods=window)
            .quantile(threshold)
        )

        feature_data[f"{column}_d4w"] = (
            weekly_data[column] - weekly_data[column].shift(4)
        )

    return feature_data