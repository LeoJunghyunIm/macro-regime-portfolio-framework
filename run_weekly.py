"""Run the complete Macro Regime-to-Portfolio workflow.

The workflow fetches FRED data, builds the two regime engines, routes the
portfolio stance, runs data-quality checks, exports the Excel workbook and
text memo, and optionally exports a two-page PDF through Microsoft Excel.
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

from export import run_export
from export_pdf import export_memo_pdf
from fetch_data import (
    add_signal_features,
    align_to_weekly_friday,
    fetch_all_series,
    load_config,
)
from portfolio_router import run_portfolio_router
from quality_checks import run_post_export_checks, run_pre_export_qa_checks
from rates_inflation_engine import run_rates_inflation
from risk_appetite_engine import run_risk_appetite


def archive_pdf_output(pdf_path: str | Path, as_of_date) -> Path:
    """Copy the current PDF into the date-stamped local archive."""
    source = Path(pdf_path)
    if not source.exists():
        raise FileNotFoundError(f"PDF memo was not created: {source}")

    date_text = as_of_date.date().isoformat()
    archive_directory = Path("outputs") / "archive" / date_text
    archive_directory.mkdir(parents=True, exist_ok=True)
    destination = archive_directory / f"Macro_Regime_Framework_Memo_{date_text}.pdf"
    shutil.copy2(source, destination)
    return destination


def run_workflow(config_path: str | Path, skip_pdf: bool = False) -> dict:
    """Run the weekly workflow and return generated output paths."""
    print("Starting Macro Regime-to-Portfolio Decision Framework...")

    print("Step 1: Load configuration")
    config = load_config(str(config_path))

    print("Step 2: Fetch raw FRED data")
    raw_data = fetch_all_series(config)

    print("Step 3: Align data to week-ending Friday")
    weekly_data = align_to_weekly_friday(raw_data)

    print("Step 4: Add signal features")
    signal_data = add_signal_features(weekly_data, config)

    print("Step 5: Run Risk Appetite Engine")
    risk_appetite_output = run_risk_appetite(signal_data, config)

    print("Step 6: Run Rates & Inflation Engine")
    rates_inflation_output = run_rates_inflation(signal_data, config)

    print("Step 7: Run Portfolio Router")
    portfolio_output = run_portfolio_router(
        risk_appetite_output,
        rates_inflation_output,
        config,
    )

    print("Step 8: Run pre-export data-quality checks")
    quality_output = run_pre_export_qa_checks(
        raw_data=raw_data,
        weekly_data=weekly_data,
        risk_appetite_output=risk_appetite_output,
        rates_inflation_output=rates_inflation_output,
        portfolio_output=portfolio_output,
        config=config,
    )

    print("Step 9: Export Excel workbook and text memo")
    output_paths = run_export(
        raw_data=raw_data,
        weekly_data=weekly_data,
        signal_data=signal_data,
        risk_appetite_output=risk_appetite_output,
        rates_inflation_output=rates_inflation_output,
        portfolio_output=portfolio_output,
        qa_output=quality_output,
        config=config,
    )

    latest_date = portfolio_output.index.max()
    pdf_path = Path(config["outputs"].get("pdf_file", "outputs/Macro_Regime_Framework_Memo.pdf"))

    should_export_pdf = not skip_pdf and sys.platform == "win32"
    if should_export_pdf:
        print("Step 10: Export two-page PDF memo")
        export_memo_pdf(output_paths["excel_file"], pdf_path)
        output_paths["pdf_file"] = str(pdf_path)
        output_paths["pdf_archive"] = str(
            archive_pdf_output(pdf_path, latest_date)
        )
    elif skip_pdf:
        print("Step 10: PDF export skipped by command-line option")
    else:
        print("Step 10: PDF export skipped; Windows and Microsoft Excel are required")

    print("Step 11: Run post-export checks")
    post_export_output = run_post_export_checks(output_paths)

    latest_risk = risk_appetite_output.loc[latest_date]
    latest_rates = rates_inflation_output.loc[latest_date]
    latest_portfolio = portfolio_output.loc[latest_date]
    real_yield_status = (
        "ON" if bool(latest_portfolio["Real_Yield_Tightening"]) else "OFF"
    )

    print("\nWeekly run completed successfully.")
    print(f"As of: {latest_date.date()}")
    print(
        "Risk Appetite: "
        f"{latest_risk['Risk_Appetite_Regime']} | "
        f"Score: {int(latest_risk['Risk_Appetite_Score'])}"
    )
    print(f"Rates & Inflation Regime: {latest_rates['Rates_Inflation_Regime']}")
    print(f"Real Yield Tightening: {real_yield_status}")
    print(f"Confidence: {quality_output['confidence']}")

    print("\nPortfolio stance:")
    for asset in ("Equity", "Duration", "Cash", "TIPS"):
        print(
            f"{asset}: {latest_portfolio[f'{asset}_Stance']} "
            f"({latest_portfolio[f'{asset}_Band']})"
        )

    print("\nOutput files:")
    for label, path in output_paths.items():
        print(f"{label}: {path}")

    print("\nPost-export checks:")
    for result in post_export_output["post_export_results"]:
        print(result)

    return output_paths


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--config",
        default="config.yaml",
        help="Path to the project configuration file.",
    )
    parser.add_argument(
        "--skip-pdf",
        action="store_true",
        help="Skip the Windows-only Excel PDF export step.",
    )
    return parser.parse_args()


def main() -> None:
    arguments = parse_args()
    run_workflow(arguments.config, skip_pdf=arguments.skip_pdf)


if __name__ == "__main__":
    main()
