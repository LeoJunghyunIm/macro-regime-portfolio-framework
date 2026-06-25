"""Export the two memo worksheets as a single two-page PDF.

PDF export requires Windows, Microsoft Excel, and pywin32. The import is
performed lazily so the rest of the framework can be imported and tested on
other operating systems.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

DEFAULT_WORKBOOK = Path("outputs/Macro_Regime_Framework_Output.xlsx")
DEFAULT_PDF = Path("outputs/Macro_Regime_Framework_Memo.pdf")
PDF_SHEETS = ["PDF_MEMO_1", "PDF_MEMO_2"]


def export_memo_pdf(
    workbook_path: str | Path = DEFAULT_WORKBOOK,
    pdf_path: str | Path = DEFAULT_PDF,
) -> None:
    """Export the two memo sheets as one PDF using Microsoft Excel."""
    if sys.platform != "win32":
        raise RuntimeError(
            "PDF export requires Windows and a locally installed copy of "
            "Microsoft Excel. Run the weekly workflow with --skip-pdf on "
            "other operating systems."
        )

    try:
        import win32com.client
    except ImportError as error:
        raise RuntimeError(
            "PDF export requires pywin32. Install project requirements and "
            "retry on Windows."
        ) from error

    workbook_path = Path(workbook_path).resolve()
    pdf_path = Path(pdf_path).resolve()

    if not workbook_path.exists():
        raise FileNotFoundError(f"Excel workbook not found: {workbook_path}")

    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    excel = None
    workbook = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        workbook = excel.Workbooks.Open(
            str(workbook_path),
            UpdateLinks=0,
            ReadOnly=True,
        )

        available_sheets = {
            workbook.Worksheets(index).Name
            for index in range(1, workbook.Worksheets.Count + 1)
        }
        missing_sheets = [
            sheet_name
            for sheet_name in PDF_SHEETS
            if sheet_name not in available_sheets
        ]
        if missing_sheets:
            raise ValueError(
                "Missing PDF memo sheets: " + ", ".join(missing_sheets)
            )

        workbook.Worksheets(PDF_SHEETS[0]).Select()
        for sheet_name in PDF_SHEETS[1:]:
            workbook.Worksheets(sheet_name).Select(False)

        excel.ActiveSheet.ExportAsFixedFormat(
            Type=0,
            Filename=str(pdf_path),
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "workbook",
        nargs="?",
        default=str(DEFAULT_WORKBOOK),
        help="Path to the generated Excel workbook.",
    )
    parser.add_argument(
        "pdf",
        nargs="?",
        default=str(DEFAULT_PDF),
        help="Destination PDF path.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    arguments = parse_args()
    export_memo_pdf(arguments.workbook, arguments.pdf)
