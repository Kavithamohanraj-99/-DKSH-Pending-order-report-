#!/usr/bin/env python3
from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))

from tc_report_processor.pipeline import run_pipeline  # noqa: E402

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Process a TC Report CSV into a 4-sheet workbook: "
        "Raw Data, New Status Alert, Final Report, Pivot Summary."
    )
    parser.add_argument("input_csv", help="Path to the TC Report CSV")
    parser.add_argument(
        "-o",
        "--output",
        default="TC_Report_Processed.xlsx",
        help="Output .xlsx path (default: TC_Report_Processed.xlsx)",
    )
    parser.add_argument(
        "--marketplace-column",
        default=None,
        help="Header name for the marketplace channel column, if it isn't "
        "literally 'marketplace_channel' in your file (see columns.py notes "
        "on why this is needed for the Part 3 pivot).",
    )
    parser.add_argument(
        "--now",
        default=None,
        help="ISO datetime to use as 'current time' for the New Status "
        "Alert cutoff (default: live system clock). Useful for backfills "
        "or reproducible test runs, e.g. --now 2026-07-16T15:00:00",
    )
    args = parser.parse_args()

    now = datetime.fromisoformat(args.now) if args.now else None

    summary = run_pipeline(
        input_csv=args.input_csv,
        output_xlsx=args.output,
        marketplace_channel_column=args.marketplace_column,
        now=now,
    )

    print("\nDone.")
    for k, v in summary.items():
        print(f"  {k}: {v}")


if __name__ == "__main__":
    main()
