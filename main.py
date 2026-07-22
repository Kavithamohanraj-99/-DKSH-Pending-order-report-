#!/usr/bin/env python3
from __future__ import annotations

import argparse
import logging
from datetime import date, datetime

from tc_pipeline import ORDER_ID_COLUMN_LETTER_DEFAULT, inspect_columns, run_pipeline

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Process a TC Report CSV into a 3-sheet workbook: "
        "Raw Data, Final Report, Pivot Summary."
    )
    parser.add_argument("input_csv", help="Path to the TC Report CSV")
    parser.add_argument(
        "-o",
        "--output",
        default="TC_Report_Processed.xlsx",
        help="Output .xlsx path (default: TC_Report_Processed.xlsx). If "
        "this file already exists, its Raw Data/Final Report/Pivot "
        "Summary sheets are overwritten; any other sheets are left alone.",
    )
    parser.add_argument(
        "--order-id-column",
        default=ORDER_ID_COLUMN_LETTER_DEFAULT,
        help=f"Spreadsheet letter for the order_id column, used only for "
        f"dedup (default: {ORDER_ID_COLUMN_LETTER_DEFAULT}) — the spec "
        f"doesn't state this explicitly, see README.",
    )
    parser.add_argument(
        "--today",
        default=None,
        help="ISO date to treat as 'today' for the Pivot Summary's "
        "red/green date-column coloring (default: live system date), "
        "e.g. --today 2026-07-16",
    )
    parser.add_argument(
        "--inspect",
        action="store_true",
        help="Don't run the pipeline — just print every expected column "
        "(letter position, expected name, actual header found, and 5 "
        "sample values) so you can eyeball a layout mismatch directly. "
        "Run this first against a new file.",
    )
    args = parser.parse_args()

    if args.inspect:
        inspect_columns(args.input_csv, order_id_letter=args.order_id_column)
        return

    today = date.fromisoformat(args.today) if args.today else None

    summary = run_pipeline(
        input_csv=args.input_csv,
        output_xlsx=args.output,
        order_id_column_letter=args.order_id_column,
        today=today,
    )

    print("\nDone.")
    for k, v in summary.items():
        if k in ("column_warnings", "order_item_status_value_counts"):
            continue  # printed separately below
        print(f"  {k}: {v}")

    if summary["column_warnings"]:
        print("\nColumn warnings:")
        for msg in summary["column_warnings"]:
            print(f"  ⚠ {msg}")

    if summary["after_status_filter"] == 0:
        print(
            f"\n0 rows survived the order-status filter (Step 2.1) out of "
            f"{summary['input_rows']} input rows."
        )
        print(f"Column read as order_item_status: {summary['order_item_status_column_resolved']}")
        if summary["order_item_status_value_counts"]:
            print("Actual values found, with counts:")
            for value, count in list(summary["order_item_status_value_counts"].items())[:10]:
                print(f"  {value!r}: {count}")


if __name__ == "__main__":
    main()
