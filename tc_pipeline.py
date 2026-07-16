"""tc_pipeline.py — everything needed to process a TC Report CSV into a
4-sheet workbook (Raw Data, New Status Alert, Final Report, Pivot Summary).

Deliberately a single flat file, not a package/subfolder. This repo
previously used a src/tc_report_processor/ package and then a top-level
tc_report_processor/ package, and both times the folder failed to make it
into the deployed GitHub repo (ModuleNotFoundError on Streamlit Cloud).
A lone file sitting next to app.py has nothing that can be "left out" —
if app.py is in the repo, this is too.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from enum import Enum
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from pandas.api.types import is_string_dtype

logger = logging.getLogger("tc_report_processor")

# ===========================================================================
# Column resolution — maps the spec's spreadsheet-letter columns (B, G, J,
# BI, BP) to whatever header text is actually at that position in the file.
# ===========================================================================

EXPECTED_COLUMNS: dict[str, str] = {
    "B": "order_number",
    "G": "payment_status",
    "J": "order_item_status",
    "BI": "payment_methods",
    "BP": "erp_reference_id",
}

# Columns retained in the Final Report (Part 2 of the spec), in order.
# NOTE: the spec's Part 3 pivot needs a "Marketplace Channel" row
# dimension, but that column isn't in Part 2's retained-column list — a
# gap in the source spec. This pipeline assumes marketplace_channel should
# also survive into the Final Report so the pivot can be built; override
# via the marketplace_channel_column argument if your file uses a
# different header.
FINAL_REPORT_COLUMNS: list[str] = [
    "order_id",
    "order_status",
    "payment_status",
    "courier_name",
    "tracking_number",
    "ordered_date",
    "accepted_date",
    "nickname",
    "payment_methods",
    "time_shippinglabel_printed",
    "erp_reference_id",
    "time_order_paid",
]

MARKETPLACE_CHANNEL_COLUMN_DEFAULT = "marketplace_channel"


def col_letter_to_index(letter: str) -> int:
    """'A' -> 0, 'B' -> 1, ..., 'Z' -> 25, 'AA' -> 26, 'BP' -> 67, etc."""
    letter = letter.strip().upper()
    if not re.fullmatch(r"[A-Z]+", letter):
        raise ValueError(f"Invalid column letter: {letter!r}")
    index = 0
    for ch in letter:
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index - 1


def _normalize(name: str) -> str:
    return re.sub(r"[\s_]+", "", name.strip().lower())


def resolve_columns(headers: list[str]) -> dict[str, str]:
    """Returns {logical_name: actual_header_text}, resolved by position
    against `headers`. Logs a warning (doesn't raise) if the header found
    there doesn't resemble the expected name."""
    resolved: dict[str, str] = {}
    for letter, expected_name in EXPECTED_COLUMNS.items():
        idx = col_letter_to_index(letter)
        if idx >= len(headers):
            raise ValueError(
                f"Expected column {letter} ({expected_name}) at position {idx}, "
                f"but the file only has {len(headers)} columns. Check the file "
                f"matches the TC Report template this pipeline was built for."
            )
        actual_header = headers[idx]
        if _normalize(actual_header) != _normalize(expected_name):
            logger.warning(
                "Column %s: expected header resembling '%s' but found '%s'. "
                "Proceeding with the actual header text found there — verify "
                "this is really the right column before trusting the output.",
                letter,
                expected_name,
                actual_header,
            )
        resolved[expected_name] = actual_header
    return resolved


# ===========================================================================
# Loading — read the CSV, keep order_number as text so Excel/pandas never
# coerce a long numeric id into scientific notation.
# ===========================================================================


def load_tc_report(csv_path: str) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, str]]:
    """Returns (raw_df, working_df, resolved_columns).
    raw_df: exact copy as read — becomes the untouched "Raw Data" sheet.
    working_df: a separate copy the rest of the pipeline mutates freely."""
    header_df = pd.read_csv(csv_path, nrows=0)
    headers = list(header_df.columns)
    resolved = resolve_columns(headers)

    order_number_header = resolved[EXPECTED_COLUMNS["B"]]
    erp_header = resolved[EXPECTED_COLUMNS["BP"]]

    raw_df = pd.read_csv(
        csv_path,
        dtype={order_number_header: str, erp_header: str},
        keep_default_na=True,
    )
    working_df = raw_df.copy(deep=True)
    return raw_df, working_df, resolved


# ===========================================================================
# Cleaning — trim whitespace, validate erp_reference_id (Step 2).
# ===========================================================================

ERP_ID_PATTERN = re.compile(r"^\d{11}$")


def trim_all_string_fields(df: pd.DataFrame) -> pd.DataFrame:
    """Trims leading/trailing whitespace on every string column. Handles
    both legacy `object` dtype and pandas's dedicated `StringDtype`
    (default for text columns as of pandas 3.0)."""
    df = df.copy()
    for col in df.columns:
        if is_string_dtype(df[col]):
            df[col] = df[col].apply(lambda v: v.strip() if isinstance(v, str) else v)
    return df


def is_blank(value) -> bool:
    if value is None:
        return True
    if value is pd.NA:
        return True
    try:
        if isinstance(value, float) and pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def classify_erp_reference(value) -> str:
    """'blank', 'valid' (exactly 11 digits), or 'invalid' (populated but
    not 11 digits)."""
    if is_blank(value):
        return "blank"
    stripped = str(value).strip()
    return "valid" if ERP_ID_PATTERN.fullmatch(stripped) else "invalid"


def add_erp_validation_column(df: pd.DataFrame, erp_column: str) -> pd.DataFrame:
    df = df.copy()
    df["_erp_reference_status"] = df[erp_column].apply(classify_erp_reference)
    return df


def invalid_erp_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Rows where erp_reference_id is populated but not exactly 11 digits.
    Doesn't affect filtering (Step 3 only checks blank-vs-not) — logged as
    a data-quality warning."""
    return df[df["_erp_reference_status"] == "invalid"]


# ===========================================================================
# Filtering — Step 3 (which records survive) and Step 4 (New Status Alert).
# ===========================================================================


def filter_records(
    df: pd.DataFrame,
    erp_column: str,
    order_item_status_column: str,
    payment_status_column: str,
    payment_methods_column: str,
) -> pd.DataFrame:
    """Step 3:
      - Only evaluate rows where erp_reference_id is blank.
      - Drop rows where order_item_status == 'Cancelled'.
      - Drop rows where order_item_status == 'New' AND payment_status ==
        'Pending' AND payment_methods != 'COD'.
      - Everything else survives, explicitly including New/Pending/COD."""
    blank_mask = df[erp_column].apply(is_blank)
    working = df[blank_mask].copy()

    status = working[order_item_status_column].astype(str).str.strip()
    payment_status = working[payment_status_column].astype(str).str.strip()
    payment_method = working[payment_methods_column].astype(str).str.strip()

    condition_1_cancelled = status.str.casefold() == "cancelled"
    condition_2_new_pending_not_cod = (
        (status.str.casefold() == "new")
        & (payment_status.str.casefold() == "pending")
        & (payment_method.str.casefold() != "cod")
    )

    remove_mask = condition_1_cancelled | condition_2_new_pending_not_cod
    return working[~remove_mask].copy()


def new_status_alert(
    df: pd.DataFrame,
    order_item_status_column: str,
    ordered_date_column: str,
    now: datetime | None = None,
) -> pd.DataFrame:
    """Step 4: orders still 'New' more than 1 hour after ordered_date.
    Operates on the already-filtered (Step 3) set."""
    now = now or datetime.now()
    cutoff = now - timedelta(hours=1)

    status = df[order_item_status_column].astype(str).str.strip()
    ordered_date = pd.to_datetime(df[ordered_date_column], errors="coerce")

    mask = (status.str.casefold() == "new") & (ordered_date < cutoff)
    return df[mask].copy()


# ===========================================================================
# Final Report — Part 2: column selection + dedup.
# ===========================================================================


def build_final_report(
    df: pd.DataFrame,
    resolved_columns: dict[str, str],
    order_id_column: str,
    marketplace_channel_column: str | None = None,
) -> pd.DataFrame:
    columns_to_keep: list[str] = []
    for logical_name in FINAL_REPORT_COLUMNS:
        actual_header = resolved_columns.get(logical_name, logical_name)
        if actual_header not in df.columns:
            raise KeyError(
                f"Final Report column '{logical_name}' (expected header "
                f"'{actual_header}') not found in the input file. Check "
                f"FINAL_REPORT_COLUMNS against your file's actual headers."
            )
        columns_to_keep.append(actual_header)

    if marketplace_channel_column:
        if marketplace_channel_column not in df.columns:
            raise KeyError(
                f"marketplace_channel column '{marketplace_channel_column}' "
                f"not found — required for the Part 3 pivot summary."
            )
        columns_to_keep.append(marketplace_channel_column)

    final_df = df[columns_to_keep].copy()
    final_df = final_df.drop_duplicates(subset=[order_id_column], keep="first")
    return final_df


# ===========================================================================
# Pivot Summary — Part 3.
# ===========================================================================


def build_pivot_summary(
    df: pd.DataFrame,
    marketplace_channel_column: str,
    ordered_date_column: str,
    order_id_column: str,
) -> pd.DataFrame:
    working = df.copy()
    working["_ordered_date_only"] = pd.to_datetime(
        working[ordered_date_column], errors="coerce"
    ).dt.date

    pivot = pd.pivot_table(
        working,
        index=marketplace_channel_column,
        columns="_ordered_date_only",
        values=order_id_column,
        aggfunc="count",
        fill_value=0,
        margins=True,
        margins_name="Grand Total",
    )
    pivot.columns = [
        c.isoformat() if hasattr(c, "isoformat") else str(c) for c in pivot.columns
    ]
    return pivot


# ===========================================================================
# Workbook writer — Part 4: assemble the four sheets in order.
# ===========================================================================

HEADER_FONT = Font(bold=True)
TEXT_FORMAT = "@"


def _write_df(ws: Worksheet, df: pd.DataFrame, text_columns: set[str] | None = None) -> None:
    text_columns = text_columns or set()
    ws.append(list(df.columns))
    for cell in ws[1]:
        cell.font = HEADER_FONT

    for row in df.itertuples(index=False):
        ws.append(list(row))

    header_to_idx = {col: i + 1 for i, col in enumerate(df.columns)}
    for col_name in text_columns:
        idx = header_to_idx.get(col_name)
        if idx is None:
            continue
        letter = get_column_letter(idx)
        for row_idx in range(2, ws.max_row + 1):
            ws[f"{letter}{row_idx}"].number_format = TEXT_FORMAT

    _autofit(ws, df)


def _autofit(ws: Worksheet, df: pd.DataFrame, max_width: int = 40) -> None:
    for i, col in enumerate(df.columns, start=1):
        sample = [str(col)] + [str(v) for v in df[col].astype(str).head(200)]
        width = min(max(len(s) for s in sample) + 2, max_width)
        ws.column_dimensions[get_column_letter(i)].width = width


def _write_pivot(ws: Worksheet, pivot_df: pd.DataFrame) -> None:
    index_name = pivot_df.index.name or "Marketplace Channel"
    header = [index_name] + list(pivot_df.columns)
    ws.append(header)
    for cell in ws[1]:
        cell.font = HEADER_FONT

    for idx_value, row in pivot_df.iterrows():
        ws.append([idx_value] + list(row.values))

    grand_total_col_idx = None
    for i, col in enumerate(header, start=1):
        if col == "Grand Total":
            grand_total_col_idx = i
    grand_total_row_idx = None
    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=1).value == "Grand Total":
            grand_total_row_idx = r

    if grand_total_col_idx:
        for r in range(1, ws.max_row + 1):
            ws.cell(row=r, column=grand_total_col_idx).font = HEADER_FONT
    if grand_total_row_idx:
        for c in range(1, ws.max_column + 1):
            ws.cell(row=grand_total_row_idx, column=c).font = HEADER_FONT

    _autofit(ws, pd.DataFrame(columns=header))


def write_workbook(
    output_path: str,
    raw_df: pd.DataFrame,
    new_status_alert_df: pd.DataFrame,
    final_report_df: pd.DataFrame,
    pivot_df: pd.DataFrame,
    order_number_header: str,
) -> None:
    wb = Workbook()

    ws_raw = wb.active
    ws_raw.title = "Raw Data"
    _write_df(ws_raw, raw_df, text_columns={order_number_header})

    ws_alert = wb.create_sheet("New Status Alert")
    _write_df(ws_alert, new_status_alert_df)

    ws_final = wb.create_sheet("Final Report")
    _write_df(ws_final, final_report_df)

    ws_pivot = wb.create_sheet("Pivot Summary")
    _write_pivot(ws_pivot, pivot_df)

    wb._sheets = [ws_raw, ws_alert, ws_final, ws_pivot]
    wb.save(output_path)


# ===========================================================================
# Orchestration — runs Parts 1-4 end to end.
# ===========================================================================


def run_pipeline(
    input_csv: str,
    output_xlsx: str,
    marketplace_channel_column: str | None = None,
    now: datetime | None = None,
) -> dict:
    raw_df, working_df, resolved = load_tc_report(input_csv)

    order_number_col = resolved[EXPECTED_COLUMNS["B"]]
    payment_status_col = resolved[EXPECTED_COLUMNS["G"]]
    order_item_status_col = resolved[EXPECTED_COLUMNS["J"]]
    payment_methods_col = resolved[EXPECTED_COLUMNS["BI"]]
    erp_col = resolved[EXPECTED_COLUMNS["BP"]]

    marketplace_col = marketplace_channel_column or MARKETPLACE_CHANNEL_COLUMN_DEFAULT
    if marketplace_col not in working_df.columns:
        raise KeyError(
            f"Marketplace channel column '{marketplace_col}' not found in the "
            f"input file. The Part 3 pivot requires it — pass the correct "
            f"header name via marketplace_channel_column."
        )

    working_df = trim_all_string_fields(working_df)
    working_df = add_erp_validation_column(working_df, erp_col)
    bad_erp_rows = invalid_erp_rows(working_df)
    if len(bad_erp_rows):
        logger.warning(
            "%d row(s) have a populated erp_reference_id that is not exactly "
            "11 digits. Not excluded by filtering alone, but worth reviewing.",
            len(bad_erp_rows),
        )

    filtered_df = filter_records(
        working_df,
        erp_column=erp_col,
        order_item_status_column=order_item_status_col,
        payment_status_column=payment_status_col,
        payment_methods_column=payment_methods_col,
    )

    alert_df = new_status_alert(
        filtered_df,
        order_item_status_column=order_item_status_col,
        ordered_date_column=resolved.get("ordered_date", "ordered_date"),
        now=now,
    )

    order_id_col = "order_id" if "order_id" in filtered_df.columns else order_number_col
    final_df = build_final_report(
        filtered_df,
        resolved_columns=resolved,
        order_id_column=order_id_col,
        marketplace_channel_column=marketplace_col,
    )

    pivot_df = build_pivot_summary(
        final_df,
        marketplace_channel_column=marketplace_col,
        ordered_date_column="ordered_date",
        order_id_column=order_id_col,
    )

    final_df = final_df.drop(columns=["_erp_reference_status"], errors="ignore")

    write_workbook(
        output_xlsx,
        raw_df=raw_df,
        new_status_alert_df=alert_df.drop(columns=["_erp_reference_status"], errors="ignore"),
        final_report_df=final_df,
        pivot_df=pivot_df,
        order_number_header=order_number_col,
    )

    return {
        "input_rows": len(raw_df),
        "after_step3_filter": len(filtered_df),
        "new_status_alert_rows": len(alert_df),
        "final_report_rows": len(final_df),
        "invalid_erp_reference_rows": len(bad_erp_rows),
        "output_path": output_xlsx,
    }
