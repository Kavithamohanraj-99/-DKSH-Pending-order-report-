# tc-report-processor

Processes a TC Report CSV into a 4-sheet workbook: **Raw Data**,
**New Status Alert**, **Final Report**, **Pivot Summary** — per the
"Prepare Pending Order Reports, Generate ERP Exception Report, Pivot
Summary" spec.

## Quick start

```bash
pip install -r requirements.txt
python main.py path/to/tc_report.csv -o output.xlsx
```

A synthetic sample file covering every filtering rule is included at
`sample_data/tc_report_sample.csv` — try it first:

```bash
python main.py sample_data/tc_report_sample.csv -o /tmp/demo.xlsx --now 2026-07-16T15:00:00
```

(`--now` pins the "current time" used for the New Status Alert cutoff —
useful for reproducible test runs; omit it in production to use the live
clock.)

## How the spec maps to code

| Spec section | Module |
|---|---|
| Part 1 / Step 1 — backup, "Raw Data" tab, never modified | `io_utils.py` loads the file twice-conceptually: `raw_df` is written straight to the Raw Data sheet and never touched again |
| Part 1 / Step 2 — order_number as text, trim whitespace, validate erp_reference_id | `io_utils.py` (text dtype on read), `cleaning.py` (trim + 11-digit validation) |
| Part 1 / Step 3 — filter records | `filtering.py: filter_records()` |
| Part 1 / Step 4 — New Status Alert | `filtering.py: new_status_alert()` |
| Part 2 — Final Report (column selection + dedup) | `final_report.py` |
| Part 3 — Pivot Summary | `pivot.py` |
| Part 4 — assemble workbook | `workbook_writer.py` |
| Everything end-to-end | `pipeline.py: run_pipeline()` |

## Column resolution: letters vs. header names

The spec identifies several columns by spreadsheet letter (`Column B` =
order_number, `Column J` = order_item_status, `Column G` = payment_status,
`Column BI` = payment_methods, `Column BP` = erp_reference_id). Real CSV
exports vary in header text, so `columns.py` resolves each letter to
whatever header is actually at that position in your file, and logs a
warning if it doesn't resemble the expected name (rather than silently
reading the wrong column, or hard-failing on a minor naming difference).

If your TC Report template doesn't put those five fields in those exact
positions, this will misread your file — check the warning log on first
run against a new export.

## Known gap in the source spec (and the assumption made to resolve it)

**Part 3 requires "Marketplace Channel" as a pivot row dimension, but
Part 2's retained-column list for the Final Report does not include a
marketplace/channel column** — so as written, the pivot can't be built
from the Final Report at all.

This pipeline assumes the Final Report should also retain a
`marketplace_channel` column so the pivot works (see the note in
`columns.py`). If your actual TC Report uses a different header for this
field, either rename it to `marketplace_channel` upstream or pass:

```bash
python main.py input.csv -o output.xlsx --marketplace-column "Your Header Name"
```

## Other interpretive decisions worth knowing about

- **New Status Alert is built from the *filtered* (Step 3) record set**,
  not the full raw file — the spec places Step 4 immediately after Step 3
  filtering and frames it as highlighting stale orders among the ones
  still being tracked (already-cancelled orders, for instance, aren't
  "stuck in New").
- **erp_reference_id validity (11-digit check) doesn't gate filtering.**
  Step 3 only checks blank-vs-populated. A populated-but-malformed value
  (not exactly 11 digits) is logged as a data-quality warning
  (`invalid_erp_reference_rows` in the run summary) but doesn't get
  excluded or included differently — the spec doesn't say what to do with
  it beyond "verify."
- **ordered_date pivot columns are bucketed to calendar date**, not full
  timestamp, since the spec asks for one pivot column per date and
  "automatically expand[ing]" based on dates present.
- **Pivot values are computed in Python and written as static numbers**,
  not native Excel `PivotTable` objects or formulas — openpyxl can't
  create a true interactive Excel pivot table. If you need a real
  refreshable pivot (right-click → Refresh in Excel), you'd need to build
  it from a template `.xlsx` with an existing PivotCache via a different
  approach (e.g. `xlwings` against a real Excel instance, or a VBA-driven
  template). Flagging this rather than quietly shipping something that
  looks like a pivot but doesn't behave like one.

## Tests

```bash
pip install pytest
pytest tests/ -v
```

Covers: full pipeline row counts against the synthetic sample, sheet
names/order, Raw Data staying untouched (including preserved whitespace),
Final Report column order, and pivot grand totals.
