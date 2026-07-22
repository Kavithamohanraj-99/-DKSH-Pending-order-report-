# tc-report-processor (v3, with diagnostics)

Processes a TC Report CSV into a 3-sheet workbook: **Raw Data**,
**Final Report**, **Pivot Summary**.

## Fixes in this update

A real run against a 2070-row file produced 0 rows after Step 2.1 with no
visible explanation. Two separate bugs caused that:

1. **Column-mismatch and status-mismatch warnings only went to server
   logs**, invisible in the Streamlit UI. Now `run_pipeline()` returns
   them in its summary dict, and both `app.py` and `main.py` surface them
   directly — including a table of the *actual* `order_item_status`
   values found in your file with their counts, whenever Step 2.1 filters
   out everything. This is the fix that actually matters for debugging:
   next time this happens, you'll see immediately whether it's a wrong
   column or unmatched status text, instead of just "0 rows."
2. **A genuine pandas bug in the pipeline itself**: `.apply()` on an
   *empty* Series returns `float64` dtype instead of `bool`, and indexing
   a DataFrame with a non-bool empty mask silently drops **all columns**,
   not just rows. This meant that if Step 2.1 ever filtered out
   everything, the very next step would crash with a confusing
   `KeyError` instead of completing with an empty (but valid) output.
   Fixed by explicitly casting the mask to `bool`.
3. **Status matching now normalizes spacing around `/`**, so
   `"Accepted / Picked"` matches the whitelist the same as
   `"ACCEPTED/PICKED"` — a formatting variant that would otherwise still
   silently drop every row even with the diagnostics in place.

If you still get 0 rows after this, the app will now tell you exactly
which column it read as `order_item_status` and exactly what values it
found there — that's enough to tell whether it's a column-position
mismatch (wrong file layout) or a genuinely different status vocabulary
than New/ACCEPTED-PICKED/READY TO SHIP.

## Single flat file for the processing logic

`tc_pipeline.py` sits directly next to `app.py`/`main.py` — no
subfolders — because a prior package-based version of this repo
repeatedly failed to fully upload to GitHub, causing `ModuleNotFoundError`
on Streamlit Cloud.

```
tc-report-processor/
├── app.py              # Streamlit dashboard
├── main.py              # CLI
├── tc_pipeline.py        # all processing logic
├── requirements.txt
├── sample_data/
│   └── tc_report_sample.csv
└── README.md
```

## What changed from the previous version of this spec

| Change | Detail |
|---|---|
| **New Status Alert sheet removed** | Part 6 now lists only 3 output sheets. This pipeline no longer builds that alert. |
| **Final Report columns completely changed** | Now `order_number, order_item_status, payment_status, courier_name, tracking_number, ordered_date, accepted_date, nickname, payment_methods, time_shippinglabel_printed, erp_reference_id, time_order_paid` — resolved by spreadsheet letter (B, J, G, AP, AQ, AS, AV, BD, BI, BO, BP, BR). Note this drops `order_id` and `marketplace_channel` as retained columns entirely. |
| **Pivot rows = nickname (BD), displayed as "Marketplace Channel"** | See assumption below — the spec says "rows = Nickname" but the example table's header reads "Marketplace Channel". |
| **New color-coding rule** | Each date column in the Pivot Summary is filled red if the date is before today, green if it's today. Future dates are left unfilled (not specified). |
| **Dedup happens before column deletion** | order_id isn't a retained Final Report column, so dedup on order_id has to happen while it's still present in the working data, before the "delete all remaining columns" step. |

## Two assumptions made to resolve gaps in the spec

1. **order_id's column isn't specified anywhere** (it's referenced only
   for dedup — "Remove duplicate records using order_id" — but never
   given a letter). Defaulted to **Column A**, the conventional
   primary-key position in these exports. Override with
   `--order-id-column` (CLI) or the sidebar field (dashboard) if your
   file uses a different column.
2. **Pivot row label vs. row values**: the spec's instruction says
   `rows = Nickname (BD)`, but the reference example's row header is
   "Marketplace Channel" with values like `lazada-Hiruscar`,
   `shopee-aquamaris`. This pipeline pivots on the actual `nickname`
   column's values (assuming that field already contains
   channel-style identifiers, which the example supports) and just
   labels the header "Marketplace Channel" to match the example. If your
   real `nickname` values are just plain shop nicknames without a
   channel prefix, the pivot will still work — the values will simply
   look different from the reference image.

## Run it

Dashboard:
```bash
pip install -r requirements.txt
streamlit run app.py
```

CLI:
```bash
python main.py path/to/tc_report.csv -o output.xlsx
python main.py sample_data/tc_report_sample.csv -o demo.xlsx --today 2026-07-16
```

`--today` controls the Pivot Summary's red/green coloring — useful for
reproducible test runs; omit it in production to use the live date.

## Deploy to Streamlit Community Cloud

1. **Push with `git push`, not the GitHub web upload UI** if at all
   possible. If you must use the web UI, upload all files in one single
   drag-and-drop.
2. Confirm `app.py`, `main.py`, `tc_pipeline.py`, and `requirements.txt`
   are all visible at your repo's root on GitHub before deploying — if
   `tc_pipeline.py` is missing, you'll get `ModuleNotFoundError`
   regardless of anything on the code side.
3. On [share.streamlit.io](https://share.streamlit.io), point the app at
   this repo, branch `main`, main file path `app.py`.

## Other interpretive notes

- **Columns are resolved by spreadsheet letter position**, per the spec
  (B, G, J, AP, AQ, AS, AV, BD, BI, BO, BP, BR) — logs a warning if the
  header found at that position doesn't resemble the expected name.
- **Step 2.1's whitelist is implemented literally**: "Retain only records
  with the following statuses" — anything not in {New, ACCEPTED/PICKED,
  READY TO SHIP} is dropped, including values not on the spec's explicit
  removal list either.
- **Overwrite-in-place preserved from the prior version**: if the output
  file already exists, only the 3 target sheets are replaced; any other
  sheet in that workbook (e.g. manual notes) is left untouched.
- **Pivot Summary is a computed static table with cell-level fills**, not
  a native interactive Excel PivotTable — openpyxl can't create those.

## Tested against `sample_data/tc_report_sample.csv`

11 synthetic rows covering: each removal status, the New+Pending+non-COD
removal, ACCEPTED/PICKED and READY TO SHIP retention, a populated-erp
row, a duplicate order_id, whitespace-padded fields, and orders dated
today/yesterday/two-days-ago (to exercise the red/green coloring).
Hand-verified results:

```
input_rows: 11
after_status_filter: 8
after_erp_filter: 6
final_report_rows: 5
```

Pivot Summary date columns confirmed colored correctly (two red past-date
columns, one green today column), and data cells — not just headers — are
filled. Overwrite-in-place was re-tested: pre-seeding the output path
with an unrelated "Notes" sheet survives a re-run.
