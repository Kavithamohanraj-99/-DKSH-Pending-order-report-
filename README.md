# tc-report-processor (v3)

Processes a TC Report CSV into a 3-sheet workbook: **Raw Data**,
**Final Report**, **Pivot Summary**.

Single flat file for the processing logic (`tc_pipeline.py`) — no
subfolders — for the same reason as before: a prior package-based version
of this repo repeatedly failed to fully upload to GitHub, causing
`ModuleNotFoundError` on Streamlit Cloud.

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
