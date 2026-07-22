# tc-report-processor (CLI-only)

Processes a TC Report CSV into a 3-sheet workbook: **Raw Data**,
**Final Report**, **Pivot Summary**. No dashboard, no Streamlit — just a
script.

```
tc-report-processor/
├── main.py              # CLI
├── tc_pipeline.py        # all processing logic
├── requirements.txt
└── sample_data/
    └── tc_report_sample.csv
```

## Setup

```bash
pip install -r requirements.txt
```

## Step 1: check your file's column layout BEFORE running the pipeline

```bash
python main.py your_real_file.csv --inspect
```

This prints every column position this pipeline reads from (by letter:
A, B, G, J, AP, AQ, AS, AV, BD, BI, BO, BP, BR), what header it expects
there, what header it actually found, and 5 real sample values — so a
misaligned column shows up immediately as either `[MISMATCH?]` or an
obviously wrong sample value, without running anything else. **Always run
this first against a new file.**

Example output:
```
    J (expects order_item_status           ): found 'order_item_status' [OK]  samples=['Cancelled', 'DELIVERED', 'New', 'New', 'ACCEPTED/PICKED']
```
If instead you saw something like:
```
    J (expects order_item_status           ): found 'warehouse_code' [MISMATCH?]  samples=['WH01', 'WH02', ...]
```
— that tells you column J in your file isn't order_item_status at all,
which is exactly the kind of thing that silently produces 0 rows with no
explanation.

## Step 2: run the pipeline

```bash
python main.py your_real_file.csv -o output.xlsx
```

If Step 2.1 (the order-status whitelist) still filters out everything,
the script prints the actual `order_item_status` values it found and
their counts — so you can see whether it's a wording mismatch (e.g.
`"Ready to Ship"` vs. the spec's `"READY TO SHIP"` — matching is
case-insensitive, so casing alone isn't the issue, but different
punctuation/wording would be) even after confirming the column itself is
right.

```bash
python main.py sample_data/tc_report_sample.csv -o demo.xlsx --today 2026-07-16
```

## Options

| Flag | Purpose |
|---|---|
| `--inspect` | Print column layout + samples, don't run the pipeline |
| `-o / --output` | Output path (default `TC_Report_Processed.xlsx`). If it already exists, only the 3 target sheets are overwritten — other sheets in that workbook are left alone. |
| `--order-id-column` | Spreadsheet letter for order_id, used only for dedup (default `A` — the spec never states this explicitly, see below) |
| `--today` | ISO date for Pivot Summary's red/green coloring (default: live system date) |

## Assumptions this pipeline makes (check these against your real file)

1. **order_id is Column A.** The spec dedups on order_id but never gives
   it a letter — it isn't one of the retained Final Report columns
   either. Override with `--order-id-column` if wrong.
2. **Pivot rows come from `nickname` (Column BD), displayed with the
   header "Marketplace Channel."** The spec literally says
   `rows = Nickname`, but its example table's header reads "Marketplace
   Channel" with values like `lazada-Hiruscar` — this assumes your
   `nickname` field already contains channel-style values.
3. **Column positions are exactly**: B=order_number, G=payment_status,
   J=order_item_status, AP=courier_name, AQ=tracking_number,
   AS=ordered_date, AV=accepted_date, BD=nickname, BI=payment_methods,
   BO=time_shippinglabel_printed, BP=erp_reference_id, BR=time_order_paid.
   Any extra or missing column anywhere before these positions in your
   real file shifts everything after it — this is the single most likely
   cause of a mismatch. `--inspect` catches this immediately.

## Processing logic, in order

1. Trim whitespace on every field; read order_number as text (no
   scientific notation).
2. **Step 2.1**: keep only rows where order_item_status is New,
   ACCEPTED/PICKED, or READY TO SHIP (case-insensitive, tolerant of
   spacing around `/`). Everything else — Cancelled, DELIVERED, SHIPPED,
   RETURN*, etc. — is dropped here.
3. **Part 2**: of the rows with a blank erp_reference_id, drop New +
   Pending + non-COD combinations. Rows with a populated
   erp_reference_id are excluded entirely (never reach the Final Report).
4. **Part 4**: select the 12 retained columns, dedup by order_id (first
   occurrence wins), save as Final Report.
5. **Part 5**: pivot nickname × ordered_date, count of orders, Grand
   Totals on every axis, date columns colored red (past) / green (today).
6. **Part 6**: write Raw Data / Final Report / Pivot Summary into the
   output workbook, preserving any other sheets already in that file.

## Known pandas gotcha already fixed here

If a filtering step ever produces 0 rows, `.apply()` on an empty pandas
Series returns `float64` dtype instead of `bool` — and using that as a
boolean mask silently drops **all columns**, not just rows, causing a
confusing `KeyError` two steps later instead of a clean empty result.
This is fixed (masks are explicitly cast to `bool`), so an empty
intermediate result now produces a valid (empty) output instead of
crashing.
