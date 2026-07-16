# tc-report-processor

Processes a TC Report CSV into a 4-sheet workbook: **Raw Data**,
**New Status Alert**, **Final Report**, **Pivot Summary**.

## Why this version is one flat file

Earlier versions put the processing logic in a `src/tc_report_processor/`
package, then a top-level `tc_report_processor/` package. Both times, the
package folder failed to actually reach the deployed GitHub repo (likely
from uploading files individually via the GitHub web UI rather than
`git push`), causing `ModuleNotFoundError` on Streamlit Cloud even though
everything worked locally.

This version has **no subfolders for code at all** — `tc_pipeline.py`,
`app.py`, and `main.py` all sit directly in the repo root. A single file
next to `app.py` can't be "partially" uploaded the way a nested folder
can: if `app.py` made it into the repo, `tc_pipeline.py` did too.

```
tc-report-processor/
├── app.py              # Streamlit dashboard
├── main.py              # CLI
├── tc_pipeline.py        # all processing logic (columns, cleaning, filtering, pivot, workbook writer)
├── requirements.txt
├── sample_data/
│   └── tc_report_sample.csv   # synthetic file covering every rule, for testing
└── README.md
```

## Run it

Dashboard:
```bash
pip install -r requirements.txt
streamlit run app.py
```

CLI:
```bash
python main.py path/to/tc_report.csv -o output.xlsx
python main.py sample_data/tc_report_sample.csv -o demo.xlsx --now 2026-07-16T15:00:00
```

## Deploy to Streamlit Community Cloud

1. **Push with `git push`, not the GitHub web upload UI.** If you must use
   the web UI, use "Add file → Upload files" and drag all files in
   *one single drop* (including `sample_data/tc_report_sample.csv`) rather
   than creating them one at a time.
2. After pushing, open the repo on GitHub and confirm you see exactly
   these files at the root: `app.py`, `main.py`, `tc_pipeline.py`,
   `requirements.txt`. If `tc_pipeline.py` isn't listed, Streamlit Cloud
   will fail with `ModuleNotFoundError: No module named 'tc_pipeline'`
   again — that check takes ten seconds and would have caught both
   previous failures before deploying.
3. On [share.streamlit.io](https://share.streamlit.io), point the app at
   this repo, branch `main`, main file path `app.py`.

## Spec-mapping and known assumptions

`tc_pipeline.py` is organized in sections mirroring the spec, top to
bottom: column resolution → loading → cleaning (Step 2) → filtering
(Step 3) → New Status Alert (Step 4) → Final Report (Part 2) → Pivot
Summary (Part 3) → workbook writer (Part 4) → orchestration.

- **Columns are resolved by spreadsheet letter** (B, G, J, BI, BP) against
  whatever header text is actually at that position in your file — logs a
  warning if it doesn't look like the expected name, rather than silently
  reading the wrong column.
- **The Part 3 pivot needs a "Marketplace Channel" column that Part 2's
  retained-column list doesn't include** — a gap in the source spec. This
  pipeline assumes the Final Report should also retain
  `marketplace_channel`; pass `--marketplace-column "Your Header"` (CLI)
  or set it in the sidebar (dashboard) if your file names it differently.
- **New Status Alert runs on the already-filtered (Step 3) record set**,
  not the raw file, since Step 4 is defined right after filtering and is
  about highlighting stale orders among ones still being tracked.
- **erp_reference_id's 11-digit check doesn't gate filtering** — Step 3
  only checks blank-vs-populated. Malformed-but-populated values are
  logged as a data-quality warning (`invalid_erp_reference_rows`), not
  excluded or included differently.
- **Pivot Summary is a computed static table**, not a native interactive
  Excel PivotTable — openpyxl can't create those. It looks like a pivot
  but won't "Refresh" in Excel; flagging so that's not a surprise.
