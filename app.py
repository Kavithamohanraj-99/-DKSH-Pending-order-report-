"""Streamlit dashboard for tc_report_processor.

Deploy: point Streamlit Community Cloud at this file (main file path =
app.py). No credentials needed — this tool only processes a CSV you upload
and hands back a workbook, nothing touches external accounts.
"""

from __future__ import annotations

import io
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

from tc_report_processor.pipeline import run_pipeline

st.set_page_config(page_title="TC Report Processor", page_icon="📊", layout="wide")

st.title("📊 TC Report Processor")
st.caption(
    "Upload a TC Report CSV to generate a 4-sheet workbook: Raw Data, "
    "New Status Alert, Final Report, Pivot Summary."
)

with st.sidebar:
    st.header("Options")
    marketplace_column = st.text_input(
        "Marketplace channel column header",
        value="marketplace_channel",
        help="The Part 3 pivot needs a marketplace/channel column. Change "
        "this if your file uses a different header for that field.",
    )
    use_custom_now = st.checkbox(
        "Override 'current time' for New Status Alert",
        value=False,
        help="Useful for reproducible backfills. Leave unchecked to use "
        "the live clock.",
    )
    custom_now = None
    if use_custom_now:
        now_date = st.date_input("Date")
        now_time = st.time_input("Time")
        custom_now = datetime.combine(now_date, now_time)

uploaded_file = st.file_uploader("TC Report CSV", type=["csv"])

if uploaded_file is None:
    st.info("Upload a CSV to get started.")
    st.stop()

run_clicked = st.button("▶ Process report", type="primary")

if run_clicked:
    with tempfile.TemporaryDirectory() as tmp:
        input_path = Path(tmp) / "input.csv"
        input_path.write_bytes(uploaded_file.getvalue())
        output_path = Path(tmp) / "output.xlsx"

        try:
            with st.spinner("Processing..."):
                summary = run_pipeline(
                    input_csv=str(input_path),
                    output_xlsx=str(output_path),
                    marketplace_channel_column=marketplace_column or None,
                    now=custom_now,
                )
        except Exception as e:  # noqa: BLE001
            st.error(f"Processing failed: {e}")
            st.stop()

        output_bytes = output_path.read_bytes()

    st.success("Done.")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Input rows", summary["input_rows"])
    m2.metric("After Step 3 filter", summary["after_step3_filter"])
    m3.metric("Final Report rows", summary["final_report_rows"])
    m4.metric("New Status Alert rows", summary["new_status_alert_rows"])

    if summary["invalid_erp_reference_rows"]:
        st.warning(
            f"{summary['invalid_erp_reference_rows']} row(s) have a populated "
            f"erp_reference_id that isn't exactly 11 digits. These aren't "
            f"excluded by filtering (see README) but are worth reviewing."
        )

    st.download_button(
        "⬇ Download workbook",
        data=output_bytes,
        file_name=f"TC_Report_Processed_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.subheader("Preview")
    wb = load_workbook(io.BytesIO(output_bytes))
    tabs = st.tabs(wb.sheetnames)
    for tab, sheet_name in zip(tabs, wb.sheetnames):
        with tab:
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                st.write("(empty)")
                continue
            preview_df = pd.DataFrame(rows[1:], columns=rows[0])
            st.dataframe(preview_df.head(200), use_container_width=True, hide_index=True)
            if len(preview_df) > 200:
                st.caption(f"Showing first 200 of {len(preview_df)} rows.")
