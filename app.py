import streamlit as st
import pandas as pd
import io
from datetime import datetime
from processor import process_report_data, find_column

# Set page config for professional aesthetics
st.set_page_config(
    page_title="TC Report Processor",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .main-title {
        font-family: 'Outfit', 'Inter', sans-serif;
        font-weight: 800;
        font-size: 2.8rem;
        background: linear-gradient(135deg, #FF4B4B, #FF8F8F);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.2rem;
    }
    .subtitle {
        font-family: 'Inter', sans-serif;
        color: #6c757d;
        font-size: 1.1rem;
        margin-bottom: 2rem;
    }
    .card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 20px;
        border-left: 5px solid #FF4B4B;
        margin-bottom: 1.5rem;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="main-title">📊 TC Report Processor</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Upload a TC Report CSV to generate a 4-sheet workbook: Raw Data, New Status Alert, Final Report, Pivot Summary.</p>', unsafe_allow_html=True)

# File uploader
uploaded_file = st.file_uploader("TC Report CSV", type=["csv"], help="Upload the original TC Report CSV file")

if uploaded_file is not None:
    try:
        # Load the CSV file
        # We read a small chunk first to get the columns and allow mappings without loading the full 30MB+ in memory repeatedly
        preview_df = pd.read_csv(uploaded_file, nrows=5)
        cols = list(preview_df.columns)
        
        # Automatically map column headers
        column_mapping = {
            'marketplace_channel': find_column(preview_df, ['Marketplace Channel', 'marketplace_channel', 'channel', 'marketplace'], 0),
            'order_number': find_column(preview_df, ['order_number', 'order number', 'ordernumber', 'order_no', 'order no'], 1),
            'payment_status': find_column(preview_df, ['payment_status', 'payment status', 'paymentstatus', 'pay_status'], 6),
            'order_id': find_column(preview_df, ['order_id', 'order id', 'orderid'], 7),
            'order_status': find_column(preview_df, ['order_status', 'order status', 'orderstatus'], 8),
            'order_item_status': find_column(preview_df, ['order_item_status', 'order item status', 'orderitemstatus', 'item_status'], 9),
            'courier_name': find_column(preview_df, ['courier_name', 'courier name', 'couriername', 'courier'], 10),
            'tracking_number': find_column(preview_df, ['tracking_number', 'tracking number', 'trackingnumber', 'tracking_no', 'awb'], 11),
            'ordered_date': find_column(preview_df, ['ordered_date', 'ordered date', 'ordereddate', 'order_date', 'created_at'], 12),
            'accepted_date': find_column(preview_df, ['accepted_date', 'accepted date', 'accepteddate', 'accept_date'], 13),
            'nickname': find_column(preview_df, ['nickname', 'nick name', 'username', 'buyer_username'], 14),
            'time_shippinglabel_printed': find_column(preview_df, ['time_shippinglabel_printed', 'time shippinglabel printed', 'label_printed_time'], 15),
            'time_order_paid': find_column(preview_df, ['time_order_paid', 'time order paid', 'order_paid_time'], 16),
            'payment_methods': find_column(preview_df, ['payment_methods', 'payment methods', 'paymentmethod', 'payment_method', 'pay_method'], 60),
            'erp_reference_id': find_column(preview_df, ['erp_reference_id', 'erp reference id', 'erpreferenceid', 'erp_ref_id'], 67)
        }

        # Process report button
        if st.button("▶ Process report", type="primary"):
            with st.spinner("Processing file... Please wait."):
                # Load the full CSV data
                # Reset uploader file pointer
                uploaded_file.seek(0)
                # Read all columns as string first to prevent losing precision/sci-notation on B (order_number)
                raw_df = pd.read_csv(uploaded_file, dtype=str)

                ref_datetime = datetime.now()

                # Process
                new_status_alert, final_report, pivot_summary = process_report_data(
                    raw_df, 
                    column_mapping, 
                    current_time=ref_datetime
                )

                # Generate excel workbook in memory
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    # 1. Raw Data (Backup of the original imported data)
                    raw_df.to_excel(writer, sheet_name="Raw Data", index=False)
                    
                    # 2. New Status Alert
                    new_status_alert.to_excel(writer, sheet_name="New Status Alert", index=False)
                    
                    # 3. Final Report
                    final_report.to_excel(writer, sheet_name="Final Report", index=False)
                    
                    # 4. Pivot Summary
                    pivot_summary.to_excel(writer, sheet_name="Pivot Summary", index=True)

                st.success("🎉 Report processed successfully!")

                # Download button
                st.download_button(
                    label="📥 Download Excel Workbook",
                    data=excel_buffer.getvalue(),
                    file_name="ERP_Pending_Order_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Show previews in tabs
                tab1, tab2, tab3, tab4 = st.tabs([
                    "📄 Raw Data Preview", 
                    "🚨 New Status Alert", 
                    "🏆 Final Report", 
                    "🧮 Pivot Summary"
                ])

                with tab1:
                    st.markdown("### Original Input Data (Top 5 rows)")
                    st.dataframe(raw_df.head(5))
                    st.metric("Total Records", len(raw_df))

                with tab2:
                    st.markdown("### Orders in 'New' Status for > 1 Hour")
                    st.dataframe(new_status_alert)
                    st.metric("Alert Records Count", len(new_status_alert))

                with tab3:
                    st.markdown("### Cleaned ERP Investigation Final Report")
                    st.dataframe(final_report)
                    st.metric("Cleaned Records Count", len(final_report))

                with tab4:
                    st.markdown("### Pivot Summary Table")
                    st.dataframe(pivot_summary)

    except Exception as e:
        st.error(f"Processing failed: {str(e)}")
        st.info("Please make sure you have mapped all required columns correctly in the sidebar.")
else:
    # Instructions card when file is not uploaded
    st.markdown("""
    <div class="card">
        <h3>ℹ️ Instructions</h3>
        <p>1. Upload your TC Report CSV using the file uploader above.</p>
        <p>2. The app will auto-detect columns based on common patterns.</p>
        <p>3. If your CSV has non-standard header names, update the mappings in the left sidebar.</p>
        <p>4. Click 'Process report' to clean, filter, and compile your worksheets.</p>
        <p>5. Download the final workbook containing all 4 required sheets.</p>
    </div>
    """, unsafe_allow_html=True)
