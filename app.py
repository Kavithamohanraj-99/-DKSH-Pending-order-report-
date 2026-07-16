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
        
        st.sidebar.markdown("### ⚙️ Column Configurations")
        st.sidebar.info("Adjust the column mapping below if the auto-detection mismatches your file headers.")

        # Robust Auto-detection helper
        def get_default_select_idx(col_key, possible_names, default_idx):
            detected = find_column(preview_df, possible_names, default_idx)
            if detected in cols:
                return cols.index(detected)
            return 0

        # Create configurations in sidebar
        m_channel_idx = get_default_select_idx('marketplace_channel', ['Marketplace Channel', 'marketplace_channel', 'channel', 'marketplace'], 0)
        marketplace_col = st.sidebar.selectbox("Marketplace Channel Column", cols, index=m_channel_idx)

        order_no_idx = get_default_select_idx('order_number', ['order_number', 'order number', 'ordernumber', 'order_no', 'order no'], 1)
        order_number_col = st.sidebar.selectbox("Order Number Column (Col B)", cols, index=order_no_idx)

        payment_status_idx = get_default_select_idx('payment_status', ['payment_status', 'payment status', 'paymentstatus', 'pay_status'], 6)
        payment_status_col = st.sidebar.selectbox("Payment Status Column (Col G)", cols, index=payment_status_idx)

        order_id_idx = get_default_select_idx('order_id', ['order_id', 'order id', 'orderid'], 7)
        order_id_col = st.sidebar.selectbox("Order ID Column", cols, index=order_id_idx)

        order_status_idx = get_default_select_idx('order_status', ['order_status', 'order status', 'orderstatus'], 8)
        order_status_col = st.sidebar.selectbox("Order Status Column", cols, index=order_status_idx)

        order_item_status_idx = get_default_select_idx('order_item_status', ['order_item_status', 'order item status', 'orderitemstatus', 'item_status'], 9)
        order_item_status_col = st.sidebar.selectbox("Order Item Status Column (Col J)", cols, index=order_item_status_idx)

        courier_name_idx = get_default_select_idx('courier_name', ['courier_name', 'courier name', 'couriername', 'courier'], 10)
        courier_name_col = st.sidebar.selectbox("Courier Name Column", cols, index=courier_name_idx)

        tracking_number_idx = get_default_select_idx('tracking_number', ['tracking_number', 'tracking number', 'trackingnumber', 'tracking_no', 'awb'], 11)
        tracking_number_col = st.sidebar.selectbox("Tracking Number Column", cols, index=tracking_number_idx)

        ordered_date_idx = get_default_select_idx('ordered_date', ['ordered_date', 'ordered date', 'ordereddate', 'order_date', 'created_at'], 12)
        ordered_date_col = st.sidebar.selectbox("Ordered Date Column", cols, index=ordered_date_idx)

        accepted_date_idx = get_default_select_idx('accepted_date', ['accepted_date', 'accepted date', 'accepteddate', 'accept_date'], 13)
        accepted_date_col = st.sidebar.selectbox("Accepted Date Column", cols, index=accepted_date_idx)

        nickname_idx = get_default_select_idx('nickname', ['nickname', 'nick name', 'username', 'buyer_username'], 14)
        nickname_col = st.sidebar.selectbox("Nickname Column", cols, index=nickname_idx)

        time_shippinglabel_printed_idx = get_default_select_idx('time_shippinglabel_printed', ['time_shippinglabel_printed', 'time shippinglabel printed', 'label_printed_time'], 15)
        time_shippinglabel_printed_col = st.sidebar.selectbox("Time Shipping Label Printed Column", cols, index=time_shippinglabel_printed_idx)

        time_order_paid_idx = get_default_select_idx('time_order_paid', ['time_order_paid', 'time order paid', 'order_paid_time'], 16)
        time_order_paid_col = st.sidebar.selectbox("Time Order Paid Column", cols, index=time_order_paid_idx)

        payment_methods_idx = get_default_select_idx('payment_methods', ['payment_methods', 'payment methods', 'paymentmethod', 'payment_method', 'pay_method'], 60)
        payment_methods_col = st.sidebar.selectbox("Payment Methods Column (Col BI)", cols, index=payment_methods_idx)

        erp_reference_id_idx = get_default_select_idx('erp_reference_id', ['erp_reference_id', 'erp reference id', 'erpreferenceid', 'erp_ref_id'], 67)
        erp_reference_id_col = st.sidebar.selectbox("ERP Reference ID Column (Col BP)", cols, index=erp_reference_id_idx)

        # Date input for Part 3 Status Alert (Defaulting to the user's current system local time)
        st.sidebar.markdown("### 🕒 Alert Threshold Reference Time")
        current_system_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        custom_ref_time = st.sidebar.text_input(
            "Reference Time (YYYY-MM-DD HH:MM:SS)", 
            value=current_system_time_str,
            help="The reference time used to check if orders are in 'New' status for more than 1 hour."
        )

        # Build column mapping dictionary
        column_mapping = {
            'marketplace_channel': marketplace_col,
            'order_number': order_number_col,
            'payment_status': payment_status_col,
            'order_id': order_id_col,
            'order_status': order_status_col,
            'order_item_status': order_item_status_col,
            'courier_name': courier_name_col,
            'tracking_number': tracking_number_col,
            'ordered_date': ordered_date_col,
            'accepted_date': accepted_date_col,
            'nickname': nickname_col,
            'time_shippinglabel_printed': time_shippinglabel_printed_col,
            'time_order_paid': time_order_paid_col,
            'payment_methods': payment_methods_col,
            'erp_reference_id': erp_reference_id_col
        }

        # Process report button
        if st.button("▶ Process report", type="primary"):
            with st.spinner("Processing file... Please wait."):
                # Load the full CSV data
                # Reset uploader file pointer
                uploaded_file.seek(0)
                # Read all columns as string first to prevent losing precision/sci-notation on B (order_number)
                raw_df = pd.read_csv(uploaded_file, dtype=str)

                # Parse custom date
                try:
                    ref_datetime = datetime.strptime(custom_ref_time.strip(), "%Y-%m-%d %H:%M:%S")
                except Exception as ex:
                    st.error(f"Invalid reference time format. Please use YYYY-MM-DD HH:MM:SS (e.g., 2026-07-16 09:01:59).")
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
