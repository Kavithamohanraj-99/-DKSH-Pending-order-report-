import pandas as pd
import numpy as np
from datetime import datetime, timedelta

def normalize_string(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    if s.lower() in ('nan', 'none', 'null', ''):
        return None
    return s

def find_column(df, possible_names, default_idx=None):
    """
    Finds a column in the DataFrame matching any of the possible names (case-insensitive, trimmed).
    Falls back to default_idx if provided and within range.
    """
    columns = list(df.columns)
    # 1. Direct match (case-insensitive, stripped)
    for name in possible_names:
        for col in columns:
            if col.strip().lower() == name.strip().lower():
                return col
            # Normalize spaces/underscores
            col_norm = col.strip().lower().replace('_', '').replace(' ', '')
            name_norm = name.strip().lower().replace('_', '').replace(' ', '')
            if col_norm == name_norm:
                return col
                
    # 2. Index fallback
    if default_idx is not None and default_idx < len(columns):
        return columns[default_idx]
        
    return None

def process_report_data(df, column_mapping, current_time=None):
    """
    Core data processing pipeline:
    - Step 1: Keep a backup (Raw Data) - handled by caller when saving to Excel.
    - Step 2: Clean the data (Trim spaces, Convert B to Text/General).
    - Step 2.1: Filter order status.
    - Step 2.2: Evaluate ERP Reference ID and filter for ERP investigation.
    - Part 3: Generate New Status Alert.
    - Part 4: Retain final columns & Remove duplicates by order_id.
    - Part 5: Generate Pivot Summary.
    """
    if current_time is None:
        current_time = datetime.now()
    elif isinstance(current_time, str):
        current_time = pd.to_datetime(current_time)

    # Make a copy of the input DataFrame
    working_df = df.copy()

    # Trim spaces from column names
    working_df.columns = [str(c).strip() for c in working_df.columns]

    # Resolve column names using mapping
    order_number_col = column_mapping.get('order_number')
    order_item_status_col = column_mapping.get('order_item_status')
    erp_reference_id_col = column_mapping.get('erp_reference_id')
    payment_status_col = column_mapping.get('payment_status')
    payment_methods_col = column_mapping.get('payment_methods')
    ordered_date_col = column_mapping.get('ordered_date')
    order_id_col = column_mapping.get('order_id')
    marketplace_channel_col = column_mapping.get('marketplace_channel')

    # Verify key columns exist
    required_keys = ['order_number', 'order_item_status', 'erp_reference_id', 
                     'payment_status', 'payment_methods', 'ordered_date', 'order_id', 'marketplace_channel']
    for k in required_keys:
        col = column_mapping.get(k)
        if not col or col not in working_df.columns:
            raise ValueError(f"Required column mapping for '{k}' ({col}) not found in data.")

    # --- Part 1, Step 2: Prepare the Data ---
    # Convert order_number to string (text/general format) and trim all string columns
    for col in working_df.columns:
        if col == order_number_col:
            working_df[col] = working_df[col].astype(str).apply(normalize_string)
        else:
            if working_df[col].dtype == 'object':
                working_df[col] = working_df[col].apply(normalize_string)

    # --- Step 2.1: Filter Order Status ---
    # Retain only: New, ACCEPTED/PICKED, READY TO SHIP
    allowed_statuses = {'New', 'ACCEPTED/PICKED', 'READY TO SHIP'}
    
    # We do a case-insensitive check and trim check to be safe
    def is_status_allowed(val):
        if val is None:
            return False
        # Normalize status comparison
        val_norm = str(val).strip().upper()
        return any(val_norm == status.upper() for status in allowed_statuses)

    status_filtered_df = working_df[working_df[order_item_status_col].apply(is_status_allowed)].copy()

    # --- Part 3: New Status Validation (Alert) ---
    # This worksheet highlights orders that have remained in New status for more than one hour.
    # We apply this to the cleaned status_filtered_df
    # Parse ordered_date as datetime
    status_filtered_df['_parsed_ordered_date'] = pd.to_datetime(status_filtered_df[ordered_date_col], errors='coerce')
    
    # Filter for: status == 'New' AND ordered_date < current_time - 1 hour
    new_status_mask = (status_filtered_df[order_item_status_col].astype(str).str.strip().str.upper() == 'NEW')
    one_hour_ago = current_time - timedelta(hours=1)
    
    # If date parsing failed, we treat as not matching
    time_mask = (status_filtered_df['_parsed_ordered_date'] < one_hour_ago) & (status_filtered_df['_parsed_ordered_date'].notna())
    
    new_status_alert_df = status_filtered_df[new_status_mask & time_mask].drop(columns=['_parsed_ordered_date']).copy()

    # --- Part 2: Filter Records for ERP Investigation ---
    # Evaluate only records where erp_reference_id is blank.
    # Treat blank or whitespace-only values as blank (already handled by normalize_string which sets them to None).
    
    # Filter records with blank erp_ref
    blank_erp_mask = status_filtered_df[erp_reference_id_col].isna()
    erp_blank_df = status_filtered_df[blank_erp_mask].copy()

    # Remove records that meet all:
    # - order_item_status == New
    # - payment_status == Pending
    # - payment_methods != COD
    def should_remove(row):
        status = str(row[order_item_status_col] or '').strip().upper()
        pay_status = str(row[payment_status_col] or '').strip().upper()
        pay_method = str(row[payment_methods_col] or '').strip().upper()
        return status == 'NEW' and pay_status == 'PENDING' and pay_method != 'COD'

    remove_mask = erp_blank_df.apply(should_remove, axis=1)
    erp_investigation_df = erp_blank_df[~remove_mask].copy()

    # --- Part 4: Generate Final Report ---
    # Retain only specific columns
    final_columns_order = [
        'Marketplace Channel', 'order_id', 'order_status', 'payment_status',
        'courier_name', 'tracking_number', 'ordered_date', 'accepted_date',
        'nickname', 'payment_methods', 'time_shippinglabel_printed',
        'erp_reference_id', 'time_order_paid'
    ]

    # Map the requested columns to the actual columns in the dataset
    final_col_mapping = {}
    for col_key in final_columns_order:
        # Standardize matching for the final columns list
        actual_col = column_mapping.get(col_key.lower().replace(' ', '_'))
        if actual_col:
            final_col_mapping[actual_col] = col_key

    # Select and rename columns
    final_report_df = erp_investigation_df[list(final_col_mapping.keys())].rename(columns=final_col_mapping)

    # Ensure all final columns are present (even if empty) in the correct order
    for col in final_columns_order:
        if col not in final_report_df.columns:
            final_report_df[col] = None
    final_report_df = final_report_df[final_columns_order]

    # Remove duplicate records using order_id only, keeping the first occurrence
    final_report_df = final_report_df.drop_duplicates(subset=['order_id'], keep='first')

    # --- Part 5: Create Pivot Summary ---
    # Create Pivot Table using Final Report:
    # Rows: Marketplace Channel
    # Columns: ordered_date
    # Values: Count of order_id
    
    # Prepare Pivot Source
    pivot_source = final_report_df.copy()
    
    # We want columns of ordered_date to group by date (YYYY-MM-DD) for clean representation
    # If ordered_date can't be parsed, it remains as string or NaT
    pivot_source['parsed_date'] = pd.to_datetime(pivot_source['ordered_date'], errors='coerce')
    pivot_source['Date'] = pivot_source['parsed_date'].dt.strftime('%Y-%m-%d')
    # Fallback to original string if parsing failed but has value
    pivot_source['Date'] = pivot_source['Date'].fillna(pivot_source['ordered_date'].fillna('Unknown Date'))

    if pivot_source.empty:
        # Create empty pivot table structure
        pivot_df = pd.DataFrame(columns=['Grand Total'])
        pivot_df.index.name = 'Marketplace Channel'
    else:
        # Build pivot table
        pivot_table = pd.pivot_table(
            pivot_source,
            index='Marketplace Channel',
            columns='Date',
            values='order_id',
            aggfunc='count',
            fill_value=0
        )
        
        # Chronological sorting is preserved since 'Date' was 'YYYY-MM-DD'.
        # Now rename the date columns to match the 'DD-MMM' format (e.g., 14-Jul)
        def format_date_header(col_val):
            try:
                # Try parsing the YYYY-MM-DD date string
                dt = datetime.strptime(str(col_val), '%Y-%m-%d')
                # Format as d-MMM (e.g. 14-Jul, with no leading zero for single digit day)
                return f"{dt.day}-{dt.strftime('%b')}"
            except Exception:
                return str(col_val)
                
        pivot_table = pivot_table.rename(columns=format_date_header)
        
        # Calculate Row Grand Totals
        pivot_table['Grand Total'] = pivot_table.sum(axis=1)
        
        # Calculate Column Grand Totals
        col_totals = pivot_table.sum(axis=0)
        col_totals.name = 'Grand Total'
        
        pivot_df = pd.concat([pivot_table, pd.DataFrame([col_totals])])

    # Clean up internal column in status_filtered_df
    status_filtered_df = status_filtered_df.drop(columns=['_parsed_ordered_date'])

    return new_status_alert_df, final_report_df, pivot_df
