import sys
import argparse
import pandas as pd
from datetime import datetime
from processor import process_report_data, find_column

def main():
    parser = argparse.ArgumentParser(description="Process TC Report CSV and generate ERP pending order report.")
    parser.add_argument("input_csv", help="Path to the input TC Report CSV file")
    parser.add_argument("output_excel", help="Path to save the output Excel workbook")
    parser.add_argument("--current-time", help="Reference time in YYYY-MM-DD HH:MM:SS format (default: current system time)")
    
    # Custom column arguments in case defaults don't match
    parser.add_argument("--marketplace-channel", help="Header name for Marketplace Channel column")
    parser.add_argument("--order-number", help="Header name for order_number column")
    parser.add_argument("--payment-status", help="Header name for payment_status column")
    parser.add_argument("--order-id", help="Header name for order_id column")
    parser.add_argument("--order-status", help="Header name for order_status column")
    parser.add_argument("--order-item-status", help="Header name for order_item_status column")
    parser.add_argument("--courier-name", help="Header name for courier_name column")
    parser.add_argument("--tracking-number", help="Header name for tracking_number column")
    parser.add_argument("--ordered-date", help="Header name for ordered_date column")
    parser.add_argument("--accepted-date", help="Header name for accepted_date column")
    parser.add_argument("--nickname", help="Header name for nickname column")
    parser.add_argument("--time-shippinglabel-printed", help="Header name for time_shippinglabel_printed column")
    parser.add_argument("--time-order-paid", help="Header name for time_order_paid column")
    parser.add_argument("--payment-methods", help="Header name for payment_methods column")
    parser.add_argument("--erp-reference-id", help="Header name for erp_reference_id column")

    args = parser.parse_args()

    print(f"Reading input file: {args.input_csv} ...")
    # Read CSV. Read all columns as string to prevent scientific notation and numeric auto-formatting
    try:
        raw_df = pd.read_csv(args.input_csv, dtype=str)
    except Exception as e:
        print(f"Error reading input CSV: {e}", file=sys.stderr)
        sys.exit(1)

    print("Detecting columns...")
    
    # Define mapping dictionary
    column_mapping = {}

    def get_column(arg_val, key, possible_names, default_idx):
        if arg_val:
            if arg_val not in raw_df.columns:
                print(f"Warning: Specified column '{arg_val}' not found in CSV. Attempting auto-detection...", file=sys.stderr)
            else:
                return arg_val
        detected = find_column(raw_df, possible_names, default_idx)
        if detected:
            return detected
        # If not found, use first column as dummy to prevent crash, or raise error
        print(f"Error: Could not locate column for '{key}'. Expected names: {possible_names}", file=sys.stderr)
        sys.exit(1)

    column_mapping['marketplace_channel'] = get_column(args.marketplace_channel, 'marketplace_channel', ['Marketplace Channel', 'marketplace_channel', 'channel', 'marketplace'], 0)
    column_mapping['order_number'] = get_column(args.order_number, 'order_number', ['order_number', 'order number', 'ordernumber', 'order_no', 'order no'], 1)
    column_mapping['payment_status'] = get_column(args.payment_status, 'payment_status', ['payment_status', 'payment status', 'paymentstatus', 'pay_status'], 6)
    column_mapping['order_id'] = get_column(args.order_id, 'order_id', ['order_id', 'order id', 'orderid'], 7)
    column_mapping['order_status'] = get_column(args.order_status, 'order_status', ['order_status', 'order status', 'orderstatus'], 8)
    column_mapping['order_item_status'] = get_column(args.order_item_status, 'order_item_status', ['order_item_status', 'order item status', 'orderitemstatus', 'item_status'], 9)
    column_mapping['courier_name'] = get_column(args.courier_name, 'courier_name', ['courier_name', 'courier name', 'couriername', 'courier'], 10)
    column_mapping['tracking_number'] = get_column(args.tracking_number, 'tracking_number', ['tracking_number', 'tracking number', 'trackingnumber', 'tracking_no', 'awb'], 11)
    column_mapping['ordered_date'] = get_column(args.ordered_date, 'ordered_date', ['ordered_date', 'ordered date', 'ordereddate', 'order_date', 'created_at'], 12)
    column_mapping['accepted_date'] = get_column(args.accepted_date, 'accepted_date', ['accepted_date', 'accepted date', 'accepteddate', 'accept_date'], 13)
    column_mapping['nickname'] = get_column(args.nickname, 'nickname', ['nickname', 'nick name', 'username', 'buyer_username'], 14)
    column_mapping['time_shippinglabel_printed'] = get_column(args.time_shippinglabel_printed, 'time_shippinglabel_printed', ['time_shippinglabel_printed', 'time shippinglabel printed', 'label_printed_time'], 15)
    column_mapping['time_order_paid'] = get_column(args.time_order_paid, 'time_order_paid', ['time_order_paid', 'time order paid', 'order_paid_time'], 16)
    column_mapping['payment_methods'] = get_column(args.payment_methods, 'payment_methods', ['payment_methods', 'payment methods', 'paymentmethod', 'payment_method', 'pay_method'], 60)
    column_mapping['erp_reference_id'] = get_column(args.erp_reference_id, 'erp_reference_id', ['erp_reference_id', 'erp reference id', 'erpreferenceid', 'erp_ref_id'], 67)

    # Print detected column mappings
    print("\n--- Mapped Column Headers ---")
    for k, v in column_mapping.items():
        print(f"  {k: <30} => '{v}'")
    print("-----------------------------\n")

    # Reference Time
    ref_time = datetime.now()
    if args.current_time:
        try:
            ref_time = datetime.strptime(args.current_time.strip(), "%Y-%m-%d %H:%M:%S")
            print(f"Using specified reference time: {ref_time}")
        except ValueError:
            print("Error: Reference time must be in YYYY-MM-DD HH:MM:SS format.", file=sys.stderr)
            sys.exit(1)
    else:
        print(f"Using default system reference time: {ref_time}")

    # Process data
    try:
        new_status_alert, final_report, pivot_summary = process_report_data(
            raw_df, 
            column_mapping, 
            current_time=ref_time
        )
    except Exception as e:
        print(f"Error during report processing: {e}", file=sys.stderr)
        sys.exit(1)

    # Save to Excel
    print(f"Saving worksheets to: {args.output_excel} ...")
    try:
        with pd.ExcelWriter(args.output_excel, engine='openpyxl') as writer:
            # 1. Raw Data (Backup)
            print("  Writing 'Raw Data'...")
            raw_df.to_excel(writer, sheet_name="Raw Data", index=False)
            
            # 2. New Status Alert
            print(f"  Writing 'New Status Alert' ({len(new_status_alert)} records)...")
            new_status_alert.to_excel(writer, sheet_name="New Status Alert", index=False)
            
            # 3. Final Report
            print(f"  Writing 'Final Report' ({len(final_report)} records)...")
            final_report.to_excel(writer, sheet_name="Final Report", index=False)
            
            # 4. Pivot Summary
            print("  Writing 'Pivot Summary'...")
            pivot_summary.to_excel(writer, sheet_name="Pivot Summary", index=True)
            
        print("\nSuccess! Worksheets compiled successfully.")
    except Exception as e:
        print(f"Error saving to Excel: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
