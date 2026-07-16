import unittest
import pandas as pd
from datetime import datetime, timedelta
from processor import process_report_data, normalize_string, find_column

class TestPendingOrderReport(unittest.TestCase):
    def setUp(self):
        # Create a mock DataFrame matching the columns layout
        self.column_mapping = {
            'marketplace_channel': 'Marketplace Channel',
            'order_number': 'order_number',
            'payment_status': 'payment_status',
            'order_id': 'order_id',
            'order_status': 'order_status',
            'order_item_status': 'order_item_status',
            'courier_name': 'courier_name',
            'tracking_number': 'tracking_number',
            'ordered_date': 'ordered_date',
            'accepted_date': 'accepted_date',
            'nickname': 'nickname',
            'time_shippinglabel_printed': 'time_shippinglabel_printed',
            'time_order_paid': 'time_order_paid',
            'payment_methods': 'payment_methods',
            'erp_reference_id': 'erp_reference_id'
        }
        
        self.ref_time = datetime(2026, 7, 16, 9, 1, 59)

    def test_normalize_string(self):
        self.assertEqual(normalize_string("  hello  "), "hello")
        self.assertEqual(normalize_string(""), None)
        self.assertEqual(normalize_string("   "), None)
        self.assertEqual(normalize_string("None"), None)
        self.assertEqual(normalize_string("nan"), None)
        self.assertEqual(normalize_string(pd.NA), None)

    def test_find_column(self):
        df = pd.DataFrame(columns=['Marketplace_Channel', 'Order_Number', 'Payment Status'])
        self.assertEqual(find_column(df, ['Marketplace Channel', 'marketplace_channel'], 0), 'Marketplace_Channel')
        self.assertEqual(find_column(df, ['order_number', 'order number'], 1), 'Order_Number')
        self.assertEqual(find_column(df, ['payment_status', 'payment status'], 2), 'Payment Status')
        # Index fallback
        self.assertEqual(find_column(df, ['non_existent'], 1), 'Order_Number')

    def test_process_report_data_filtering(self):
        # Create mock dataset
        data = {
            'Marketplace Channel': ['Shopee', 'Lazada', 'TikTok', 'Shopee', 'Lazada', 'Shopee', 'Shopee'],
            'order_number': ['123', ' 456 ', '789', '101', '102', '103', '123'],
            'payment_status': ['Pending', 'Pending', 'Paid', 'Pending', 'Paid', 'Pending', 'Pending'],
            'order_id': ['ORD001', 'ORD002', 'ORD003', 'ORD004', 'ORD005', 'ORD006', 'ORD001'],
            'order_status': ['Unfulfilled', 'Unfulfilled', 'Unfulfilled', 'Unfulfilled', 'Unfulfilled', 'Unfulfilled', 'Unfulfilled'],
            'order_item_status': ['New', 'New', 'ACCEPTED/PICKED', 'READY TO SHIP', 'Shipped', 'New', 'New'],
            'courier_name': ['DHL', 'J&T', 'Ninja', 'DHL', 'J&T', 'DHL', 'DHL'],
            'tracking_number': ['TRK001', 'TRK002', 'TRK003', 'TRK004', 'TRK005', 'TRK006', 'TRK001'],
            'ordered_date': [
                '2026-07-16 08:30:00', # 31 mins ago (status New) -> No alert
                '2026-07-16 07:30:00', # 1 hr 31 mins ago (status New) -> Alert
                '2026-07-16 08:00:00',
                '2026-07-16 08:15:00',
                '2026-07-16 06:00:00', # Shipped -> Filtered out
                '2026-07-16 07:00:00', # 2 hr 1 min ago (status New) -> Alert
                '2026-07-16 08:30:00'
            ],
            'accepted_date': ['', '', '', '', '', '', ''],
            'nickname': ['UserA', 'UserB', 'UserC', 'UserD', 'UserE', 'UserF', 'UserA'],
            'time_shippinglabel_printed': ['', '', '', '', '', '', ''],
            'time_order_paid': ['', '', '', '', '', '', ''],
            'payment_methods': ['COD', 'Credit Card', 'COD', 'COD', 'COD', 'COD', 'COD'],
            'erp_reference_id': [None, None, None, None, None, 'ERP-123', None] # ORD006 has ERP -> Excluded from final report
        }
        df = pd.DataFrame(data)
        
        # Run processing
        new_status_alert, final_report, pivot_summary = process_report_data(
            df, 
            self.column_mapping, 
            current_time=self.ref_time
        )
        
        # 1. Check order item status filtering
        # 'Shipped' should be removed, so ORD005 is excluded.
        self.assertNotIn('ORD005', final_report['order_id'].values)
        
        # 2. Check ERP Reference filter
        # ORD006 has ERP-123 (not blank) -> Should be excluded from final report
        self.assertNotIn('ORD006', final_report['order_id'].values)
        
        # 3. Check Remove Records condition (Part 2)
        # ORD002: New, Pending, Credit Card (non-COD) -> Should be removed
        self.assertNotIn('ORD002', final_report['order_id'].values)
        # ORD001: New, Pending, COD -> Should be kept
        self.assertIn('ORD001', final_report['order_id'].values)
        # ORD003: ACCEPTED/PICKED -> Should be kept
        self.assertIn('ORD003', final_report['order_id'].values)
        # ORD004: READY TO SHIP -> Should be kept
        self.assertIn('ORD004', final_report['order_id'].values)
        
        # 4. Check duplicate removal
        # ORD001 is duplicate in rows index 0 and 6. Duplicate order_id should be removed, keeping first.
        # Length of ORD001 occurrences in final report should be 1
        ord001_rows = final_report[final_report['order_id'] == 'ORD001']
        self.assertEqual(len(ord001_rows), 1)

        # 5. Check New Status Alert worksheet
        # Conditions: status = New, ordered_date < ref_time - 1 hour
        # Row 1 (ORD002): status New, 07:30 (1.5 hours ago) -> Alert (Yes)
        # Row 5 (ORD006): status New, 07:00 (2 hours ago) -> Alert (Yes)
        # Row 0 (ORD001): status New, 08:30 (31 mins ago) -> No Alert (No)
        # Row 4 (ORD005): status Shipped (not New) -> No Alert
        self.assertIn('ORD002', new_status_alert['order_id'].values)
        self.assertIn('ORD006', new_status_alert['order_id'].values)
        self.assertNotIn('ORD001', new_status_alert['order_id'].values)
        self.assertNotIn('ORD005', new_status_alert['order_id'].values)
        
        # 6. Check Pivot Summary structure
        # Must have grand total columns and rows
        self.assertIn('Grand Total', pivot_summary.columns)
        self.assertIn('Grand Total', pivot_summary.index)
        # Ensure date columns are formatted as d-MMM (e.g., '16-Jul')
        self.assertIn('16-Jul', pivot_summary.columns)

if __name__ == '__main__':
    unittest.main()
