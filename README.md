# TC Report Processor

A complete Python-based automated solution to process TC Reports (CSV), identify pending orders requiring ERP investigation, generate a cleaned final report, generate new status alerts, and create a pivot summary. 

This repository includes both a **Streamlit Web Application** for interactive browser-based processing, and a **Command Line Interface (CLI)** script for automation.

## 📊 Features

1. **Raw Data Backup**: Automatically preserves the original imported data inside a `Raw Data` sheet without modifications.
2. **Robust Data Cleaning**: Automatically trims whitespace and converts order numbers to text to avoid scientific notation or auto-formatting.
3. **Smart Column Mapping**: Automatically detects required columns based on common names and allows custom overrides via UI dropdowns or CLI arguments.
4. **Order Filter for ERP Investigation**: Retains only `New`, `ACCEPTED/PICKED`, and `READY TO SHIP` orders with blank ERP IDs (filtering out specific payment and status combinations like non-COD pending orders).
5. **New Status Alert**: Automatically identifies orders that have been in the `New` status for more than 1 hour.
6. **Pivot Table Summary**: Automatically generates a pivot summary table grouping order counts by Marketplace Channel (rows) and order dates (columns, formatted as `DD-MMM` like `14-Jul`). It includes row totals, column totals, and overall grand totals.

---

## 📁 Repository Structure

```text
├── .gitignore               # Standard git ignore definitions
├── README.md                # Project documentation
├── requirements.txt         # Package dependencies
├── processor.py             # Core data processing pipeline (Parts 1-5)
├── app.py                   # Streamlit web dashboard (Part 6)
├── process_report.py        # Command-line interface script (Part 6)
└── test_process_report.py   # Unit test suite
```

---

## ⚙️ Installation & Setup

1. **Ensure Python 3.8+** is installed on your computer.
2. Navigate to the project directory:
   ```bash
   cd order_processor
   ```
3. Create a Python virtual environment:
   ```bash
   python -m venv venv
   ```
4. Activate the virtual environment:
   * **Windows (Command Prompt / PowerShell)**:
     ```powershell
     .\venv\Scripts\activate
     ```
   * **macOS / Linux**:
     ```bash
     source venv/bin/activate
     ```
5. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

---

## 🖥️ How to Run

### Option A: Streamlit Web Dashboard (Recommended)

To run the interactive web interface:
```bash
streamlit run app.py
```
This will automatically launch the app in your default web browser (usually at `http://localhost:8501`).
* **Upload** your TC Report CSV file.
* **Review Column Mappings** in the sidebar. The app auto-detects them, but you can manually adjust them if needed.
* **Process & Download**: Click "Process report" and download the complete multi-sheet Excel file.

### Option B: Command Line Interface (CLI)

To process files from your terminal or scripts:
```bash
python process_report.py <path_to_input.csv> <path_to_output.xlsx>
```

#### CLI Options:
* `--current-time "YYYY-MM-DD HH:MM:SS"`: Specify a custom date/time to calculate the 1-hour threshold for the New Status Alert. Defaults to the current system time.
* Column override flags (e.g., `--marketplace-channel channel` or `--payment-methods payment_methods`) to map custom headers.

Example:
```bash
python process_report.py sample_tc_report.csv output.xlsx --current-time "2026-07-16 09:01:59"
```

---

## 🧪 Running Unit Tests

To run the automated unit tests verifying the processing rules, filtering, and data integrity:
```bash
python -m unittest test_process_report.py
```
*(All tests should pass successfully)*

---

## 📋 CSV Data Schema

The tool identifies columns by their header name (case-insensitive and trimmed of space). By default, it looks for standard names or fallback indices:

| Column | Fallback Index | Default Search Patterns | Description |
| :--- | :--- | :--- | :--- |
| **Marketplace Channel** | 0 | `Marketplace Channel`, `channel`, `marketplace` | Used for row grouping in the Pivot Summary |
| **order_number** | 1 | `order_number`, `order_no`, `order number` | Treated as text/general format |
| **payment_status** | 6 | `payment_status`, `payment status` | Used for Part 2 filtering rules |
| **order_id** | 7 | `order_id`, `order id` | Used to remove duplicates |
| **order_status** | 8 | `order_status`, `order status` | Included in final report |
| **order_item_status** | 9 | `order_item_status`, `order_item status` | Retains only `New`, `ACCEPTED/PICKED`, `READY TO SHIP` |
| **payment_methods** | 60 | `payment_methods`, `payment_method` | Used for Part 2 filtering rules |
| **erp_reference_id** | 67 | `erp_reference_id`, `erp_ref_id` | Evaluates only blank/whitespace values |
