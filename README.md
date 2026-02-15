# Financial Data Reconciliation Tool

A professional web application for reconciling financial data between PDF statements and Excel reference files.

## Features

- **PDF Data Extraction**: Automatically extracts financial values from PDF statements
- **Excel Integration**: Reads and compares against Excel reference data
- **Zero-Tolerance Matching**: Exact numerical matching for financial accuracy
- **Professional Interface**: Clean white theme with black text for corporate use
- **Export Functionality**: Download detailed reconciliation reports as Excel files
- **Real-time Processing**: Live progress tracking during reconciliation

## How to Use

1. **Upload PDF**: Upload your financial statement PDF file
2. **Upload Excel**: Upload your Excel reference data file
3. **Process Files**: Click "Process Files" to start reconciliation
4. **Review Results**: View matches, mismatches, and summary statistics
5. **Download Report**: Export detailed reconciliation report

## Installation

1. Clone this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the application:
   ```bash
   streamlit run app.py
   ```

## Requirements

- Python 3.7+
- Streamlit
- pdfplumber
- pandas
- openpyxl
- xlsxwriter

## File Structure

- `app.py` - Main Streamlit web application
- `final_reconciliation_tool.py` - Core reconciliation engine
- `requirements.txt` - Python dependencies
- `start_app.sh` - Launch script

## Professional Features

- Clean, emoji-free interface suitable for corporate environments
- Professional black and white theme
- Comprehensive error handling and validation
- Detailed logging and reporting

## License

This project is private and proprietary.