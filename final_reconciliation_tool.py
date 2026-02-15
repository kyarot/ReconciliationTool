#!/usr/bin/env python3
"""
Financial Data Reconciliation Tool - Complete Solution
Extract PDF financial data, compare with Excel, generate detailed report
"""

import pandas as pd
import pdfplumber
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import re
import os
import warnings
warnings.filterwarnings('ignore')


class FinancialReconciliation:
    def __init__(self):
        # File paths
        self.pdf_path = "06302025_mason_resources,_inc._cfg-ias_2025ye.pdf"
        self.excel_path = "data.xlsx"
        self.output_path = "Financial_Reconciliation_Report.xlsx"
        
        # Data containers
        self.pdf_data = {}
        self.excel_data = {}
        self.matches = []
        self.mismatches = []
    
    def extract_pdf_data(self):
        """Extract financial data from PDF"""
        print("🔍 Extracting data from PDF...")
        
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                # Get specific pages based on our earlier analysis
                balance_sheet_text = ""
                income_statement_text = ""
                
                # Page 7: Balance Sheet (Statement of Financial Position)
                if len(pdf.pages) >= 7:
                    text = pdf.pages[6].extract_text() or ""  # Page 7 (0-indexed)
                    if "statement of financial position" in text.lower() or "total assets" in text.lower():
                        balance_sheet_text = text
                        print(f"📊 Found balance sheet on page 7")
                
                # Page 8: Income Statement (Statement of Loss)
                if len(pdf.pages) >= 8:
                    text = pdf.pages[7].extract_text() or ""  # Page 8 (0-indexed)
                    if "statement of loss" in text.lower() or "expenses" in text.lower():
                        income_statement_text = text
                        print(f"💰 Found income statement on page 8")
                
                # Also check page 9 for cash flow statement
                cash_flow_text = ""
                if len(pdf.pages) >= 9:
                    text = pdf.pages[8].extract_text() or ""  # Page 9 (0-indexed)
                    if "cash flow" in text.lower():
                        cash_flow_text = text
                        print(f"💸 Found cash flow statement on page 9")
                
                # Extract financial data
                data = {}
                
                if balance_sheet_text:
                    # Total Assets - look for exact pattern from Mason Resources
                    match = re.search(r'Total assets\s*\$\s*([\d,]+)', balance_sheet_text, re.IGNORECASE)
                    if match:
                        data["Total Assets"] = self.clean_number(match.group(1))
                    
                    # Current Liabilities - Accounts payable and accrued liabilities
                    match = re.search(r'Accounts payable and accrued liabilities.*?\$\s*([\d,]+)', balance_sheet_text, re.IGNORECASE)
                    if match:
                        data["Current Liabilities"] = self.clean_number(match.group(1))
                    
                    # Total Liabilities
                    match = re.search(r'Total liabilities\s*([\d,]+)', balance_sheet_text, re.IGNORECASE)
                    if match:
                        data["Total Liabilities"] = self.clean_number(match.group(1))
                    
                    # Total Equity
                    match = re.search(r'Total equity\s*([\d,]+)', balance_sheet_text, re.IGNORECASE)
                    if match:
                        data["Total Equity"] = self.clean_number(match.group(1))
                    
                    # Cash
                    match = re.search(r'Cash\s*\$\s*([\d,]+)', balance_sheet_text, re.IGNORECASE)
                    if match:
                        data["Cash"] = self.clean_number(match.group(1))
                    
                    # Investments
                    match = re.search(r'Investments.*?\$\s*([\d,]+)', balance_sheet_text, re.IGNORECASE)
                    if match:
                        data["Investments"] = self.clean_number(match.group(1))
                
                # Extract revenue from income statement
                if income_statement_text:
                    # Look for revenue patterns
                    patterns = [
                        r'Revenue\s*\$\s*([\d,]+)',
                        r'Sales\s*\$\s*([\d,]+)',
                        r'Net\s+sales\s*\$\s*([\d,]+)',
                        r'Total\s+revenue\s*\$\s*([\d,]+)'
                    ]
                    for pattern in patterns:
                        match = re.search(pattern, income_statement_text, re.IGNORECASE)
                        if match:
                            data["Revenue"] = self.clean_number(match.group(1))
                            break
                
                # Add company metadata
                data["Company"] = "Mason Resources Inc."
                data["Currency"] = "CAD"
                data["Period"] = "June 30, 2025"
                
                self.pdf_data = data
                print(f"✅ Extracted {len([k for k, v in data.items() if isinstance(v, (int, float))])} financial metrics from PDF")
                return data
                
        except Exception as e:
            print(f"❌ Error reading PDF: {e}")
            return {}
    
    def clean_number(self, num_str):
        """Convert string number to float"""
        if not num_str:
            return None
        cleaned = re.sub(r'[,\s]', '', str(num_str))
        try:
            return float(cleaned)
        except:
            return None
    
    def load_excel_data(self):
        """Load Excel data and find matching items"""
        print("📊 Loading Excel data...")
        
        try:
            df = pd.read_excel(self.excel_path)
            print(f"✅ Loaded Excel: {df.shape[0]} rows, {df.shape[1]} columns")
            
            excel_data = {}
            
            # Map PDF items to Excel data
            mappings = {
                "Total Assets": ["total assets", "assets"],
                "Current Liabilities": ["current liabilities", "payables", "accounts payable"],
                "Total Liabilities": ["total liabilities", "liabilities"],
                "Total Equity": ["equity", "total equity"],
                "Cash": ["cash"],
                "Investments": ["investments"],
                "Revenue": ["revenue", "sales", "net sales", "total revenue"]
            }
            
            # Search through Excel data
            for pdf_item, excel_patterns in mappings.items():
                value = self.find_excel_value(df, excel_patterns)
                if value is not None:
                    excel_data[pdf_item] = value
            
            self.excel_data = excel_data
            print(f"✅ Found {len(excel_data)} matching items in Excel")
            return excel_data
            
        except Exception as e:
            print(f"❌ Error reading Excel: {e}")
            return {}
    
    def find_excel_value(self, df, patterns):
        """Find value in Excel based on patterns"""
        for pattern in patterns:
            # Search column headers
            for col in df.columns:
                if pattern.lower() in str(col).lower():
                    for value in df[col].dropna():
                        if isinstance(value, (int, float)):
                            return float(value)
                        elif isinstance(value, str):
                            num = self.extract_number_from_text(value)
                            if num:
                                return num
            
            # Search cell values
            for col in df.columns:
                for idx, cell_value in df[col].items():
                    if isinstance(cell_value, str) and pattern.lower() in cell_value.lower():
                        # Found matching text, look for number in same row
                        for other_col in df.columns:
                            row_val = df.loc[idx, other_col]
                            if isinstance(row_val, (int, float)) and not pd.isna(row_val):
                                return float(row_val)
                            elif isinstance(row_val, str):
                                num = self.extract_number_from_text(row_val)
                                if num:
                                    return num
        return None
    
    def extract_number_from_text(self, text):
        """Extract number from text string"""
        if not isinstance(text, str):
            return None
        
        # Remove currency symbols
        cleaned = re.sub(r'[\$€£]', '', text)
        match = re.search(r'-?[\d,]+\.?\d*', cleaned)
        if match:
            try:
                return float(match.group().replace(',', ''))
            except:
                pass
        return None
    
    def compare_data(self):
        """Compare PDF and Excel data"""
        print("🔄 Comparing PDF vs Excel data...")
        
        self.matches = []
        self.mismatches = []
        
        for item_name in self.pdf_data:
            if item_name in ["Company", "Currency", "Period"]:
                continue  # Skip metadata
            
            pdf_val = self.pdf_data.get(item_name)
            excel_val = self.excel_data.get(item_name)
            
            if pdf_val is not None and excel_val is not None:
                difference = abs(pdf_val - excel_val)
                percentage_diff = (difference / pdf_val * 100) if pdf_val != 0 else 0
                
                if percentage_diff == 0.0:  # Exact match only
                    self.matches.append({
                        "Item": item_name,
                        "PDF": pdf_val,
                        "Excel": excel_val,
                        "Difference": difference,
                        "Diff%": percentage_diff
                    })
                else:
                    self.mismatches.append({
                        "Item": item_name,
                        "PDF": pdf_val,
                        "Excel": excel_val,
                        "Difference": difference,
                        "Diff%": percentage_diff
                    })
        
        print(f"✅ Comparison complete: {len(self.matches)} matches, {len(self.mismatches)} mismatches")
        return self.matches, self.mismatches
    
    def create_report(self):
        """Create detailed Excel report"""
        print("📝 Creating detailed Excel report...")
        
        wb = Workbook()
        
        # 1. SUMMARY SHEET
        summary = wb.active
        summary.title = "📊 Summary"
        
        # Title and company info
        summary["A1"] = "Financial Data Reconciliation Report"
        summary["A1"].font = Font(size=16, bold=True)
        summary["A3"] = f"Company: {self.pdf_data.get('Company', 'Unknown')}"
        summary["A4"] = f"Currency: {self.pdf_data.get('Currency', 'Unknown')}"
        summary["A5"] = f"Period: {self.pdf_data.get('Period', 'Unknown')}"
        
        # Statistics
        total_compared = len(self.matches) + len(self.mismatches)
        match_rate = (len(self.matches) / total_compared * 100) if total_compared > 0 else 0
        
        summary["A7"] = "RECONCILIATION RESULTS"
        summary["A7"].font = Font(size=14, bold=True)
        summary["A8"] = f"Total Items Compared: {total_compared}"
        summary["A9"] = f"✅ Matches: {len(self.matches)}"
        summary["A10"] = f"❌ Mismatches: {len(self.mismatches)}"
        summary["A11"] = f"Match Rate: {match_rate:.1f}%"
        
        # Color code match rate
        if match_rate >= 90:
            summary["A11"].fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        elif match_rate >= 70:
            summary["A11"].fill = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")
        else:
            summary["A11"].fill = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")
        
        # 2. MATCHES SHEET
        if self.matches:
            matches_sheet = wb.create_sheet("✅ Matches")
            headers = ["Item", "PDF Value", "Excel Value", "Difference", "Diff %"]
            
            # Headers
            for col, header in enumerate(headers, 1):
                cell = matches_sheet.cell(1, col, header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
            
            # Data
            for row, match in enumerate(self.matches, 2):
                matches_sheet.cell(row, 1, match["Item"])
                matches_sheet.cell(row, 2, f"${match['PDF']:,.0f}")
                matches_sheet.cell(row, 3, f"${match['Excel']:,.0f}")
                matches_sheet.cell(row, 4, f"${match['Difference']:,.0f}")
                matches_sheet.cell(row, 5, f"{match['Diff%']:.2f}%")
        
        # 3. MISMATCHES SHEET
        if self.mismatches:
            mismatches_sheet = wb.create_sheet("❌ Mismatches")
            headers = ["Item", "PDF Value", "Excel Value", "Difference", "Diff %"]
            
            # Headers
            for col, header in enumerate(headers, 1):
                cell = mismatches_sheet.cell(1, col, header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")
            
            # Data
            for row, mismatch in enumerate(self.mismatches, 2):
                mismatches_sheet.cell(row, 1, mismatch["Item"])
                mismatches_sheet.cell(row, 2, f"${mismatch['PDF']:,.0f}")
                mismatches_sheet.cell(row, 3, f"${mismatch['Excel']:,.0f}")
                mismatches_sheet.cell(row, 4, f"${mismatch['Difference']:,.0f}")
                mismatches_sheet.cell(row, 5, f"{mismatch['Diff%']:.2f}%")
        
        # Auto-adjust column widths
        for sheet in wb.worksheets:
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                sheet.column_dimensions[column_letter].width = min(max_length + 2, 50)
        
        # Save report
        wb.save(self.output_path)
        print(f"✅ Report saved: {self.output_path}")
    
    def run(self):
        """Main execution method"""
        print("🚀 Starting Financial Data Reconciliation")
        print("=" * 60)
        
        # Step 1: Extract PDF data
        pdf_data = self.extract_pdf_data()
        if not pdf_data:
            print("❌ Failed to extract PDF data")
            return
        
        # Step 2: Load Excel data
        excel_data = self.load_excel_data()
        
        # Step 3: Show extracted data
        print(f"\n📋 EXTRACTED DATA ({pdf_data.get('Company', 'Unknown')}):")
        print("-" * 50)
        for key, value in pdf_data.items():
            if isinstance(value, (int, float)):
                print(f"  {key:.<30} ${value:,.0f}")
            else:
                print(f"  {key:.<30} {value}")
        
        # Step 4: Compare data
        matches, mismatches = self.compare_data()
        
        # Step 5: Show results
        print(f"\n🔍 COMPARISON RESULTS:")
        print("-" * 50)
        
        if matches:
            print(f"\n✅ MATCHES ({len(matches)}):")
            for match in matches:
                print(f"  {match['Item']:.<25} PDF: ${match['PDF']:,.0f} | Excel: ${match['Excel']:,.0f}")
        
        if mismatches:
            print(f"\n❌ MISMATCHES ({len(mismatches)}):")
            for mismatch in mismatches:
                print(f"  {mismatch['Item']:.<25} PDF: ${mismatch['PDF']:,.0f} | Excel: ${mismatch['Excel']:,.0f} | Diff: {mismatch['Diff%']:.1f}%")
        
        # Items only in PDF
        pdf_only = [k for k in pdf_data.keys() if k not in ["Company", "Currency", "Period"] and k not in excel_data]
        if pdf_only:
            print(f"\n📄 PDF ONLY:")
            for item in pdf_only:
                if isinstance(pdf_data[item], (int, float)):
                    print(f"  {item:.<25} ${pdf_data[item]:,.0f}")
        
        # Items only in Excel  
        excel_only = [k for k in excel_data.keys() if k not in pdf_data]
        if excel_only:
            print(f"\n📊 EXCEL ONLY:")
            for item in excel_only:
                print(f"  {item:.<25} ${excel_data[item]:,.0f}")
        
        # Step 6: Create report
        self.create_report()
        
        # Final summary
        total_compared = len(matches) + len(mismatches)
        match_rate = (len(matches) / total_compared * 100) if total_compared > 0 else 0
        
        print(f"\n🎉 RECONCILIATION COMPLETE!")
        print("=" * 60)
        print(f"📊 {len(matches)} matches, {len(mismatches)} mismatches ({match_rate:.1f}% match rate)")
        print(f"📁 Detailed report: {self.output_path}")
        print("=" * 60)


def main():
    """Main function"""
    # Check if files exist
    current_dir = os.getcwd()
    pdf_file = "06302025_mason_resources,_inc._cfg-ias_2025ye.pdf"
    excel_file = "data.xlsx"
    
    if not os.path.exists(pdf_file):
        print(f"❌ PDF file not found: {pdf_file}")
        print("Please ensure the PDF file is in the same directory as this script")
        return
    
    if not os.path.exists(excel_file):
        print(f"❌ Excel file not found: {excel_file}")
        print("Please ensure the Excel file is in the same directory as this script")
        return
    
    print(f"📄 PDF: {pdf_file}")
    print(f"📊 Excel: {excel_file}")
    
    # Run reconciliation
    reconciler = FinancialReconciliation()
    reconciler.run()


if __name__ == "__main__":
    main()