import pandas as pd
import pdfplumber
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import re
import os
import json
import warnings
from dotenv import load_dotenv
from google import genai

load_dotenv()
warnings.filterwarnings('ignore')

class FinancialReconciliation:
    def __init__(self):
        self.pdf_path = None
        self.excel_path = None
        self.output_path = None
        self.custom_rules = ""
        
        self.pdf_data = {}
        self.excel_data = {}
        self.matches = []
        self.mismatches = []
        
        # Core requested 13 items
        self.expected_fields = [
            "Total assets",
            "Current Liability",
            "Common Equity",
            "Minority Interest / Non Controlling Interest",
            "Total Liability",
            "Common Dividend",
            "Sales",
            "Shares outstanding",
            "Shares diluted eps",
            "Preferred dividend",
            "Deferred tax",
            "Share capital",
            "Par value per share"
        ]
        
    def _parse_custom_rules(self):
        """Parse text box rules into simple keyword overrides. (Format: FIELD_NAME=keyword1,keyword2)"""
        rule_dict = {}
        if not self.custom_rules: return rule_dict
        for line in self.custom_rules.split('\n'):
            if '=' in line:
                k, v = line.split('=', 1)
                rule_dict[k.strip().lower()] = [x.strip().lower() for x in v.split(',')]
        return rule_dict

    def _extract_number(self, text, keywords, text_block):
        # A basic heuristic: search for keywords in text block, get the next number
        for keyword in keywords:
            # We look for the keyword, then capture the next formatted number
            pattern = re.compile(rf'{keyword}.*?([\(\-]?[\$\s]*[\d,]+\.?\d*[\)]?)', re.IGNORECASE | re.DOTALL)
            match = pattern.search(text_block)
            if match:
                num_str = match.group(1).replace('$', '').replace(',', '').strip()
                if not num_str: continue
                is_negative = False
                if '(' in num_str and ')' in num_str:
                    is_negative = True
                    num_str = num_str.replace('(', '').replace(')', '')
                if num_str.startswith('-'):
                    is_negative = True
                    num_str = num_str[1:]
                try:
                    val = float(num_str)
                    return -val if is_negative else val
                except ValueError:
                    continue
        return None

    def extract_pdf_data(self):
        print(f"🔍 Extracting data from PDF via Gemini 1.5 Pro: {self.pdf_path}...")
        try:
            full_text = ""
            with pdfplumber.open(self.pdf_path) as pdf:
                for page in pdf.pages:
                    full_text += (page.extract_text() or "") + "\n"
                    
            text = full_text.replace('\xa0', ' ')
            
            # Setup Gemini client
            api_key = os.getenv("API_KEY") or os.getenv("GEMINI_API_KEY")
            if not api_key:
                print("❌ ERROR: API_KEY not found in .env file.")
                return {}
                
            client = genai.Client(api_key=api_key)
            
            prompt = f"""
You are an expert financial analyst. Your task is to extract ONE specific reporting period of data from the provided financial statement text. 
CRITICAL RULE: Always extract the MOST RECENT period provided. Prioritize 3-month (quarterly) data over 6/9/12 months if both are presented on the Income statement/Cash Flow.
CRITICAL RULE: Extract ONLY consolidated data if both consolidated and unconsolidated are present.

You must extract EXACTLY the 13 line items below, following the strict spreading mappings provided.

MAPPINGS TO FOLLOW:
1. Total assets = Total assets from face of balance sheet + Netting of Deferred tax assets
2. Current Liability = Current Liability from face of balance sheet (No adjustments needed)
3. Common Equity = Shareholders' funds (excluding non-controlling & minority interests) minus Preferred Stock minus Other Shares including Hybrids as Part of Equity
4. Minority Interest / Non Controlling Interest = Combined value of all non-controlling/minority interests regardless of location (equity, liability, or independent).
5. Total Liability = Total liabilities + Non-controlling & minority interests (all types) + Loans from non-controlling interest shareholders + Other Shares including Hybrids + Netting of DTL minus Non-Equity Reserves
6. Common Dividend = Cash dividends paid to Common/Equity Shareholders ONLY from the Cash flow statement (if available) or Changes in Equity. Exclude minority interest dividends unless impossible to bifurcate.
7. Sales = Operating revenue (Turnover) minus Gains/losses from hedging, Foreign exchange, Government grants, and Excise duties. (If "Revenue" is presented, use that).
8. Shares outstanding = Number of common/ordinary/basic shares issued and outstanding (excluding treasury). If in units, convert to actual shares.
9. Shares diluted eps = Diluted weighted average shares.
10. Preferred dividend = Total cash dividends to preferred shareholders.
11. Deferred tax = Deferred tax from Consolidated financial statement.
12. Share capital = From balance sheet or table of equity.
13. Par value per share = Par value.

CUSTOM RULES OVERRIDE:
The user has provided the following custom rules that override or assist the above logic. Pay careful attention to these:
{self.custom_rules}

FINANCIAL STATEMENT TEXT:
{text[:90000]} # Limiting to ~90k chars to be safe

Output your response as a STRICT JSON object with NO markdown formatting, NO backticks, and NO extra text. 
Use EXACTLY these keys:
"Total assets" (number or null)
"Current Liability" (number or null)
"Common Equity" (number or null)
"Minority Interest / Non Controlling Interest" (number or null)
"Total Liability" (number or null)
"Common Dividend" (number or null)
"Sales" (number or null)
"Shares outstanding" (number or null)
"Shares diluted eps" (number or null)
"Preferred dividend" (number or null)
"Deferred tax" (number or null)
"Share capital" (number or null)
"Par value per share" (number or null)
"Company" (string)
"Period" (string)
"Currency" (string, e.g., "USD")
"Months_IS" (number of months presented in income statement, e.g., 3)
"Months_CF" (number of months presented in cash flow, e.g., 3)
"Units" (string, e.g., "Millions")
"""

            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=prompt
            )
            
            raw_json = response.text.replace('```json', '').replace('```', '').strip()
            data = json.loads(raw_json)
            self.pdf_data = data
            print(f"✅ Gemini Extraction Successful!")
            return data
            
        except Exception as e:
            print(f"❌ Error reading PDF via Gemini: {e}")
            return {}



    def load_excel_data(self):
        print("📊 Loading Excel data...")
        try:
            df = pd.read_excel(self.excel_path)
            excel_data = {}
            rules = self._parse_custom_rules()
            
            def get_kw(default_kws, rule_key):
                if rule_key in rules:
                    return rules[rule_key] + default_kws
                return default_kws

            excel_mappings = {
                "Total assets": get_kw(["total assets", "assets"], 'total assets'),
                "Current Liability": get_kw(["current liability", "current liabilities"], 'current liability'),
                "Common Equity": get_kw(["common equity", "shareholders equity", "total equity"], 'common equity'),
                "Minority Interest / Non Controlling Interest": get_kw(["minority interest", "non controlling", "nci"], 'minority interest'),
                "Total Liability": get_kw(["total liability", "total liabilities"], 'total liability'),
                "Common Dividend": get_kw(["common dividend", "dividend paid"], 'common dividend'),
                "Sales": get_kw(["sales", "revenue", "net sales"], 'sales'),
                "Shares outstanding": get_kw(["shares outstanding", "basic shares"], 'shares outstanding'),
                "Shares diluted eps": get_kw(["diluted shares", "diluted eps"], 'shares diluted eps'),
                "Preferred dividend": get_kw(["preferred dividend"], 'preferred dividend'),
                "Deferred tax": get_kw(["deferred tax"], 'deferred tax'),
                "Share capital": get_kw(["share capital", "capital stock"], 'share capital'),
                "Par value per share": get_kw(["par value"], 'par value')
            }

            for item, patterns in excel_mappings.items():
                val = self.find_excel_value(df, patterns)
                if val is not None:
                    excel_data[item] = val

            self.excel_data = excel_data
            return excel_data
        except Exception as e:
            print(f"❌ Error reading Excel: {e}")
            return {}

    def find_excel_value(self, df, patterns):
        for pattern in patterns:
            for col in df.columns:
                if pattern.lower() in str(col).lower():
                    for value in df[col].dropna():
                        if isinstance(value, (int, float)):
                            return float(value)
                        elif isinstance(value, str):
                            num = self._extract_number_from_str(value)
                            if num is not None: return num
                        
            for col in df.columns:
                for idx, cell_value in df[col].items():
                    if isinstance(cell_value, str) and pattern.lower() in cell_value.lower():
                        for other_col in df.columns:
                            row_val = df.loc[idx, other_col]
                            if isinstance(row_val, (int, float)) and pd.notna(row_val):
                                return float(row_val)
                            elif isinstance(row_val, str):
                                num = self._extract_number_from_str(row_val)
                                if num is not None: return num
        return None

    def _extract_number_from_str(self, text):
        if not isinstance(text, str): return None
        cleaned = re.sub(r'[\$€£,]', '', text)
        match = re.search(r'-?\d+\.?\d*', cleaned)
        if match:
            try: return float(match.group())
            except: pass
        return None

    def compare_data(self):
        self.matches = []
        self.mismatches = []
        
        for item_name in self.expected_fields:
            pdf_val = self.pdf_data.get(item_name)
            excel_val = self.excel_data.get(item_name)
            
            p_val = pdf_val if pdf_val is not None else 0
            e_val = excel_val if excel_val is not None else 0

            difference = abs(p_val - e_val)
            percentage_diff = (difference / abs(p_val) * 100) if p_val != 0 else 0
            
            if difference < 0.01:
                self.matches.append({
                    "Item": item_name, "PDF": p_val, "Excel": e_val, 
                    "PDF_Raw": pdf_val, "Excel_Raw": excel_val,
                    "Difference": difference, "Diff%": percentage_diff
                })
            else:
                self.mismatches.append({
                    "Item": item_name, "PDF": p_val, "Excel": e_val, 
                    "PDF_Raw": pdf_val, "Excel_Raw": excel_val,
                    "Difference": difference, "Diff%": percentage_diff
                })
        return self.matches, self.mismatches

    def create_report(self):
        wb = Workbook()
        summary = wb.active
        summary.title = "📊 Summary"
        summary["A1"] = f"Reconciliation Report: {self.pdf_data.get('Company', 'Unknown')}"
        summary["A1"].font = Font(size=14, bold=True)
        summary["A3"] = f"Period: {self.pdf_data.get('Period', '')}"
        summary["A4"] = f"Currency: {self.pdf_data.get('Currency', '')}"
        summary["A5"] = f"Months_IS: {self.pdf_data.get('Months_IS', '')}"
        summary["A6"] = f"Months_CF: {self.pdf_data.get('Months_CF', '')}"
        summary["A7"] = f"Units: {self.pdf_data.get('Units', '')}"
        
        total = len(self.matches) + len(self.mismatches)
        rate = (len(self.matches)/total*100) if total>0 else 0
        summary["A9"] = f"Match Rate: {rate:.1f}%"

        details = wb.create_sheet("Detailed Comparison")
        headers = ["Item", "PDF Value", "Excel Value", "Difference", "Status"]
        for col, h in enumerate(headers, 1):
            cell = details.cell(1, col, h)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
            
        row = 2
        for item in self.expected_fields:
            pdf_v = self.pdf_data.get(item)
            exc_v = self.excel_data.get(item)
            
            p_val = pdf_v if pdf_v is not None else 0
            e_val = exc_v if exc_v is not None else 0
            diff = abs(p_val - e_val)
            status = "MATCH" if diff < 0.01 else "MISMATCH"
            
            p_out = pdf_v if pdf_v is not None else "NA"
            e_out = exc_v if exc_v is not None else "NA"
            
            details.cell(row, 1, item)
            c2 = details.cell(row, 2, p_out)
            if p_out != "NA": c2.number_format = '#,##0.00'
            c3 = details.cell(row, 3, e_out)
            if e_out != "NA": c3.number_format = '#,##0.00'
            
            details.cell(row, 4, diff).number_format = '#,##0.00'
            cell = details.cell(row, 5, status)
            cell.font = Font(color="00A300" if status=="MATCH" else "FF0000", bold=True)
            row += 1
            
        # Format column width
        for col in details.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            details.column_dimensions[column].width = min(max_length + 2, 40)
            
        wb.save(self.output_path)
        print(f"✅ Report saved: {self.output_path}")

    def run(self):
        self.extract_pdf_data()
        self.load_excel_data()
        self.compare_data()
        self.create_report()