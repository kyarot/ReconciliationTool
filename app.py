"""
Financial Reconciliation Web App
Professional frontend for financial data reconciliation tool
Features bulk uploads, custom mapping rules, and side-by-side comparison
"""

import streamlit as st
import pandas as pd
import os
import tempfile
import zipfile
from io import BytesIO
import base64
from final_reconciliation_tool import FinancialReconciliation
import time

def configure_page():
    st.set_page_config(
        page_title="Financial Data Reconciliation",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )

def apply_custom_css():
    st.markdown("""
    <style>
    .stApp { background-color: #ffffff; color: #000000; }
    .main-header { background-color: #ffffff; border: 2px solid #e9ecef; padding: 2rem; border-radius: 10px; margin-bottom: 2rem; text-align: center; }
    .main-header h1 { color: #000000; }
    .main-header p { color: #555555; }
    .upload-section { background-color: #ffffff; padding: 1.5rem; border-radius: 8px; border: 2px solid #e9ecef; margin-bottom: 1rem; }
    .stButton > button { background-color: #000000 !important; color: white !important; font-weight: 600 !important; }
    .stFileUploader label { color: #000000 !important; font-weight: 600 !important; }
    .metric-card { background-color: white; padding: 1rem; border-radius: 6px; border: 1px solid #dee2e6; text-align: center; }
    .metric-value { font-size: 2rem; font-weight: 700; color: #000000; }
    .match-item { background-color: #d4f6d4; border-left: 4px solid #28a745; padding: 0.75rem; margin: 0.25rem 0; border-radius: 4px; color:#000000; }
    .mismatch-item { background-color: #f8d7da; border-left: 4px solid #dc3545; padding: 0.75rem; margin: 0.25rem 0; border-radius: 4px; color:#000000; }
    .stTextArea label { color: #000000 !important; font-weight:bold; }
    p, h1, h2, h3, h4, span { color: #000000 !important; }
    </style>
    """, unsafe_allow_html=True)

def display_header():
    st.markdown("""
    <div class="main-header">
        <h1>Financial Data Reconciliation</h1>
        <p>Professional PDF & Excel Bulk Reconciliation Tool</p>
    </div>
    """, unsafe_allow_html=True)

def build_combined_excel(all_results):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for result in all_results:
            # Sheet names max length is 31 and cannot have certain special characters
            sheet_name = "".join(c for c in result['company'] if c.isalnum() or c in " _-")[:31] 
            if not sheet_name: sheet_name = "Company"
            
            # Create a dataframe for the matches and mismatches
            data = []
            for m in result['matches']:
                p_out = m['PDF'] if m.get('PDF_Raw') is not None else "NA"
                e_out = m['Excel'] if m.get('Excel_Raw') is not None else "NA"
                data.append({"Item": m['Item'], "PDF": p_out, "Excel": e_out, "Diff": m['Difference'], "Status": "MATCH"})
            for m in result['mismatches']:
                p_out = m['PDF'] if m.get('PDF_Raw') is not None else "NA"
                e_out = m['Excel'] if m.get('Excel_Raw') is not None else "NA"
                data.append({"Item": m['Item'], "PDF": p_out, "Excel": e_out, "Diff": m['Difference'], "Status": "MISMATCH"})
                
            df = pd.DataFrame(data)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Color code rules
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            match_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
            mismatch_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
            
            for row in range(1, len(data) + 1):
                status = data[row-1]["Status"]
                fmt = match_format if status == "MATCH" else mismatch_format
                worksheet.write(row, 4, status, fmt)
                
            # Formatting cols width
            for i, col in enumerate(df.columns):
                column_len = max(df[col].astype(str).map(len).max(), len(col))
                worksheet.set_column(i, i, column_len + 5)
                
    output.seek(0)
    return output.getvalue()

def display_results_ui(result):
    st.markdown(f"### Results for: {result['company']}")
    st.markdown(f"**Period**: {result['period']} | **Currency**: {result['currency']} | **Months_IS**: {result['months_is']} | **Months_CF**: {result['months_cf']} | **Units**: {result['units']}")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### Matches")
        if not result['matches']:
            st.info("No matches found.")
        for m in result['matches']:
            pdf_str = f"{m['PDF']:,.2f}" if m.get('PDF_Raw') is not None else "NA"
            exc_str = f"{m['Excel']:,.2f}" if m.get('Excel_Raw') is not None else "NA"
            st.markdown(f"""
            <div class="match-item">
                <strong>{m['Item']}</strong><br>
                PDF: {pdf_str} | Excel: {exc_str}
            </div>
            """, unsafe_allow_html=True)
            
    with col2:
        st.markdown("#### Mismatches")
        if not result['mismatches']:
            st.info("No mismatches found.")
        for m in result['mismatches']:
            pdf_str = f"{m['PDF']:,.2f}" if m.get('PDF_Raw') is not None else "NA"
            exc_str = f"{m['Excel']:,.2f}" if m.get('Excel_Raw') is not None else "NA"
            st.markdown(f"""
            <div class="mismatch-item">
                <strong>{m['Item']}</strong><br>
                PDF: {pdf_str} | Excel: {exc_str}<br>
                <small>Diff: {m['Difference']:,.2f}</small>
            </div>
            """, unsafe_allow_html=True)
    st.markdown("---")

def main():
    configure_page()
    apply_custom_css()
    display_header()
    
    with st.sidebar:
        st.markdown("### Settings")
        st.markdown("Upload multiple PDFs and Excels for bulk processing. The tool will parse pair iteratively.")
        
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### PDF Files")
        pdf_files = st.file_uploader("Upload PDF files", type=['pdf'], accept_multiple_files=True)
        
    with col2:
        st.markdown("### Excel Files")
        excel_files = st.file_uploader("Upload Excel files", type=['xlsx', 'xls'], accept_multiple_files=True)
        
    st.markdown("### AI Extraction Guidelines")
    st.markdown("Write natural language instructions for the Google Gemini AI parser.")
    st.markdown("Example: `For the Revenue, only look for 'Net Gaming Sales', ignore other revenues.`")
    custom_rules = st.text_area("Optional Gemini Extraction Rules", height=150)
    
    if st.button("Process Bulk Files", width="stretch"):
        if not pdf_files or not excel_files:
            st.error("Please upload at least one PDF and one Excel file.")
            return
            
        if len(pdf_files) != len(excel_files):
            st.warning(f"Uploaded {len(pdf_files)} PDFs and {len(excel_files)} Excel files. They will be paired in input order up to the lowest count.")
            
        process_count = min(len(pdf_files), len(excel_files))
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        all_results = []
        zip_buffer = BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i in range(process_count):
                pdf = pdf_files[i]
                exc = excel_files[i]
                status_text.text(f"Processing Pair {i+1}: {pdf.name} & {exc.name}...")
                
                with tempfile.TemporaryDirectory() as tmp_dir:
                    pdf_path = os.path.join(tmp_dir, pdf.name)
                    exc_path = os.path.join(tmp_dir, exc.name)
                    
                    with open(pdf_path, 'wb') as f: f.write(pdf.getbuffer())
                    with open(exc_path, 'wb') as f: f.write(exc.getbuffer())
                    
                    reconciler = FinancialReconciliation()
                    reconciler.pdf_path = pdf_path
                    reconciler.excel_path = exc_path
                    reconciler.custom_rules = custom_rules
                    
                    # Process
                    pdf_data = reconciler.extract_pdf_data()
                    reconciler.load_excel_data()
                    matches, mismatches = reconciler.compare_data()
                    
                    company_name = pdf_data.get('Company', f'Company_{i+1}')
                    # Clean company name for folder paths
                    clean_company_name = "".join(c for c in company_name if c.isalnum() or c in " _-")
                    if not clean_company_name: clean_company_name = f'Company_{i+1}'
                    
                    # Create report
                    report_name = f"{clean_company_name}_Reconciliation_Report.xlsx"
                    report_path = os.path.join(tmp_dir, report_name)
                    reconciler.output_path = report_path
                    reconciler.create_report()
                    
                    # Add to zip inside an individual folder named after the company
                    zip_file.write(report_path, arcname=f"{clean_company_name}/{report_name}")
                    
                    all_results.append({
                        "company": company_name,
                        "period": pdf_data.get('Period', 'Unknown'),
                        "currency": pdf_data.get('Currency', 'Unknown'),
                        "months_is": pdf_data.get('Months_IS', '3'),
                        "months_cf": pdf_data.get('Months_CF', '3'),
                        "units": pdf_data.get('Units', 'Unknown'),
                        "matches": matches,
                        "mismatches": mismatches
                    })
                    
                progress_bar.progress((i + 1) / process_count)
                
        status_text.text("Processing complete!")
        
        # Display side-by-side results in UI
        st.markdown("## Reconciliation Overview")
        for res in all_results:
            display_results_ui(res)
            
        # Compile total master Excel and ZIP
        combined_excel_data = build_combined_excel(all_results)
        
        st.markdown("### Download Results")
        dcol1, dcol2 = st.columns(2)
        with dcol1:
            st.download_button(
                label="Download Combined Summary Excel",
                data=combined_excel_data,
                file_name="Combined_Summary_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch"
            )
        with dcol2:
            st.download_button(
                label="Download Individual Report Folders (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Reconciliation_Reports.zip",
                mime="application/zip",
                width="stretch"
            )

if __name__ == "__main__":
    main()