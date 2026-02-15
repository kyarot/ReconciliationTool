"""
Financial Reconciliation Web App
Professional frontend for financial data reconciliation tool
"""

import streamlit as st
import pandas as pd
import os
import tempfile
from io import BytesIO
import base64
from final_reconciliation_tool import FinancialReconciliation
import time


def configure_page():
    """Configure Streamlit page settings"""
    st.set_page_config(
        page_title="Financial Data Reconciliation",
        page_icon=None,
        layout="wide",
        initial_sidebar_state="expanded"
    )


def apply_custom_css():
    """Apply professional black and white theme"""
    st.markdown("""
    <style>
    /* Main theme colors */
    .stApp {
        background-color: #ffffff;
        color: #000000;
    }
    
    /* Header styling - White theme */
    .main-header {
        background-color: #ffffff;
        border: 2px solid #e9ecef;
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        text-align: center;
        color: #000000;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    }
    
    .main-header h1 {
        color: #000000;
        font-size: 2.5rem;
        margin-bottom: 0.5rem;
        font-weight: 700;
    }
    
    .main-header p {
        color: #666666;
        font-size: 1.1rem;
        margin: 0;
    }
    
    /* Upload section */
    .upload-section {
        background-color: #ffffff;
        padding: 1.5rem;
        border-radius: 8px;
        border: 2px solid #e9ecef;
        margin-bottom: 1rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    
    .upload-section h3 {
        color: #000000;
        margin-bottom: 1rem;
        font-weight: 600;
    }
    
    /* Process button */
    .stButton > button {
        background-color: #000000 !important;
        color: white !important;
        font-weight: 600 !important;
        font-size: 1.1rem !important;
        padding: 0.75rem 2rem !important;
        border-radius: 8px !important;
        border: none !important;
        width: 100% !important;
        transition: all 0.3s !important;
    }
    
    .stButton > button:hover {
        background-color: #333333 !important;
        color: white !important;
    }
    
    .stButton > button:focus {
        background-color: #000000 !important;
        color: white !important;
        box-shadow: 0 0 0 2px rgba(0,0,0,0.3) !important;
    }
    
    .stButton > button:active {
        background-color: #000000 !important;
        color: white !important;
    }
    
    /* File uploader styling */
    .stFileUploader label {
        color: #000000 !important;
        font-weight: 600 !important;
    }
    
    .stFileUploader button {
        background-color: #000000 !important;
        color: white !important;
        border: 1px solid #000000 !important;
        font-weight: 500 !important;
    }
    
    .stFileUploader button:hover {
        background-color: #333333 !important;
        color: white !important;
        border: 1px solid #333333 !important;
    }
    
    .stFileUploader button:focus {
        background-color: #000000 !important;
        color: white !important;
        border: 1px solid #000000 !important;
    }
    
    /* File uploader browse files text */
    .stFileUploader button span {
        color: white !important;
    }
    
    /* Download button styling */
    .stDownloadButton > button {
        background-color: #000000 !important;
        color: white !important;
        font-weight: 600 !important;
        border: none !important;
        border-radius: 6px !important;
    }
    
    .stDownloadButton > button:hover {
        background-color: #333333 !important;
        color: white !important;
    }
    
    .stDownloadButton > button:focus {
        background-color: #000000 !important;
        color: white !important;
    }
    
    .stDownloadButton > button span {
        color: white !important;
    }
    
    /* Additional button text overrides */
    button[kind="primary"] {
        background-color: #000000 !important;
        color: white !important;
    }
    
    button[kind="secondary"] {
        background-color: #000000 !important;
        color: white !important;
    }
    
    /* Force all button text to be white */
    .stButton button, .stDownloadButton button, .stFileUploader button {
        color: white !important;
    }
    
    .stButton button *, .stDownloadButton button *, .stFileUploader button * {
        color: white !important;
    }
    
    /* Results section */
    .results-header {
        background-color: #ffffff;
        color: #000000;
        padding: 1rem;
        border-radius: 8px 8px 0 0;
        margin-bottom: 0;
        border: 2px solid #e9ecef;
        border-bottom: 1px solid #e9ecef;
    }
    
    .results-content {
        background-color: #ffffff;
        padding: 1.5rem;
        border: 2px solid #e9ecef;
        border-top: none;
        border-radius: 0 0 8px 8px;
    }
    
    /* Metrics styling */
    .metric-card {
        background-color: white;
        padding: 1rem;
        border-radius: 6px;
        border: 1px solid #dee2e6;
        text-align: center;
        margin: 0.5rem 0;
    }
    
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        color: #000000;
    }
    
    .metric-label {
        font-size: 0.9rem;
        color: #666666;
        text-transform: uppercase;
        font-weight: 500;
    }
    
    /* Match/Mismatch styling */
    .match-item {
        background-color: #d4f6d4;
        border-left: 4px solid #28a745;
        padding: 0.75rem;
        margin: 0.25rem 0;
        border-radius: 4px;
    }
    
    .mismatch-item {
        background-color: #f8d7da;
        border-left: 4px solid #dc3545;
        padding: 0.75rem;
        margin: 0.25rem 0;
        border-radius: 4px;
    }
    
    /* Download button */
    .download-button {
        background-color: #333333;
        color: white;
        padding: 0.75rem 1.5rem;
        border-radius: 6px;
        text-decoration: none;
        font-weight: 600;
        display: inline-block;
        margin: 1rem 0;
    }
    
    /* Sidebar */
    .css-1d391kg {
        background-color: #ffffff;
    }
    
    /* Sidebar text visibility */
    .css-1d391kg .markdown-text-container {
        color: #000000 !important;
    }
    
    /* Tables */
    .stDataFrame {
        border: 1px solid #dee2e6;
        border-radius: 6px;
    }
    
    /* Progress bar */
    .stProgress > div > div {
        background-color: #000000;
    }
    
    /* Text visibility improvements */
    .stMarkdown {
        color: #000000;
    }
    
    .stText {
        color: #000000;
    }
    
    /* Ensure all text is visible */
    p, h1, h2, h3, h4, h5, h6, span, div {
        color: #000000 !important;
    }
    
    /* Ensure all headers are black text */
    .main-header h1, 
    .main-header p,
    .results-header h3 {
        color: #000000 !important;
    }
    
    /* Info boxes - Clean white theme */
    .stInfo {
        background-color: #ffffff;
        border: 2px solid #0066cc;
        border-radius: 6px;
    }
    
    .stInfo .markdown-text-container {
        color: #000000 !important;
    }
    
    .stSuccess {
        background-color: #ffffff;
        border: 2px solid #28a745;
        border-radius: 6px;
    }
    
    .stSuccess .markdown-text-container {
        color: #000000 !important;
    }
    
    .stError {
        background-color: #ffffff;
        border: 2px solid #dc3545;
        border-radius: 6px;
    }
    
    .stError .markdown-text-container {
        color: #000000 !important;
    }
    
    .stWarning {
        background-color: #ffffff;
        border: 2px solid #ffc107;
        border-radius: 6px;
    }
    
    /* Spinner text */
    .stSpinner > div {
        color: #000000 !important;
    }
    </style>
    """, unsafe_allow_html=True)


def create_download_link(file_path, filename):
    """Create download link for files"""
    with open(file_path, 'rb') as f:
        data = f.read()
    b64 = base64.b64encode(data).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}" class="download-button">📥 Download {filename}</a>'
    return href


def display_header():
    """Display the main header"""
    st.markdown("""
    <div class="main-header">
        <h1>Financial Data Reconciliation</h1>
        <p>Professional PDF & Excel Reconciliation Tool</p>
    </div>
    """, unsafe_allow_html=True)


def upload_section():
    """File upload section"""
    st.markdown("## File Upload")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="upload-section">
            <h3>PDF Financial Statement</h3>
        </div>
        """, unsafe_allow_html=True)
        
        pdf_file = st.file_uploader(
            "Upload PDF file",
            type=['pdf'],
            key="pdf_uploader",
            help="Upload your financial statement PDF file"
        )
    
    with col2:
        st.markdown("""
        <div class="upload-section">
            <h3>Excel Reference Data</h3>
        </div>
        """, unsafe_allow_html=True)
        
        excel_file = st.file_uploader(
            "Upload Excel file",
            type=['xlsx', 'xls'],
            key="excel_uploader",
            help="Upload your Excel reference data file"
        )
    
    return pdf_file, excel_file


def display_metrics(matches, mismatches, total_compared, match_rate):
    """Display key metrics"""
    st.markdown("### Reconciliation Metrics")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{total_compared}</div>
            <div class="metric-label">Items Compared</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value" style="color: #28a745;">{len(matches)}</div>
            <div class="metric-label">Matches</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value" style="color: #dc3545;">{len(mismatches)}</div>
            <div class="metric-label">Mismatches</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        color = "#28a745" if match_rate >= 80 else "#ffc107" if match_rate >= 60 else "#dc3545"
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value" style="color: {color};">{match_rate:.1f}%</div>
            <div class="metric-label">Match Rate</div>
        </div>
        """, unsafe_allow_html=True)


def display_results(reconciler):
    """Display detailed results"""
    matches = reconciler.matches
    mismatches = reconciler.mismatches
    
    if not matches and not mismatches:
        st.warning("No data to compare. Please check your files.")
        return
    
    total_compared = len(matches) + len(mismatches)
    match_rate = (len(matches) / total_compared * 100) if total_compared > 0 else 0
    
    # Display metrics
    display_metrics(matches, mismatches, total_compared, match_rate)
    
    # Results section
    st.markdown("""
    <div class="results-header">
        <h3 style="margin: 0; color: white;">Detailed Results</h3>
    </div>
    <div class="results-content">
    """, unsafe_allow_html=True)
    
    # Matches
    if matches:
        st.markdown("#### Perfect Matches")
        for match in matches:
            st.markdown(f"""
            <div class="match-item">
                <strong>{match['Item']}</strong><br>
                PDF: ${match['PDF']:,.0f} | Excel: ${match['Excel']:,.0f}
                <span style="float: right; color: #28a745; font-weight: bold;">MATCH</span>
            </div>
            """, unsafe_allow_html=True)
    
    # Mismatches
    if mismatches:
        st.markdown("#### Mismatches")
        for mismatch in mismatches:
            st.markdown(f"""
            <div class="mismatch-item">
                <strong>{mismatch['Item']}</strong><br>
                PDF: ${mismatch['PDF']:,.0f} | Excel: ${mismatch['Excel']:,.0f}<br>
                <small>Difference: ${mismatch['Difference']:,.0f} ({mismatch['Diff%']:.2f}%)</small>
                <span style="float: right; color: #dc3545; font-weight: bold;">MISMATCH</span>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("</div>", unsafe_allow_html=True)
    
    # Data tables
    col1, col2 = st.columns(2)
    
    if matches:
        with col1:
            st.markdown("#### Matches Data")
            matches_df = pd.DataFrame(matches)
            matches_df['PDF'] = matches_df['PDF'].apply(lambda x: f"${x:,.0f}")
            matches_df['Excel'] = matches_df['Excel'].apply(lambda x: f"${x:,.0f}")
            matches_df['Difference'] = matches_df['Difference'].apply(lambda x: f"${x:,.0f}")
            matches_df['Diff%'] = matches_df['Diff%'].apply(lambda x: f"{x:.2f}%")
            st.dataframe(matches_df, use_container_width=True)
    
    if mismatches:
        with col2 if matches else st:
            st.markdown("#### Mismatches Data")
            mismatches_df = pd.DataFrame(mismatches)
            mismatches_df['PDF'] = mismatches_df['PDF'].apply(lambda x: f"${x:,.0f}")
            mismatches_df['Excel'] = mismatches_df['Excel'].apply(lambda x: f"${x:,.0f}")
            mismatches_df['Difference'] = mismatches_df['Difference'].apply(lambda x: f"${x:,.0f}")
            mismatches_df['Diff%'] = mismatches_df['Diff%'].apply(lambda x: f"{x:.2f}%")
            st.dataframe(mismatches_df, use_container_width=True)


def process_files(pdf_file, excel_file):
    """Process uploaded files"""
    try:
        # Create temporary directory
        with tempfile.TemporaryDirectory() as temp_dir:
            # Save uploaded files
            pdf_path = os.path.join(temp_dir, "uploaded.pdf")
            excel_path = os.path.join(temp_dir, "uploaded.xlsx")
            output_path = os.path.join(temp_dir, "reconciliation_result.xlsx")
            
            # Write files
            with open(pdf_path, "wb") as f:
                f.write(pdf_file.getbuffer())
            
            with open(excel_path, "wb") as f:
                f.write(excel_file.getbuffer())
            
            # Create modified reconciler instance
            reconciler = FinancialReconciliation()
            reconciler.pdf_path = pdf_path
            reconciler.excel_path = excel_path  
            reconciler.output_path = output_path
            
            # Process files
            with st.spinner("Extracting data from PDF..."):
                pdf_data = reconciler.extract_pdf_data()
                if not pdf_data:
                    st.error("Failed to extract data from PDF")
                    return None, None
            
            with st.spinner("Loading Excel data..."):
                excel_data = reconciler.load_excel_data()
            
            with st.spinner("Comparing data..."):
                reconciler.compare_data()
            
            with st.spinner("Generating report..."):
                reconciler.create_report()
            
            # Read the generated file
            with open(output_path, 'rb') as f:
                result_data = f.read()
            
            return reconciler, result_data
            
    except Exception as e:
        st.error(f"Error processing files: {str(e)}")
        return None, None


def main():
    """Main application"""
    configure_page()
    apply_custom_css()
    
    # Header
    display_header()
    
    # Sidebar
    with st.sidebar:
        st.markdown("### Settings")
        st.markdown("""
        **File Requirements:**
        - PDF: Financial statements
        - Excel: Reference data (.xlsx)
        
        **Matching Criteria:**
        - Exact matches only (0% tolerance)
        - Strict numerical comparison
        """)
        
        st.markdown("### Process Steps")
        st.markdown("""
        1. Upload PDF & Excel files
        2. Extract financial data
        3. Compare values
        4. Generate detailed report
        5. Download results
        """)
    
    # Main content
    pdf_file, excel_file = upload_section()
    
    # Process button
    if pdf_file and excel_file:
        st.markdown("---")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("Process Files", key="process_btn"):
                # Show progress
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                status_text.text("Starting reconciliation process...")
                progress_bar.progress(20)
                
                # Process files
                reconciler, result_data = process_files(pdf_file, excel_file)
                
                if reconciler and result_data:
                    progress_bar.progress(100)
                    status_text.text("Processing complete!")
                    
                    time.sleep(1)
                    progress_bar.empty()
                    status_text.empty()
                    
                    # Display results
                    st.markdown("---")
                    display_results(reconciler)
                    
                    # Download section
                    st.markdown("---")
                    st.markdown("### Download Results")
                    
                    col1, col2 = st.columns([1, 1])
                    with col1:
                        st.download_button(
                            label="Download Excel Report",
                            data=result_data,
                            file_name="Financial_Reconciliation_Report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    
                    with col2:
                        if st.button("New Analysis", use_container_width=True):
                            st.rerun()
                
                else:
                    progress_bar.empty()
                    status_text.empty()
    
    else:
        st.info("Please upload both PDF and Excel files to begin reconciliation.")
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666666; padding: 1rem;">
        <p>Financial Data Reconciliation Tool | Professional Grade Analysis</p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()