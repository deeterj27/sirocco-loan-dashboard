import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from dateutil.relativedelta import relativedelta
import numpy as np
import os

st.set_page_config(page_title="Sirocco I LP Portfolio Dashboard", layout="wide", initial_sidebar_state="expanded")

# Custom CSS for Sirocco branding
st.markdown("""
<style>
    /* Main background */
    .stApp {
        background-color: #1a1a1a;
    }
    
    /* Headers */
    h1, h2, h3 {
        color: #FFFFFF !important;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Helvetica', 'Arial', sans-serif;
    }
    
    /* Metrics */
    [data-testid="metric-container"] {
        background-color: #2d2d2d;
        border: 1px solid #3d3d3d;
        padding: 1rem;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.3);
    }
    
    [data-testid="metric-container"] [data-testid="stMetricLabel"] {
        color: #FDB813 !important;
        font-weight: 600;
        font-size: 0.9rem;
    }
    
    [data-testid="metric-container"] [data-testid="stMetricValue"] {
        color: #FFFFFF !important;
        font-weight: 700;
    }
    
    /* Sidebar */
    .css-1d391kg, [data-testid="stSidebar"] {
        background-color: #242424;
    }
    
    .css-1d391kg .stMarkdown, [data-testid="stSidebar"] .stMarkdown {
        color: #FFFFFF;
    }
    
    /* Info boxes */
    .stAlert {
        background-color: #2d2d2d;
        color: #FFFFFF;
        border: 1px solid #3d3d3d;
    }
    
    /* Tables */
    .dataframe {
        background-color: #2d2d2d !important;
        color: #FFFFFF !important;
    }
    
    .dataframe th {
        background-color: #FDB813 !important;
        color: #1a1a1a !important;
        font-weight: 600;
        border: none !important;
    }
    
    .dataframe td {
        background-color: #2d2d2d !important;
        color: #FFFFFF !important;
        border-color: #3d3d3d !important;
    }
    
    .dataframe tr:hover td {
        background-color: #3d3d3d !important;
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background-color: #2d2d2d;
        color: #FDB813 !important;
        border-radius: 8px;
    }
    
    .streamlit-expanderContent {
        background-color: #242424;
        border: 1px solid #3d3d3d;
    }
    
    /* File uploader */
    [data-testid="stFileUploadDropzone"] {
        background-color: #2d2d2d;
        border: 2px dashed #FDB813;
    }
    
    /* Checkbox */
    .stCheckbox label {
        color: #FFFFFF !important;
    }
    
    /* General text */
    .stMarkdown, .stText, p, span, div {
        color: #FFFFFF;
    }
    
    /* Column headers with Sirocco yellow */
    .css-1kyxreq {
        color: #FDB813 !important;
    }
    
    /* Buttons */
    .stButton button {
        background-color: #FDB813;
        color: #1a1a1a;
        font-weight: 600;
        border: none;
        border-radius: 4px;
        transition: all 0.3s;
    }
    
    .stButton button:hover {
        background-color: #fcc944;
        transform: translateY(-1px);
        box-shadow: 0 4px 8px rgba(253, 184, 19, 0.3);
    }
</style>
""", unsafe_allow_html=True)

# Helper functions
def safe_float(value):
    """Safely convert a value to float"""
    try:
        if isinstance(value, str):
            if value.lower() in ['interest only', 'n/a', '']:
                return 0.0
            value = value.replace('$', '').replace(',', '')
        return float(value)
    except (ValueError, TypeError):
        return 0.0

def excel_date_to_datetime(serial_date):
    """Convert Excel serial date to datetime"""
    if pd.isna(serial_date):
        return pd.NaT
    if isinstance(serial_date, datetime):
        return serial_date
    if isinstance(serial_date, str):
        try:
            return pd.to_datetime(serial_date)
        except:
            try:
                serial_date = float(serial_date)
            except:
                return pd.NaT
    try:
        return pd.to_datetime('1899-12-30') + pd.to_timedelta(serial_date, unit='D')
    except:
        return pd.NaT

def format_currency(value):
    """Format value as currency"""
    return f"${value:,.2f}" if pd.notna(value) and value != 0 else "$0.00"

def format_percent(value):
    """Format value as percentage"""
    return f"{value:.2%}" if pd.notna(value) and value != 0 else "0.00%"

def get_cell_value(sheet, locations, default=None):
    """Try to get value from multiple cell locations"""
    for location in locations:
        if sheet[location].value is not None:
            return sheet[location].value
    return default

# Main app
st.markdown("""
<div style='background-color: #1a1a1a; padding: 2rem 0; margin: -2rem -2rem 2rem -2rem; border-bottom: 4px solid #FDB813;'>
    <h1 style='text-align: center; color: #FFFFFF; font-size: 2.5rem; margin: 0;'>
        <span style='color: #FDB813;'>⚡</span> Sirocco I LP Portfolio Dashboard
    </h1>
    <p style='text-align: center; color: #999999; margin-top: 0.5rem;'>Loan Participation & Life Settlement Portfolio Management System</p>
</div>
""", unsafe_allow_html=True)

# File upload section
col1, col2, col3 = st.columns(3)
with col1:
    master_file = st.file_uploader("Upload Master Excel File", type=["xlsx"])
with col2:
    ls_file = st.file_uploader("Upload LS Portfolio File", type=["xlsx"])
with col3:
    remittance_file = st.file_uploader("Upload Monthly Remittance File (CSV or XLSX)", type=["csv", "xlsx"])

# Process loan data
if master_file:
    try:
        # Load workbook
        wb = load_workbook(master_file, data_only=True)
        
        # Get all loan sheets (sheets starting with '#')
        loan_sheets = [s for s in wb.sheetnames if s.startswith('#') and s != '#AddSheet']
        
        # Get as-of date from Dashboard
        dashboard_sheet = wb['Dashboard']
        as_of_date = dashboard_sheet['E3'].value
        if isinstance(as_of_date, str):
            as_of_date = pd.to_datetime(as_of_date)
        elif isinstance(as_of_date, (int, float)):
            as_of_date = excel_date_to_datetime(as_of_date)
        
        # Sidebar with Sirocco branding
        with st.sidebar:
            st.markdown("""
            <div style='text-align: center; padding: 1rem 0;'>
                <h2 style='color: #FDB813; margin: 0;'>⚡ Sirocco Partners</h2>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown(f"""
            <div style='background-color: #2d2d2d; padding: 1rem; border-radius: 8px; margin-bottom: 1rem;'>
                <p style='color: #FDB813; margin: 0; font-weight: 600;'>📅 Data as of</p>
                <p style='color: #FFFFFF; margin: 0; font-size: 1.2rem;'>{as_of_date.strftime('%B %d, %Y') if pd.notna(as_of_date) else 'Unknown'}</p>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown(f"""
            <div style='background-color: #2d2d2d; padding: 1rem; border-radius: 8px;'>
                <p style='color: #FDB813; margin: 0; font-weight: 600;'>📊 Total Loans</p>
                <p style='color: #FFFFFF; margin: 0; font-size: 1.2rem;'>{len(loan_sheets)}</p>
            </div>
            """, unsafe_allow_html=True)
        
        st.success(f"✅ Master file loaded: {len(loan_sheets)} loan sheets found")
        
    except Exception as e:
        st.error(f"Error processing master file: {str(e)}")
        st.error("Please ensure the Excel file has the expected structure with loan sheets starting with '#'")
        
        with st.expander("Debug Information"):
            st.code(str(e))

else:
    # Landing page
    st.markdown("""
    <div style='text-align: center; padding: 3rem 0;'>
        <div style='font-size: 5rem; color: #FDB813;'>⚡</div>
        <h2 style='color: #FFFFFF; margin-top: 1rem;'>Welcome to the Sirocco I LP Portfolio Dashboard</h2>
        <p style='color: #999999; font-size: 1.2rem; margin-top: 1rem;'>
            Upload your Master Excel file to begin analyzing your portfolio
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Show expected file structure
    with st.expander("📋 Expected Excel File Structure"):
        st.markdown("""
        <div style='color: #FFFFFF;'>
        The Master Excel file should contain:
        
        **Dashboard sheet**: Summary of all loans
        
        **Loan sheets**: Named with # prefix (e.g., #1, #2, etc.)
        
        Each loan sheet should have loan information in either:
        
        **Format 1** (Data in column B):
        - A2 or B2: Borrower name
        - B3: Original loan amount
        - B4: Annual interest rate
        - B5: Loan period in months
        - B6: Payment amount
        - B7: Loan start date
        
        **Format 2** (Data in column C):
        - A2 or B2: Borrower name
        - C3: Original loan amount (when B3 contains label)
        - C4: Annual interest rate
        - C5: Loan period in months
        - C6: Payment amount
        - C7: Loan start date
        
        **Amortization schedule** starting from row 11 with columns:
        - A: Month
        - B: Repayment number
        - C: Opening balance
        - D: Loan repayment
        - E: Interest charged
        - F: Capital repaid
        - G: Closing balance
        - J: Payment date
        - K: Amount paid
        </div>
        """, unsafe_allow_html=True)

# Footer
st.markdown("""
<div style='margin-top: 3rem; padding-top: 2rem; border-top: 1px solid #3d3d3d; text-align: center; color: #666666;'>
    <p>Sirocco Partners - Portfolio Management System</p>
</div>
""", unsafe_allow_html=True) 