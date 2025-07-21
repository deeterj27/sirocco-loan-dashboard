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
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        background-color: #2d2d2d;
        color: #FFFFFF;
        border-radius: 4px 4px 0px 0px;
        gap: 1px;
        padding-top: 10px;
        padding-bottom: 10px;
        border: 1px solid #3d3d3d;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #FDB813;
        color: #1a1a1a;
        font-weight: 600;
    }
    
    /* General text */
    .stMarkdown, .stText, p, span, div {
        color: #FFFFFF;
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
            # Handle special cases like "Interest Only"
            if value.lower() in ['interest only', 'n/a', '']:
                return 0.0
            # Remove currency symbols and commas
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
            # Try parsing as date string first
            return pd.to_datetime(serial_date)
        except:
            # If that fails, try as serial number
            try:
                serial_date = float(serial_date)
            except:
                return pd.NaT
    try:
        # Excel date serial number (days since 1899-12-30)
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

# Header with Sirocco branding
st.markdown("""
<div style='background-color: #1a1a1a; padding: 2rem 0; margin: -2rem -2rem 2rem -2rem; border-bottom: 4px solid #FDB813;'>
    <h1 style='text-align: center; color: #FFFFFF; font-size: 2.5rem; margin: 0;'>
        <span style='color: #FDB813;'>⚡</span> Sirocco I LP Portfolio Dashboard
    </h1>
    <p style='text-align: center; color: #999999; margin-top: 0.5rem;'>Loan Participation Portfolio Management System</p>
</div>
""", unsafe_allow_html=True)

# File upload section
col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader("Upload Master Excel File", type=["xlsx"])
with col2:
    remittance_file = st.file_uploader("Upload Monthly Remittance File (CSV or XLSX)", type=["csv", "xlsx"])

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
        
        # Process each loan sheet
        loans = []
        loan_details = {}  # Store detailed amortization data
        
        for sheet_name in loan_sheets:
            sheet = wb[sheet_name]
            
            # Extract loan header information - try multiple locations
            # Borrower can be in A2 or B2
            borrower = get_cell_value(sheet, ['B2', 'A2'], f"Unknown ({sheet_name})")
            if borrower == "" or borrower is None:
                borrower = f"Unknown ({sheet_name})"
            
            # For some sheets, data is in column C instead of B
            # Check if B3 has a label (like "Loan Principle Amount") - if so, data is in C3
            b3_value = sheet['B3'].value
            if isinstance(b3_value, str) and 'loan' in str(b3_value).lower():
                # Data is in column C
                loan_amount = safe_float(sheet['C3'].value)
                interest_rate = safe_float(sheet['C4'].value)
                loan_period = safe_float(sheet['C5'].value)
                payment_amount_val = sheet['C6'].value
                # Check both C6 and C7 for loan start date (some sheets have it in C6)
                loan_start = excel_date_to_datetime(sheet['C6'].value)
                if pd.isna(loan_start) or (isinstance(sheet['C6'].value, str) and 'loan' not in str(sheet['C6'].value).lower()):
                    loan_start = excel_date_to_datetime(sheet['C7'].value)
                # If C6 has the date, payment amount might need to come from elsewhere
                if pd.notna(loan_start) and isinstance(sheet['C6'].value, (datetime, str, int, float)) and not isinstance(sheet['C6'].value, str):
                    # Look for payment amount in other cells or calculate it
                    payment_amount_val = None
            else:
                # Data is in column B
                loan_amount = safe_float(sheet['B3'].value)
                interest_rate = safe_float(sheet['B4'].value)
                loan_period = safe_float(sheet['B5'].value)
                payment_amount_val = sheet['B6'].value
                loan_start = excel_date_to_datetime(sheet['B7'].value)
            
            # If still no loan amount, it might be empty row format - check C3 directly
            if loan_amount == 0:
                loan_amount = safe_float(sheet['C3'].value)
                if loan_amount > 0:
                    # Data is in column C with empty B column
                    interest_rate = safe_float(sheet['C4'].value)
                    loan_period = safe_float(sheet['C5'].value)
                    payment_amount_val = sheet['C6'].value
                    # For sheets with data in C column, loan start date is usually in C7
                    loan_start = excel_date_to_datetime(sheet['C7'].value)
                    if pd.isna(loan_start):
                        loan_start = excel_date_to_datetime(sheet['C6'].value)
            
            # Handle payment amount
            if isinstance(payment_amount_val, str) and payment_amount_val.lower() == 'interest only':
                payment_amount = loan_amount * (interest_rate / 12)  # Calculate monthly interest
            else:
                payment_amount = safe_float(payment_amount_val)
            
            # If still no payment amount, try to get from amortization table
            if payment_amount == 0:
                # Check row 11 column D for first payment
                first_payment = safe_float(sheet['D11'].value)
                if first_payment > 0:
                    payment_amount = first_payment
            
            # Check if loan is interest only
            is_interest_only = False
            if isinstance(payment_amount_val, str) and 'interest only' in payment_amount_val.lower():
                is_interest_only = True
            
            # Basic loan information
            loan_info = {
                'Sheet': sheet_name,
                'Borrower': borrower,
                'Original Loan Balance': loan_amount,
                'Annual Interest Rate': interest_rate,
                'Loan Period (months)': loan_period,
                'Payment Amount': payment_amount,
                'Loan Start Date': loan_start,
                'Last Payment Amount': 0,  # Will be updated from amortization
                'Notes': '',  # Will be populated from amortization
                'Is Interest Only': is_interest_only,
            }
            
            # Read amortization schedule
            # Starting from row 11 (after headers in row 10)
            amort_data = []
            row = 11
            
            while row < 100:  # Reasonable max rows
                month_cell = sheet[f'A{row}']
                if month_cell.value is None:
                    break
                    
                # Skip if the row has header text instead of data
                opening_val = sheet[f'C{row}'].value
                if isinstance(opening_val, str) and 'balance' in opening_val.lower():
                    row += 1
                    continue
                    
                amort_row = {
                    'Month': excel_date_to_datetime(month_cell.value),
                    'Repayment Number': safe_float(sheet[f'B{row}'].value),
                    'Opening Balance': safe_float(sheet[f'C{row}'].value),
                    'Loan Repayment': safe_float(sheet[f'D{row}'].value),
                    'Interest Charged': safe_float(sheet[f'E{row}'].value),
                    'Capital Repaid': safe_float(sheet[f'F{row}'].value),
                    'Closing Balance': safe_float(sheet[f'G{row}'].value),
                    'Payment Date': excel_date_to_datetime(sheet[f'J{row}'].value),
                    'Amount Paid': safe_float(sheet[f'K{row}'].value),
                    'Notes': str(sheet[f'L{row}'].value) if sheet[f'L{row}'].value and sheet[f'L{row}'].value != 'Notes' else '',
                }
                
                # Only add valid data rows
                if amort_row['Opening Balance'] > 0 or amort_row['Closing Balance'] >= 0:
                    amort_data.append(amort_row)
                row += 1
            
            if amort_data:
                amort_df = pd.DataFrame(amort_data)
                
                # Collect all notes from the amortization schedule
                all_notes = [note for note in amort_df['Notes'] if note and note.strip()]
                if all_notes:
                    loan_info['Notes'] = '; '.join(all_notes)
                
                # Get the last payment amount from most recent payment
                if pd.notna(as_of_date):
                    past_payments = amort_df[amort_df['Month'] <= as_of_date]
                    if not past_payments.empty:
                        last_payment = past_payments.iloc[-1]
                        loan_info['Last Payment Amount'] = last_payment['Loan Repayment'] if last_payment['Loan Repayment'] > 0 else last_payment['Amount Paid']
                
                # Find the row closest to as_of_date
                if pd.notna(as_of_date) and 'Month' in amort_df.columns:
                    amort_df['Month'] = pd.to_datetime(amort_df['Month'])
                    # Get the last row where Month <= as_of_date
                    current_rows = amort_df[amort_df['Month'] <= as_of_date]
                    if not current_rows.empty:
                        current_row = current_rows.iloc[-1]
                        first_row = amort_df.iloc[0]
                        
                        loan_info['Opening Loan Balance'] = first_row['Opening Balance']
                        loan_info['Current Loan Balance'] = current_row['Closing Balance']
                        
                        # Calculate capital repaid properly
                        # It's the sum of all capital repaid up to the as_of_date
                        loan_info['Total Principal Repaid'] = current_rows['Capital Repaid'].sum()
                        
                        # If capital repaid sum is 0, calculate from balance difference
                        if loan_info['Total Principal Repaid'] == 0:
                            loan_info['Total Principal Repaid'] = loan_info['Opening Loan Balance'] - loan_info['Current Loan Balance']
                            if loan_info['Total Principal Repaid'] < 0:
                                loan_info['Total Principal Repaid'] = 0
                        
                        loan_info['Total Interest Repaid'] = current_rows['Interest Charged'].sum()
                    else:
                        # If no payments have been made yet
                        loan_info['Opening Loan Balance'] = loan_info['Original Loan Balance']
                        loan_info['Current Loan Balance'] = loan_info['Original Loan Balance']
                        loan_info['Total Principal Repaid'] = 0
                        loan_info['Total Interest Repaid'] = 0
                else:
                    # Fallback to last available data
                    first_row = amort_df.iloc[0]
                    last_row = amort_df.iloc[-1]
                    
                    loan_info['Opening Loan Balance'] = first_row['Opening Balance']
                    loan_info['Current Loan Balance'] = last_row['Closing Balance']
                    
                    # Sum all capital repaid
                    loan_info['Total Principal Repaid'] = amort_df['Capital Repaid'].sum()
                    
                    # If capital repaid sum is 0, calculate from balance difference
                    if loan_info['Total Principal Repaid'] == 0:
                        loan_info['Total Principal Repaid'] = loan_info['Opening Loan Balance'] - loan_info['Current Loan Balance']
                        if loan_info['Total Principal Repaid'] < 0:
                            loan_info['Total Principal Repaid'] = 0
                    
                    loan_info['Total Interest Repaid'] = amort_df['Interest Charged'].sum()
                    loan_info['Last Payment Amount'] = last_row['Loan Repayment'] if last_row['Loan Repayment'] > 0 else last_row['Amount Paid']
                
                # Calculate maturity date
                if pd.notna(loan_info['Loan Start Date']) and loan_info['Loan Period (months)'] > 0:
                    loan_info['Maturity Date'] = loan_info['Loan Start Date'] + relativedelta(months=int(loan_info['Loan Period (months)']))
                else:
                    loan_info['Maturity Date'] = amort_df['Month'].iloc[-1] if not amort_df.empty else pd.NaT
                
                # Store detailed data
                loan_details[borrower] = amort_df
            else:
                # No amortization data - use header info only
                loan_info['Opening Loan Balance'] = loan_info['Original Loan Balance']
                loan_info['Current Loan Balance'] = loan_info['Original Loan Balance']
                loan_info['Total Principal Repaid'] = 0
                loan_info['Total Interest Repaid'] = 0
                
                if pd.notna(loan_info['Loan Start Date']) and loan_info['Loan Period (months)'] > 0:
                    loan_info['Maturity Date'] = loan_info['Loan Start Date'] + relativedelta(months=int(loan_info['Loan Period (months)']))
                else:
                    loan_info['Maturity Date'] = pd.NaT
            
            # Always add loans with valid original balance (even if zero for other fields)
            if loan_info['Original Loan Balance'] > 0:
                # Check if loan hasn't started yet based on current date
                today = pd.Timestamp.now()
                if pd.notna(loan_info['Loan Start Date']):
                    loan_start_timestamp = pd.Timestamp(loan_info['Loan Start Date'])
                    if loan_start_timestamp > today:
                        loan_info['Status'] = 'Not Started'
                        # For not started loans, set current balance to 0
                        loan_info['Current Loan Balance'] = 0
                        loan_info['Opening Loan Balance'] = 0
                    elif loan_info['Current Loan Balance'] == 0:
                        loan_info['Status'] = 'Closed'
                    else:
                        loan_info['Status'] = 'Active'
                else:
                    # No start date, determine by balance
                    loan_info['Status'] = 'Active' if loan_info['Current Loan Balance'] > 0 else 'Closed'
                
                # Add Interest Only indicator to notes if applicable
                if loan_info['Is Interest Only']:
                    if loan_info['Notes']:
                        loan_info['Notes'] = 'Interest Only; ' + loan_info['Notes']
                    else:
                        loan_info['Notes'] = 'Interest Only'
                
                loans.append(loan_info)
        
        # Create main dataframe
        loans_df = pd.DataFrame(loans)
        
        # Debug: Show which sheets were processed
        if st.checkbox("Show debug info", value=False):
            st.write(f"Total sheets found: {len(loan_sheets)}")
            st.write(f"Total loans processed: {len(loans_df)}")
            st.write("Sheets processed:", loan_sheets)
            missing_sheets = set(loan_sheets) - set(loans_df['Sheet'].tolist())
            if missing_sheets:
                st.warning(f"Sheets not showing in tables: {missing_sheets}")
        
        # Separate loans by status
        active_loans = loans_df[loans_df['Status'] == 'Active'].copy()
        closed_loans = loans_df[loans_df['Status'] == 'Closed'].copy()
        not_started_loans = loans_df[loans_df['Status'] == 'Not Started'].copy()
        
        # Sort by current balance or original balance
        active_loans = active_loans.sort_values('Current Loan Balance', ascending=False)
        closed_loans = closed_loans.sort_values('Original Loan Balance', ascending=False)
        not_started_loans = not_started_loans.sort_values('Loan Start Date')
        
        # Display portfolio summary with key metrics only
        st.markdown("<h2 style='color: #FDB813; margin-top: 2rem;'>📊 Portfolio Summary</h2>", unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        
        total_original = loans_df['Original Loan Balance'].sum()
        total_repaid_principal = loans_df['Total Principal Repaid'].sum()
        total_repaid_interest = loans_df['Total Interest Repaid'].sum()
        total_collected = total_repaid_principal + total_repaid_interest
        collection_rate = total_collected / total_original if total_original > 0 else 0
        
        with col1:
            st.metric("Total Principal Repaid", format_currency(total_repaid_principal))
            st.metric("Active Loans", len(active_loans))
        with col2:
            st.metric("Total Interest Earned", format_currency(total_repaid_interest))
            st.metric("Closed Loans", len(closed_loans))
        with col3:
            st.metric("Total Collections", format_currency(total_collected))
            st.metric("Not Started", len(not_started_loans))
        with col4:
            st.metric("Collection Rate", format_percent(collection_rate))
            st.metric("Total Loans", len(loans_df))
        
        # Display active loans
        st.markdown("<h2 style='color: #FDB813; margin-top: 2rem;'>💰 Active Loans</h2>", unsafe_allow_html=True)
        
        # Format display columns - changed Payment Amount to Last Payment Amount
        display_columns = ['Sheet', 'Borrower', 'Original Loan Balance', 'Current Loan Balance', 
                          'Total Principal Repaid', 'Total Interest Repaid', 'Last Payment Amount',
                          'Annual Interest Rate', 'Loan Start Date', 'Maturity Date', 'Notes']
        
        active_display = active_loans[display_columns].copy()
        
        # Format columns for display
        for col in ['Original Loan Balance', 'Current Loan Balance', 'Total Principal Repaid', 
                   'Total Interest Repaid', 'Last Payment Amount']:
            active_display[col] = active_display[col].apply(format_currency)
        
        active_display['Annual Interest Rate'] = active_display['Annual Interest Rate'].apply(format_percent)
        active_display['Loan Start Date'] = pd.to_datetime(active_display['Loan Start Date']).dt.strftime('%Y-%m-%d')
        active_display['Maturity Date'] = pd.to_datetime(active_display['Maturity Date']).dt.strftime('%Y-%m-%d')
        
        st.dataframe(active_display, use_container_width=True, hide_index=True)
        
        # Show detailed amortization in expanders
        if st.checkbox("Show loan details", key="active_details"):
            for _, loan in active_loans.iterrows():
                borrower = loan['Borrower']
                with st.expander(f"📋 {borrower} - {loan['Sheet']}"):
                    if borrower in loan_details:
                        detail_df = loan_details[borrower].copy()
                        
                        # Format detail columns
                        for col in ['Opening Balance', 'Loan Repayment', 'Interest Charged', 
                                   'Capital Repaid', 'Closing Balance', 'Amount Paid']:
                            if col in detail_df.columns:
                                detail_df[col] = detail_df[col].apply(format_currency)
                        
                        # Format dates
                        if 'Month' in detail_df.columns:
                            detail_df['Month'] = pd.to_datetime(detail_df['Month']).dt.strftime('%Y-%m-%d')
                        if 'Payment Date' in detail_df.columns:
                            detail_df['Payment Date'] = pd.to_datetime(detail_df['Payment Date']).dt.strftime('%Y-%m-%d')
                        
                        # Remove empty Notes rows
                        if 'Notes' in detail_df.columns:
                            detail_df['Notes'] = detail_df['Notes'].fillna('')
                        
                        st.dataframe(detail_df, use_container_width=True, hide_index=True)
        
        # Display closed loans
        if len(closed_loans) > 0:
            st.markdown("<h2 style='color: #FDB813; margin-top: 2rem;'>✅ Closed Loans</h2>", unsafe_allow_html=True)
            
            closed_display = closed_loans[display_columns].copy()
            
            # Format columns
            for col in ['Original Loan Balance', 'Current Loan Balance', 'Total Principal Repaid', 
                       'Total Interest Repaid', 'Last Payment Amount']:
                closed_display[col] = closed_display[col].apply(format_currency)
            
            closed_display['Annual Interest Rate'] = closed_display['Annual Interest Rate'].apply(format_percent)
            closed_display['Loan Start Date'] = pd.to_datetime(closed_display['Loan Start Date']).dt.strftime('%Y-%m-%d')
            closed_display['Maturity Date'] = pd.to_datetime(closed_display['Maturity Date']).dt.strftime('%Y-%m-%d')
            
            st.dataframe(closed_display, use_container_width=True, hide_index=True)
        
        # Display not started loans
        if len(not_started_loans) > 0:
            st.markdown("<h2 style='color: #FDB813; margin-top: 2rem;'>🕒 Not Started Loans</h2>", unsafe_allow_html=True)
            
            not_started_display = not_started_loans[display_columns].copy()
            
            # Format columns
            for col in ['Original Loan Balance', 'Current Loan Balance', 'Total Principal Repaid', 
                       'Total Interest Repaid', 'Last Payment Amount']:
                not_started_display[col] = not_started_display[col].apply(format_currency)
            
            not_started_display['Annual Interest Rate'] = not_started_display['Annual Interest Rate'].apply(format_percent)
            not_started_display['Loan Start Date'] = pd.to_datetime(not_started_display['Loan Start Date']).dt.strftime('%Y-%m-%d')
            not_started_display['Maturity Date'] = pd.to_datetime(not_started_display['Maturity Date']).dt.strftime('%Y-%m-%d')
            
            st.dataframe(not_started_display, use_container_width=True, hide_index=True)
        
        # Cash flow projection
        st.markdown("<h2 style='color: #FDB813; margin-top: 2rem;'>💸 Cash Flow Projection (Next 6 Months)</h2>", unsafe_allow_html=True)
        
        # Calculate upcoming payments - exclude not started loans
        today = datetime.now()
        cashflow_data = []
        
        for borrower, amort_df in loan_details.items():
            # Check if this borrower is in not started loans
            if borrower in not_started_loans['Borrower'].values:
                continue
                
            if 'Month' in amort_df.columns and 'Loan Repayment' in amort_df.columns:
                amort_df['Month'] = pd.to_datetime(amort_df['Month'])
                upcoming = amort_df[(amort_df['Month'] > today) & 
                                  (amort_df['Month'] <= today + relativedelta(months=6))]
                
                for _, row in upcoming.iterrows():
                    cashflow_data.append({
                        'Borrower': borrower,
                        'Payment Date': row['Month'],
                        'Payment Amount': row['Loan Repayment'],
                        'Interest': row.get('Interest Charged', 0),
                        'Principal': row.get('Capital Repaid', 0)
                    })
        
        if cashflow_data:
            cashflow_df = pd.DataFrame(cashflow_data)
            cashflow_df = cashflow_df.sort_values('Payment Date')
            
            # Create a pivot table for month-over-month view by borrower
            cashflow_df['Month'] = cashflow_df['Payment Date'].dt.to_period('M')
            
            # First show summary by month
            st.markdown("<h3 style='color: #FFFFFF;'>Monthly Summary</h3>", unsafe_allow_html=True)
            monthly_summary = cashflow_df.groupby('Month').agg({
                'Payment Amount': 'sum',
                'Interest': 'sum',
                'Principal': 'sum'
            }).reset_index()
            
            monthly_summary['Month'] = monthly_summary['Month'].astype(str)
            
            # Display summary with totals
            col1, col2 = st.columns([3, 1])
            
            with col1:
                # Format for display
                display_summary = monthly_summary.copy()
                for col in ['Payment Amount', 'Interest', 'Principal']:
                    display_summary[col] = display_summary[col].apply(format_currency)
                
                st.dataframe(display_summary, use_container_width=True, hide_index=True)
            
            with col2:
                total_expected = monthly_summary['Payment Amount'].sum()
                total_interest = monthly_summary['Interest'].sum()
                total_principal = monthly_summary['Principal'].sum()
                st.metric("Total Expected", format_currency(total_expected))
                st.metric("Total Interest", format_currency(total_interest))
                st.metric("Total Principal", format_currency(total_principal))
            
            # Show detailed breakdown by borrower
            st.markdown("<h3 style='color: #FFFFFF;'>Detailed Breakdown by Borrower</h3>", unsafe_allow_html=True)
            
            # Create pivot table for borrower view
            borrower_pivot = cashflow_df.pivot_table(
                index='Borrower',
                columns='Month',
                values=['Payment Amount', 'Interest', 'Principal'],
                aggfunc='sum',
                fill_value=0
            )
            
            # Flatten column names and format
            borrower_display = pd.DataFrame()
            borrower_display['Borrower'] = borrower_pivot.index
            
            # Add columns for each month
            months = sorted(cashflow_df['Month'].unique())
            for month in months:
                month_str = str(month)
                if ('Payment Amount', month) in borrower_pivot.columns:
                    borrower_display[f'{month_str} Payment'] = borrower_pivot[('Payment Amount', month)].values
                    borrower_display[f'{month_str} Interest'] = borrower_pivot[('Interest', month)].values
                    borrower_display[f'{month_str} Principal'] = borrower_pivot[('Principal', month)].values
            
            # Add total columns
            borrower_display['Total Payment'] = borrower_pivot['Payment Amount'].sum(axis=1).values
            borrower_display['Total Interest'] = borrower_pivot['Interest'].sum(axis=1).values
            borrower_display['Total Principal'] = borrower_pivot['Principal'].sum(axis=1).values
            
            # Format currency columns
            currency_cols = [col for col in borrower_display.columns if col != 'Borrower']
            for col in currency_cols:
                borrower_display[col] = borrower_display[col].apply(format_currency)
            
            st.dataframe(borrower_display, use_container_width=True, hide_index=True)
        else:
            st.info("No upcoming payments in the next 6 months")
        
        # Life Insurance Portfolio Statistics
        st.markdown("<h2 style='color: #FDB813; margin-top: 2rem;'>🏦 Life Insurance Portfolio Statistics</h2>", unsafe_allow_html=True)
        
        # Read life insurance stats from Dashboard sheet
        life_stats = {}
        monthly_premiums_data = []
        
        try:
            # Read specific cells for life insurance stats
            life_stats['Face Value'] = safe_float(dashboard_sheet['E39'].value)
            life_stats['AUM'] = safe_float(dashboard_sheet['H39'].value) if dashboard_sheet['H39'].value else safe_float(dashboard_sheet['E39'].value)
            life_stats['Premiums'] = safe_float(dashboard_sheet['E40'].value)
            life_stats['Mgmt Fee'] = safe_float(dashboard_sheet['H40'].value) if dashboard_sheet['H40'].value else 0
            life_stats['Annual Premium Load as %'] = safe_float(dashboard_sheet['E41'].value)
            life_stats['Annual Mgmt Fee'] = safe_float(dashboard_sheet['H41'].value) if dashboard_sheet['H41'].value else 0
            
            # Display life insurance stats
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Face Value", format_currency(life_stats['Face Value']))
                st.metric("Premiums", format_currency(life_stats['Premiums']))
            
            with col2:
                st.metric("AUM", format_currency(life_stats['AUM']))
                st.metric("Mgmt Fee", format_currency(life_stats['Mgmt Fee']))
            
            with col3:
                st.metric("Annual Premium Load", format_percent(life_stats['Annual Premium Load as %']))
                st.metric("Annual Mgmt Fee", format_currency(life_stats['Annual Mgmt Fee']))
            
            # Monthly Premiums Table
            st.markdown("<h3 style='color: #FFFFFF; margin-top: 2rem;'>Monthly Premiums</h3>", unsafe_allow_html=True)
            
            # Read monthly premiums data
            # Headers are in row 43, data starts in row 44
            months = []
            total_premiums = []
            mgmt_fees = []
            total_shortfall = []
            
            # Read month headers from row 43 (columns E, F, G, etc.)
            for col_idx in range(5, 12):  # E through K
                col_letter = chr(65 + col_idx - 1)  # Convert to letter
                month_val = dashboard_sheet[f'{col_letter}43'].value
                if month_val:
                    months.append(excel_date_to_datetime(month_val))
            
            # Read data rows
            for idx, month in enumerate(months):
                col_letter = chr(65 + 4 + idx)  # E, F, G, etc.
                total_premiums.append(safe_float(dashboard_sheet[f'{col_letter}44'].value))
                mgmt_fees.append(safe_float(dashboard_sheet[f'{col_letter}45'].value))
                total_shortfall.append(safe_float(dashboard_sheet[f'{col_letter}46'].value))
            
            if months:
                # Create monthly premiums dataframe
                monthly_df = pd.DataFrame({
                    'Month': months,
                    'Total Premiums': total_premiums,
                    'Management Fee': mgmt_fees,
                    'Total Shortfall': total_shortfall
                })
                
                # Format for display
                monthly_display = monthly_df.copy()
                monthly_display['Month'] = pd.to_datetime(monthly_display['Month']).dt.strftime('%b %Y')
                for col in ['Total Premiums', 'Management Fee', 'Total Shortfall']:
                    monthly_display[col] = monthly_display[col].apply(format_currency)
                
                st.dataframe(monthly_display, use_container_width=True, hide_index=True)
                
                # Summary metrics
                col1, col2, col3 = st.columns(3)
                with col1:
                    avg_premium = sum(total_premiums) / len(total_premiums) if total_premiums else 0
                    st.metric("Average Monthly Premium", format_currency(avg_premium))
                with col2:
                    total_annual_premiums = sum(total_premiums[:12]) if len(total_premiums) >= 12 else sum(total_premiums) * (12 / len(total_premiums))
                    st.metric("Projected Annual Premiums", format_currency(total_annual_premiums))
                with col3:
                    avg_shortfall = sum(total_shortfall) / len(total_shortfall) if total_shortfall else 0
                    st.metric("Average Monthly Shortfall", format_currency(avg_shortfall))
            
            # Additional Portfolio Stats
            st.markdown("<h3 style='color: #FFFFFF; margin-top: 2rem;'>Additional Portfolio Statistics</h3>", unsafe_allow_html=True)
            
            add_stats = {}
            add_stats['% of Capital Collected'] = safe_float(dashboard_sheet['E49'].value)
            add_stats['Avg Loan Size'] = safe_float(dashboard_sheet['E50'].value)
            add_stats['Avg Term'] = safe_float(dashboard_sheet['E51'].value)
            add_stats['Avg Interest Rate'] = safe_float(dashboard_sheet['E52'].value)
            add_stats['Projected Cash Flow Coverage'] = safe_float(dashboard_sheet['E53'].value)
            
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                st.metric("Capital Collected", format_percent(add_stats['% of Capital Collected']))
            with col2:
                st.metric("Avg Loan Size", format_currency(add_stats['Avg Loan Size']))
            with col3:
                st.metric("Avg Term", f"{add_stats['Avg Term']:.1f} months" if add_stats['Avg Term'] > 0 else "N/A")
            with col4:
                st.metric("Avg Interest Rate", format_percent(add_stats['Avg Interest Rate']))
            with col5:
                st.metric("Cash Flow Coverage", format_percent(add_stats['Projected Cash Flow Coverage']))
            
        except Exception as e:
            st.warning("Life insurance portfolio data not available in the Dashboard sheet")
        
        # Process remittance file if uploaded
        if remittance_file:
            st.markdown("<h2 style='color: #FDB813; margin-top: 2rem;'>📄 Monthly Remittance Analysis</h2>", unsafe_allow_html=True)
            
            try:
                if remittance_file.type == "text/csv":
                    rem_df = pd.read_csv(remittance_file)
                else:
                    rem_df = pd.read_excel(remittance_file)
                
                st.dataframe(rem_df.head(), use_container_width=True)
                st.info("Remittance file loaded successfully. Add custom processing logic based on your file format.")
                
            except Exception as e:
                st.error(f"Error processing remittance file: {str(e)}")
    
    except Exception as e:
        st.error(f"Error processing master file: {str(e)}")
        st.error("Please ensure the Excel file has the expected structure with loan sheets starting with '#'")
        
        # Show debug information
        with st.expander("Debug Information"):
            st.code(str(e))
        
else:
    # Landing page with Sirocco branding
    st.markdown("""
    <div style='text-align: center; padding: 3rem 0;'>
        <div style='font-size: 5rem; color: #FDB813;'>⚡</div>
        <h2 style='color: #FFFFFF; margin-top: 1rem;'>Welcome to the Sirocco I LP Portfolio Dashboard</h2>
        <p style='color: #999999; font-size: 1.2rem; margin-top: 1rem;'>
            Upload your Master Excel file to begin analyzing your loan portfolio
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