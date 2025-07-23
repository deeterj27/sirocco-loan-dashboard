import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from dateutil.relativedelta import relativedelta
import numpy as np
import os

try:
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

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

def process_life_settlement_data(ls_file):
    """Process Life Settlement Excel file and return summary data"""
    try:
        ls_wb = load_workbook(ls_file, data_only=True)
        
        if 'Valuation Summary' not in ls_wb.sheetnames or 'Premium Stream' not in ls_wb.sheetnames:
            return None
        
        val_sheet = ls_wb['Valuation Summary']
        premium_sheet = ls_wb['Premium Stream']
        
        policies = []
        
        # Process Valuation Summary sheet
        # Headers are in row 2, data starts from row 3
        for row in range(3, 200):
            policy_id_cell = val_sheet[f'A{row}']  # Policy ID is in column A
            if not policy_id_cell.value:
                break
                
            try:
                policy_data = {
                    'Policy_ID': str(policy_id_cell.value),
                    'Insured_ID': str(val_sheet[f'B{row}'].value or ''),  # Column B
                    'Name': str(val_sheet[f'C{row}'].value or ''),  # Column C
                    'Age': safe_float(val_sheet[f'E{row}'].value),  # Column E
                    'Gender': str(val_sheet[f'F{row}'].value or ''),  # Column F
                    'NDB': safe_float(str(val_sheet[f'U{row}'].value or '0').replace('$', '').replace(',', '').strip()),  # Column U
                    'Valuation': safe_float(str(val_sheet[f'Y{row}'].value or '0').replace('$', '').replace(',', '').strip()),  # Column Y
                    'Cost_Basis': safe_float(str(val_sheet[f'Z{row}'].value or '0').replace('$', '').replace(',', '').strip()),  # Column Z (Purchase Price)
                    'Remaining_LE': safe_float(val_sheet[f'AB{row}'].value or val_sheet[f'AC{row}'].value),  # Column AB or AC (Calc LE)
                }
                policies.append(policy_data)
            except:
                continue
        
        if len(policies) == 0:
            return None
        
        # Calculate summary statistics
        total_policies = len(policies)
        total_ndb = sum(p['NDB'] for p in policies)
        total_valuation = sum(p['Valuation'] for p in policies)
        total_cost_basis = sum(p['Cost_Basis'] for p in policies)
        
        valid_ages = [p['Age'] for p in policies if p['Age'] > 0]
        avg_age = sum(valid_ages) / len(valid_ages) if valid_ages else 0
        
        male_count = sum(1 for p in policies if 'male' in p['Gender'].lower() and 'female' not in p['Gender'].lower())
        female_count = sum(1 for p in policies if 'female' in p['Gender'].lower())
        male_percentage = (male_count / (male_count + female_count)) * 100 if (male_count + female_count) > 0 else 0
        
        valid_les = [p['Remaining_LE'] for p in policies if p['Remaining_LE'] > 0]
        avg_remaining_le = sum(valid_les) / len(valid_les) if valid_les else 0
        
        # Process monthly premiums from Premium Stream sheet
        monthly_premiums = {}
        policy_premiums = {}
        
        # Premium columns start at column L (index 11) based on the structure
        # Get month headers from row 2
        month_headers = []
        col_index = 11  # Start at column L
        
        while col_index < 300:  # Check many columns as there are many months
            cell = premium_sheet.cell(row=2, column=col_index+1)  # openpyxl uses 1-based indexing
            if cell.value and isinstance(cell.value, str) and '-' in str(cell.value):
                month_headers.append((col_index, str(cell.value)))
            col_index += 1
            
        # Read premium data
        for month_col, month_name in month_headers:
            month_total = 0
            for prem_row in range(3, 200):  # Start from row 3
                try:
                    lyric_id_cell = premium_sheet.cell(row=prem_row, column=1)  # Column A - Lyric ID
                    if lyric_id_cell.value:
                        lyric_id = str(lyric_id_cell.value)
                        premium_cell = premium_sheet.cell(row=prem_row, column=month_col+1)
                        premium_val = safe_float(premium_cell.value) if premium_cell.value else 0
                        month_total += premium_val
                        
                        if lyric_id not in policy_premiums:
                            policy_premiums[lyric_id] = {}
                        policy_premiums[lyric_id][month_name] = premium_val
                except:
                    continue
            
            monthly_premiums[month_name] = month_total
        
        # Calculate policy-level metrics
        for policy in policies:
            policy_id = policy['Policy_ID']
            if policy_id in policy_premiums:
                annual_premium = sum(policy_premiums[policy_id].values())
                policy['Annual_Premium'] = annual_premium
                if policy['NDB'] > 0:
                    policy['Premium_Pct_Face'] = (annual_premium / policy['NDB']) * 100
                else:
                    policy['Premium_Pct_Face'] = 0
            else:
                policy['Annual_Premium'] = 0
                policy['Premium_Pct_Face'] = 0
        
        total_annual_premiums = sum(monthly_premiums.values())
        premiums_as_pct_face = (total_annual_premiums / total_ndb) * 100 if total_ndb > 0 else 0
        
        return {
            'policies': policies,
            'summary': {
                'total_policies': total_policies,
                'total_ndb': total_ndb,
                'total_valuation': total_valuation,
                'total_cost_basis': total_cost_basis,
                'avg_age': avg_age,
                'male_count': male_count,
                'female_count': female_count,
                'male_percentage': male_percentage,
                'avg_remaining_le': avg_remaining_le,
                'total_annual_premiums': total_annual_premiums,
                'premiums_as_pct_face': premiums_as_pct_face,
            },
            'monthly_premiums': monthly_premiums,
            'policy_premiums': policy_premiums
        }
        
    except:
        return None 

# Main app

# Header with Sirocco branding
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

# Process LS data if uploaded
ls_data = None
if ls_file:
    ls_data = process_life_settlement_data(ls_file)
    if ls_data:
        st.success(f"✅ Life Settlement data loaded: {ls_data['summary']['total_policies']} policies, {format_currency(ls_data['summary']['total_ndb'])} face value")
    else:
        st.error("❌ Failed to load Life Settlement data. Please check file format.")

# Process loan data (keep original logic)
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
        
        # Process each loan sheet (keep existing logic)
        loans = []
        loan_details = {}
        
        for sheet_name in loan_sheets:
            sheet = wb[sheet_name]
            
            # Extract loan header information - try multiple locations
            borrower = get_cell_value(sheet, ['B2', 'A2'], f"Unknown ({sheet_name})")
            if borrower == "" or borrower is None:
                borrower = f"Unknown ({sheet_name})"
            
            # Check if B3 has a label (like "Loan Principle Amount") - if so, data is in C3
            b3_value = sheet['B3'].value
            if isinstance(b3_value, str) and 'loan' in str(b3_value).lower():
                # Data is in column C
                loan_amount = safe_float(sheet['C3'].value)
                interest_rate = safe_float(sheet['C4'].value)
                loan_period = safe_float(sheet['C5'].value)
                payment_amount_val = sheet['C6'].value
                # Check both C6 and C7 for loan start date
                loan_start = excel_date_to_datetime(sheet['C6'].value)
                if pd.isna(loan_start) or (isinstance(sheet['C6'].value, str) and 'loan' not in str(sheet['C6'].value).lower()):
                    loan_start = excel_date_to_datetime(sheet['C7'].value)
                if pd.notna(loan_start) and isinstance(sheet['C6'].value, (datetime, str, int, float)) and not isinstance(sheet['C6'].value, str):
                    payment_amount_val = None
            else:
                # Data is in column B
                loan_amount = safe_float(sheet['B3'].value)
                interest_rate = safe_float(sheet['B4'].value)
                loan_period = safe_float(sheet['B5'].value)
                payment_amount_val = sheet['B6'].value
                loan_start = excel_date_to_datetime(sheet['B7'].value)
            
            # If still no loan amount, try C3 directly
            if loan_amount == 0:
                loan_amount = safe_float(sheet['C3'].value)
                if loan_amount > 0:
                    interest_rate = safe_float(sheet['C4'].value)
                    loan_period = safe_float(sheet['C5'].value)
                    payment_amount_val = sheet['C6'].value
                    loan_start = excel_date_to_datetime(sheet['C7'].value)
                    if pd.isna(loan_start):
                        loan_start = excel_date_to_datetime(sheet['C6'].value)
            
            # Handle payment amount
            if isinstance(payment_amount_val, str) and payment_amount_val.lower() == 'interest only':
                payment_amount = loan_amount * (interest_rate / 12)
            else:
                payment_amount = safe_float(payment_amount_val)
            
            # If still no payment amount, try from amortization table
            if payment_amount == 0:
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
                'Last Payment Amount': 0,
                'Notes': '',
                'Is Interest Only': is_interest_only,
            }
            
            # Read amortization schedule
            amort_data = []
            row = 11
            
            while row < 100:
                month_cell = sheet[f'A{row}']
                if month_cell.value is None:
                    break
                    
                # Skip header rows
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
                
                if amort_row['Opening Balance'] > 0 or amort_row['Closing Balance'] >= 0:
                    amort_data.append(amort_row)
                row += 1
            
            if amort_data:
                amort_df = pd.DataFrame(amort_data)
                
                # Collect notes
                all_notes = [note for note in amort_df['Notes'] if note and note.strip()]
                if all_notes:
                    loan_info['Notes'] = '; '.join(all_notes)
                
                # Get the last payment amount
                if pd.notna(as_of_date):
                    past_payments = amort_df[amort_df['Month'] <= as_of_date]
                    if not past_payments.empty:
                        last_payment = past_payments.iloc[-1]
                        loan_info['Last Payment Amount'] = last_payment['Loan Repayment'] if last_payment['Loan Repayment'] > 0 else last_payment['Amount Paid']
                
                # Find current position
                if pd.notna(as_of_date) and 'Month' in amort_df.columns:
                    amort_df['Month'] = pd.to_datetime(amort_df['Month'])
                    current_rows = amort_df[amort_df['Month'] <= as_of_date]
                    if not current_rows.empty:
                        current_row = current_rows.iloc[-1]
                        first_row = amort_df.iloc[0]
                        
                        loan_info['Opening Loan Balance'] = first_row['Opening Balance']
                        loan_info['Current Loan Balance'] = current_row['Closing Balance']
                        loan_info['Total Principal Repaid'] = current_rows['Capital Repaid'].sum()
                        loan_info['Total Interest Repaid'] = current_rows['Interest Charged'].sum()
                        
                        # If capital repaid sum is 0, calculate from balance difference
                        if loan_info['Total Principal Repaid'] == 0:
                            loan_info['Total Principal Repaid'] = loan_info['Opening Loan Balance'] - loan_info['Current Loan Balance']
                            if loan_info['Total Principal Repaid'] < 0:
                                loan_info['Total Principal Repaid'] = 0
                    else:
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
                    loan_info['Total Principal Repaid'] = amort_df['Capital Repaid'].sum()
                    loan_info['Total Interest Repaid'] = amort_df['Interest Charged'].sum()
                    loan_info['Last Payment Amount'] = last_row['Loan Repayment'] if last_row['Loan Repayment'] > 0 else last_row['Amount Paid']
                    
                    if loan_info['Total Principal Repaid'] == 0:
                        loan_info['Total Principal Repaid'] = loan_info['Opening Loan Balance'] - loan_info['Current Loan Balance']
                        if loan_info['Total Principal Repaid'] < 0:
                            loan_info['Total Principal Repaid'] = 0
                
                # Calculate maturity date
                if pd.notna(loan_info['Loan Start Date']) and loan_info['Loan Period (months)'] > 0:
                    loan_info['Maturity Date'] = loan_info['Loan Start Date'] + relativedelta(months=int(loan_info['Loan Period (months)']))
                else:
                    loan_info['Maturity Date'] = amort_df['Month'].iloc[-1] if not amort_df.empty else pd.NaT
                
                loan_details[borrower] = amort_df
            else:
                # No amortization data
                loan_info['Opening Loan Balance'] = loan_info['Original Loan Balance']
                loan_info['Current Loan Balance'] = loan_info['Original Loan Balance']
                loan_info['Total Principal Repaid'] = 0
                loan_info['Total Interest Repaid'] = 0
                
                if pd.notna(loan_info['Loan Start Date']) and loan_info['Loan Period (months)'] > 0:
                    loan_info['Maturity Date'] = loan_info['Loan Start Date'] + relativedelta(months=int(loan_info['Loan Period (months)']))
                else:
                    loan_info['Maturity Date'] = pd.NaT
            
            # Add loans with valid original balance
            if loan_info['Original Loan Balance'] > 0:
                # Check status
                today = pd.Timestamp.now()
                if pd.notna(loan_info['Loan Start Date']):
                    loan_start_timestamp = pd.Timestamp(loan_info['Loan Start Date'])
                    if loan_start_timestamp > today:
                        loan_info['Status'] = 'Not Started'
                        loan_info['Current Loan Balance'] = 0
                        loan_info['Opening Loan Balance'] = 0
                    elif loan_info['Current Loan Balance'] == 0:
                        loan_info['Status'] = 'Closed'
                    else:
                        loan_info['Status'] = 'Active'
                else:
                    loan_info['Status'] = 'Active' if loan_info['Current Loan Balance'] > 0 else 'Closed'
                
                # Add Interest Only indicator to notes
                if loan_info['Is Interest Only']:
                    if loan_info['Notes']:
                        loan_info['Notes'] = 'Interest Only; ' + loan_info['Notes']
                    else:
                        loan_info['Notes'] = 'Interest Only'
                
                loans.append(loan_info)
        
        # Create main dataframe
        loans_df = pd.DataFrame(loans)
        
        # Debug info
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
        
        # Display portfolio summary
        st.markdown("<h2 style='color: #FDB813; margin-top: 2rem;'>📊 Portfolio Summary</h2>", unsafe_allow_html=True)
        
        # Overall Portfolio Summary
        st.markdown("<h3 style='color: #FFFFFF;'>Overall Portfolio</h3>", unsafe_allow_html=True)
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
        
        # Active Loans Summary
        if len(active_loans) > 0:
            st.markdown("<h3 style='color: #FFFFFF; margin-top: 2rem;'>Active Loans Summary</h3>", unsafe_allow_html=True)
            col1, col2, col3, col4, col5 = st.columns(5)
            
            active_original_balance = active_loans['Original Loan Balance'].sum()
            active_current_balance = active_loans['Current Loan Balance'].sum()
            active_avg_interest = active_loans['Annual Interest Rate'].mean()
            
            # Calculate average days to maturity for active loans
            maturity_dates = pd.to_datetime(active_loans['Maturity Date'])
            valid_maturity_dates = maturity_dates[maturity_dates.notna()]
            if len(valid_maturity_dates) > 0:
                today = pd.Timestamp.now()
                days_to_maturity = (valid_maturity_dates - today).dt.days
                avg_days_to_maturity = days_to_maturity[days_to_maturity > 0].mean()
                if pd.notna(avg_days_to_maturity):
                    avg_months_to_maturity = avg_days_to_maturity / 30.44  # Average days in a month
                else:
                    avg_months_to_maturity = 0
            else:
                avg_months_to_maturity = 0
            
            with col1:
                st.metric("Original Balance", format_currency(active_original_balance))
            with col2:
                st.metric("Current Balance", format_currency(active_current_balance))
            with col3:
                st.metric("Avg Interest Rate", format_percent(active_avg_interest))
            with col4:
                st.metric("Avg Maturity", f"{avg_months_to_maturity:.1f} months" if avg_months_to_maturity > 0 else "N/A")
            with col5:
                st.metric("# of Active Loans", len(active_loans))
        
        # Closed Loans Summary
        if len(closed_loans) > 0:
            st.markdown("<h3 style='color: #FFFFFF; margin-top: 2rem;'>Closed Loans Summary</h3>", unsafe_allow_html=True)
            col1, col2, col3, col4, col5 = st.columns(5)
            
            closed_original_balance = closed_loans['Original Loan Balance'].sum()
            closed_principal_repaid = closed_loans['Total Principal Repaid'].sum()
            closed_interest_earned = closed_loans['Total Interest Repaid'].sum()
            closed_avg_interest = closed_loans['Annual Interest Rate'].mean()
            
            with col1:
                st.metric("Original Balance", format_currency(closed_original_balance))
            with col2:
                st.metric("Principal Repaid", format_currency(closed_principal_repaid))
            with col3:
                st.metric("Interest Earned", format_currency(closed_interest_earned))
            with col4:
                st.metric("Avg Interest Rate", format_percent(closed_avg_interest))
            with col5:
                st.metric("# of Closed Loans", len(closed_loans))
        
        # Display active loans
        st.markdown("<h2 style='color: #FDB813; margin-top: 2rem;'>💰 Active Loans</h2>", unsafe_allow_html=True)
        
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
        
        # Show loan details in expanders
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
                        
                        if 'Notes' in detail_df.columns:
                            detail_df['Notes'] = detail_df['Notes'].fillna('')
                        
                        st.dataframe(detail_df, use_container_width=True, hide_index=True)
        
        # Display closed loans
        if len(closed_loans) > 0:
            st.markdown("<h2 style='color: #FDB813; margin-top: 2rem;'>✅ Closed Loans</h2>", unsafe_allow_html=True)
            
            closed_display = closed_loans[display_columns].copy()
            
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
            
            for col in ['Original Loan Balance', 'Current Loan Balance', 'Total Principal Repaid', 
                       'Total Interest Repaid', 'Last Payment Amount']:
                not_started_display[col] = not_started_display[col].apply(format_currency)
            
            not_started_display['Annual Interest Rate'] = not_started_display['Annual Interest Rate'].apply(format_percent)
            not_started_display['Loan Start Date'] = pd.to_datetime(not_started_display['Loan Start Date']).dt.strftime('%Y-%m-%d')
            not_started_display['Maturity Date'] = pd.to_datetime(not_started_display['Maturity Date']).dt.strftime('%Y-%m-%d')
            
            st.dataframe(not_started_display, use_container_width=True, hide_index=True)
        
        # Cash flow projection
        st.markdown("<h2 style='color: #FDB813; margin-top: 2rem;'>💸 Cash Flow Projection (Next 6 Months)</h2>", unsafe_allow_html=True)
        
        today = datetime.now()
        cashflow_data = []
        
        for borrower, amort_df in loan_details.items():
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
            
            # Monthly summary
            st.markdown("<h3 style='color: #FFFFFF;'>Monthly Summary</h3>", unsafe_allow_html=True)
            monthly_summary = cashflow_df.groupby('Month').agg({
                'Payment Amount': 'sum',
                'Interest': 'sum',
                'Principal': 'sum'
            }).reset_index()
            
            monthly_summary['Month'] = monthly_summary['Month'].astype(str)
            
            col1, col2 = st.columns([3, 1])
            
            with col1:
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
        else:
            st.info("No upcoming payments in the next 6 months")
        
        # Display LS data if available (but after all loan data)
        if ls_data:
            st.markdown("<h2 style='color: #FDB813; margin-top: 3rem;'>🏥 Life Settlement Portfolio</h2>", unsafe_allow_html=True)
            
            # Key Metrics
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Policies", ls_data['summary']['total_policies'])
                st.metric("Average Age", f"{ls_data['summary']['avg_age']:.1f} years")
            
            with col2:
                st.metric("Total NDB (Face Value)", format_currency(ls_data['summary']['total_ndb']))
                st.metric("% Male", f"{ls_data['summary']['male_percentage']:.1f}%")
            
            with col3:
                st.metric("Total Valuation", format_currency(ls_data['summary']['total_valuation']))
                st.metric("Avg Remaining LE", f"{ls_data['summary']['avg_remaining_le']:.1f} months")
            
            with col4:
                st.metric("Cost Basis", format_currency(ls_data['summary']['total_cost_basis']))
                st.metric("Premiums % of Face", f"{ls_data['summary']['premiums_as_pct_face']:.2f}%")
            
            # Monthly Premium Projections
            if ls_data['monthly_premiums']:
                st.markdown("<h3 style='color: #FFFFFF; margin-top: 2rem;'>Monthly Premium Projections</h3>", unsafe_allow_html=True)
                
                premium_items = list(ls_data['monthly_premiums'].items())
                months_per_row = 6
                
                for i in range(0, len(premium_items), months_per_row):
                    cols = st.columns(months_per_row)
                    for j in range(months_per_row):
                        if i + j < len(premium_items):
                            month, amount = premium_items[i + j]
                            with cols[j]:
                                st.metric(month, format_currency(amount))
            
            # Policy Details Table
            st.markdown("<h3 style='color: #FFFFFF; margin-top: 2rem;'>Policy Details</h3>", unsafe_allow_html=True)
            
            if ls_data['policies']:
                policies_df = pd.DataFrame(ls_data['policies'])
                
                # Format for display
                display_policy_df = policies_df.copy()
                display_policy_df['NDB'] = display_policy_df['NDB'].apply(format_currency)
                display_policy_df['Valuation'] = display_policy_df['Valuation'].apply(format_currency)
                display_policy_df['Cost_Basis'] = display_policy_df['Cost_Basis'].apply(format_currency)
                display_policy_df['Annual_Premium'] = display_policy_df['Annual_Premium'].apply(format_currency)
                display_policy_df['Premium_Pct_Face'] = display_policy_df['Premium_Pct_Face'].apply(lambda x: f"{x:.2f}%")
                display_policy_df['Remaining_LE'] = display_policy_df['Remaining_LE'].apply(lambda x: f"{x:.1f} months" if x > 0 else "N/A")
                
                # Rename columns
                display_policy_df = display_policy_df.rename(columns={
                    'Policy_ID': 'Policy ID',
                    'Name': 'Name',
                    'Age': 'Age',
                    'Gender': 'Gender',
                    'NDB': 'Face Value',
                    'Valuation': 'Valuation',
                    'Cost_Basis': 'Cost Basis',
                    'Remaining_LE': 'Remaining LE',
                    'Annual_Premium': 'Annual Premium',
                    'Premium_Pct_Face': 'Premium % Face'
                })
                
                st.dataframe(display_policy_df[['Policy ID', 'Name', 'Age', 'Gender', 'Face Value', 'Valuation', 'Annual Premium', 'Premium % Face']], 
                            use_container_width=True, hide_index=True)
        
        # Cashflow vs Premium Analysis (if both data sources are available)
        if cashflow_data and ls_data and ls_data['monthly_premiums']:
            st.markdown("<h2 style='color: #FDB813; margin-top: 3rem;'>📈 Cashflow vs Premium Analysis</h2>", unsafe_allow_html=True)
            
            try:
                # Prepare cashflow data
                cashflow_monthly = cashflow_df.groupby('Month')['Payment Amount'].sum()
                
                # Prepare premium data - need to convert month names to periods
                premium_months = []
                premium_amounts = []
                
                for month_str, amount in ls_data['monthly_premiums'].items():
                    try:
                        # Parse month string (assuming format like "Jul-25", "Aug-25", etc.)
                        if '-' in month_str:
                            month_parts = month_str.split('-')
                            month_name = month_parts[0]
                            year = '20' + month_parts[1] if len(month_parts[1]) == 2 else month_parts[1]
                            
                            # Convert to datetime
                            month_date = pd.to_datetime(f"{month_name} {year}", format='%b %Y')
                            month_period = month_date.to_period('M')
                            
                            premium_months.append(month_period)
                            premium_amounts.append(amount)
                    except:
                        continue
                
                # Create premium series
                premium_series = pd.Series(premium_amounts, index=premium_months)
                
                # Align the data - get common months
                all_months = sorted(set(cashflow_monthly.index) | set(premium_series.index))
                
                # Create aligned dataframes
                comparison_data = []
                for month in all_months:
                    cashflow_amt = cashflow_monthly.get(month, 0)
                    premium_amt = premium_series.get(month, 0)
                    net_flow = cashflow_amt - premium_amt
                    
                    comparison_data.append({
                        'Month': str(month),
                        'Loan Cashflows': cashflow_amt,
                        'LS Premiums': premium_amt,
                        'Net Cash Flow': net_flow
                    })
                
                comparison_df = pd.DataFrame(comparison_data)
                
                # Only proceed if we have data
                if len(comparison_df) == 0:
                    st.warning("No overlapping months found between loan cashflows and life settlement premiums.")
                else:
                    # Create visualization
                    col1, col2 = st.columns([3, 1])
                    
                    with col1:
                        # Create a line graph using plotly if available
                        if PLOTLY_AVAILABLE:
                            fig = go.Figure()
                            
                            # Add loan cashflows line
                            fig.add_trace(go.Scatter(
                                name='Loan Cashflows',
                                x=comparison_df['Month'],
                                y=comparison_df['Loan Cashflows'],
                                mode='lines+markers',
                                line=dict(color='#FDB813', width=3),
                                marker=dict(size=8),
                                text=[f'${v:,.0f}' for v in comparison_df['Loan Cashflows']],
                                hovertemplate='Loan Collections: %{text}<extra></extra>'
                            ))
                            
                            # Add LS premiums line
                            fig.add_trace(go.Scatter(
                                name='LS Premiums',
                                x=comparison_df['Month'],
                                y=comparison_df['LS Premiums'],
                                mode='lines+markers',
                                line=dict(color='#FF6B6B', width=3),
                                marker=dict(size=8),
                                text=[f'${v:,.0f}' for v in comparison_df['LS Premiums']],
                                hovertemplate='LS Premiums: %{text}<extra></extra>'
                            ))
                            
                            # Add net cashflow line with emphasis
                            fig.add_trace(go.Scatter(
                                name='Net Cash Flow',
                                x=comparison_df['Month'],
                                y=comparison_df['Net Cash Flow'],
                                mode='lines+markers',
                                line=dict(color='#4ECDC4', width=5, dash='solid'),
                                marker=dict(
                                    size=12, 
                                    symbol='diamond',
                                    color=comparison_df['Net Cash Flow'].apply(lambda x: '#4ECDC4' if x >= 0 else '#FF6B6B'),
                                    line=dict(width=2, color='white')
                                ),
                                text=[f'${v:,.0f}' for v in comparison_df['Net Cash Flow']],
                                hovertemplate='Net: %{text}<extra></extra>'
                            ))
                            
                            # Add zero line for reference
                            fig.add_hline(y=0, line_dash="dash", line_color="white", opacity=0.5)
                            
                            # Find min and max values for better scaling
                            all_values = list(comparison_df['Loan Cashflows']) + list(comparison_df['LS Premiums']) + list(comparison_df['Net Cash Flow'])
                            y_min = min(all_values) * 1.1
                            y_max = max(all_values) * 1.1
                            
                            # Update layout
                            fig.update_layout(
                                title={
                                    'text': 'Monthly Cash Flow Analysis',
                                    'font': {'size': 20, 'color': '#FDB813'},
                                    'x': 0.5,
                                    'xanchor': 'center'
                                },
                                xaxis_title='Month',
                                yaxis_title='Amount ($)',
                                height=500,
                                plot_bgcolor='#1a1a1a',
                                paper_bgcolor='#1a1a1a',
                                font=dict(color='#FFFFFF', size=12),
                                xaxis=dict(
                                    gridcolor='#3d3d3d',
                                    showgrid=True,
                                    tickangle=-45,
                                    tickfont=dict(size=10)
                                ),
                                yaxis=dict(
                                    gridcolor='#3d3d3d',
                                    showgrid=True,
                                    tickformat='$,.0f',
                                    range=[y_min, y_max]
                                ),
                                legend=dict(
                                    bgcolor='#2d2d2d',
                                    bordercolor='#FDB813',
                                    borderwidth=1,
                                    orientation='h',
                                    yanchor='bottom',
                                    y=1.02,
                                    xanchor='center',
                                    x=0.5,
                                    font=dict(size=11)
                                ),
                                hovermode='x unified',
                                margin=dict(l=80, r=20, t=80, b=80)
                            )
                            
                            # Add annotations for negative net cash flows
                            for idx, row in comparison_df.iterrows():
                                if row['Net Cash Flow'] < 0:
                                    fig.add_annotation(
                                        x=row['Month'],
                                        y=row['Net Cash Flow'],
                                        text="⚠️",
                                        showarrow=False,
                                        yshift=-20,
                                        font=dict(size=16)
                                    )
                            
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            # Fallback to simple line chart if plotly not available
                            st.line_chart(comparison_df.set_index('Month')[['Loan Cashflows', 'LS Premiums', 'Net Cash Flow']])
                    
                    with col2:
                        # Summary metrics - Current Month Focus
                        st.markdown("<h3 style='color: #FFFFFF;'>Current Month Stats</h3>", unsafe_allow_html=True)
                        
                        # Get current month data
                        current_date = datetime.now()
                        current_period_str = current_date.strftime('%Y-%m')
                        
                        # Find the current month in the comparison data
                        current_month_data = None
                        for idx, row in comparison_df.iterrows():
                            if current_period_str in row['Month']:
                                current_month_data = row
                                break
                        
                        # If current month not found, use the first month in the data
                        if current_month_data is None and len(comparison_df) > 0:
                            current_month_data = comparison_df.iloc[0]
                            month_label = comparison_df.iloc[0]['Month']
                        else:
                            month_label = current_period_str if current_month_data is not None else "N/A"
                        
                        if current_month_data is not None:
                            # Current month metrics
                            st.markdown(f"<p style='color: #FDB813; font-weight: 600; margin-bottom: 1rem;'>📅 {month_label}</p>", unsafe_allow_html=True)
                            
                            current_cashflow = current_month_data['Loan Cashflows']
                            current_premium = current_month_data['LS Premiums']
                            current_net = current_month_data['Net Cash Flow']
                            
                            st.metric("Loan Collections", format_currency(current_cashflow))
                            st.metric("LS Premiums", format_currency(current_premium))
                            st.metric("Net Position", format_currency(current_net), 
                                     delta=f"{'Surplus' if current_net > 0 else 'Deficit'}",
                                     delta_color="normal" if current_net > 0 else "inverse")
                            
                            # Coverage ratio
                            if current_premium > 0:
                                coverage_ratio = (current_cashflow / current_premium) * 100
                                st.metric("Coverage Ratio", f"{coverage_ratio:.1f}%",
                                         help="Loan collections as % of LS premiums")
                            else:
                                st.metric("Coverage Ratio", "N/A",
                                         help="No premiums this month")
                        else:
                            st.info("No data available for current month")
                        
                        # Period averages
                        st.markdown("<h4 style='color: #FFFFFF; margin-top: 2rem;'>Period Averages</h4>", unsafe_allow_html=True)
                        avg_cashflow = comparison_df['Loan Cashflows'].mean()
                        avg_premium = comparison_df['LS Premiums'].mean()
                        avg_net = comparison_df['Net Cash Flow'].mean()
                        
                        st.metric("Avg Collections", format_currency(avg_cashflow))
                        st.metric("Avg Premiums", format_currency(avg_premium))
                        st.metric("Avg Net", format_currency(avg_net),
                                 delta_color="normal" if avg_net > 0 else "inverse")
                    
                    # Detailed comparison table
                    with st.expander("📊 Detailed Monthly Comparison"):
                        display_comparison = comparison_df.copy()
                        for col in ['Loan Cashflows', 'LS Premiums', 'Net Cash Flow']:
                            display_comparison[col] = display_comparison[col].apply(format_currency)
                        
                        # Add highlighting for negative net flows
                        def highlight_negative(val):
                            if isinstance(val, str) and val.startswith('$'):
                                num_val = float(val.replace('$', '').replace(',', ''))
                                if num_val < 0:
                                    return 'color: #FF6B6B'
                            return ''
                        
                        styled_df = display_comparison.style.applymap(highlight_negative, subset=['Net Cash Flow'])
                        st.dataframe(styled_df, use_container_width=True, hide_index=True)
            
            except Exception as e:
                st.error(f"Error creating cashflow vs premium analysis: {str(e)}")
                st.info("Please check that both loan cashflow data and life settlement premium data are properly loaded.")

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