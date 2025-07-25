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

def process_life_settlement_data(ls_file):
    """Process Life Settlement Excel file and return summary data"""
    try:
        ls_wb = load_workbook(ls_file, data_only=True)
        
        if 'Valuation Summary' not in ls_wb.sheetnames or 'Premium Stream' not in ls_wb.sheetnames:
            return None
        
        val_sheet = ls_wb['Valuation Summary']
        premium_sheet = ls_wb['Premium Stream']
        
        policies = []
        
        for row in range(3, 200):
            policy_id_cell = val_sheet[f'B{row}']
            if not policy_id_cell.value:
                break
                
            try:
                # Get NDB value first
                ndb_cell_value = val_sheet[f'V{row}'].value
                ndb_value = safe_float(str(ndb_cell_value or '0').replace('$', '').replace(',', ''))
                
                # If NDB is 0, check for Face Amount column (try common locations)
                if ndb_value == 0:
                    # Try column W first (next to V)
                    face_cell_value = val_sheet[f'W{row}'].value
                    face_amount = safe_float(str(face_cell_value or '0').replace('$', '').replace(',', ''))
                    if face_amount == 0:
                        # Try other possible columns for Face Amount
                        for col in ['X', 'Y', 'U', 'T']:
                            face_cell_value = val_sheet[f'{col}{row}'].value
                            face_amount = safe_float(str(face_cell_value or '0').replace('$', '').replace(',', ''))
                            if face_amount > 0:
                                break
                    if face_amount > 0:
                        ndb_value = face_amount
                
                policy_data = {
                    'Policy_ID': str(policy_id_cell.value),
                    'Insured_ID': str(val_sheet[f'C{row}'].value or ''),
                    'Name': str(val_sheet[f'D{row}'].value or ''),
                    'Age': safe_float(val_sheet[f'F{row}'].value),
                    'Gender': str(val_sheet[f'G{row}'].value or ''),
                    'NDB': ndb_value,
                    'Valuation': safe_float(str(val_sheet[f'Z{row}'].value or '0').replace('$', '').replace(',', '')),
                    'Cost_Basis': safe_float(str(val_sheet[f'AB{row}'].value or '0').replace('$', '').replace(',', '')),
                    'Remaining_LE': safe_float(val_sheet[f'AC{row}'].value),
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
        
        # Process monthly premiums
        monthly_premiums = {}
        policy_premiums = {}
        month_columns = ['M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X']
        
        month_headers = []
        for col in month_columns:
            header = premium_sheet[f'{col}2'].value
            if header:
                month_headers.append((col, str(header)))
        
        for col_letter, month_name in month_headers:
            month_total = 0
            for prem_row in range(3, len(policies) + 3):
                try:
                    lyric_id_cell = premium_sheet[f'B{prem_row}']
                    if lyric_id_cell.value:
                        lyric_id = str(lyric_id_cell.value)
                        premium_cell = premium_sheet[f'{col_letter}{prem_row}']
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
            
            # Show loan status breakdown
            st.write("\nLoan Status Breakdown:")
            status_counts = loans_df['Status'].value_counts()
            st.write(status_counts)
            
            # Show loans by status with key fields
            st.write("\nLoans by Status:")
            for status in ['Active', 'Closed', 'Not Started']:
                st.write(f"\n{status} Loans:")
                status_loans = loans_df[loans_df['Status'] == status][['Sheet', 'Borrower', 'Original Loan Balance', 'Current Loan Balance', 'Status']].copy()
                st.dataframe(status_loans)
            
            missing_sheets = set(loan_sheets) - set(loans_df['Sheet'].tolist())
            if missing_sheets:
                st.warning(f"Sheets not showing in tables: {missing_sheets}")
            
            # Check for any loans that might be miscategorized
            st.write("\nPotential Issues:")
            # Active loans with zero balance
            active_zero_balance = loans_df[(loans_df['Status'] == 'Active') & (loans_df['Current Loan Balance'] == 0)]
            if len(active_zero_balance) > 0:
                st.warning("Active loans with zero balance found:")
                st.dataframe(active_zero_balance[['Sheet', 'Borrower', 'Current Loan Balance']])
            
            # Closed loans with non-zero balance
            closed_nonzero = loans_df[(loans_df['Status'] == 'Closed') & (loans_df['Current Loan Balance'] > 0)]
            if len(closed_nonzero) > 0:
                st.warning("Closed loans with non-zero balance found:")
                st.dataframe(closed_nonzero[['Sheet', 'Borrower', 'Current Loan Balance']])
        
        # Separate loans by status
        active_loans = loans_df[loans_df['Status'] == 'Active'].copy()
        closed_loans = loans_df[loans_df['Status'] == 'Closed'].copy()
        not_started_loans = loans_df[loans_df['Status'] == 'Not Started'].copy()
        
        # Fix the count issue - ensure we're counting correctly
        # Remove any duplicate counting
        active_loans = active_loans.drop_duplicates(subset=['Sheet'])
        closed_loans = closed_loans.drop_duplicates(subset=['Sheet'])
        not_started_loans = not_started_loans.drop_duplicates(subset=['Sheet'])
        
        # Sort by current balance or original balance
        active_loans = active_loans.sort_values('Current Loan Balance', ascending=False)
        closed_loans = closed_loans.sort_values('Original Loan Balance', ascending=False)
        not_started_loans = not_started_loans.sort_values('Loan Start Date')
        
        # Display portfolio summary
        st.markdown("<h2 style='color: #FDB813; margin-top: 2rem; font-size: 2rem;'>📊 Portfolio Summary</h2>", unsafe_allow_html=True)
        
        # Calculate all metrics first
        total_original = loans_df['Original Loan Balance'].sum()
        active_orig_balance = active_loans['Original Loan Balance'].sum()
        active_current_balance = active_loans['Current Loan Balance'].sum()
        total_repaid_principal = loans_df['Total Principal Repaid'].sum()
        total_repaid_interest = loans_df['Total Interest Repaid'].sum()
        total_collected = total_repaid_principal + total_repaid_interest
        
        # Create styled summary boxes
        st.markdown("""
        <style>
        .summary-box {
            background-color: #2d2d2d;
            border: 1px solid #3d3d3d;
            border-radius: 8px;
            padding: 2rem;
            margin-bottom: 1.5rem;
            overflow: visible;
        }
        .summary-title {
            color: #FDB813;
            font-size: 1.4rem;
            font-weight: 700;
            margin-bottom: 1.5rem;
            display: flex;
            align-items: center;
            letter-spacing: 0.5px;
        }
        .summary-metrics {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 2rem;
        }
        .metric-item {
            text-align: center;
        }
        .metric-label {
            color: #AAAAAA;
            font-size: 0.95rem;
            font-weight: 500;
            margin-bottom: 0.5rem;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        .metric-value {
            color: #FFFFFF;
            font-size: 2.2rem;
            font-weight: 800;
            line-height: 1.2;
        }
        .metric-subvalue {
            color: #999999;
            font-size: 1.1rem;
            font-weight: 500;
            margin-top: 0.25rem;
        }
        </style>
        """, unsafe_allow_html=True)
        
        # Portfolio Overview Box
        st.markdown("""
        <div class='summary-box'>
            <div class='summary-title'>📁 Portfolio Overview</div>
            <div class='summary-metrics'>
                <div class='metric-item'>
                    <div class='metric-label'>Active Loans</div>
                    <div class='metric-value'>{}</div>
                </div>
                <div class='metric-item'>
                    <div class='metric-label'>Closed Loans</div>
                    <div class='metric-value'>{}</div>
                </div>
                <div class='metric-item'>
                    <div class='metric-label'>Not Started</div>
                    <div class='metric-value'>{}</div>
                </div>
                <div class='metric-item'>
                    <div class='metric-label'>Total Loans</div>
                    <div class='metric-value'>{}</div>
                </div>
            </div>
        </div>
        """.format(len(active_loans), len(closed_loans), len(not_started_loans), len(loans_df)), 
        unsafe_allow_html=True)
        
        # Active Loans Summary Box
        if len(active_loans) > 0:
            # Calculate weighted average interest rate (excluding opportunity fund)
            # Filter out loans with unusually high interest rates (>30%)
            reasonable_rate_loans = active_loans[active_loans['Annual Interest Rate'] <= 0.30].copy()
            if len(reasonable_rate_loans) > 0:
                reasonable_rate_loans['Weighted_Rate'] = reasonable_rate_loans['Original Loan Balance'] * reasonable_rate_loans['Annual Interest Rate']
                weighted_avg_rate = reasonable_rate_loans['Weighted_Rate'].sum() / reasonable_rate_loans['Original Loan Balance'].sum()
            else:
                weighted_avg_rate = 0
            
            # Count amortizing vs interest only loans
            amortizing_loans = len(active_loans[~active_loans['Is Interest Only']])
            interest_only_loans = len(active_loans[active_loans['Is Interest Only']])
            
            # Calculate average maturity and average age
            today = pd.Timestamp.now()
            active_loans['Maturity Date'] = pd.to_datetime(active_loans['Maturity Date'])
            active_loans['Loan Start Date'] = pd.to_datetime(active_loans['Loan Start Date'])
            
            # Average maturity calculation
            valid_maturities = active_loans[active_loans['Maturity Date'] > today]['Maturity Date']
            if len(valid_maturities) > 0:
                avg_months_to_maturity = (valid_maturities - today).dt.days.mean() / 30.44
                avg_years_to_maturity = avg_months_to_maturity / 12
            else:
                avg_months_to_maturity = 0
                avg_years_to_maturity = 0
            
            # Average age calculation
            valid_start_dates = active_loans[active_loans['Loan Start Date'] <= today]['Loan Start Date']
            if len(valid_start_dates) > 0:
                avg_months_since_start = (today - valid_start_dates).dt.days.mean() / 30.44
                avg_years_since_start = avg_months_since_start / 12
            else:
                avg_months_since_start = 0
                avg_years_since_start = 0
            
            # Debug output for average loan age
            st.write(f"DEBUG - Average Loan Age: {avg_years_since_start:.1f} years ({avg_months_since_start:.0f} months)")
            
            # Create the HTML with all values pre-formatted
            active_loans_html = f"""
            <div class='summary-box'>
                <div class='summary-title'>💰 Active Loans Summary</div>
                <div class='summary-metrics'>
                    <div class='metric-item'>
                        <div class='metric-label'>Original Balance</div>
                        <div class='metric-value'>{format_currency(active_orig_balance)}</div>
                    </div>
                    <div class='metric-item'>
                        <div class='metric-label'>Current Balance</div>
                        <div class='metric-value'>{format_currency(active_current_balance)}</div>
                    </div>
                    <div class='metric-item'>
                        <div class='metric-label'>Avg Interest Rate</div>
                        <div class='metric-value'>{format_percent(weighted_avg_rate)}</div>
                        <div class='metric-subvalue' style='color: #888888; font-size: 0.9rem;'>(excl. high-rate loans)</div>
                    </div>
                    <div class='metric-item'>
                        <div class='metric-label'>Avg Maturity</div>
                        <div class='metric-value'>{avg_years_to_maturity:.1f} yrs</div>
                        <div class='metric-subvalue'>({avg_months_to_maturity:.0f} months)</div>
                    </div>
                </div>
                <div style='margin-top: 2rem; padding-top: 2rem; border-top: 1px solid #3d3d3d; min-height: 120px;'>
                    <div style='display: grid; grid-template-columns: repeat(3, 1fr); gap: 2rem; width: 100%;'>
                        <div style='text-align: center; padding: 0.5rem;'>
                            <div class='metric-label' style='margin-bottom: 0.5rem;'>Amortizing Loans</div>
                            <div style='color: #FFFFFF; font-size: 1.8rem; font-weight: 700;'>{amortizing_loans}</div>
                        </div>
                        <div style='text-align: center; padding: 0.5rem;'>
                            <div class='metric-label' style='margin-bottom: 0.5rem;'>Interest Only</div>
                            <div style='color: #FFFFFF; font-size: 1.8rem; font-weight: 700;'>{interest_only_loans}</div>
                        </div>
                        <div style='text-align: center; padding: 0.5rem;'>
                            <div class='metric-label' style='margin-bottom: 0.5rem;'>Average Loan Age</div>
                            <div style='color: #FFFFFF; font-size: 1.8rem; font-weight: 700;'>{avg_years_since_start:.1f} yrs</div>
                            <div style='color: #999999; font-size: 1rem; margin-top: 0.25rem;'>({avg_months_since_start:.0f} months)</div>
                        </div>
                    </div>
                </div>
            </div>
            """
            
            st.markdown(active_loans_html, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class='summary-box'>
                <div class='summary-title'>💰 Active Loans Summary</div>
                <div style='text-align: center; color: #999999; padding: 3rem; font-size: 1.2rem;'>
                    No active loans in portfolio
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        # Historical Performance Box
        st.markdown("""
        <div class='summary-box'>
            <div class='summary-title'>📈 Historical Performance</div>
            <div style='display: grid; grid-template-columns: repeat(3, 1fr); gap: 2rem;'>
                <div class='metric-item'>
                    <div class='metric-label'>Principal Repaid</div>
                    <div class='metric-value'>{}</div>
                </div>
                <div class='metric-item'>
                    <div class='metric-label'>Interest Earned</div>
                    <div class='metric-value'>{}</div>
                </div>
                <div class='metric-item'>
                    <div class='metric-label'>Total Collections</div>
                    <div class='metric-value'>{}</div>
                </div>
            </div>
        </div>
        """.format(
            format_currency(total_repaid_principal),
            format_currency(total_repaid_interest),
            format_currency(total_collected)
        ), unsafe_allow_html=True)
        
        # Active Loans Breakdown
        if len(active_loans) > 0:
            # Additional active loans insights
            st.markdown("<h2 style='color: #FDB813; margin-top: 2rem;'>📈 Active Loans Breakdown</h2>", unsafe_allow_html=True)
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                # Loan size distribution
                st.markdown("**Loan Size Distribution**")
                size_bins = [0, 100000, 250000, 500000, 1000000, float('inf')]
                size_labels = ['< $100K', '$100K-$250K', '$250K-$500K', '$500K-$1M', '> $1M']
                active_loans['Size_Category'] = pd.cut(active_loans['Current Loan Balance'], bins=size_bins, labels=size_labels)
                size_dist = active_loans['Size_Category'].value_counts().sort_index()
                
                for category, count in size_dist.items():
                    st.markdown(f"<div style='display: flex; justify-content: space-between; padding: 0.25rem 0; border-bottom: 1px solid #3d3d3d;'>"
                               f"<span style='color: #FDB813;'>{category}:</span>"
                               f"<span style='color: #FFFFFF; font-weight: 600;'>{count} loans</span>"
                               f"</div>", unsafe_allow_html=True)
            
            with col2:
                # Interest rate distribution
                st.markdown("**Interest Rate Distribution**")
                rate_bins = [0, 0.05, 0.075, 0.10, 0.125, float('inf')]
                rate_labels = ['< 5%', '5%-7.5%', '7.5%-10%', '10%-12.5%', '> 12.5%']
                active_loans['Rate_Category'] = pd.cut(active_loans['Annual Interest Rate'], bins=rate_bins, labels=rate_labels)
                rate_dist = active_loans['Rate_Category'].value_counts().sort_index()
                
                for category, count in rate_dist.items():
                    st.markdown(f"<div style='display: flex; justify-content: space-between; padding: 0.25rem 0; border-bottom: 1px solid #3d3d3d;'>"
                               f"<span style='color: #FDB813;'>{category}:</span>"
                               f"<span style='color: #FFFFFF; font-weight: 600;'>{count} loans</span>"
                               f"</div>", unsafe_allow_html=True)
            
            with col3:
                # Maturity distribution
                st.markdown("**Maturity Distribution**")
                active_loans['Months_to_Maturity'] = (active_loans['Maturity Date'] - today).dt.days / 30.44
                maturity_bins = [0, 6, 12, 24, 36, float('inf')]
                maturity_labels = ['< 6 months', '6-12 months', '1-2 years', '2-3 years', '> 3 years']
                active_loans['Maturity_Category'] = pd.cut(active_loans['Months_to_Maturity'], bins=maturity_bins, labels=maturity_labels)
                maturity_dist = active_loans['Maturity_Category'].value_counts().sort_index()
                
                for category, count in maturity_dist.items():
                    if pd.notna(count):
                        st.markdown(f"<div style='display: flex; justify-content: space-between; padding: 0.25rem 0; border-bottom: 1px solid #3d3d3d;'>"
                                   f"<span style='color: #FDB813;'>{category}:</span>"
                                   f"<span style='color: #FFFFFF; font-weight: 600;'>{count} loans</span>"
                                   f"</div>", unsafe_allow_html=True)
        
        # Display active loans table
        st.markdown("<h2 style='color: #FDB813; margin-top: 2rem;'>💰 Active Loans Detail</h2>", unsafe_allow_html=True)
        
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
        
        # Cash flow projection - Updated to 12 months
        st.markdown("<h2 style='color: #FDB813; margin-top: 2rem;'>💸 Cash Flow Projection (Next 12 Months)</h2>", unsafe_allow_html=True)
        
        today = datetime.now()
        cashflow_data = []
        
        for borrower, amort_df in loan_details.items():
            if borrower in not_started_loans['Borrower'].values:
                continue
                
            if 'Month' in amort_df.columns and 'Loan Repayment' in amort_df.columns:
                amort_df['Month'] = pd.to_datetime(amort_df['Month'])
                upcoming = amort_df[(amort_df['Month'] > today) & 
                                  (amort_df['Month'] <= today + relativedelta(months=12))]
                
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
            
            # Identify months with large payments (>$500k)
            monthly_summary['Is_Large'] = monthly_summary['Payment Amount'] > 500000
            
            # Create styled display
            st.markdown("""
            <style>
            .cashflow-highlight {
                background-color: #3d3d3d !important;
                border: 2px solid #FDB813 !important;
            }
            </style>
            """, unsafe_allow_html=True)
            
            col1, col2 = st.columns([3, 1])
            
            with col1:
                display_summary = monthly_summary.copy()
                
                # Create HTML table with highlighting
                table_html = "<div style='background-color: #2d2d2d; padding: 1rem; border-radius: 8px;'>"
                table_html += "<table style='width: 100%; color: white;'>"
                table_html += "<thead><tr style='border-bottom: 2px solid #FDB813;'>"
                table_html += "<th style='padding: 0.75rem; text-align: left;'>Month</th>"
                table_html += "<th style='padding: 0.75rem; text-align: right;'>Payment Amount</th>"
                table_html += "<th style='padding: 0.75rem; text-align: right;'>Interest</th>"
                table_html += "<th style='padding: 0.75rem; text-align: right;'>Principal</th>"
                table_html += "</tr></thead><tbody>"
                
                for idx, row in display_summary.iterrows():
                    row_class = "cashflow-highlight" if row['Is_Large'] else ""
                    table_html += f"<tr class='{row_class}' style='border-bottom: 1px solid #3d3d3d;'>"
                    table_html += f"<td style='padding: 0.75rem;'>{row['Month']}</td>"
                    table_html += f"<td style='padding: 0.75rem; text-align: right; font-weight: {'bold' if row['Is_Large'] else 'normal'}; color: {'#FDB813' if row['Is_Large'] else '#FFFFFF'};'>{format_currency(row['Payment Amount'])}</td>"
                    table_html += f"<td style='padding: 0.75rem; text-align: right;'>{format_currency(row['Interest'])}</td>"
                    table_html += f"<td style='padding: 0.75rem; text-align: right;'>{format_currency(row['Principal'])}</td>"
                    table_html += "</tr>"
                
                table_html += "</tbody></table></div>"
                st.markdown(table_html, unsafe_allow_html=True)
                
                # Add note about highlighted months
                large_months = monthly_summary[monthly_summary['Is_Large']]['Month'].tolist()
                if large_months:
                    st.info(f"⭐ Highlighted months have payments exceeding $500,000: {', '.join(large_months)}")
            
            with col2:
                total_expected = monthly_summary['Payment Amount'].sum()
                total_interest = monthly_summary['Interest'].sum()
                total_principal = monthly_summary['Principal'].sum()
                avg_monthly = total_expected / len(monthly_summary) if len(monthly_summary) > 0 else 0
                
                st.metric("12-Month Total", format_currency(total_expected))
                st.metric("Total Interest", format_currency(total_interest))
                st.metric("Total Principal", format_currency(total_principal))
                st.metric("Monthly Average", format_currency(avg_monthly))
                
                # Add breakdown by quarter
                st.markdown("<h4 style='color: #FDB813; margin-top: 2rem;'>Quarterly View</h4>", unsafe_allow_html=True)
                
                # Calculate quarters
                cashflow_df['Quarter'] = cashflow_df['Payment Date'].dt.to_period('Q')
                quarterly_summary = cashflow_df.groupby('Quarter')['Payment Amount'].sum()
                
                for quarter, amount in quarterly_summary.items():
                    st.metric(f"{quarter}", format_currency(amount))
        else:
            st.info("No upcoming payments in the next 12 months")

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
                        # Create a larger line chart that fills the available space
                        chart_data = comparison_df.set_index('Month')[['Loan Cashflows', 'LS Premiums', 'Net Cash Flow']]
                        
                        # Use container with custom height
                        chart_container = st.container()
                        with chart_container:
                            # Add custom CSS to make the chart taller
                            st.markdown("""
                            <style>
                            [data-testid="stLineChart"] {
                                height: 600px !important;
                            }
                            </style>
                            """, unsafe_allow_html=True)
                            
                            st.line_chart(chart_data, height=600, use_container_width=True)
                    
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

        # Display LS data if available (but after all loan data)
        if ls_data:
            st.markdown("<h2 style='color: #FDB813; margin-top: 3rem; font-size: 2rem;'>🏥 Life Settlement Portfolio</h2>", unsafe_allow_html=True)
            
            # Key Metrics Box with Unrealized Gain/Loss
            st.markdown("""
            <div class='summary-box'>
                <div class='summary-title'>📊 Portfolio Metrics</div>
                <div class='summary-metrics'>
                    <div class='metric-item'>
                        <div class='metric-label'>Total Policies</div>
                        <div class='metric-value'>{}</div>
                    </div>
                    <div class='metric-item'>
                        <div class='metric-label'>Total NDB (Face Value)</div>
                        <div class='metric-value'>{}</div>
                    </div>
                    <div class='metric-item'>
                        <div class='metric-label'>Cost Basis</div>
                        <div class='metric-value'>{}</div>
                    </div>
                    <div class='metric-item'>
                        <div class='metric-label'>Total Valuation</div>
                        <div class='metric-value'>{}</div>
                    </div>
                </div>
                <div style='margin-top: 1.5rem; padding: 1.5rem; background-color: #3d3d3d; border-radius: 6px;'>
                    <div style='display: grid; grid-template-columns: 1fr 2fr; gap: 2rem; align-items: center;'>
                        <div style='text-align: center;'>
                            <div class='metric-label'>Unrealized Gain/(Loss)</div>
                            <div style='color: {}; font-size: 2rem; font-weight: 700;'>{}</div>
                            <div style='color: #999999; font-size: 1rem; margin-top: 0.25rem;'>{:.1f}% return</div>
                        </div>
                        <div style='display: grid; grid-template-columns: repeat(3, 1fr); gap: 2rem;'>
                            <div class='metric-item'>
                                <div class='metric-label'>Average Age</div>
                                <div style='color: #FFFFFF; font-size: 1.6rem; font-weight: 700;'>{:.1f} years</div>
                            </div>
                            <div class='metric-item'>
                                <div class='metric-label'>% Male</div>
                                <div style='color: #FFFFFF; font-size: 1.6rem; font-weight: 700;'>{:.1f}%</div>
                            </div>
                            <div class='metric-item'>
                                <div class='metric-label'>Avg Remaining LE</div>
                                <div style='color: #FFFFFF; font-size: 1.6rem; font-weight: 700;'>{:.1f} months</div>
                            </div>
                        </div>
                    </div>
                </div>
                <div style='margin-top: 1.5rem; text-align: center;'>
                    <div class='metric-label'>Premiums % of Face</div>
                    <div style='color: #FFFFFF; font-size: 1.8rem; font-weight: 700;'>{:.2f}%</div>
                </div>
            </div>
            """.format(
                ls_data['summary']['total_policies'],
                format_currency(ls_data['summary']['total_ndb']),
                format_currency(ls_data['summary']['total_cost_basis']),
                format_currency(ls_data['summary']['total_valuation']),
                '#4ECDC4' if ls_data['summary']['total_valuation'] >= ls_data['summary']['total_cost_basis'] else '#FF6B6B',
                format_currency(ls_data['summary']['total_valuation'] - ls_data['summary']['total_cost_basis']),
                ((ls_data['summary']['total_valuation'] - ls_data['summary']['total_cost_basis']) / ls_data['summary']['total_cost_basis'] * 100) if ls_data['summary']['total_cost_basis'] > 0 else 0,
                ls_data['summary']['avg_age'],
                ls_data['summary']['male_percentage'],
                ls_data['summary']['avg_remaining_le'],
                ls_data['summary']['premiums_as_pct_face']
            ), unsafe_allow_html=True)
            
            # Monthly Premium Projections
            if ls_data['monthly_premiums']:
                st.markdown("""
                <div class='summary-box'>
                    <div class='summary-title'>💵 Monthly Premium Projections</div>
                    <div style='display: grid; grid-template-columns: repeat(6, 1fr); gap: 2rem; padding: 1rem 0;'>
                """, unsafe_allow_html=True)
                
                premium_items = list(ls_data['monthly_premiums'].items())
                
                # Display all months in a grid
                for month, amount in premium_items:
                    st.markdown(f"""
                        <div class='metric-item'>
                            <div class='metric-label'>{month}</div>
                            <div class='metric-value'>{format_currency(amount)}</div>
                        </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("</div></div>", unsafe_allow_html=True)
            
            # Policy Details Table
            st.markdown("<h3 style='color: #FFFFFF; margin-top: 2rem; font-size: 1.4rem;'>📋 Policy Details</h3>", unsafe_allow_html=True)
            
            if ls_data['policies']:
                policies_df = pd.DataFrame(ls_data['policies'])
                
                # Calculate unrealized gain/loss for each policy
                policies_df['Unrealized_Gain_Loss'] = policies_df['Valuation'] - policies_df['Cost_Basis']
                policies_df['Gain_Loss_Pct'] = (policies_df['Unrealized_Gain_Loss'] / policies_df['Cost_Basis'] * 100).fillna(0)
                
                # Format for display
                display_policy_df = policies_df.copy()
                display_policy_df['NDB'] = display_policy_df['NDB'].apply(format_currency)
                display_policy_df['Cost_Basis'] = display_policy_df['Cost_Basis'].apply(format_currency)
                display_policy_df['Valuation'] = display_policy_df['Valuation'].apply(format_currency)
                
                # Format unrealized gain/loss with color coding
                def format_gain_loss(value):
                    if value >= 0:
                        return f"<span style='color: #4ECDC4;'>{format_currency(value)}</span>"
                    else:
                        return f"<span style='color: #FF6B6B;'>{format_currency(value)}</span>"
                
                display_policy_df['Unrealized_Gain_Loss_Display'] = display_policy_df['Unrealized_Gain_Loss'].apply(format_gain_loss)
                display_policy_df['Annual_Premium'] = display_policy_df['Annual_Premium'].apply(format_currency)
                display_policy_df['Premium_Pct_Face'] = display_policy_df['Premium_Pct_Face'].apply(lambda x: f"{x:.2f}%")
                
                # Rename columns
                display_policy_df = display_policy_df.rename(columns={
                    'Policy_ID': 'Policy ID',
                    'Name': 'Name',
                    'Age': 'Age',
                    'Gender': 'Gender',
                    'NDB': 'Face Value',
                    'Cost_Basis': 'Cost Basis',
                    'Valuation': 'Valuation',
                    'Unrealized_Gain_Loss_Display': 'Unrealized Gain/(Loss)',
                    'Annual_Premium': 'Annual Premium',
                    'Premium_Pct_Face': 'Premium % Face'
                })
                
                # Create HTML table for better formatting
                table_html = """
                <div style='background-color: #2d2d2d; padding: 1rem; border-radius: 8px; overflow-x: auto;'>
                    <table style='width: 100%; color: white; border-collapse: collapse;'>
                        <thead>
                            <tr style='background-color: #FDB813; color: #1a1a1a;'>
                                <th style='padding: 0.75rem; text-align: left; font-weight: 600;'>Policy ID</th>
                                <th style='padding: 0.75rem; text-align: left; font-weight: 600;'>Name</th>
                                <th style='padding: 0.75rem; text-align: center; font-weight: 600;'>Age</th>
                                <th style='padding: 0.75rem; text-align: center; font-weight: 600;'>Gender</th>
                                <th style='padding: 0.75rem; text-align: right; font-weight: 600;'>Face Value</th>
                                <th style='padding: 0.75rem; text-align: right; font-weight: 600;'>Cost Basis</th>
                                <th style='padding: 0.75rem; text-align: right; font-weight: 600;'>Valuation</th>
                                <th style='padding: 0.75rem; text-align: right; font-weight: 600;'>Unrealized Gain/(Loss)</th>
                                <th style='padding: 0.75rem; text-align: right; font-weight: 600;'>Annual Premium</th>
                                <th style='padding: 0.75rem; text-align: right; font-weight: 600;'>Premium % Face</th>
                            </tr>
                        </thead>
                        <tbody>
                """
                
                for _, row in display_policy_df.iterrows():
                    table_html += f"""
                        <tr style='border-bottom: 1px solid #3d3d3d;'>
                            <td style='padding: 0.75rem;'>{row['Policy ID']}</td>
                            <td style='padding: 0.75rem;'>{row['Name']}</td>
                            <td style='padding: 0.75rem; text-align: center;'>{row['Age']:.0f}</td>
                            <td style='padding: 0.75rem; text-align: center;'>{row['Gender']}</td>
                            <td style='padding: 0.75rem; text-align: right;'>{row['Face Value']}</td>
                            <td style='padding: 0.75rem; text-align: right;'>{row['Cost Basis']}</td>
                            <td style='padding: 0.75rem; text-align: right;'>{row['Valuation']}</td>
                            <td style='padding: 0.75rem; text-align: right;'>{row['Unrealized Gain/(Loss)']}</td>
                            <td style='padding: 0.75rem; text-align: right;'>{row['Annual Premium']}</td>
                            <td style='padding: 0.75rem; text-align: right;'>{row['Premium % Face']}</td>
                        </tr>
                    """
                
                table_html += """
                        </tbody>
                    </table>
                </div>
                """
                
                st.markdown(table_html, unsafe_allow_html=True)
        
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