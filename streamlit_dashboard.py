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
    try:
        if isinstance(value, str):
            if value.lower() in ['interest only', 'n/a', '']:
                return 0.0
            value = value.replace('$', '').replace(',', '')
        return float(value)
    except (ValueError, TypeError):
        return 0.0

def excel_date_to_datetime(serial_date):
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
    return f"${value:,.2f}" if pd.notna(value) and value != 0 else "$0.00"

def format_percent(value):
    return f"{value:.2%}" if pd.notna(value) and value != 0 else "0.00%"

def get_cell_value(sheet, locations, default=None):
    for location in locations:
        if sheet[location].value is not None:
            return sheet[location].value
    return default

def process_life_settlement_data(ls_file):
    try:
        ls_wb = load_workbook(ls_file, data_only=True)
        available_sheets = ls_wb.sheetnames
        if 'Valuation Summary' not in available_sheets or 'Premium Stream' not in available_sheets:
            return None
        val_sheet = ls_wb['Valuation Summary']
        premium_sheet = ls_wb['Premium Stream']
        policies = []
        for row in range(3, 200):
            policy_id_cell = val_sheet[f'B{row}']
            if not policy_id_cell.value:
                break
            try:
                policy_data = {
                    'Policy_ID': str(policy_id_cell.value),
                    'Insured_ID': str(val_sheet[f'C{row}'].value or ''),
                    'Name': str(val_sheet[f'D{row}'].value or ''),
                    'Age': safe_float(val_sheet[f'F{row}'].value),
                    'Gender': str(val_sheet[f'G{row}'].value or ''),
                    'NDB': safe_float(str(val_sheet[f'V{row}'].value or '0').replace('$', '').replace(',', '')),
                    'Valuation': safe_float(str(val_sheet[f'Z{row}'].value or '0').replace('$', '').replace(',', '')),
                    'Cost_Basis': safe_float(str(val_sheet[f'AB{row}'].value or '0').replace('$', '').replace(',', '')),
                    'Remaining_LE': safe_float(val_sheet[f'AC{row}'].value),
                }
                policies.append(policy_data)
            except Exception:
                continue
        if len(policies) == 0:
            return None
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
                except Exception:
                    continue
            monthly_premiums[month_name] = month_total
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
    except Exception:
        return None

# Main app
st.markdown("""
<div style='background-color: #1a1a1a; padding: 2rem 0; margin: -2rem -2rem 2rem -2rem; border-bottom: 4px solid #FDB813;'>
    <h1 style='text-align: center; color: #FFFFFF; font-size: 2.5rem; margin: 0;'>
        <span style='color: #FDB813;'>⚡</span> Sirocco I LP Portfolio Dashboard
    </h1>
    <p style='text-align: center; color: #999999; margin-top: 0.5rem;'>Comprehensive Portfolio Management System</p>
</div>
""", unsafe_allow_html=True)

st.markdown("<h2 style='color: #FDB813;'>📁 File Upload</h2>", unsafe_allow_html=True)
col1, col2, col3 = st.columns(3)
with col1:
    master_file = st.file_uploader("📊 Loan Portfolio File", type=["xlsx"], help="Upload the master Excel file containing loan data")
with col2:
    ls_file = st.file_uploader("🏥 Life Settlement File", type=["xlsx"], help="Upload the LS portfolio file")
with col3:
    remittance_file = st.file_uploader("📄 Remittance File", type=["csv", "xlsx"], help="Upload monthly remittance file")

ls_data = None
loan_summary = {
    'total_original': 0,
    'total_repaid_principal': 0,
    'total_repaid_interest': 0,
    'active_loans': 0,
    'closed_loans': 0,
    'loans_df': pd.DataFrame()
}

if ls_file:
    ls_data = process_life_settlement_data(ls_file)
    if ls_data:
        st.success(f"✅ Life Settlement data loaded: {ls_data['summary']['total_policies']} policies, {format_currency(ls_data['summary']['total_ndb'])} face value")
    else:
        st.error("❌ Failed to load Life Settlement data. Please check file format.")

if master_file:
    try:
        wb = load_workbook(master_file, data_only=True)
        loan_sheets = [s for s in wb.sheetnames if s.startswith('#') and s != '#AddSheet']
        dashboard_sheet = wb['Dashboard']
        as_of_date = dashboard_sheet['E3'].value
        if isinstance(as_of_date, str):
            as_of_date = pd.to_datetime(as_of_date)
        elif isinstance(as_of_date, (int, float)):
            as_of_date = excel_date_to_datetime(as_of_date)
        loans = []
        loan_details = {}
        for sheet_name in loan_sheets:
            sheet = wb[sheet_name]
            borrower = get_cell_value(sheet, ['B2', 'A2'], f"Unknown ({sheet_name})")
            if borrower == "" or borrower is None:
                borrower = f"Unknown ({sheet_name})"
            b3_value = sheet['B3'].value
            if isinstance(b3_value, str) and 'loan' in str(b3_value).lower():
                loan_amount = safe_float(sheet['C3'].value)
                interest_rate = safe_float(sheet['C4'].value)
                loan_period = safe_float(sheet['C5'].value)
                payment_amount_val = sheet['C6'].value
                loan_start = excel_date_to_datetime(sheet['C7'].value)
            else:
                loan_amount = safe_float(sheet['B3'].value)
                interest_rate = safe_float(sheet['B4'].value)
                loan_period = safe_float(sheet['B5'].value)
                payment_amount_val = sheet['B6'].value
                loan_start = excel_date_to_datetime(sheet['B7'].value)
            if isinstance(payment_amount_val, str) and payment_amount_val.lower() == 'interest only':
                payment_amount = loan_amount * (interest_rate / 12)
            else:
                payment_amount = safe_float(payment_amount_val)
            loan_info = {
                'Sheet': sheet_name,
                'Borrower': borrower,
                'Original Loan Balance': loan_amount,
                'Annual Interest Rate': interest_rate,
                'Loan Period (months)': loan_period,
                'Payment Amount': payment_amount,
                'Loan Start Date': loan_start,
                'Last Payment Amount': 0,
                'Current Loan Balance': loan_amount,
                'Total Principal Repaid': 0,
                'Total Interest Repaid': 0,
                'Status': 'Active'
            }
            amort_data = []
            row = 11
            while row < 100:
                month_cell = sheet[f'A{row}']
                if month_cell.value is None:
                    break
                amort_row = {
                    'Month': excel_date_to_datetime(month_cell.value),
                    'Opening Balance': safe_float(sheet[f'C{row}'].value),
                    'Loan Repayment': safe_float(sheet[f'D{row}'].value),
                    'Interest Charged': safe_float(sheet[f'E{row}'].value),
                    'Capital Repaid': safe_float(sheet[f'F{row}'].value),
                    'Closing Balance': safe_float(sheet[f'G{row}'].value),
                    'Amount Paid': safe_float(sheet[f'K{row}'].value),
                }
                if amort_row['Opening Balance'] > 0 or amort_row['Closing Balance'] >= 0:
                    amort_data.append(amort_row)
                row += 1
            if amort_data:
                amort_df = pd.DataFrame(amort_data)
                if pd.notna(as_of_date):
                    amort_df['Month'] = pd.to_datetime(amort_df['Month'])
                    current_rows = amort_df[amort_df['Month'] <= as_of_date]
                    if not current_rows.empty:
                        current_row = current_rows.iloc[-1]
                        loan_info['Current Loan Balance'] = current_row['Closing Balance']
                        loan_info['Total Principal Repaid'] = current_rows['Capital Repaid'].sum()
                        loan_info['Total Interest Repaid'] = current_rows['Interest Charged'].sum()
                        loan_info['Last Payment Amount'] = current_row['Loan Repayment']
                if loan_info['Current Loan Balance'] == 0:
                    loan_info['Status'] = 'Closed'
                loan_details[borrower] = amort_df
            if loan_info['Original Loan Balance'] > 0:
                loans.append(loan_info)
        loans_df = pd.DataFrame(loans)
        if not loans_df.empty:
            loan_summary.update({
                'total_original': loans_df['Original Loan Balance'].sum(),
                'total_repaid_principal': loans_df['Total Principal Repaid'].sum(),
                'total_repaid_interest': loans_df['Total Interest Repaid'].sum(),
                'active_loans': len(loans_df[loans_df['Status'] == 'Active']),
                'closed_loans': len(loans_df[loans_df['Status'] == 'Closed']),
                'loans_df': loans_df
            })
            st.success(f"✅ Loan data loaded: {len(loans_df)} loans, {format_currency(loan_summary['total_original'])} total principal")
        else:
            st.error("❌ No loan data found in file")
    except Exception as e:
        st.error("❌ Failed to load loan data. Please check file format.")

if ls_data or not loan_summary['loans_df'].empty:
    st.markdown("<h2 style='color: #FDB813; margin-top: 2rem;'>📈 Portfolio Summary</h2>", unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        if ls_data:
            st.metric("LS Face Value", format_currency(ls_data['summary']['total_ndb']))
        st.metric("Loan Principal", format_currency(loan_summary['total_original']))
    with col2:
        if ls_data:
            st.metric("LS Valuation", format_currency(ls_data['summary']['total_valuation']))
        st.metric("Interest Collected", format_currency(loan_summary['total_repaid_interest']))
    with col3:
        if ls_data:
            st.metric("LS Policies", ls_data['summary']['total_policies'])
        st.metric("Active Loans", loan_summary['active_loans'])
    with col4:
        if ls_data:
            st.metric("Annual Premiums", format_currency(ls_data['summary']['total_annual_premiums']))
        st.metric("Closed Loans", loan_summary['closed_loans'])
    if ls_data and loan_summary['total_original'] > 0:
        st.markdown("<h3 style='color: #FFFFFF; margin-top: 2rem;'>Portfolio Allocation</h3>", unsafe_allow_html=True)
        total_ls_value = ls_data['summary']['total_valuation']
        total_loan_value = loan_summary['total_original']
        total_portfolio = total_ls_value + total_loan_value
        col1, col2 = st.columns(2)
        with col1:
            allocation_data = {
                'Asset Class': ['Life Settlements', 'Loans'],
                'Current Value': [format_currency(total_ls_value), format_currency(total_loan_value)],
                'Face/Principal': [format_currency(ls_data['summary']['total_ndb']), format_currency(total_loan_value)],
                'Allocation %': [f"{(total_ls_value/total_portfolio*100):.1f}%", f"{(total_loan_value/total_portfolio*100):.1f}%"]
            }
            allocation_df = pd.DataFrame(allocation_data)
            st.dataframe(allocation_df, use_container_width=True, hide_index=True)
        with col2:
            ls_cost_to_face = (ls_data['summary']['total_cost_basis'] / ls_data['summary']['total_ndb']) * 100
            portfolio_yield = ((ls_data['summary']['total_annual_premiums'] + loan_summary['total_repaid_interest']) / total_portfolio) * 100
            perf_col1, perf_col2 = st.columns(2)
            with perf_col1:
                st.metric("LS Cost/Face Ratio", f"{ls_cost_to_face:.1f}%")
                st.metric("Portfolio Yield", f"{portfolio_yield:.2f}%")
            with perf_col2:
                avg_policy = ls_data['summary']['total_ndb'] / ls_data['summary']['total_policies']
                avg_loan = loan_summary['total_original'] / (loan_summary['active_loans'] + loan_summary['closed_loans']) if (loan_summary['active_loans'] + loan_summary['closed_loans']) > 0 else 0
                st.metric("Avg Policy Size", format_currency(avg_policy))
                st.metric("Avg Loan Size", format_currency(avg_loan))
    if ls_data:
        st.markdown("<h2 style='color: #FDB813; margin-top: 3rem;'>🏥 Life Settlement Portfolio</h2>", unsafe_allow_html=True)
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("Total NDB", format_currency(ls_data['summary']['total_ndb']))
        with col2:
            st.metric("Total Valuation", format_currency(ls_data['summary']['total_valuation']))
        with col3:
            st.metric("Cost Basis", format_currency(ls_data['summary']['total_cost_basis']))
        with col4:
            st.metric("Avg Age", f"{ls_data['summary']['avg_age']:.1f} years")
        with col5:
            st.metric("% Male", f"{ls_data['summary']['male_percentage']:.1f}%")
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("Avg Remaining LE", f"{ls_data['summary']['avg_remaining_le']:.1f} months")
        with col2:
            st.metric("Annual Premiums", format_currency(ls_data['summary']['total_annual_premiums']))
        with col3:
            st.metric("Premiums % Face", f"{ls_data['summary']['premiums_as_pct_face']:.2f}%")
        with col4:
            cost_to_face = (ls_data['summary']['total_cost_basis'] / ls_data['summary']['total_ndb']) * 100
            st.metric("Cost to Face", f"{cost_to_face:.1f}%")
        with col5:
            value_to_cost = (ls_data['summary']['total_valuation'] / ls_data['summary']['total_cost_basis']) * 100
            st.metric("Value to Cost", f"{value_to_cost:.1f}%")
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
        st.markdown("<h3 style='color: #FFFFFF; margin-top: 2rem;'>Policy Details</h3>", unsafe_allow_html=True)
        if ls_data['policies']:
            policies_df = pd.DataFrame(ls_data['policies'])
            display_policy_df = policies_df.copy()
            display_policy_df['NDB'] = display_policy_df['NDB'].apply(format_currency)
            display_policy_df['Valuation'] = display_policy_df['Valuation'].apply(format_currency)
            display_policy_df['Cost_Basis'] = display_policy_df['Cost_Basis'].apply(format_currency)
            display_policy_df['Annual_Premium'] = display_policy_df['Annual_Premium'].apply(format_currency)
            display_policy_df['Premium_Pct_Face'] = display_policy_df['Premium_Pct_Face'].apply(lambda x: f"{x:.2f}%")
            display_policy_df['Remaining_LE'] = display_policy_df['Remaining_LE'].apply(lambda x: f"{x:.1f} months" if x > 0 else "N/A")
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
            top_policies = display_policy_df.nlargest(10, 'Face Value')
            st.dataframe(top_policies[['Policy ID', 'Name', 'Age', 'Gender', 'Face Value', 'Valuation', 'Annual Premium', 'Premium % Face']], use_container_width=True, hide_index=True)
    if not loan_summary['loans_df'].empty:
        st.markdown("<h2 style='color: #FDB813; margin-top: 3rem;'>💰 Loan Portfolio</h2>", unsafe_allow_html=True)
        col1, col2, col3, col4, col5 = st.columns(5)
        collection_rate = (loan_summary['total_repaid_principal'] + loan_summary['total_repaid_interest']) / loan_summary['total_original'] * 100 if loan_summary['total_original'] > 0 else 0
        with col1:
            st.metric("Total Principal", format_currency(loan_summary['total_original']))
        with col2:
            st.metric("Principal Repaid", format_currency(loan_summary['total_repaid_principal']))
        with col3:
            st.metric("Interest Earned", format_currency(loan_summary['total_repaid_interest']))
        with col4:
            st.metric("Collection Rate", f"{collection_rate:.1f}%")
        with col5:
            avg_loan_size = loan_summary['total_original'] / len(loan_summary['loans_df']) if len(loan_summary['loans_df']) > 0 else 0
            st.metric("Avg Loan Size", format_currency(avg_loan_size))
        active_loans = loan_summary['loans_df'][loan_summary['loans_df']['Status'] == 'Active'].copy()
        if not active_loans.empty:
            st.markdown("<h3 style='color: #FFFFFF; margin-top: 2rem;'>Active Loans</h3>", unsafe_allow_html=True)
            active_display = active_loans[['Borrower', 'Original Loan Balance', 'Current Loan Balance', 'Total Principal Repaid', 'Total Interest Repaid', 'Annual Interest Rate']].copy()
            for col in ['Original Loan Balance', 'Current Loan Balance', 'Total Principal Repaid', 'Total Interest Repaid']:
                active_display[col] = active_display[col].apply(format_currency)
            active_display['Annual Interest Rate'] = active_display['Annual Interest Rate'].apply(format_percent)
            st.dataframe(active_display, use_container_width=True, hide_index=True)
        closed_loans = loan_summary['loans_df'][loan_summary['loans_df']['Status'] == 'Closed']
        if not closed_loans.empty:
            with st.expander(f"View {len(closed_loans)} Closed Loans"):
                closed_display = closed_loans[['Borrower', 'Original Loan Balance', 'Total Principal Repaid', 'Total Interest Repaid']].copy()
                for col in ['Original Loan Balance', 'Total Principal Repaid', 'Total Interest Repaid']:
                    closed_display[col] = closed_display[col].apply(format_currency)
                st.dataframe(closed_display, use_container_width=True, hide_index=True)
else:
    st.info("📁 Upload loan and/or life settlement files to view portfolio data")

with st.sidebar:
    st.markdown("""
    <div style='text-align: center; padding: 1rem 0;'>
        <h2 style='color: #FDB813; margin: 0;'>⚡ Sirocco Partners</h2>
        <p style='color: #999999; margin: 0.5rem 0 0 0;'>Portfolio Dashboard</p>
    </div>
    """, unsafe_allow_html=True)
    if master_file:
        st.success("✅ Loan file loaded")
    if ls_file:
        st.success("✅ LS file loaded")
    if remittance_file:
        st.success("✅ Remittance file loaded")
    if not (master_file or ls_file):
        st.info("📤 Upload files to view data")
    if ls_data or loan_summary['total_original'] > 0:
        st.markdown("### 📊 Quick Stats")
        if ls_data:
            st.markdown(f"**LS Policies:** {ls_data['summary']['total_policies']}")
            st.markdown(f"**Face Value:** {format_currency(ls_data['summary']['total_ndb'])}")
            st.markdown(f"**Annual Premiums:** {format_currency(ls_data['summary']['total_annual_premiums'])}")
        if loan_summary['total_original'] > 0:
            total_loans = loan_summary['active_loans'] + loan_summary['closed_loans']
            st.markdown(f"**Total Loans:** {total_loans}")
            st.markdown(f"**Loan Principal:** {format_currency(loan_summary['total_original'])}")
            st.markdown(f"**Active:** {loan_summary['active_loans']} | **Closed:** {loan_summary['closed_loans']}")
    st.markdown("---")
    st.markdown("### 📋 File Requirements")
    st.markdown("**Loan File:** Excel with loan sheets (#1, #2, etc.)")
    st.markdown("**LS File:** Excel with 'Valuation Summary' and 'Premium Stream' sheets")

st.markdown("""
<div style='margin-top: 3rem; padding-top: 2rem; border-top: 1px solid #3d3d3d; text-align: center; color: #666666;'>
    <p>Sirocco Partners Portfolio Management System</p>
</div>
""", unsafe_allow_html=True)
