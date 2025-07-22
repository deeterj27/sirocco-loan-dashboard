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
    """Process Life Settlement Excel file and return summary data with debugging"""
    try:
        st.info("🔄 Loading LS workbook...")
        # Load the LS workbook
        ls_wb = load_workbook(ls_file, data_only=True)
        
        # Check if required sheets exist
        available_sheets = ls_wb.sheetnames
        st.info(f"Available sheets: {available_sheets}")
        
        if 'Valuation Summary' not in available_sheets:
            st.error("❌ 'Valuation Summary' sheet not found")
            return None
        if 'Premium Stream' not in available_sheets:
            st.error("❌ 'Premium Stream' sheet not found") 
            return None
        
        st.info("✅ Required sheets found, processing data...")
        
        # Process Valuation Summary sheet
        val_sheet = ls_wb['Valuation Summary']
        premium_sheet = ls_wb['Premium Stream']
        
        # Initialize data collection
        policies = []
        
        # Read policy data starting from row 3 (row 2 has headers)
        st.info("📊 Reading Valuation Summary data...")
        
        for row in range(3, 200):  # Start from row 3, reasonable limit
            policy_id_cell = val_sheet[f'B{row}']
            if not policy_id_cell.value:
                break
                
            # Extract policy data from specific columns
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
                
            except Exception as e:
                st.warning(f"Error processing row {row}: {str(e)}")
                continue
        
        st.success(f"✅ Processed {len(policies)} policies from Valuation Summary")
        
        if len(policies) == 0:
            st.error("❌ No policies found in Valuation Summary")
            return None
        
        # Calculate basic summary statistics
        total_policies = len(policies)
        total_ndb = sum(p['NDB'] for p in policies)
        total_valuation = sum(p['Valuation'] for p in policies)
        total_cost_basis = sum(p['Cost_Basis'] for p in policies)
        
        # Age statistics
        valid_ages = [p['Age'] for p in policies if p['Age'] > 0]
        avg_age = sum(valid_ages) / len(valid_ages) if valid_ages else 0
        
        # Gender statistics
        male_count = sum(1 for p in policies if 'male' in p['Gender'].lower() and 'female' not in p['Gender'].lower())
        female_count = sum(1 for p in policies if 'female' in p['Gender'].lower())
        male_percentage = (male_count / (male_count + female_count)) * 100 if (male_count + female_count) > 0 else 0
        
        # Life expectancy statistics
        valid_les = [p['Remaining_LE'] for p in policies if p['Remaining_LE'] > 0]
        avg_remaining_le = sum(valid_les) / len(valid_les) if valid_les else 0
        
        # Process monthly premiums from Premium Stream
        st.info("💰 Processing Premium Stream data...")
        monthly_premiums = {}
        policy_premiums = {}
        
        # Month columns M through X (12 months)
        month_columns = ['M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X']
        
        # Get month headers first
        month_headers = []
        for col in month_columns:
            header = premium_sheet[f'{col}2'].value
            if header:
                month_headers.append((col, str(header)))
        
        st.info(f"Found {len(month_headers)} monthly columns: {[h[1] for h in month_headers]}")
        
        # Process each month
        for col_letter, month_name in month_headers:
            month_total = 0
            
            # Process each policy row in premium sheet
            for prem_row in range(3, len(policies) + 3):
                try:
                    # Get Lyric ID 
                    lyric_id_cell = premium_sheet[f'B{prem_row}']
                    if lyric_id_cell.value:
                        lyric_id = str(lyric_id_cell.value)
                        
                        # Get premium for this month
                        premium_cell = premium_sheet[f'{col_letter}{prem_row}']
                        premium_val = safe_float(premium_cell.value) if premium_cell.value else 0
                        month_total += premium_val
                        
                        # Store individual policy premium
                        if lyric_id not in policy_premiums:
                            policy_premiums[lyric_id] = {}
                        policy_premiums[lyric_id][month_name] = premium_val
                        
                except Exception as e:
                    continue
            
            monthly_premiums[month_name] = month_total
        
        st.success(f"✅ Processed premium data for {len(policy_premiums)} policies")
        
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
        
        # Calculate final summary
        total_annual_premiums = sum(monthly_premiums.values())
        premiums_as_pct_face = (total_annual_premiums / total_ndb) * 100 if total_ndb > 0 else 0
        
        result = {
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
        
        st.success("✅ LS data processing completed successfully!")
        return result
        
    except Exception as e:
        st.error(f"❌ Error processing Life Settlement file: {str(e)}")
        st.error("Please check that your file has the correct structure")
        return None

# Main app

# Header with Sirocco branding
st.markdown("""
<div style='background-color: #1a1a1a; padding: 2rem 0; margin: -2rem -2rem 2rem -2rem; border-bottom: 4px solid #FDB813;'>
    <h1 style='text-align: center; color: #FFFFFF; font-size: 2.5rem; margin: 0;'>
        <span style='color: #FDB813;'>⚡</span> Sirocco I LP Portfolio Dashboard
    </h1>
    <p style='text-align: center; color: #999999; margin-top: 0.5rem;'>Loan Participation & Life Settlement Portfolio Management</p>
</div>
""", unsafe_allow_html=True)

# File upload section with status indicators
st.markdown("<h2 style='color: #FDB813;'>📁 File Upload</h2>", unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)
with col1:
    master_file = st.file_uploader("📊 Master Excel File (Loans)", type=["xlsx"], 
                                  help="Upload the master Excel file containing loan data with sheets starting with '#'")
    if master_file:
        st.success("✅ Loan file uploaded")

with col2:
    ls_file = st.file_uploader("🏥 LS Portfolio File", type=["xlsx"],
                              help="Upload the Life Settlement portfolio file with 'Valuation Summary' and 'Premium Stream' sheets")
    if ls_file:
        st.success("✅ LS file uploaded")
        st.info(f"Filename: {ls_file.name}")

with col3:
    remittance_file = st.file_uploader("📄 Monthly Remittance File", type=["csv", "xlsx"],
                                     help="Upload monthly remittance file for reconciliation")
    if remittance_file:
        st.success("✅ Remittance file uploaded")

# Process Life Settlement data if file is uploaded
ls_data = None
if ls_file:
    st.markdown("### 🔄 Processing Life Settlement File...")
    
    try:
        # Load workbook and check sheets
        wb_ls = load_workbook(ls_file, data_only=True)
        st.info(f"Sheets found: {wb_ls.sheetnames}")
        
        if 'Valuation Summary' not in wb_ls.sheetnames:
            st.error("❌ 'Valuation Summary' sheet not found!")
        if 'Premium Stream' not in wb_ls.sheetnames:
            st.error("❌ 'Premium Stream' sheet not found!")
        
        if 'Valuation Summary' in wb_ls.sheetnames and 'Premium Stream' in wb_ls.sheetnames:
            ls_data = process_life_settlement_data(ls_file)
            if ls_data:
                st.success(f"✅ Successfully processed {ls_data['summary']['total_policies']} LS policies")
                st.success(f"✅ Total Face Value: {format_currency(ls_data['summary']['total_ndb'])}")
            else:
                st.error("❌ Failed to process LS data - check function")
        
    except Exception as e:
        st.error(f"❌ Error loading LS file: {str(e)}")
        st.error("Please ensure you've uploaded the correct Excel file with LS data")

# Create tabs for different sections
tab1, tab2, tab3 = st.tabs(["📊 Portfolio Overview", "💰 Loan Portfolio", "🏥 Life Settlement Portfolio"])

# Initialize loan_summary at the module level so it can be updated by other tabs
loan_summary = {
    'total_original': 0,
    'total_repaid_principal': 0,
    'total_repaid_interest': 0,
    'active_loans': 0,
    'closed_loans': 0
}

# Tab 1: Portfolio Overview
with tab1:
    if master_file or ls_data:
        st.markdown("<h2 style='color: #FDB813;'>📈 Combined Portfolio Summary</h2>", unsafe_allow_html=True)
        
        # Quick processing of loan data for overview (simplified)
        if master_file and not loan_summary['total_original']:  # Only process if not already done
            try:
                wb = load_workbook(master_file, data_only=True)
                loan_sheets = [s for s in wb.sheetnames if s.startswith('#') and s != '#AddSheet']
                
                temp_total = 0
                temp_active = 0
                temp_closed = 0
                
                for sheet_name in loan_sheets:
                    sheet = wb[sheet_name]
                    loan_amount = safe_float(get_cell_value(sheet, ['B3', 'C3'], 0))
                    temp_total += loan_amount
                    
                    # Quick status check - if there's amortization data, determine status
                    current_balance = safe_float(sheet['G11'].value) if sheet['G11'].value else loan_amount
                    if current_balance > 0:
                        temp_active += 1
                    else:
                        temp_closed += 1
                
                loan_summary.update({
                    'total_original': temp_total,
                    'active_loans': temp_active,
                    'closed_loans': temp_closed
                })
                
            except Exception as e:
                st.warning(f"Error processing loan file: {str(e)}")
        
        # Display combined metrics
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
        
        # Portfolio allocation and performance metrics
        if ls_data and loan_summary['total_original'] > 0:
            st.markdown("<h3 style='color: #FFFFFF; margin-top: 2rem;'>Portfolio Allocation</h3>", unsafe_allow_html=True)
            
            # Calculate total portfolio value
            total_ls_value = ls_data['summary']['total_valuation']
            total_loan_value = loan_summary['total_original']  # Could use current balance if available
            total_portfolio = total_ls_value + total_loan_value
            
            allocation_data = {
                'Asset Class': ['Life Settlements', 'Loans'],
                'Current Value': [total_ls_value, total_loan_value],
                'Face/Principal': [ls_data['summary']['total_ndb'], loan_summary['total_original']],
                'Allocation %': [
                    (total_ls_value / total_portfolio * 100) if total_portfolio > 0 else 0,
                    (total_loan_value / total_portfolio * 100) if total_portfolio > 0 else 0
                ]
            }
            
            allocation_df = pd.DataFrame(allocation_data)
            allocation_df['Current Value Formatted'] = allocation_df['Current Value'].apply(format_currency)
            allocation_df['Face/Principal Formatted'] = allocation_df['Face/Principal'].apply(format_currency)
            allocation_df['Allocation % Formatted'] = allocation_df['Allocation %'].apply(lambda x: f"{x:.1f}%")
            
            display_df = allocation_df[['Asset Class', 'Current Value Formatted', 'Face/Principal Formatted', 'Allocation % Formatted']]
            display_df.columns = ['Asset Class', 'Current Value', 'Face/Principal Value', 'Allocation %']
            st.dataframe(display_df, use_container_width=True, hide_index=True)
            
            # Key performance metrics
            st.markdown("<h3 style='color: #FFFFFF; margin-top: 2rem;'>Performance Metrics</h3>", unsafe_allow_html=True)
            
            perf_col1, perf_col2, perf_col3 = st.columns(3)
            
            with perf_col1:
                # LS metrics
                ls_cost_to_face = (ls_data['summary']['total_cost_basis'] / ls_data['summary']['total_ndb']) * 100 if ls_data['summary']['total_ndb'] > 0 else 0
                st.metric("LS Cost to Face Ratio", f"{ls_cost_to_face:.1f}%")
                
                ls_current_yield = (ls_data['summary']['total_annual_premiums'] / ls_data['summary']['total_valuation']) * 100 if ls_data['summary']['total_valuation'] > 0 else 0
                st.metric("LS Current Yield", f"{ls_current_yield:.2f}%")
            
            with perf_col2:
                # Combined metrics
                total_annual_income = ls_data['summary']['total_annual_premiums'] + loan_summary['total_repaid_interest']
                portfolio_yield = (total_annual_income / total_portfolio) * 100 if total_portfolio > 0 else 0
                st.metric("Portfolio Current Yield", f"{portfolio_yield:.2f}%")
                
                diversification_ratio = min(total_ls_value, total_loan_value) / max(total_ls_value, total_loan_value) if max(total_ls_value, total_loan_value) > 0 else 0
                st.metric("Diversification Ratio", f"{diversification_ratio:.2f}")
            
            with perf_col3:
                # Risk metrics
                avg_loan_size = loan_summary['total_original'] / loan_summary['active_loans'] if loan_summary['active_loans'] > 0 else 0
                avg_policy_size = ls_data['summary']['total_ndb'] / ls_data['summary']['total_policies'] if ls_data['summary']['total_policies'] > 0 else 0
                
                st.metric("Avg Loan Size", format_currency(avg_loan_size))
                st.metric("Avg Policy Size", format_currency(avg_policy_size))
        
        elif ls_data:
            # Show LS-only metrics
            st.markdown("<h3 style='color: #FFFFFF; margin-top: 2rem;'>Life Settlement Performance</h3>", unsafe_allow_html=True)
            
            perf_col1, perf_col2 = st.columns(2)
            
            with perf_col1:
                ls_cost_to_face = (ls_data['summary']['total_cost_basis'] / ls_data['summary']['total_ndb']) * 100 if ls_data['summary']['total_ndb'] > 0 else 0
                st.metric("Cost to Face Ratio", f"{ls_cost_to_face:.1f}%")
                
                premium_load = (ls_data['summary']['total_annual_premiums'] / ls_data['summary']['total_ndb']) * 100 if ls_data['summary']['total_ndb'] > 0 else 0
                st.metric("Premium Load (% of Face)", f"{premium_load:.2f}%")
            
            with perf_col2:
                value_to_cost = (ls_data['summary']['total_valuation'] / ls_data['summary']['total_cost_basis']) * 100 if ls_data['summary']['total_cost_basis'] > 0 else 0
                st.metric("Value to Cost Ratio", f"{value_to_cost:.1f}%")
                
                avg_policy_size = ls_data['summary']['total_ndb'] / ls_data['summary']['total_policies'] if ls_data['summary']['total_policies'] > 0 else 0
                st.metric("Avg Policy Size", format_currency(avg_policy_size))
        
        elif loan_summary['total_original'] > 0:
            # Show loan-only metrics
            st.markdown("<h3 style='color: #FFFFFF; margin-top: 2rem;'>Loan Portfolio Performance</h3>", unsafe_allow_html=True)
            
            perf_col1, perf_col2 = st.columns(2)
            
            with perf_col1:
                collection_rate = ((loan_summary['total_repaid_principal'] + loan_summary['total_repaid_interest']) / loan_summary['total_original']) * 100 if loan_summary['total_original'] > 0 else 0
                st.metric("Collection Rate", f"{collection_rate:.1f}%")
            
            with perf_col2:
                avg_loan_size = loan_summary['total_original'] / (loan_summary['active_loans'] + loan_summary['closed_loans']) if (loan_summary['active_loans'] + loan_summary['closed_loans']) > 0 else 0
                st.metric("Average Loan Size", format_currency(avg_loan_size))
    else:
        st.info("Upload loan and/or life settlement files to see portfolio overview")

# Tab 2: Loan Portfolio
with tab2:
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
            
            st.markdown("<h2 style='color: #FDB813;'>💰 Loan Portfolio Analysis</h2>", unsafe_allow_html=True)
            
            # Process each loan sheet
            loans = []
            loan_details = {}
            
            for sheet_name in loan_sheets:
                sheet = wb[sheet_name]
                
                # Extract loan header information
                borrower = get_cell_value(sheet, ['B2', 'A2'], f"Unknown ({sheet_name})")
                if borrower == "" or borrower is None:
                    borrower = f"Unknown ({sheet_name})"
                
                # Try multiple locations for loan data
                b3_value = sheet['B3'].value
                if isinstance(b3_value, str) and 'loan' in str(b3_value).lower():
                    # Data is in column C
                    loan_amount = safe_float(sheet['C3'].value)
                    interest_rate = safe_float(sheet['C4'].value)
                    loan_period = safe_float(sheet['C5'].value)
                    payment_amount_val = sheet['C6'].value
                    loan_start = excel_date_to_datetime(sheet['C7'].value)
                    if pd.isna(loan_start):
                        loan_start = excel_date_to_datetime(sheet['C6'].value)
                else:
                    # Data is in column B
                    loan_amount = safe_float(sheet['B3'].value)
                    interest_rate = safe_float(sheet['B4'].value)
                    loan_period = safe_float(sheet['B5'].value)
                    payment_amount_val = sheet['B6'].value
                    loan_start = excel_date_to_datetime(sheet['B7'].value)
                
                # Handle payment amount
                if isinstance(payment_amount_val, str) and payment_amount_val.lower() == 'interest only':
                    payment_amount = loan_amount * (interest_rate / 12)
                else:
                    payment_amount = safe_float(payment_amount_val)
                
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
                }
                
                # Read amortization schedule (simplified for space)
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
                    
                    # Calculate current balance and totals
                    if pd.notna(as_of_date):
                        amort_df['Month'] = pd.to_datetime(amort_df['Month'])
                        current_rows = amort_df[amort_df['Month'] <= as_of_date]
                        if not current_rows.empty:
                            current_row = current_rows.iloc[-1]
                            loan_info['Current Loan Balance'] = current_row['Closing Balance']
                            loan_info['Total Principal Repaid'] = current_rows['Capital Repaid'].sum()
                            loan_info['Total Interest Repaid'] = current_rows['Interest Charged'].sum()
                            loan_info['Last Payment Amount'] = current_row['Loan Repayment']
                        else:
                            loan_info['Current Loan Balance'] = loan_info['Original Loan Balance']
                            loan_info['Total Principal Repaid'] = 0
                            loan_info['Total Interest Repaid'] = 0
                    
                    # Determine status
                    if loan_info.get('Current Loan Balance', 0) == 0:
                        loan_info['Status'] = 'Closed'
                    else:
                        loan_info['Status'] = 'Active'
                    
                    loan_details[borrower] = amort_df
                
                if loan_info['Original Loan Balance'] > 0:
                    loans.append(loan_info)
            
            # Create loans dataframe
            loans_df = pd.DataFrame(loans)
            
            # Separate by status
            active_loans = loans_df[loans_df['Status'] == 'Active'].copy() if 'Status' in loans_df.columns else loans_df.copy()
            closed_loans = loans_df[loans_df['Status'] == 'Closed'].copy() if 'Status' in loans_df.columns else pd.DataFrame()
            
            # Display loan metrics
            col1, col2, col3, col4 = st.columns(4)
            
            total_original = loans_df['Original Loan Balance'].sum()
            total_repaid_principal = loans_df.get('Total Principal Repaid', pd.Series([0])).sum()
            total_repaid_interest = loans_df.get('Total Interest Repaid', pd.Series([0])).sum()
            
            with col1:
                st.metric("Total Principal Repaid", format_currency(total_repaid_principal))
                st.metric("Active Loans", len(active_loans))
            with col2:
                st.metric("Total Interest Earned", format_currency(total_repaid_interest))
                st.metric("Closed Loans", len(closed_loans))
            with col3:
                st.metric("Total Collections", format_currency(total_repaid_principal + total_repaid_interest))
                st.metric("Total Loans", len(loans_df))
            with col4:
                collection_rate = (total_repaid_principal + total_repaid_interest) / total_original if total_original > 0 else 0
                st.metric("Collection Rate", format_percent(collection_rate))
                avg_loan_size = total_original / len(loans_df) if len(loans_df) > 0 else 0
                st.metric("Avg Loan Size", format_currency(avg_loan_size))
            
            # Display active loans table
            if not active_loans.empty:
                st.markdown("<h3 style='color: #FFFFFF; margin-top: 2rem;'>Active Loans</h3>", unsafe_allow_html=True)
                
                display_columns = ['Borrower', 'Original Loan Balance', 'Current Loan Balance', 
                                 'Total Principal Repaid', 'Total Interest Repaid', 'Annual Interest Rate']
                
                active_display = active_loans[display_columns].copy()
                
                # Format for display
                for col in ['Original Loan Balance', 'Current Loan Balance', 'Total Principal Repaid', 'Total Interest Repaid']:
                    if col in active_display.columns:
                        active_display[col] = active_display[col].apply(format_currency)
                
                if 'Annual Interest Rate' in active_display.columns:
                    active_display['Annual Interest Rate'] = active_display['Annual Interest Rate'].apply(format_percent)
                
                st.dataframe(active_display, use_container_width=True, hide_index=True)
            
            # Update loan_summary for tab 1
            loan_summary.update({
                'total_original': total_original,
                'total_repaid_principal': total_repaid_principal,
                'total_repaid_interest': total_repaid_interest,
                'active_loans': len(active_loans),
                'closed_loans': len(closed_loans)
            })
            
        except Exception as e:
            st.error(f"Error processing loan file: {str(e)}")
    else:
        st.info("Upload master Excel file to view loan portfolio")

# Tab 3: Life Settlement Portfolio
with tab3:
    if ls_data:
        st.markdown("<h2 style='color: #FDB813;'>🏥 Life Settlement Portfolio Analysis</h2>", unsafe_allow_html=True)
        # Key Metrics Section
        st.markdown("<h3 style='color: #FFFFFF;'>Portfolio Summary</h3>", unsafe_allow_html=True)
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Policies", ls_data['summary']['total_policies'])
            st.metric("Average Age", f"{ls_data['summary']['avg_age']:.1f} years")
        
        with col2:
            st.metric("Total NDB (Face Value)", format_currency(ls_data['summary']['total_ndb']))
            st.metric("Male Policies", ls_data['summary']['male_count'])
        
        with col3:
            st.metric("Total Valuation", format_currency(ls_data['summary']['total_valuation']))
            st.metric("Female Policies", ls_data['summary']['female_count'])
        
        with col4:
            st.metric("Cost Basis", format_currency(ls_data['summary']['total_cost_basis']))
            st.metric("% Male", f"{ls_data['summary']['male_percentage']:.1f}%")
        
        # Additional Key Metrics Row
        col5, col6, col7, col8 = st.columns(4)
        
        with col5:
            st.metric("Avg Remaining LE", f"{ls_data['summary']['avg_remaining_le']:.1f} months")
        
        with col6:
            st.metric("Annual Premiums", format_currency(ls_data['summary']['total_annual_premiums']))
        
        with col7:
            st.metric("Premiums % of Face", f"{ls_data['summary']['premiums_as_pct_face']:.2f}%")
        
        with col8:
            cost_to_face = (ls_data['summary']['total_cost_basis'] / ls_data['summary']['total_ndb']) * 100 if ls_data['summary']['total_ndb'] > 0 else 0
            st.metric("Cost to Face Ratio", f"{cost_to_face:.1f}%")
        
        # Monthly Premium Projections
        st.markdown("<h3 style='color: #FFFFFF; margin-top: 2rem;'>Monthly Premium Projections</h3>", unsafe_allow_html=True)
        
        if ls_data['monthly_premiums']:
            # Display monthly premiums in a responsive grid
            premium_items = list(ls_data['monthly_premiums'].items())
            
            # Show 4 months per row for better readability
            months_per_row = 4
            for i in range(0, len(premium_items), months_per_row):
                cols = st.columns(months_per_row)
                for j in range(months_per_row):
                    if i + j < len(premium_items):
                        month, amount = premium_items[i + j]
                        with cols[j]:
                            st.metric(month, format_currency(amount))
            
            # Premium summary metrics
            st.markdown("<h4 style='color: #FFFFFF; margin-top: 2rem;'>Premium Summary</h4>", unsafe_allow_html=True)
            
            sum_col1, sum_col2, sum_col3 = st.columns(3)
            
            with sum_col1:
                avg_monthly = ls_data['summary']['total_annual_premiums'] / 12
                st.metric("Average Monthly Premium", format_currency(avg_monthly))
            
            with sum_col2:
                st.metric("Total Annual Premiums", format_currency(ls_data['summary']['total_annual_premiums']))
            
            with sum_col3:
                premium_yield = (ls_data['summary']['total_annual_premiums'] / ls_data['summary']['total_valuation']) * 100 if ls_data['summary']['total_valuation'] > 0 else 0
                st.metric("Premium Yield on Value", f"{premium_yield:.2f}%")
        # ... (rest of policy details, analytics, sidebar, and footer) ...

# Sidebar with portfolio summary
with st.sidebar:
    st.markdown("""
    <div style='text-align: center; padding: 1rem 0;'>
        <h2 style='color: #FDB813; margin: 0;'>⚡ Sirocco Partners</h2>
        <p style='color: #999999; margin: 0.5rem 0 0 0;'>Portfolio Analytics</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Show file upload status
    if master_file:
        st.success("✅ Loan file uploaded")
    if ls_file:
        st.success("✅ LS file uploaded")
    if remittance_file:
        st.success("✅ Remittance file uploaded")
    
    if not (master_file or ls_file):
        st.warning("📤 Upload files to begin analysis")
    
    # Quick stats in sidebar
    if ls_data or loan_summary['total_original'] > 0:
        st.markdown("""
        <div style='background-color: #2d2d2d; padding: 1rem; border-radius: 8px; margin-top: 1rem;'>
            <p style='color: #FDB813; margin: 0; font-weight: 600;'>📊 Quick Stats</p>
        </div>
        """, unsafe_allow_html=True)
        
        if ls_data:
            st.markdown(f"""
            <div style='background-color: #242424; padding: 0.5rem; margin: 0.5rem 0; border-radius: 4px;'>
                <p style='color: #FFFFFF; margin: 0; font-size: 0.9rem;'>LS Policies: {ls_data['summary']['total_policies']}</p>
                <p style='color: #FFFFFF; margin: 0; font-size: 0.9rem;'>Face Value: {format_currency(ls_data['summary']['total_ndb'])}</p>
            </div>
            """, unsafe_allow_html=True)
        
        if loan_summary['total_original'] > 0:
            st.markdown(f"""
            <div style='background-color: #242424; padding: 0.5rem; margin: 0.5rem 0; border-radius: 4px;'>
                <p style='color: #FFFFFF; margin: 0; font-size: 0.9rem;'>Total Loans: {loan_summary['active_loans'] + loan_summary['closed_loans']}</p>
                <p style='color: #FFFFFF; margin: 0; font-size: 0.9rem;'>Principal: {format_currency(loan_summary['total_original'])}</p>
            </div>
            """, unsafe_allow_html=True)

# Footer
st.markdown("""
<div style='margin-top: 3rem; padding-top: 2rem; border-top: 1px solid #3d3d3d; text-align: center; color: #666666;'>
    <p>Sirocco Partners Portfolio Management System</p>
</div>
""", unsafe_allow_html=True)
