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

# ... rest of the user's provided code ... 