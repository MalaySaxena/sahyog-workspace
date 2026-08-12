import streamlit as st
import pandas as pd
import datetime
import os
import openpyxl
import plotly.express as px
import json
import re
import numpy as np

# -----------------------------------------------------------------------------
# 1. APP CONFIGURATION & STYLING
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Sahyog Financials Console", layout="wide", page_icon="🏦")

st.markdown("""
    <style>
    .section-header {
        background-color: #1F4E79;
        color: white;
        padding: 10px;
        border-radius: 5px;
        font-weight: bold;
        margin-bottom: 15px;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #1F4E79;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
        margin-bottom: 15px;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #f1f3f1;
        border-radius: 4px 4px 0px 0px;
        padding: 8px 16px;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1F4E79 !important;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# 1.5 HELPER FUNCTIONS & AUTHENTICATION ENGINE
# -----------------------------------------------------------------------------
def clean_int_string(val):
    """Converts numbers/floats to clean whole integer strings (e.g. 12345.0 -> '12345')."""
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if val_str.lower() in ["nan", "none", ""]:
        return ""
    if val_str.endswith(".0"):
        return val_str[:-2]
    try:
        if "." in val_str:
            f_val = float(val_str)
            if f_val.is_integer():
                return str(int(f_val))
    except ValueError:
        pass
    return val_str


def check_credentials(username, password):
    clean_username = username.strip().lower()
    try:
        if "users" in st.secrets:
            users = st.secrets["users"]
            if clean_username in users and users[clean_username] == password:
                return True
    except Exception:
        pass

    default_users = {
        "admin": "sahyog2026",
        "sahyog": "wealth123",
        "malay": "malay@sahyog",
        "shishir": "shishir@sahyog",
        "neelesh": "neelesh@sahyog",
        "krati": "krati@sahyog",
        "ayush": "ayush@sahyog",
        "aayushi": "aayushi@sahyog",
        "amit": "amit@sahyog",
        "ruchi": "ruchi@sahyog"
    }
    return clean_username in default_users and default_users[clean_username] == password


def render_login_screen():
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("""
            <div style="background-color: #f8f9fa; padding: 30px; border-radius: 12px; border-top: 6px solid #1F4E79; box-shadow: 0 4px 12px rgba(0,0,0,0.1); margin-top: 40px;">
                <h2 style="color: #1F4E79; text-align: center; margin-bottom: 5px;">🏦 Sahyog Financials</h2>
                <p style="text-align: center; color: #6c757d; font-size: 14px;">Enterprise Workspace Console Login (v10.3)</p>
                <hr style="margin-bottom: 20px;">
            </div>
        """, unsafe_allow_html=True)
        
        with st.form("login_form"):
            username_input = st.text_input("👤 Username")
            password_input = st.text_input("🔑 Password", type="password")
            submit_login = st.form_submit_button("Sign In to Workspace", use_container_width=True)
            
            if submit_login:
                if check_credentials(username_input, password_input.strip()):
                    st.session_state["authenticated"] = True
                    st.session_state["username"] = username_input.strip()
                    st.success("✅ Login successful! Loading workspace...")
                    st.rerun()
                else:
                    st.error("❌ Invalid Username or Password. Please try again.")

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    render_login_screen()
    st.stop()


# -----------------------------------------------------------------------------
# 2. AUTO-PATH RESOLUTION & DATABASE ENGINE
# -----------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else '.'
EXCEL_FILE = os.path.join(SCRIPT_DIR, "Live_Wealth_Manager_Book.xlsx")
SYNC_LOG_FILE = os.path.join(SCRIPT_DIR, "sync_log.json")

# Master Schemas matching Live_Wealth_Manager_Book.xlsx exactly
SCHEMA_COLUMNS = [
    "Date", "Name", "Amount", "Ucc / APP./FOLIO NO.", "SCHEME", "TYPE", 
    "Top Up", "FROM", "TO", "PAYMENT DETAIL", "BANK DETAIL", "REMARK", 
    "Code", "STNAME", "EXIT LOAD", "Other Remarks", "Entry checked by",
    "Entry created by", "Last edited by", "Last edit timestamp"
]

CLIENT_MASTER_SCHEMA = [
    "Code", "Partner/Employee", "Group Name", "Investor Name", "UCC", "PAN",
    "Email Id", "Mobile No", "Tax Status", "Invested Equity", "Invested Debt",
    "Invested Hybrid", "Invested Total", "Current Equity AUM", "Current Debt AUM",
    "Current Hybrid AUM", "Current Total AUM", "Total XIRR (%)"
]

SCHEME_MASTER_SCHEMA = [
    "Code", "AMC Name", "Mutual Fund Scheme Name", "NAV (₹)",
    "Invested Total (₹)", "Current AUM Total (₹)", "Equity AUM (₹)",
    "Debt AUM (₹)", "Hybrid AUM (₹)", "Brokerage Weight"
]

def load_scheme_weight_map():
    weight_files = ["weight_3.xlsx", "weight_2.xlsx", "weight.xlsx"]
    weight_map = {}
    check_dirs = [
        os.path.dirname(__file__) if '__file__' in locals() else '.',
        os.path.join(os.path.dirname(__file__) if '__file__' in locals() else '.', "temp_downloads")
    ]
    for wf_name in weight_files:
        for cdir in check_dirs:
            fpath = os.path.join(cdir, wf_name)
            if os.path.exists(fpath):
                try:
                    df_w = pd.read_excel(fpath)
                    df_w_clean = df_w.dropna(how='all', axis=1)
                    alias_s, alias_w = None, None
                    for col in df_w_clean.columns:
                        c_clean = str(col).strip().lower()
                        if 'scheme' in c_clean or 'product' in c_clean:
                            alias_s = col
                        elif 'weight' in c_clean or 'brokerage' in c_clean:
                            alias_w = col
                    if alias_s and alias_w:
                        for _, row in df_w_clean.dropna(subset=[alias_s]).iterrows():
                            s_name = str(row[alias_s]).strip().lower()
                            w_val = row[alias_w]
                            if pd.isna(w_val) or w_val is None:
                                w_val = 1.0
                            else:
                                try:
                                    w_val = float(w_val)
                                except Exception:
                                    w_val = 1.0
                            weight_map[s_name] = w_val
                        return weight_map
                except Exception:
                    pass
    return weight_map

def lookup_brokerage_weight(scheme_name, weight_map):
    if not scheme_name or str(scheme_name).strip() == "Total Book":
        return None
    s_clean = str(scheme_name).strip().lower()
    if s_clean in weight_map:
        return weight_map[s_clean]
    for k, v in weight_map.items():
        if k == s_clean or k in s_clean or s_clean in k:
            return v
    return 1.0

# Authorized Users List for Book Summary & Analytics Desk
ALLOWED_ANALYTICS_USERS = ["admin", "malay", "shishir", "sahyog", "ruchi", "neelesh"]

def is_gsheets_configured():
    try:
        return ("gcp_service_account" in st.secrets) and ("gsheets" in st.secrets and "spreadsheet_url" in st.secrets["gsheets"])
    except Exception:
        return False

def get_gspread_worksheet(worksheet_name="4. Transaction Ledger"):
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
        client = gspread.authorize(creds)
        spreadsheet_url = st.secrets["gsheets"]["spreadsheet_url"]
        sheet = client.open_by_url(spreadsheet_url)
        
        try:
            return sheet.worksheet(worksheet_name)
        except gspread.exceptions.WorksheetNotFound:
            ws = sheet.add_worksheet(title=worksheet_name, rows=1000, cols=30)
            ws.append_row(SCHEMA_COLUMNS)
            return ws
    except Exception as e:
        st.error(f"⚠️ Google Sheets Connection Error: {e}")
        return None

def init_excel_ledger():
    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        if "4. Transaction Ledger" not in wb.sheetnames:
            ws = wb.create_sheet("4. Transaction Ledger")
            ws.append(SCHEMA_COLUMNS)
            wb.save(EXCEL_FILE)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "4. Transaction Ledger"
        ws.append(SCHEMA_COLUMNS)
        wb.save(EXCEL_FILE)

init_excel_ledger()

@st.cache_data(ttl=5)
def load_all_data():
    try:
        df_c = pd.read_excel(EXCEL_FILE, sheet_name="2. Client AUM Master")
        df_s = pd.read_excel(EXCEL_FILE, sheet_name="3. Scheme Allocation Master")
        df_c = df_c[df_c['Code'].notna() & (df_c['Group Name'] != 'Total Book')]
        df_s = df_s[df_s['Code'].notna() & (df_s['Mutual Fund Scheme Name'] != 'Total Book')]
        
        if 'UCC' in df_c.columns:
            df_c['UCC'] = df_c['UCC'].apply(clean_int_string)
    except Exception:
        df_c = pd.DataFrame(columns=CLIENT_MASTER_SCHEMA)
        df_s = pd.DataFrame(columns=SCHEME_MASTER_SCHEMA)

    df_l = pd.DataFrame(columns=SCHEMA_COLUMNS)
    if is_gsheets_configured():
        ws = get_gspread_worksheet("4. Transaction Ledger")
        if ws:
            try:
                records = ws.get_all_records()
                df_l = pd.DataFrame(records) if records else pd.DataFrame(columns=SCHEMA_COLUMNS)
            except Exception:
                if os.path.exists(EXCEL_FILE):
                    df_l = pd.read_excel(EXCEL_FILE, sheet_name="4. Transaction Ledger")
    else:
        if os.path.exists(EXCEL_FILE):
            try:
                df_l = pd.read_excel(EXCEL_FILE, sheet_name="4. Transaction Ledger")
            except Exception:
                df_l = pd.DataFrame(columns=SCHEMA_COLUMNS)
                
    for col in SCHEMA_COLUMNS:
        if col not in df_l.columns:
            df_l[col] = ""
        elif col == "Ucc / APP./FOLIO NO.":
            df_l[col] = df_l[col].apply(clean_int_string)
        elif col not in ["Amount", "Date"]:
            df_l[col] = df_l[col].astype(str).replace(["nan", "NaN", "None"], "")

    return df_c, df_s, df_l

df_master_clients, df_master_schemes, df_ledger = load_all_data()

client_options = sorted(df_master_clients['Investor Name'].dropna().unique().tolist()) if not df_master_clients.empty else []
scheme_options = sorted(df_master_schemes['Mutual Fund Scheme Name'].dropna().unique().tolist()) if not df_master_schemes.empty else []

def save_transaction_record(record):
    row_values = [record.get(col, "") for col in SCHEMA_COLUMNS]
    saved_to_gsheets = False
    
    if is_gsheets_configured():
        ws = get_gspread_worksheet("4. Transaction Ledger")
        if ws:
            try:
                ws.append_row(row_values)
                saved_to_gsheets = True
            except Exception as e:
                st.error(f"❌ Failed to write to Google Sheets: {e}")
                
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        if "4. Transaction Ledger" not in wb.sheetnames:
            ws_ex = wb.create_sheet("4. Transaction Ledger")
            ws_ex.append(SCHEMA_COLUMNS)
        else:
            ws_ex = wb["4. Transaction Ledger"]
        ws_ex.append(row_values)
        wb.save(EXCEL_FILE)
    except Exception:
        pass
        
    st.cache_data.clear()
    return saved_to_gsheets

def overwrite_ledger_database(df_updated):
    saved_to_gsheets = False
    df_updated = df_updated[SCHEMA_COLUMNS]
    
    if is_gsheets_configured():
        ws = get_gspread_worksheet("4. Transaction Ledger")
        if ws:
            try:
                ws.clear()
                ws.update([SCHEMA_COLUMNS] + df_updated.fillna("").values.tolist())
                saved_to_gsheets = True
            except Exception as e:
                st.error(f"❌ Failed to update Google Sheets ledger: {e}")
                
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        if "4. Transaction Ledger" in wb.sheetnames:
            del wb["4. Transaction Ledger"]
        ws_ex = wb.create_sheet("4. Transaction Ledger")
        ws_ex.append(SCHEMA_COLUMNS)
        for row in df_updated.fillna("").values.tolist():
            ws_ex.append(row)
        wb.save(EXCEL_FILE)
    except Exception as e:
        st.error(f"❌ Excel Update Error: {e}")
        
    st.cache_data.clear()
    return saved_to_gsheets


# -----------------------------------------------------------------------------
# 2.5 MULTI-CODE CONCATENATION & AGGREGATION LOGIC (WITH OTHERS SHEETS MERGE)
# -----------------------------------------------------------------------------
def parse_client_dataframe(df_raw, code_val):
    """
    Parses a single Client DataFrame and returns a list of dictionaries.
    """
    rows = []
    if df_raw is None or df_raw.empty:
        return rows

    c_val = float(code_val) if str(code_val).replace('.','').isdigit() else str(code_val).strip()
    df_c = df_raw.copy()

    # Exclude Total Book if present
    if 'Group' in df_c.columns:
        df_c = df_c[df_c['Group'] != 'Total Book']
    if 'Group Name' in df_c.columns:
        df_c = df_c[df_c['Group Name'] != 'Total Book']

    # Alias mapping for client master fields
    alias_map = {
        'Partner/Employee': ['partner/employee', 'partner name', 'rm name', 'employee', 'sub broker', 'partner'],
        'Group Name': ['group name', 'group', 'family group', 'family'],
        'Investor Name': ['investor name', 'investor', 'client name', 'applicant name', 'name', 'customer name'],
        'UCC': ['ucc', 'client code', 'ucc code', 'client ucc'],
        'PAN': ['pan', 'pan no', 'pan number', 'investor pan'],
        'Email Id': ['email id', 'email', 'investor e-mail id.', 'email id.', 'email address'],
        'Mobile No': ['mobile no', 'mobile', 'investor mobile no.', 'mobile number', 'phone', 'contact no'],
        'Tax Status': ['tax status', 'tax slab', 'status', 'holding nature'],
        'Invested Equity': ['invested equity', 'equity investment'],
        'Invested Debt': ['invested debt', 'debt investment'],
        'Invested Hybrid': ['invested hybrid', 'hybrid investment'],
        'Invested Total': ['invested total', 'total mf investment (₹)', 'total mf investment', 'total investment', 'invested amount'],
        'Current Equity AUM': ['current equity aum', 'equity aum (₹)', 'equity aum', 'current equity'],
        'Current Debt AUM': ['current debt aum', 'debt aum (₹)', 'debt aum', 'current debt'],
        'Current Hybrid AUM': ['current hybrid aum', 'cash aum (₹)', 'gold  aum (₹)', 'gold aum (₹)', 'hybrid aum', 'current hybrid'],
        'Current Total AUM': ['current total aum', 'total mf aum (₹)', 'total mf aum', 'total aum', 'current aum', 'total valuation'],
        'Total XIRR (%)': ['total xirr (%)', 'investor xirr - mf total', 'xirr (%)', 'xirr', 'total xirr']
    }

    col_rename = {}
    for col in df_c.columns:
        col_clean = str(col).strip().lower()
        for target_k, aliases in alias_map.items():
            if col_clean in aliases:
                col_rename[col] = target_k
                break

    df_c = df_c.rename(columns=col_rename)

    # Filter active accounts
    aum_series = df_c['Current Total AUM'] if 'Current Total AUM' in df_c.columns else pd.Series(0, index=df_c.index)
    inv_series = df_c['Invested Total'] if 'Invested Total' in df_c.columns else pd.Series(0, index=df_c.index)
    pms_series = df_c['PMS AUM (₹)'] if 'PMS AUM (₹)' in df_c.columns else pd.Series(0, index=df_c.index)

    aum_s = pd.to_numeric(aum_series, errors='coerce').fillna(0)
    inv_s = pd.to_numeric(inv_series, errors='coerce').fillna(0)
    pms_s = pd.to_numeric(pms_series, errors='coerce').fillna(0)

    df_act = df_c[(aum_s > 0) | (inv_s > 0) | (pms_s > 0)].copy()
    if df_act.empty:
        df_act = df_c.copy()

    for idx, row in df_act.iterrows():
        def clean_str(col_name):
            v = row.get(col_name, "")
            if isinstance(v, pd.Series): v = v.iloc[0]
            return "" if pd.isna(v) or str(v).strip() in ['-', 'nan', 'None'] else str(v).strip()

        def clean_num(col_name):
            v = row.get(col_name, 0)
            if isinstance(v, pd.Series): v = v.iloc[0]
            if pd.isna(v): return 0.0
            try: return float(str(v).replace(',', '').strip())
            except Exception: return 0.0

        investor_val = clean_str('Investor Name')
        if not investor_val or investor_val.lower() in ['nan', 'none', '', 'total book']:
            continue

        ucc_val = clean_str('UCC')
        if ucc_val.endswith('.0'): ucc_val = ucc_val[:-2]

        email_val = clean_str('Email Id')
        mob_val = row.get('Mobile No', '')
        if isinstance(mob_val, pd.Series): mob_val = mob_val.iloc[0]
        if pd.notna(mob_val):
            try:
                mob_float = float(mob_val)
                mob_val = mob_float if mob_float > 0 else ""
            except Exception:
                mob_val = str(mob_val)
        else:
            mob_val = ""

        pan_val = clean_str('PAN').upper()
        partner_val = clean_str('Partner/Employee')
        group_val = clean_str('Group Name') or investor_val
        tax_val = clean_str('Tax Status') or "individual"

        tot_inv = clean_num('Invested Total')
        eq_aum = clean_num('Current Equity AUM')
        debt_aum = clean_num('Current Debt AUM')
        hyb_aum = clean_num('Current Hybrid AUM')
        tot_aum = clean_num('Current Total AUM')
        xirr_val = clean_num('Total XIRR (%)')

        if tot_aum > 0:
            inv_eq = round(tot_inv * (eq_aum / tot_aum))
            inv_debt = round(tot_inv * (debt_aum / tot_aum))
            inv_hyb = round(tot_inv - inv_eq - inv_debt)
        else:
            inv_eq = tot_inv
            inv_debt = 0.0
            inv_hyb = 0.0

        rows.append({
            'Code': c_val,
            'Partner/Employee': partner_val,
            'Group Name': group_val,
            'Investor Name': investor_val,
            'UCC': ucc_val,
            'PAN': pan_val,
            'Email Id': email_val,
            'Mobile No': mob_val,
            'Tax Status': tax_val,
            'Invested Equity': inv_eq,
            'Invested Debt': inv_debt,
            'Invested Hybrid': inv_hyb,
            'Invested Total': tot_inv,
            'Current Equity AUM': eq_aum,
            'Current Debt AUM': debt_aum,
            'Current Hybrid AUM': hyb_aum,
            'Current Total AUM': tot_aum,
            'Total XIRR (%)': xirr_val
        })

    return rows


def build_sheet2_from_client_files(client_files_input):
    """
    Safely handles dict or list input, parses Client report DataFrames and extra sheets,
    concatenates code by code, and appends a single 'Total Book' summary row.
    """
    all_rows = []
    default_codes = ["31241", "19221", "43854", "72021"]
    processed_file_sheets = set()

    if isinstance(client_files_input, list):
        items = []
        for idx, item in enumerate(client_files_input):
            c_code = None
            if isinstance(item, pd.DataFrame) and 'Code' in item.columns:
                codes = item['Code'].dropna().astype(str).unique()
                if len(codes) > 0:
                    c_code = clean_int_string(codes[0])
            if not c_code or c_code.startswith("code_"):
                c_code = default_codes[idx] if idx < len(default_codes) else f"code_{idx+1}"
            items.append((c_code, item))
    elif isinstance(client_files_input, dict):
        items = list(client_files_input.items())
    else:
        items = []

    for code_str, raw_input in items:
        if isinstance(raw_input, str) and os.path.exists(raw_input):
            try:
                xls = pd.ExcelFile(raw_input)
                for sname in xls.sheet_names:
                    s_lower = sname.strip().lower()

                    if "31241" in s_lower:
                        target_code = "31241"
                    elif "43854" in s_lower:
                        target_code = "43854"
                    elif "19221" in s_lower:
                        target_code = "19221"
                    elif "72021" in s_lower:
                        target_code = "72021"
                    else:
                        target_code = code_str

                    df_sheet = pd.read_excel(raw_input, sheet_name=sname)
                    parsed_rows = parse_client_dataframe(df_sheet, target_code)
                    all_rows.extend(parsed_rows)
                    processed_file_sheets.add(f"{os.path.basename(raw_input)}::{sname}")
            except Exception as e:
                print(f"Error reading client file {raw_input}: {e}")
        elif isinstance(raw_input, pd.DataFrame) and not raw_input.empty:
            parsed_rows = parse_client_dataframe(raw_input, code_str)
            all_rows.extend(parsed_rows)

    df_res = pd.DataFrame(all_rows, columns=CLIENT_MASTER_SCHEMA)

    if not df_res.empty:
        code_order = {31241.0: 1, 19221.0: 2, 43854.0: 3, 72021.0: 4}
        df_res['Sort_Order'] = df_res['Code'].map(lambda x: code_order.get(float(x) if str(x).replace('.','').isdigit() else x, 99))
        df_res = df_res.sort_values(by=['Sort_Order']).drop(columns=['Sort_Order']).reset_index(drop=True)

    # Recalculate single Total Book row
    tot_row = {col: None for col in CLIENT_MASTER_SCHEMA}
    tot_row['Group Name'] = "Total Book"
    tot_row['Invested Equity'] = df_res['Invested Equity'].sum() if not df_res.empty else 0
    tot_row['Invested Debt'] = df_res['Invested Debt'].sum() if not df_res.empty else 0
    tot_row['Invested Hybrid'] = df_res['Invested Hybrid'].sum() if not df_res.empty else 0
    tot_row['Invested Total'] = df_res['Invested Total'].sum() if not df_res.empty else 0
    tot_row['Current Equity AUM'] = df_res['Current Equity AUM'].sum() if not df_res.empty else 0
    tot_row['Current Debt AUM'] = df_res['Current Debt AUM'].sum() if not df_res.empty else 0
    tot_row['Current Hybrid AUM'] = df_res['Current Hybrid AUM'].sum() if not df_res.empty else 0
    tot_row['Current Total AUM'] = df_res['Current Total AUM'].sum() if not df_res.empty else 0

    df_res = pd.concat([df_res, pd.DataFrame([tot_row])], ignore_index=True)
    return df_res


def parse_aum_dataframe(df_raw, code_val):
    """
    Parses a single AUM DataFrame (raw NJ export or clean layout/Others) and returns a list of dictionaries.
    """
    rows = []
    if df_raw is None or df_raw.empty:
        return rows

    c_val = float(code_val) if str(code_val).replace('.','').isdigit() else str(code_val).strip()
    weight_map = load_scheme_weight_map()

    header_row_idx = None
    is_raw_nj = False
    is_others = False

    for r_idx in range(min(5, len(df_raw))):
        row_vals = [str(x).strip() for x in df_raw.iloc[r_idx].values]
        if "Sr. No." in row_vals or ("AMC" in row_vals and "Scheme" in row_vals and "Sr. No." in row_vals):
            is_raw_nj = True
            header_row_idx = r_idx
            break
        elif "AMC Name" in row_vals or "Scheme Name" in row_vals or "Inv. Cost" in row_vals:
            is_others = True
            header_row_idx = r_idx
            break

    if is_raw_nj:
        for idx in range(header_row_idx + 1, len(df_raw)):
            sr_no = df_raw.iloc[idx, 0]
            if pd.isna(sr_no) or str(sr_no).strip().lower() == 'total':
                continue
            
            amc = df_raw.iloc[idx, 1]
            scheme = df_raw.iloc[idx, 2]
            if pd.isna(amc) or pd.isna(scheme) or str(amc).strip().lower() in ['', 'nan', 'none']:
                continue
                
            def clean_num(val):
                if pd.isna(val): return 0.0
                try: return float(str(val).replace(',', '').strip())
                except Exception: return 0.0

            nav = clean_num(df_raw.iloc[idx, 3])
            inv = clean_num(df_raw.iloc[idx, 9])
            eq_aum = clean_num(df_raw.iloc[idx, 10])
            debt_aum = clean_num(df_raw.iloc[idx, 11])
            hybrid_aum = clean_num(df_raw.iloc[idx, 12])
            total_aum = clean_num(df_raw.iloc[idx, 15])
            
            w_val = lookup_brokerage_weight(scheme, weight_map)
            
            rows.append({
                'Code': c_val,
                'AMC Name': str(amc).strip(),
                'Mutual Fund Scheme Name': str(scheme).strip(),
                'NAV (₹)': nav,
                'Invested Total (₹)': inv,
                'Current AUM Total (₹)': total_aum,
                'Equity AUM (₹)': eq_aum,
                'Debt AUM (₹)': debt_aum,
                'Hybrid AUM (₹)': hybrid_aum,
                'Brokerage Weight': w_val
            })
    elif is_others or header_row_idx is not None:
        for idx in range(header_row_idx + 1, len(df_raw)):
            amc = df_raw.iloc[idx, 0]
            scheme = df_raw.iloc[idx, 1]
            category = str(df_raw.iloc[idx, 2]).strip() if pd.notna(df_raw.iloc[idx, 2]) else ""

            if pd.isna(amc) or pd.isna(scheme) or str(amc).strip().lower() in ['', 'nan', 'none', 'amc name']:
                continue
            if str(df_raw.iloc[idx, 3]).strip().lower().startswith('total'):
                continue

            def clean_num(val):
                if pd.isna(val): return 0.0
                try: return float(str(val).replace(',', '').strip())
                except Exception: return 0.0

            nav = clean_num(df_raw.iloc[idx, 3])
            inv = clean_num(df_raw.iloc[idx, 4])
            total_aum = clean_num(df_raw.iloc[idx, 5])

            cat_lower = category.lower()
            s_lower = str(scheme).lower()

            eq_aum, debt_aum, hybrid_aum = 0.0, 0.0, 0.0
            if 'hybrid' in cat_lower or 'hybrid' in s_lower or 'silver' in cat_lower or 'gold' in cat_lower or 'fof' in cat_lower:
                if 'domestic-equity' in cat_lower or 'all cap' in s_lower:
                    eq_aum = total_aum
                elif 'domestic-hybrid' in cat_lower or 'aggressive hybrid' in s_lower or 'dynamic asset' in s_lower:
                    hybrid_aum = total_aum
                elif 'silver' in cat_lower or 'gold' in cat_lower:
                    hybrid_aum = total_aum
                else:
                    hybrid_aum = total_aum
            elif 'debt' in cat_lower or 'liquid' in cat_lower or 'bond' in cat_lower:
                debt_aum = total_aum
            else:
                eq_aum = total_aum

            w_val = lookup_brokerage_weight(scheme, weight_map)

            rows.append({
                'Code': c_val,
                'AMC Name': str(amc).strip(),
                'Mutual Fund Scheme Name': str(scheme).strip(),
                'NAV (₹)': nav,
                'Invested Total (₹)': inv,
                'Current AUM Total (₹)': total_aum,
                'Equity AUM (₹)': eq_aum,
                'Debt AUM (₹)': debt_aum,
                'Hybrid AUM (₹)': hybrid_aum,
                'Brokerage Weight': w_val
            })
    else:
        df_c = df_raw.copy()
        if 'Mutual Fund Scheme Name' in df_c.columns:
            df_c = df_c[df_c['Mutual Fund Scheme Name'] != 'Total Book']
            
        alias_map = {
            'AMC Name': ['amc name', 'amc', 'fund house', 'amc_name'],
            'Mutual Fund Scheme Name': ['mutual fund scheme name', 'scheme name', 'scheme', 'fund name', 'scheme_name'],
            'NAV (₹)': ['nav (₹)', 'nav(₹)', 'nav', 'current nav', 'latest nav'],
            'Invested Total (₹)': ['invested total (₹)', 'invested total', 'invested amount', 'purchase value', 'invested_total', 'inv. cost'],
            'Current AUM Total (₹)': ['current aum total (₹)', 'current aum total', 'current aum', 'aum(₹)', 'aum', 'total aum', 'market value', 'current value'],
            'Equity AUM (₹)': ['equity aum (₹)', 'equity aum', 'equity'],
            'Debt AUM (₹)': ['debt aum (₹)', 'debt aum', 'debt'],
            'Hybrid AUM (₹)': ['hybrid aum (₹)', 'hybrid aum', 'hybrid']
        }
        col_rename = {}
        for col in df_c.columns:
            col_clean = str(col).strip().lower()
            for target_k, aliases in alias_map.items():
                if col_clean in aliases:
                    col_rename[col] = target_k
                    break
        df_c = df_c.rename(columns=col_rename)

        for _, r in df_c.iterrows():
            scheme_name = str(r.get('Mutual Fund Scheme Name', '')).strip()
            if not scheme_name or scheme_name.lower() in ['nan', 'none', '', 'total', 'total book', 'scheme name']:
                continue
                
            amc_name = str(r.get('AMC Name', '')).strip()
            if not amc_name or amc_name.lower() in ['nan', 'none', '', 'amc name']:
                amc_name = scheme_name.split()[0] + " Mutual Fund"

            def clean_val(v):
                if pd.isna(v): return 0.0
                try: return float(str(v).replace(',', '').strip())
                except Exception: return 0.0

            nav = clean_val(r.get('NAV (₹)', 0.0))
            inv = clean_val(r.get('Invested Total (₹)', 0.0))
            aum = clean_val(r.get('Current AUM Total (₹)', 0.0))
            eq = clean_val(r.get('Equity AUM (₹)', 0.0))
            debt = clean_val(r.get('Debt AUM (₹)', 0.0))
            hyb = clean_val(r.get('Hybrid AUM (₹)', 0.0))

            if aum > 0 and (eq + debt + hyb) == 0:
                s_low = scheme_name.lower()
                if 'debt' in s_low or 'bond' in s_low or 'liquid' in s_low:
                    debt = aum
                elif 'hybrid' in s_low or 'balanced' in s_low or 'arbitrage' in s_low:
                    hyb = aum
                else:
                    eq = aum

            w_val = lookup_brokerage_weight(scheme_name, weight_map)

            rows.append({
                'Code': c_val,
                'AMC Name': amc_name,
                'Mutual Fund Scheme Name': scheme_name,
                'NAV (₹)': nav,
                'Invested Total (₹)': inv,
                'Current AUM Total (₹)': aum,
                'Equity AUM (₹)': eq,
                'Debt AUM (₹)': debt,
                'Hybrid AUM (₹)': hyb,
                'Brokerage Weight': w_val
            })

    return rows


def build_sheet3_from_aum_files(aum_files_input):
    """
    Safely handles dict or list input, parses AUM report DataFrames and extra sheets
    (aum_31241_Others -> appended under code 31241_ot, aum_43854_Others -> appended under code 43854_ot),
    and appends a single 'Total Book' summary row.
    """
    all_rows = []
    default_codes = ["31241", "19221", "43854", "72021"]
    processed_file_sheets = set()

    if isinstance(aum_files_input, list):
        items = []
        for idx, item in enumerate(aum_files_input):
            c_code = None
            if isinstance(item, pd.DataFrame) and 'Code' in item.columns:
                codes = item['Code'].dropna().astype(str).unique()
                if len(codes) > 0:
                    c_code = clean_int_string(codes[0])
            if not c_code or c_code.startswith("code_"):
                c_code = default_codes[idx] if idx < len(default_codes) else f"code_{idx+1}"
            items.append((c_code, item))
    elif isinstance(aum_files_input, dict):
        items = list(aum_files_input.items())
    else:
        items = []

    for code_str, raw_input in items:
        if isinstance(raw_input, str) and os.path.exists(raw_input):
            try:
                xls = pd.ExcelFile(raw_input)
                for sname in xls.sheet_names:
                    s_lower = sname.strip().lower()

                    if "31241_others" in s_lower or "31241_ot" in s_lower or "31241_other" in s_lower:
                        target_code = "31241_ot"
                    elif "43854_others" in s_lower or "43854_ot" in s_lower or "43854_other" in s_lower:
                        target_code = "43854_ot"
                    elif "31241" in s_lower:
                        target_code = "31241"
                    elif "43854" in s_lower:
                        target_code = "43854"
                    elif "19221" in s_lower:
                        target_code = "19221"
                    elif "72021" in s_lower:
                        target_code = "72021"
                    else:
                        target_code = code_str

                    df_sheet = pd.read_excel(raw_input, sheet_name=sname)
                    parsed_rows = parse_aum_dataframe(df_sheet, target_code)
                    all_rows.extend(parsed_rows)
                    processed_file_sheets.add(f"{os.path.basename(raw_input)}::{sname}")
            except Exception as e:
                print(f"Error reading AUM file {raw_input}: {e}")
        elif isinstance(raw_input, pd.DataFrame) and not raw_input.empty:
            parsed_rows = parse_aum_dataframe(raw_input, code_str)
            all_rows.extend(parsed_rows)

    check_dirs = [".", os.path.join(os.path.dirname(__file__), "temp_downloads")]
    additional_other_files = [
        ("31241_ot", ["aum_31241_Others.xlsx", "aum_31241_others.xlsx", "31241_Others.xlsx", "31241_others.xlsx"]),
        ("43854_ot", ["aum_43854_Others.xlsx", "aum_43854_others.xlsx", "43854_Others.xlsx", "43854_others.xlsx"])
    ]

    for target_code, fnames in additional_other_files:
        for cdir in check_dirs:
            for fn in fnames:
                fpath = os.path.join(cdir, fn)
                if os.path.exists(fpath):
                    try:
                        xls = pd.ExcelFile(fpath)
                        for sname in xls.sheet_names:
                            key = f"{os.path.basename(fpath)}::{sname}"
                            if key not in processed_file_sheets:
                                df_sheet = pd.read_excel(fpath, sheet_name=sname)
                                parsed_rows = parse_aum_dataframe(df_sheet, target_code)
                                all_rows.extend(parsed_rows)
                                processed_file_sheets.add(key)
                    except Exception as e:
                        print(f"Error reading additional file {fpath}: {e}")

    df_res = pd.DataFrame(all_rows, columns=SCHEME_MASTER_SCHEMA)

    if not df_res.empty:
        code_order = {31241.0: 1, '31241_ot': 2, 19221.0: 3, 43854.0: 4, '43854_ot': 5, 72021.0: 6}
        def get_sort_key(c):
            if c in code_order: return code_order[c]
            try:
                f_c = float(c)
                if f_c in code_order: return code_order[f_c]
            except Exception:
                pass
            return 99

        df_res['Sort_Order'] = df_res['Code'].map(get_sort_key)
        df_res = df_res.sort_values(by=['Sort_Order']).drop(columns=['Sort_Order']).reset_index(drop=True)

    # Append single Total Book summary row
    tot_row = {col: None for col in SCHEME_MASTER_SCHEMA}
    tot_row['Mutual Fund Scheme Name'] = "Total Book"
    tot_row['Invested Total (₹)'] = df_res['Invested Total (₹)'].sum() if not df_res.empty else 0
    tot_row['Current AUM Total (₹)'] = df_res['Current AUM Total (₹)'].sum() if not df_res.empty else 0
    tot_row['Equity AUM (₹)'] = df_res['Equity AUM (₹)'].sum() if not df_res.empty else 0
    tot_row['Debt AUM (₹)'] = df_res['Debt AUM (₹)'].sum() if not df_res.empty else 0
    tot_row['Hybrid AUM (₹)'] = df_res['Hybrid AUM (₹)'].sum() if not df_res.empty else 0
    tot_row['Brokerage Weight'] = None

    df_res = pd.concat([df_res, pd.DataFrame([tot_row])], ignore_index=True)
    return df_res
def purge_and_write_sheet(wb, sheet_name, df, expected_schema):
    """
    Completely deletes existing worksheet object to clear historical rows,
    re-creates the sheet at the exact tab index, and writes schema headers + fresh data.
    """
    sheet_index = None
    if sheet_name in wb.sheetnames:
        sheet_index = wb.sheetnames.index(sheet_name)
        del wb[sheet_name]

    if sheet_index is not None:
        ws = wb.create_sheet(title=sheet_name, index=sheet_index)
    else:
        ws = wb.create_sheet(title=sheet_name)

    ws.append(expected_schema)

    if not df.empty:
        df_out = df[expected_schema]
        for row in df_out.itertuples(index=False, name=None):
            clean_row = ["" if pd.isna(val) or val is None else val for val in row]
            ws.append(clean_row)


def update_master_sheets_from_all_codes(client_files_input, aum_files_input):
    """
    Only executed when feeds for all codes are ready.
    Constructs Sheet 2 and Sheet 3 and overwrites workbook.
    """
    try:
        df_clients_final = build_sheet2_from_client_files(client_files_input)
        df_schemes_final = build_sheet3_from_aum_files(aum_files_input)

        if os.path.exists(EXCEL_FILE):
            wb = openpyxl.load_workbook(EXCEL_FILE)
        else:
            wb = openpyxl.Workbook()

        purge_and_write_sheet(wb, "2. Client AUM Master", df_clients_final, CLIENT_MASTER_SCHEMA)
        purge_and_write_sheet(wb, "3. Scheme Allocation Master", df_schemes_final, SCHEME_MASTER_SCHEMA)

        wb.save(EXCEL_FILE)
        wb.close()

        st.cache_data.clear()
        return True

    except PermissionError:
        st.error("❌ **File Locked:** Close `Live_Wealth_Manager_Book.xlsx` in Excel and try again.")
        return False
    except Exception as e:
        st.error(f"❌ **Master Sheets Update Error:** {e}")
        return False


# -----------------------------------------------------------------------------
# 3. INTERFACE NAVIGATION PANEL & SIDEBAR CONTROLS
# -----------------------------------------------------------------------------
st.sidebar.title("🛠️ Sahyog Financials Workspace")
st.sidebar.caption("Enterprise Consolidated Architecture v10.3")

current_user = st.session_state.get("username", "User")
st.sidebar.info(f"👤 Logged in as: **{current_user}**")
if st.sidebar.button("🚪 Logout", use_container_width=True):
    st.session_state["authenticated"] = False
    st.session_state["username"] = None
    st.rerun()

if is_gsheets_configured():
    st.sidebar.success("🟢 Backend: Live Google Sheets API")
else:
    st.sidebar.warning("🟡 Backend: Local Excel (Add Secrets for Cloud Sync)")

st.sidebar.markdown("---")

# DYNAMIC NAVIGATION (RBAC WITHOUT WARNING BANNERS)
nav_options = [
    "📝 Log Transaction",
    "🔍 Advanced Demographics Desk",
    "🗃️ View & Edit Master Ledger"
]

if current_user.lower() in [u.lower() for u in ALLOWED_ANALYTICS_USERS]:
    nav_options.append("📊 Book Summary & Analytics")

workspace = st.sidebar.radio("Navigate Module Panel:", nav_options)

st.sidebar.markdown("---")

# 2-STEP DECOUPLED SYNC CONTROLLER (FETCH & MERGE SEPARATED)
st.sidebar.markdown("#### 🤖 Live Portal Data Sync")

from portal_scraper import (
    get_last_sync_status, 
    record_sync_completed, 
    fetch_all_portal_sheets, 
    scan_downloaded_feed_files
)

sync_info = get_last_sync_status()
last_sync_date_str = sync_info.get("last_sync_date")
last_sync_time_str = sync_info.get("last_sync_time")

already_synced_this_week = False
days_since_sync = None

if last_sync_date_str:
    try:
        last_date = datetime.datetime.strptime(last_sync_date_str, "%Y-%m-%d").date()
        days_since_sync = (datetime.date.today() - last_date).days
        already_synced_this_week = (days_since_sync < 7)
    except Exception:
        already_synced_this_week = False

if already_synced_this_week:
    st.sidebar.info(f"📅 **Synced {days_since_sync} day(s) ago:** \n`{last_sync_time_str}`")
    force_sync = st.sidebar.checkbox("Force re-sync before 1-week window", value=False)
else:
    if last_sync_time_str:
        st.sidebar.caption(f"Last Sync: {last_sync_time_str} ({days_since_sync} days ago)")
    else:
        st.sidebar.caption("Last Sync: Never executed")
    force_sync = True

col_s1, col_s2 = st.sidebar.columns(2)

# Step 1: Download from portal only
if col_s1.button("📥 1. Fetch Feeds", use_container_width=True, help="Download latest sheets from portal into temp_downloads directory"):
    if already_synced_this_week and not force_sync:
        st.sidebar.warning("⚠️ Portal sheets were downloaded recently. Check 'Force re-sync' to download again.")
    else:
        with st.spinner("Logging into portals across all codes & pulling reports..."):
            try:
                res = fetch_all_portal_sheets(headless=True)
                ok_count = sum(1 for v in res.values() if v["client_ok"] or v["aum_ok"])
                st.sidebar.success(f"✅ Fetched reports for {ok_count} code(s) into `temp_downloads/`!")
            except Exception as e:
                st.sidebar.error(f"❌ Portal Fetch Error: {e}")

# Step 2: Merge local & downloaded feeds into Live Wealth Manager Book
if col_s2.button("💾 2. Merge Book", use_container_width=True, help="Ingest all files from temp_downloads/ & root into Live Wealth Manager Book"):
    with st.spinner("Scanning local feeds and merging into Master Sheets 2 & 3..."):
        try:
            c_files, a_files = scan_downloaded_feed_files()
            if c_files or a_files:
                success = update_master_sheets_from_all_codes(c_files, a_files)
                if success:
                    record_sync_completed()
                    st.sidebar.success("✅ Master Sheets 2 & 3 Updated & Merged Successfully!")
                    st.rerun()
                else:
                    st.sidebar.error("❌ Failed to write merged data to workbook.")
            else:
                st.sidebar.warning("⚠️ No feed files found in `temp_downloads/` or root directory.")
        except Exception as e:
            st.sidebar.error(f"❌ Ingestion / Merge Error: {e}")

# Full Automated Sync (Step 1 + Step 2 in sequence)
if st.sidebar.button("🔄 Full Sync (Fetch + Merge)", use_container_width=True):
    if already_synced_this_week and not force_sync:
        st.sidebar.warning("⚠️ Master sheets were updated within 7 days. Check 'Force re-sync' to overwrite.")
    else:
        with st.spinner("Fetching portal sheets & merging into Master Workbook..."):
            try:
                fetch_all_portal_sheets(headless=True)
                c_files, a_files = scan_downloaded_feed_files()
                if c_files or a_files:
                    success = update_master_sheets_from_all_codes(c_files, a_files)
                    if success:
                        record_sync_completed()
                        st.sidebar.success("✅ Full Automated Sync & Merge Completed!")
                        st.rerun()
                    else:
                        st.sidebar.error("❌ Failed to write scraped data to workbook.")
                else:
                    st.sidebar.warning("⚠️ Scraper completed, but no data files were found.")
            except Exception as e:
                st.sidebar.error(f"❌ Full Sync Error: {e}")


def find_matching_code_index(raw_val, code_list):
    if pd.isna(raw_val) or not raw_val:
        return None
    val_clean = clean_int_string(raw_val)
    val_digits = "".join(filter(str.isdigit, val_clean))
    
    for idx, opt in enumerate(code_list):
        opt_digits = "".join(filter(str.isdigit, opt))
        if val_digits and opt_digits and val_digits == opt_digits:
            return idx
        if val_clean in opt or opt in str(raw_val):
            return idx
    return None


# -----------------------------------------------------------------------------
# MODULE 1: LOG TRANSACTION FORM DESK (WITH SMART AUTOFILL)
# -----------------------------------------------------------------------------
if workspace == "📝 Log Transaction":
    st.title("Log Transaction")
    st.caption("Transactions entered here are written permanently back to the Master Ledger.")
    st.markdown("---")

    st.markdown("#### 👤 Select Investor (Triggers Intelligent Data Autofill)")
    client_select = st.selectbox(
        "Choose Existing Investor from Database:", 
        ["-- Custom / New Investor Entry --"] + client_options,
        key="autofill_client_selector"
    )

    autofill_name = ""
    autofill_ucc = ""
    autofill_code_idx = 0
    autofill_bank = ""
    autofill_payment = "ACH"
    code_list = ["NJ-31241", "NJ-19221", "NJ-43854", "NJ-72021", "ARN-274672"]
    payment_options = ["ACH", "Netbanking", "Cheque", "UPI"]

    if client_select != "-- Custom / New Investor Entry --":
        autofill_name = client_select
        
        if not df_master_clients.empty:
            matches = df_master_clients[df_master_clients['Investor Name'] == client_select]
            if not matches.empty:
                c_row = matches.iloc[0]
                if 'UCC' in c_row and pd.notna(c_row['UCC']):
                    autofill_ucc = clean_int_string(c_row['UCC'])
                if 'Code' in c_row and pd.notna(c_row['Code']):
                    matched_idx = find_matching_code_index(c_row['Code'], code_list)
                    if matched_idx is not None:
                        autofill_code_idx = matched_idx

        if not df_ledger.empty:
            past_txns = df_ledger[df_ledger['Name'] == client_select]
            if not past_txns.empty:
                last_txn = past_txns.iloc[-1]
                if pd.notna(last_txn.get('BANK DETAIL')) and str(last_txn.get('BANK DETAIL')).strip() != "":
                    autofill_bank = str(last_txn.get('BANK DETAIL'))
                if pd.notna(last_txn.get('PAYMENT DETAIL')) and str(last_txn.get('PAYMENT DETAIL')).strip() in payment_options:
                    autofill_payment = str(last_txn.get('PAYMENT DETAIL'))
                if not autofill_ucc and pd.notna(last_txn.get('Ucc / APP./FOLIO NO.')):
                    autofill_ucc = clean_int_string(last_txn.get('Ucc / APP./FOLIO NO.'))
                if autofill_code_idx == 0 and pd.notna(last_txn.get('Code')):
                    matched_idx = find_matching_code_index(last_txn.get('Code'), code_list)
                    if matched_idx is not None:
                        autofill_code_idx = matched_idx

        st.success(f"⚡ Smart Autofill Active for **{client_select}**: UCC/Folio, Bank Details & Default License Code auto-loaded!")

    with st.form("persistent_form", clear_on_submit=True):
        st.markdown("<div class='section-header'>1. Core Transaction Details</div>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            txn_date = st.date_input("Date", datetime.date.today())
            final_name = st.text_input("Investor Name", value=autofill_name, placeholder="Type investor full name")
        with c2:
            txn_amount = st.number_input("Amount (₹)", min_value=0.0, step=1000.0)
            txn_type = st.selectbox("TYPE", [
                "Purchase", "Redemption", "SIP", "SWP", "STP", "Switch",
                "PMS Subscription", "PMS Top-Up", "PMS Redemption"
            ])
        with c3:
            ucc_folio = st.text_input("Ucc / APP./FOLIO NO.", value=autofill_ucc)
            top_up = st.selectbox("Top Up", ["No", "Yes"])

        st.markdown("<div class='section-header'>2. Asset Class & Routing Channels</div>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            scheme_select = st.selectbox("SCHEME (Select Pre-loaded Product/Fund)", ["-- Custom / PMS Entry --"] + scheme_options)
            scheme_custom = st.text_input("SCHEME (Type custom product / strategy name)")
            final_scheme = scheme_custom if scheme_select == "-- Custom / PMS Entry --" else scheme_select
        with c2:
            from_target = st.text_input("FROM (Source scheme for Switch/STP)")
        with c3:
            to_target = st.text_input("TO (Destination target scheme for Switch/STP)")

        st.markdown("<div class='section-header'>3. Banking Rails & Compliance Audit Details</div>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            pay_idx = payment_options.index(autofill_payment) if autofill_payment in payment_options else 0
            payment_detail = st.selectbox("PAYMENT DETAIL", payment_options, index=pay_idx)
            bank_detail = st.text_input("BANK DETAIL (Bank Name & A/c details)", value=autofill_bank)
        with c2:
            prod_code = st.selectbox("Code", code_list, index=autofill_code_idx)
            stname_subbroker = st.selectbox("STNAME (Sub Broker Name)", ["NJ", "MFU"])
            exit_load_val = st.text_input("EXIT LOAD (e.g. 1%)")
        with c3:
            remark = st.text_input("REMARK")
            other_remarks = st.text_area("Other Remarks", height=68)

        st.markdown("<div class='section-header'>4. Audit Tracking & Creator Signature</div>", unsafe_allow_html=True)
        c_audit1, c_audit2 = st.columns(2)
        with c_audit1:
            logged_by = st.text_input("Entry created by (Logged by)", value=current_user)
        with c_audit2:
            st.info("ℹ️ **Verification Note:** 'Entry checked by' is saved as empty initially. Auditors can check entries using the tick box in the Master Ledger.")

        submit = st.form_submit_button("Save Permanently to Database")

        if submit:
            if not final_name or final_name.strip() == "":
                st.error("❌ Save Refused: 'Name' is a required field.")
            elif txn_amount <= 0:
                st.error("❌ Save Refused: 'Amount' must be greater than zero.")
            elif "PMS" in txn_type and "Subscription" in txn_type and txn_amount < 5000000:
                st.error("❌ Regulatory Halt: Fresh PMS mandates cannot fall below the SEBI mandate of ₹50 Lakhs.")
            else:
                new_record = {
                    "Date": str(txn_date),
                    "Name": final_name,
                    "Amount": float(txn_amount),
                    "Ucc / APP./FOLIO NO.": clean_int_string(ucc_folio),
                    "SCHEME": final_scheme,
                    "TYPE": txn_type,
                    "Top Up": top_up,
                    "FROM": from_target,
                    "TO": to_target,
                    "PAYMENT DETAIL": payment_detail,
                    "BANK DETAIL": bank_detail,
                    "REMARK": remark,
                    "Code": prod_code,
                    "STNAME": stname_subbroker,
                    "EXIT LOAD": exit_load_val,
                    "Other Remarks": other_remarks,
                    "Entry checked by": "",
                    "Entry created by": logged_by,
                    "Last edited by": "",
                    "Last edit timestamp": ""
                }
                was_gs = save_transaction_record(new_record)
                if was_gs:
                    st.success("⚡ Transaction appended directly to your Live Google Sheet in real time!")
                else:
                    st.success("💾 Transaction saved locally inside the central workbook!")
                st.rerun()


# -----------------------------------------------------------------------------
# MODULE 2: ADVANCED DEMOGRAPHICS LOOKUPS DESK
# -----------------------------------------------------------------------------
elif workspace == "🔍 Advanced Demographics Desk":
    st.title("Advanced Demographics & Investor Lookups Desk")
    st.caption("Query the fully enriched multi-branch master database using micro-demographic tracking keys.")
    st.markdown("---")
    
    if df_master_clients.empty:
        st.info("No master data loaded yet.")
    else:
        search_filter = st.radio("Select Lookup Key Method:", ["Search by PAN", "Search by UCC", "Search by Distribution Code", "Search by Client Group/Name"], horizontal=True)
        
        filtered_df = df_master_clients.copy()
        
        if search_filter == "Search by PAN":
            pan_input = st.text_input("Enter 10-Digit PAN:").strip().upper()
            if pan_input:
                filtered_df = filtered_df[filtered_df['PAN'].astype(str).str.contains(pan_input, na=False)]
        elif search_filter == "Search by UCC":
            ucc_input = st.text_input("Enter UCC Code:").strip()
            if ucc_input:
                filtered_df = filtered_df[filtered_df['UCC'].astype(str).str.contains(ucc_input, na=False)]
        elif search_filter == "Search by Distribution Code":
            code_select = st.selectbox("Select Branch License Code:", sorted(filtered_df['Code'].astype(str).unique().tolist()))
            filtered_df = filtered_df[filtered_df['Code'].astype(str) == code_select]
        elif search_filter == "Search by Client Group/Name":
            name_input = st.text_input("Enter Investor / Group Name keyword:").strip().lower()
            if name_input:
                filtered_df = filtered_df[
                    filtered_df['Investor Name'].astype(str).str.lower().str.contains(name_input, na=False) |
                    filtered_df['Group Name'].astype(str).str.lower().str.contains(name_input, na=False)
                ]
        
        st.markdown(f"#### Matching Accounts Records: ({len(filtered_df)} profiles found)")
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)
        
        if len(filtered_df) == 1:
            st.markdown("---")
            st.markdown("### 📈 Individual Portfolio Breakdown Snapshot")
            row = filtered_df.iloc[0]
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Investor Name", str(row['Investor Name']))
            c2.metric("Total Current AUM", f"₹{row['Current Total AUM']:,.0f}")
            c3.metric("Personal XIRR", f"{row['Total XIRR (%)']:.2f}%")
            c4.metric("Tax Status Slab", str(row['Tax Status']))
            
            labels = ['Equity AUM', 'Debt AUM', 'Hybrid AUM']
            vals = [row['Current Equity AUM'], row['Current Debt AUM'], row['Current Hybrid AUM']]
            if sum(vals) > 0:
                fig_ind = px.pie(names=labels, values=vals, hole=0.4, color_discrete_sequence=px.colors.sequential.Plotly3)
                st.plotly_chart(fig_ind, use_container_width=True)


# -----------------------------------------------------------------------------
# MODULE 3: VIEW & EDIT MASTER LEDGER (WITH AUDITOR CHECKBOX)
# -----------------------------------------------------------------------------
elif workspace == "🗃️ View & Edit Master Ledger":
    st.title("Permanent Transaction Audit Ledger & History Editor")
    st.caption("Direct streaming logs retrieved from '4. Transaction Ledger' tab with live auditor verification tick box.")
    st.markdown("---")

    if df_ledger.empty:
        st.info("No saved records found in the transaction log sheet yet. Use the input form to create rows.")
    else:
        amt_series = pd.to_numeric(df_ledger["Amount"], errors="coerce").fillna(0)
        total_volume = amt_series.sum()
        total_rows = len(df_ledger)
        checked_rows = df_ledger["Entry checked by"].apply(lambda x: pd.notna(x) and str(x).strip() != "").sum()
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Logged Entries", f"{total_rows} Rows")
        with col2:
            st.metric("Gross Logged Volume", f"₹{total_volume:,.2f}")
        with col3:
            st.metric("Audited & Verified", f"{checked_rows} Rows")

        df_ledger_display = df_ledger.copy()
        
        for c in df_ledger_display.columns:
            if c == "Ucc / APP./FOLIO NO.":
                df_ledger_display[c] = df_ledger_display[c].apply(clean_int_string)
            elif c not in ["Amount", "Date"]:
                df_ledger_display[c] = df_ledger_display[c].astype(str).replace(["nan", "NaN", "None"], "")

        df_ledger_display.insert(
            0, 
            "Verified (Tick Box)", 
            df_ledger_display["Entry checked by"].apply(lambda x: True if pd.notna(x) and str(x).strip() != "" else False)
        )

        st.markdown("### ✏️ Interactive Editable Ledger & Auditor Verification Desk")
        st.info(f"💡 **Instructions:** Tick the **'Verified (Tick Box)'** column to mark an entry as checked by auditor (**{current_user}**). Click **'💾 Save Ledger Modifications'** below when finished.")
        
        edited_df = st.data_editor(
            df_ledger_display, 
            use_container_width=True, 
            num_rows="dynamic",
            column_config={
                "Verified (Tick Box)": st.column_config.CheckboxColumn(
                    "Verified (Tick Box)",
                    help="Tick to mark this entry as checked/verified by auditor",
                    default=False,
                )
            },
            key="master_ledger_editor"
        )

        if st.button("💾 Save Ledger Modifications & Verification", type="primary"):
            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            for c in edited_df.columns:
                if c not in ["Amount", "Verified (Tick Box)"]:
                    edited_df[c] = edited_df[c].astype(object)

            for idx in edited_df.index:
                is_ticked = edited_df.loc[idx, "Verified (Tick Box)"]
                orig_checked = df_ledger.loc[idx, "Entry checked by"] if idx in df_ledger.index else ""
                
                if is_ticked and (pd.isna(orig_checked) or str(orig_checked).strip() == ""):
                    edited_df.at[idx, "Entry checked by"] = current_user
                elif not is_ticked:
                    edited_df.at[idx, "Entry checked by"] = ""

                if idx in df_ledger.index:
                    row_orig = df_ledger.loc[idx].fillna("").astype(str).tolist()
                    row_edit = edited_df.drop(columns=["Verified (Tick Box)"]).loc[idx].fillna("").astype(str).tolist()
                    if row_orig != row_edit:
                        edited_df.at[idx, "Last edited by"] = current_user
                        edited_df.at[idx, "Last edit timestamp"] = current_time

            final_save_df = edited_df.drop(columns=["Verified (Tick Box)"])
            
            was_gs = overwrite_ledger_database(final_save_df)
            if was_gs:
                st.success("⚡ Saved edits & verifications directly to Live Google Sheet!")
            else:
                st.success("💾 Ledger edits & verifications saved locally inside workbook!")
            st.rerun()


# -----------------------------------------------------------------------------
# MODULE 4: ENRICHED BOOK SUMMARY & ANALYTICS (WITH SCHEME TRIMMING DESK)
# -----------------------------------------------------------------------------
elif workspace == "📊 Book Summary & Analytics":
    st.title("Sahyog Financials Advanced Intelligence Desk")
    st.caption("Consolidated macro insights, distribution code breakdowns, sales leaderboards, AMC weights, and scheme trimming analytics.")
    st.markdown("---")
    
    if df_master_clients.empty or df_master_schemes.empty:
        st.error("Historical base sheets are missing from the workbook. Re-verify Excel file integrity.")
    else:
        # EXPANDED 7-TAB ANALYTICS MATRIX INCLUDING INTERACTIVE SCHEME EXPLORER
        tab1, tab_schemes, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "🏠 Executive Overview", 
            "📈 Interactive Scheme Explorer",
            "🎯 Scheme Rationalization Desk",
            "📌 Distribution Code Breakdown", 
            "👥 Executive Leaderboard", 
            "💎 Client Segment Matrix", 
            "🏢 Asset House Concentration"
        ])
        
        total_aum_base = df_master_clients['Current Total AUM'].sum()
        equity_base = df_master_clients['Current Equity AUM'].sum()
        hybrid_base = df_master_clients['Current Hybrid AUM'].sum()
        debt_base = df_master_clients['Current Debt AUM'].sum()
        
        with tab1:
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.markdown(f"<div class='metric-card'><h5>Consolidated AUA</h5><h3>₹{total_aum_base:,.2f}</h3></div>", unsafe_allow_html=True)
            with c2:
                st.markdown(f"<div class='metric-card'><h5>Monitored Investors</h5><h3>{len(df_master_clients)} Accounts</h3></div>", unsafe_allow_html=True)
            with c3:
                st.markdown(f"<div class='metric-card'><h5>Active Product Matrix</h5><h3>{len(df_master_schemes)} Schemes</h3></div>", unsafe_allow_html=True)
            with c4:
                avg_ticket = total_aum_base / len(df_master_clients) if len(df_master_clients) > 0 else 0
                st.markdown(f"<div class='metric-card'><h5>Average Ticket Size</h5><h3>₹{avg_ticket:,.0f}</h3></div>", unsafe_allow_html=True)
                
            ch_col1, ch_col2 = st.columns(2)
            with ch_col1:
                st.markdown("#### Top 10 High Net Worth Investor Accounts")
                top10_c = df_master_clients.sort_values(by='Current Total AUM', ascending=False).head(10)
                fig1 = px.bar(top10_c, x='Current Total AUM', y='Investor Name', orientation='h', color_discrete_sequence=['#1F4E79'])
                fig1.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig1, use_container_width=True)
            with ch_col2:
                st.markdown("#### Consolidated Asset Class Weights Matrix")
                labels = ['Equity', 'Hybrid', 'Debt']
                values = [equity_base, hybrid_base, debt_base]
                fig2 = px.pie(names=labels, values=values, hole=0.45, color_discrete_sequence=px.colors.sequential.Blues_r)
                st.plotly_chart(fig2, use_container_width=True)

            st.markdown("---")
            st.markdown("#### 📈 Portfolio Return (XIRR) Distribution Matrix")
            
            def xirr_bracket(val):
                if pd.isna(val): return "N/A"
                if val >= 15: return "🚀 High Return (>15%)"
                elif val >= 10: return "📈 Good Return (10% - 15%)"
                elif val >= 5: return "🟡 Moderate Return (5% - 10%)"
                elif val >= 0: return "🟠 Low Return (0% - 5%)"
                else: return "🔻 Negative Return (<0%)"

            df_c_copy = df_master_clients.copy()
            df_c_copy['XIRR Bracket'] = df_c_copy['Total XIRR (%)'].apply(xirr_bracket)
            xirr_dist = df_c_copy.groupby('XIRR Bracket').agg(
                Accounts=('Investor Name', 'count'),
                Combined_AUM=('Current Total AUM', 'sum')
            ).reset_index()
            xirr_dist['AUM Share %'] = (xirr_dist['Combined_AUM'] / total_aum_base) * 100

            st.dataframe(
                xirr_dist.style.format({
                    "Combined_AUM": "₹{:,.2f}",
                    "Accounts": "{:,.0f}",
                    "AUM Share %": "{:.2f}%"
                }), use_container_width=True, hide_index=True
            )

        
        # TAB FOR INTERACTIVE SCHEME INSIGHTS & EXPLORER
        with tab_schemes:
            st.markdown("### 📈 Interactive Scheme Explorer & Revenue Analytics")
            st.caption("Filter, slice, and analyze scheme performance, concentration risk, yield optimization, asset allocation, and brokerage multipliers across fund houses and license codes.")

            df_s_explorer = df_master_schemes[df_master_schemes['Mutual Fund Scheme Name'] != 'Total Book'].copy()

            if 'Brokerage Weight' not in df_s_explorer.columns:
                df_s_explorer['Brokerage Weight'] = 1.0
            df_s_explorer['Brokerage Weight'] = df_s_explorer['Brokerage Weight'].fillna(1.0)

            df_s_explorer['Unrealized Gain (₹)'] = df_s_explorer['Current AUM Total (₹)'] - df_s_explorer['Invested Total (₹)']
            df_s_explorer['Absolute Return (%)'] = np.where(
                df_s_explorer['Invested Total (₹)'] > 0,
                ((df_s_explorer['Unrealized Gain (₹)'] / df_s_explorer['Invested Total (₹)']) * 100),
                0.0
            )
            df_s_explorer['Weighted AUM (₹)'] = df_s_explorer['Current AUM Total (₹)'] * df_s_explorer['Brokerage Weight']
            df_s_explorer['Code_Str'] = df_s_explorer['Code'].apply(clean_int_string)

            def detect_asset_class(r):
                eq = r.get('Equity AUM (₹)', 0)
                debt = r.get('Debt AUM (₹)', 0)
                hyb = r.get('Hybrid AUM (₹)', 0)
                if hyb > 0 and hyb >= eq and hyb >= debt: return "Hybrid"
                if debt > 0 and debt >= eq and debt >= hyb: return "Debt"
                return "Equity"

            df_s_explorer['Asset Class'] = df_s_explorer.apply(detect_asset_class, axis=1)

            def detect_quadrant(r):
                aum = r.get('Current AUM Total (₹)', 0)
                ret = r.get('Absolute Return (%)', 0)
                if aum >= 10000000 and ret >= 10:
                    return "⭐ Core Stars (High AUM, High Return)"
                elif aum >= 10000000 and ret < 10:
                    return "⚠️ Underperformers (High AUM, Low Return)"
                elif aum < 10000000 and ret >= 10:
                    return "🚀 Growth Opportunities (Low AUM, High Return)"
                else:
                    return "🔴 Tail Holdings (Low AUM, Low Return)"

            df_s_explorer['Performance Quadrant'] = df_s_explorer.apply(detect_quadrant, axis=1)

            def detect_yield_band(w):
                if w < 0.70: return "🔻 Low Yield (< 0.70x)"
                elif w < 0.90: return "🟡 Moderate Yield (0.70x - 0.90x)"
                elif w <= 1.10: return "🟢 Benchmark Yield (0.90x - 1.10x)"
                else: return "🚀 Premium Yield (> 1.10x)"

            df_s_explorer['Yield Band'] = df_s_explorer['Brokerage Weight'].apply(detect_yield_band)

            st.markdown("#### ⚙️ Interactive Scheme Filter Desk")
            f_col1, f_col2, f_col3, f_col4 = st.columns(4)

            with f_col1:
                amc_list = sorted(df_s_explorer['AMC Name'].dropna().unique().tolist())
                selected_amcs = st.multiselect("Filter by Fund House (AMC):", amc_list, default=[], placeholder="All AMCs selected")

            with f_col2:
                code_list = sorted(df_s_explorer['Code_Str'].dropna().unique().tolist())
                selected_codes = st.multiselect("Filter by Distribution Code:", code_list, default=[], placeholder="All Codes selected")

            with f_col3:
                selected_asset_class = st.selectbox("Filter by Asset Class:", ["All Asset Classes", "Equity", "Hybrid", "Debt"])

            with f_col4:
                scheme_search = st.text_input("Search Scheme Name Keyword:", value="", placeholder="e.g. Flexi, Midcap, FOF")

            sf_col1, sf_col2, sf_col3 = st.columns(3)
            with sf_col1:
                max_w_val = float(df_s_explorer['Brokerage Weight'].max())
                weight_range = st.slider("Brokerage Weight Range:", min_value=0.0, max_value=max(2.5, round(max_w_val, 1)), value=(0.0, max(2.5, round(max_w_val, 1))), step=0.1)

            with sf_col2:
                aum_filter = st.selectbox("AUM Size Bracket Filter:", [
                    "All AUM Sizes",
                    "💎 Tier 1 (≥ ₹1 Crore)",
                    "🟡 Tier 2 (₹50 Lakhs - ₹1 Crore)",
                    "🟠 Tier 3 (₹10 Lakhs - ₹50 Lakhs)",
                    "🔴 Tier 4 (< ₹10 Lakhs)",
                    "🔬 Micro Holdings (< ₹2 Lakhs)"
                ])

            with sf_col3:
                quadrant_filter = st.selectbox("Performance Quadrant Filter:", [
                    "All Quadrants",
                    "⭐ Core Stars (High AUM, High Return)",
                    "⚠️ Underperformers (High AUM, Low Return)",
                    "🚀 Growth Opportunities (Low AUM, High Return)",
                    "🔴 Tail Holdings (Low AUM, Low Return)"
                ])

            df_filtered = df_s_explorer.copy()

            if selected_amcs:
                df_filtered = df_filtered[df_filtered['AMC Name'].isin(selected_amcs)]

            if selected_codes:
                df_filtered = df_filtered[df_filtered['Code_Str'].isin(selected_codes)]

            if selected_asset_class != "All Asset Classes":
                df_filtered = df_filtered[df_filtered['Asset Class'] == selected_asset_class]

            if scheme_search.strip():
                df_filtered = df_filtered[df_filtered['Mutual Fund Scheme Name'].str.contains(scheme_search.strip(), case=False, na=False)]

            df_filtered = df_filtered[(df_filtered['Brokerage Weight'] >= weight_range[0]) & (df_filtered['Brokerage Weight'] <= weight_range[1])]

            if aum_filter == "💎 Tier 1 (≥ ₹1 Crore)":
                df_filtered = df_filtered[df_filtered['Current AUM Total (₹)'] >= 10000000]
            elif aum_filter == "🟡 Tier 2 (₹50 Lakhs - ₹1 Crore)":
                df_filtered = df_filtered[(df_filtered['Current AUM Total (₹)'] >= 5000000) & (df_filtered['Current AUM Total (₹)'] < 10000000)]
            elif aum_filter == "🟠 Tier 3 (₹10 Lakhs - ₹50 Lakhs)":
                df_filtered = df_filtered[(df_filtered['Current AUM Total (₹)'] >= 1000000) & (df_filtered['Current AUM Total (₹)'] < 5000000)]
            elif aum_filter == "🔴 Tier 4 (< ₹10 Lakhs)":
                df_filtered = df_filtered[df_filtered['Current AUM Total (₹)'] < 1000000]
            elif aum_filter == "🔬 Micro Holdings (< ₹2 Lakhs)":
                df_filtered = df_filtered[df_filtered['Current AUM Total (₹)'] < 200000]

            if quadrant_filter != "All Quadrants":
                df_filtered = df_filtered[df_filtered['Performance Quadrant'] == quadrant_filter]

            st.markdown("---")
            m_col1, m_col2, m_col3, m_col4, m_col5 = st.columns(5)

            filt_aum = df_filtered['Current AUM Total (₹)'].sum()
            filt_inv = df_filtered['Invested Total (₹)'].sum()
            filt_gain = filt_aum - filt_inv
            filt_gain_pct = (filt_gain / filt_inv * 100) if filt_inv > 0 else 0.0
            filt_count = len(df_filtered)

            weighted_aum_sum = df_filtered['Weighted AUM (₹)'].sum()
            avg_weight_factor = (weighted_aum_sum / filt_aum) if filt_aum > 0 else 1.0

            m_col1.metric("Filtered AUA Volume", f"₹{filt_aum:,.0f}")
            m_col2.metric("Invested Capital", f"₹{filt_inv:,.0f}")
            m_col3.metric("Unrealized Gain / Return", f"₹{filt_gain:,.0f}", f"{filt_gain_pct:+.2f}%")
            m_col4.metric("Avg Brokerage Weight", f"{avg_weight_factor:.4f}")
            m_col5.metric("Active Scheme Count", f"{filt_count} Funds")

            st.markdown("---")
            chart_sub1, chart_sub2, chart_sub3, chart_sub4, chart_sub5, chart_sub6 = st.tabs([
                "📊 Risk/Return Quadrant Plot",
                "🏢 AMC & Scheme Hierarchical Map",
                "🎯 Performance Quadrant Matrix",
                "💰 Brokerage Revenue & Yield Bands",
                "⚖️ Portfolio Health & Concentration",
                "🔍 Side-by-Side Scheme Comparison"
            ])

            with chart_sub1:
                if not df_filtered.empty:
                    df_plot = df_filtered[df_filtered['Current AUM Total (₹)'] > 0].copy()
                    if not df_plot.empty:
                        fig_scatter = px.scatter(
                            df_plot,
                            x='Current AUM Total (₹)',
                            y='Absolute Return (%)',
                            size='Current AUM Total (₹)',
                            color='Performance Quadrant',
                            hover_name='Mutual Fund Scheme Name',
                            hover_data=['AMC Name', 'Code_Str', 'NAV (₹)', 'Invested Total (₹)', 'Brokerage Weight'],
                            title="Scheme AUM Volume vs Absolute Return % Matrix",
                            color_discrete_map={
                                "⭐ Core Stars (High AUM, High Return)": "#2ea043",
                                "⚠️ Underperformers (High AUM, Low Return)": "#d97706",
                                "🚀 Growth Opportunities (Low AUM, High Return)": "#2563eb",
                                "🔴 Tail Holdings (Low AUM, Low Return)": "#dc2626"
                            },
                            labels={'Current AUM Total (₹)': 'Current AUM Total (₹)', 'Absolute Return (%)': 'Absolute Return (%)'}
                        )
                        fig_scatter.add_vline(x=10000000, line_dash="dash", line_color="gray", annotation_text="₹1 Cr Threshold")
                        fig_scatter.add_hline(y=10.0, line_dash="dash", line_color="gray", annotation_text="10% Return Benchmark")
                        fig_scatter.update_layout(height=540)
                        st.plotly_chart(fig_scatter, use_container_width=True)
                    else:
                        st.info("No scheme with positive AUM available for scatter plot.")
                else:
                    st.info("No matching schemes found for current filters.")

            with chart_sub2:
                if not df_filtered.empty:
                    df_plot_tm = df_filtered[df_filtered['Current AUM Total (₹)'] > 0].copy()
                    if not df_plot_tm.empty:
                        fig_treemap = px.treemap(
                            df_plot_tm,
                            path=['AMC Name', 'Mutual Fund Scheme Name'],
                            values='Current AUM Total (₹)',
                            color='Current AUM Total (₹)',
                            color_continuous_scale='Blues',
                            title="AMC & Scheme Portfolio Hierarchical Map"
                        )
                        fig_treemap.update_layout(height=540)
                        st.plotly_chart(fig_treemap, use_container_width=True)
                    else:
                        st.info("No scheme with positive AUM available for treemap.")

            with chart_sub3:
                if not df_filtered.empty:
                    st.markdown("#### 🎯 Performance Quadrant Portfolio Matrix")
                    quad_summary = df_filtered.groupby('Performance Quadrant').agg(
                        Scheme_Count=('Mutual Fund Scheme Name', 'count'),
                        Total_AUM=('Current AUM Total (₹)', 'sum'),
                        Total_Invested=('Invested Total (₹)', 'sum'),
                        Avg_Brokerage_Weight=('Brokerage Weight', 'mean')
                    ).reset_index()

                    quad_summary['Unrealized Gain'] = quad_summary['Total_AUM'] - quad_summary['Total_Invested']
                    quad_summary['Avg Absolute Return %'] = np.where(
                        quad_summary['Total_Invested'] > 0,
                        (quad_summary['Unrealized Gain'] / quad_summary['Total_Invested']) * 100,
                        0.0
                    )
                    quad_summary['AUM Contribution %'] = (quad_summary['Total_AUM'] / (filt_aum if filt_aum > 0 else 1) * 100)

                    st.dataframe(
                        quad_summary.style.format({
                            "Total_AUM": "₹{:,.0f}",
                            "Total_Invested": "₹{:,.0f}",
                            "Unrealized Gain": "₹{:,.0f}",
                            "Avg Absolute Return %": "{:+.2f}%",
                            "Avg_Brokerage_Weight": "{:.4f}",
                            "AUM Contribution %": "{:.2f}%"
                        }),
                        use_container_width=True,
                        hide_index=True
                    )

            with chart_sub4:
                if not df_filtered.empty:
                    st.markdown("#### 💰 Brokerage Yield Band Breakdown")
                    st.caption("Categorization of portfolio schemes by brokerage multiplier tiers to pinpoint fee yield optimization opportunities.")
                    
                    yield_summary = df_filtered.groupby('Yield Band').agg(
                        Scheme_Count=('Mutual Fund Scheme Name', 'count'),
                        Total_AUM=('Current AUM Total (₹)', 'sum'),
                        Weighted_AUM=('Weighted AUM (₹)', 'sum'),
                        Avg_Return_Pct=('Absolute Return (%)', 'mean')
                    ).reset_index()

                    yield_summary['AUM Contribution %'] = (yield_summary['Total_AUM'] / (filt_aum if filt_aum > 0 else 1)) * 100
                    yield_summary['Effective Multiplier'] = yield_summary['Weighted_AUM'] / yield_summary['Total_AUM'].replace(0, 1)

                    st.dataframe(
                        yield_summary[['Yield Band', 'Scheme_Count', 'Total_AUM', 'AUM Contribution %', 'Effective Multiplier', 'Avg_Return_Pct']].style.format({
                            "Total_AUM": "₹{:,.0f}",
                            "AUM Contribution %": "{:.2f}%",
                            "Effective Multiplier": "{:.4f}",
                            "Avg_Return_Pct": "{:+.2f}%",
                            "Scheme_Count": "{:,.0f}"
                        }),
                        use_container_width=True,
                        hide_index=True
                    )

                    st.markdown("---")
                    st.markdown("#### 🎯 Revenue Leakage Watchlist (High AUM Schemes with Brokerage Multiplier < 0.90x)")
                    low_yield_high_aum = df_filtered[(df_filtered['Current AUM Total (₹)'] >= 5000000) & (df_filtered['Brokerage Weight'] < 0.90)].sort_values(by='Current AUM Total (₹)', ascending=False)
                    
                    if not low_yield_high_aum.empty:
                        st.dataframe(
                            low_yield_high_aum[['AMC Name', 'Mutual Fund Scheme Name', 'Code_Str', 'Current AUM Total (₹)', 'Absolute Return (%)', 'Brokerage Weight']].rename(columns={'Code_Str': 'Code'}).style.format({
                                "Current AUM Total (₹)": "₹{:,.0f}",
                                "Absolute Return (%)": "{:+.2f}%",
                                "Brokerage Weight": "{:.4f}"
                            }),
                            use_container_width=True,
                            hide_index=True
                        )
                    else:
                        st.info("No high-AUM schemes found with low brokerage multipliers in current selection.")

            with chart_sub5:
                if not df_filtered.empty:
                    st.markdown("#### ⚖️ Scheme Portfolio Concentration & Health Metrics")
                    
                    c_aum_sorted = df_filtered.sort_values(by='Current AUM Total (₹)', ascending=False)['Current AUM Total (₹)'].values
                    top10_val = c_aum_sorted[:10].sum() if len(c_aum_sorted) >= 10 else filt_aum
                    top20_val = c_aum_sorted[:20].sum() if len(c_aum_sorted) >= 20 else filt_aum

                    top10_share = (top10_val / filt_aum * 100) if filt_aum > 0 else 0.0
                    top20_share = (top20_val / filt_aum * 100) if filt_aum > 0 else 0.0

                    # Herfindahl-Hirschman Index (HHI) for concentration
                    hhi_val = ((df_filtered['Current AUM Total (₹)'] / (filt_aum if filt_aum > 0 else 1)) ** 2).sum() * 10000

                    gainer_count = len(df_filtered[df_filtered['Unrealized Gain (₹)'] >= 0])
                    loser_count = len(df_filtered[df_filtered['Unrealized Gain (₹)'] < 0])

                    hc1, hc2, hc3, hc4 = st.columns(4)
                    hc1.metric("Top 10 Schemes Volume Share", f"{top10_share:.2f}%", f"₹{top10_val:,.0f}")
                    hc2.metric("Top 20 Schemes Volume Share", f"{top20_share:.2f}%", f"₹{top20_val:,.0f}")
                    hc3.metric("Concentration HHI Index", f"{hhi_val:.1f}", "Balanced" if hhi_val < 1000 else "Concentrated")
                    hc4.metric("Gainers vs Loss-Makers", f"{gainer_count} Gain / {loser_count} Loss")

                    st.markdown("---")
                    st.markdown("#### Top 15 Schemes by AUM Concentration")
                    top15_schemes = df_filtered.sort_values(by='Current AUM Total (₹)', ascending=False).head(15)
                    fig_top15 = px.bar(
                        top15_schemes,
                        x='Current AUM Total (₹)',
                        y='Mutual Fund Scheme Name',
                        orientation='h',
                        color='Absolute Return (%)',
                        title="Top 15 Largest Schemes by AUM (Color = Absolute Return %)",
                        color_continuous_scale='RdYlGn'
                    )
                    fig_top15.update_layout(yaxis={'categoryorder': 'total ascending'}, height=500)
                    st.plotly_chart(fig_top15, use_container_width=True)

            with chart_sub6:
                st.markdown("#### 🔍 Side-by-Side Scheme Comparison Tool")
                all_schemes_options = sorted(df_s_explorer['Mutual Fund Scheme Name'].dropna().unique().tolist())
                compared_schemes = st.multiselect(
                    "Select 2 to 5 Schemes to Compare Side-by-Side:",
                    options=all_schemes_options,
                    default=all_schemes_options[:3] if len(all_schemes_options) >= 3 else all_schemes_options
                )

                if compared_schemes:
                    df_comp = df_s_explorer[df_s_explorer['Mutual Fund Scheme Name'].isin(compared_schemes)].copy()
                    
                    comp_metrics = df_comp[[
                        'Mutual Fund Scheme Name', 'AMC Name', 'Code_Str', 'Asset Class',
                        'NAV (₹)', 'Invested Total (₹)', 'Current AUM Total (₹)',
                        'Unrealized Gain (₹)', 'Absolute Return (%)', 'Brokerage Weight',
                        'Weighted AUM (₹)', 'Performance Quadrant'
                    ]].copy().rename(columns={'Code_Str': 'Code'})

                    st.dataframe(
                        comp_metrics.style.format({
                            "NAV (₹)": "₹{:,.2f}",
                            "Invested Total (₹)": "₹{:,.2f}",
                            "Current AUM Total (₹)": "₹{:,.2f}",
                            "Unrealized Gain (₹)": "₹{:,.2f}",
                            "Absolute Return (%)": "{:+.2f}%",
                            "Brokerage Weight": "{:.4f}",
                            "Weighted AUM (₹)": "₹{:,.2f}"
                        }),
                        use_container_width=True,
                        hide_index=True
                    )

            st.markdown("---")
            st.markdown("#### 📋 Filtered Schemes Master Data Table")

            df_display = df_filtered[[
                'Code_Str', 'AMC Name', 'Mutual Fund Scheme Name', 'NAV (₹)',
                'Invested Total (₹)', 'Current AUM Total (₹)', 'Unrealized Gain (₹)',
                'Absolute Return (%)', 'Brokerage Weight', 'Weighted AUM (₹)',
                'Asset Class', 'Performance Quadrant'
            ]].copy()

            df_display = df_display.rename(columns={'Code_Str': 'Code'})

            st.dataframe(
                df_display.style.format({
                    "NAV (₹)": "₹{:,.2f}",
                    "Invested Total (₹)": "₹{:,.2f}",
                    "Current AUM Total (₹)": "₹{:,.2f}",
                    "Unrealized Gain (₹)": "₹{:,.2f}",
                    "Absolute Return (%)": "{:+.2f}%",
                    "Brokerage Weight": "{:.4f}",
                    "Weighted AUM (₹)": "₹{:,.2f}"
                }),
                use_container_width=True,
                hide_index=True
            )

            csv_data = df_display.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Filtered Schemes Report (CSV)",
                data=csv_data,
                file_name=f"Sahyog_Filtered_Schemes_{datetime.date.today()}.csv",
                mime="text/csv"
            )


        # TAB 2: SCHEME RATIONALIZATION & PORTFOLIO TRIMMING DESK
        with tab2:
            st.markdown("### 🎯 Scheme Rationalization & Portfolio Trimming Desk")
            st.caption("Identify core performing schemes and eliminate long-tail holdings to bring the active fund count under 100.")

            df_s_calc = df_master_schemes.copy()
            df_s_calc['Unrealized Gain (₹)'] = df_s_calc['Current AUM Total (₹)'] - df_s_calc['Invested Total (₹)']
            df_s_calc['Absolute Return (%)'] = (
                (df_s_calc['Unrealized Gain (₹)'] / df_s_calc['Invested Total (₹)'].replace(0, 1)) * 100
            )

            # Categorization Logic
            def categorize_scheme_tier(row):
                aum = row['Current AUM Total (₹)']
                ret = row['Absolute Return (%)']
                if aum >= 10000000 and ret >= 10:
                    return "🟢 Tier 1: Core Performers (AUM ≥ ₹1 Cr & Gain ≥ 10%)"
                elif aum >= 5000000:
                    return "🟡 Tier 2: Maintain & Monitor (AUM ₹50L - ₹1 Cr)"
                elif aum >= 1000000:
                    return "🟠 Tier 3: Review & Rebalance (AUM ₹10L - ₹50L)"
                else:
                    return "🔴 Tier 4: Exit Candidates (AUM < ₹10L)"

            df_s_calc['Rationalization Tier'] = df_s_calc.apply(categorize_scheme_tier, axis=1)

            # Macro Metric Cards
            m1, m2, m3, m4 = st.columns(4)
            t1_t2_count = len(df_s_calc[df_s_calc['Rationalization Tier'].str.contains('Tier 1|Tier 2')])
            t1_t2_aum = df_s_calc[df_s_calc['Rationalization Tier'].str.contains('Tier 1|Tier 2')]['Current AUM Total (₹)'].sum()
            t1_t2_share = (t1_t2_aum / total_aum_base) * 100 if total_aum_base > 0 else 0
            t4_count = len(df_s_calc[df_s_calc['Rationalization Tier'].str.contains('Tier 4')])

            m1.metric("Recommended Core Schemes", f"{t1_t2_count} Funds")
            m2.metric("Core Matrix AUM Share", f"{t1_t2_share:.2f}%")
            m3.metric("Core Matrix AUA Volume", f"₹{t1_t2_aum:,.0f}")
            m4.metric("Exit Candidates (< ₹10L AUM)", f"{t4_count} Funds")

            st.markdown("---")

            # Tier Summary Table
            tier_summary = df_s_calc.groupby('Rationalization Tier').agg(
                Scheme_Count=('Mutual Fund Scheme Name', 'count'),
                Total_AUM=('Current AUM Total (₹)', 'sum'),
                Total_Invested=('Invested Total (₹)', 'sum')
            ).reset_index()

            tier_summary['AUM Contribution %'] = (tier_summary['Total_AUM'] / total_aum_base) * 100
            tier_summary['Unrealized Gain'] = tier_summary['Total_AUM'] - tier_summary['Total_Invested']
            tier_summary['Avg Gain %'] = (tier_summary['Unrealized Gain'] / tier_summary['Total_Invested'].replace(0, 1)) * 100

            st.markdown("#### 📊 Portfolio Tiering Matrix")
            st.dataframe(
                tier_summary.style.format({
                    "Total_AUM": "₹{:,.0f}",
                    "Total_Invested": "₹{:,.0f}",
                    "Unrealized Gain": "₹{:,.0f}",
                    "Scheme_Count": "{:,.0f}",
                    "AUM Contribution %": "{:.2f}%",
                    "Avg Gain %": "{:.2f}%"
                }), use_container_width=True, hide_index=True
            )

            st.markdown("---")
            col_t1, col_t4 = st.columns(2)

            with col_t1:
                st.markdown("#### 🟢 Recommended Core Approved List (Top Performing Funds)")
                core_list = df_s_calc[df_s_calc['Rationalization Tier'].str.contains('Tier 1')].sort_values(by='Current AUM Total (₹)', ascending=False)
                st.dataframe(
                    core_list[['AMC Name', 'Mutual Fund Scheme Name', 'Current AUM Total (₹)', 'Absolute Return (%)']].head(25).style.format({
                        "Current AUM Total (₹)": "₹{:,.0f}",
                        "Absolute Return (%)": "{:.2f}%"
                    }), use_container_width=True, hide_index=True
                )

            with col_t4:
                st.markdown("#### 🔴 Exit / Consolidation List (Micro Holdings to Switch)")
                exit_list = df_s_calc[df_s_calc['Rationalization Tier'].str.contains('Tier 4')].sort_values(by='Current AUM Total (₹)', ascending=True)
                st.dataframe(
                    exit_list[['AMC Name', 'Mutual Fund Scheme Name', 'Current AUM Total (₹)', 'Absolute Return (%)']].head(25).style.format({
                        "Current AUM Total (₹)": "₹{:,.0f}",
                        "Absolute Return (%)": "{:.2f}%"
                    }), use_container_width=True, hide_index=True
                )

        # TAB 3: CODE-WISE BREAKDOWN & COMPARATIVE INSIGHTS (INCLUDING 31241_ot AND 43854_ot)
        with tab3:
            st.markdown("### 📌 Distribution License Code Breakdown & Insights")
            st.caption("Consolidated volume and asset class distribution across primary licenses and supplementary 'Others' holdings (31241_ot & 43854_ot).")
            
            df_s_act_tab3 = df_master_schemes[df_master_schemes['Code'].notna() & (df_master_schemes['Mutual Fund Scheme Name'] != 'Total Book')].copy()
            df_s_act_tab3['Code_Str'] = df_s_act_tab3['Code'].apply(clean_int_string)
            
            total_scheme_aum_tab3 = df_s_act_tab3['Current AUM Total (₹)'].sum()
            
            code_summary_schemes = df_s_act_tab3.groupby('Code_Str').agg(
                Active_Schemes=('Mutual Fund Scheme Name', 'count'),
                Total_AUM=('Current AUM Total (₹)', 'sum'),
                Invested_Total=('Invested Total (₹)', 'sum'),
                Equity_AUM=('Equity AUM (₹)', 'sum'),
                Debt_AUM=('Debt AUM (₹)', 'sum'),
                Hybrid_AUM=('Hybrid AUM (₹)', 'sum')
            ).reset_index()
            
            code_summary_schemes['AUM Contribution %'] = (code_summary_schemes['Total_AUM'] / (total_scheme_aum_tab3 if total_scheme_aum_tab3 > 0 else 1)) * 100
            
            order_map_t3 = {'31241': 1, '31241_ot': 2, '19221': 3, '43854': 4, '43854_ot': 5, '72021': 6}
            code_summary_schemes['Sort_Order'] = code_summary_schemes['Code_Str'].map(lambda x: order_map_t3.get(str(x), 99))
            code_summary_schemes = code_summary_schemes.sort_values(by='Sort_Order').drop(columns=['Sort_Order']).reset_index(drop=True)
            
            st.dataframe(
                code_summary_schemes.rename(columns={
                    'Code_Str': 'Distribution Code',
                    'Active_Schemes': 'Active Schemes',
                    'Total_AUM': 'Current Total AUM (₹)',
                    'Invested_Total': 'Invested Total (₹)',
                    'Equity_AUM': 'Equity AUM (₹)',
                    'Debt_AUM': 'Debt AUM (₹)',
                    'Hybrid_AUM': 'Hybrid AUM (₹)'
                }).style.format({
                    "Current Total AUM (₹)": "₹{:,.2f}",
                    "Invested Total (₹)": "₹{:,.2f}",
                    "Equity AUM (₹)": "₹{:,.2f}",
                    "Debt AUM (₹)": "₹{:,.2f}",
                    "Hybrid AUM (₹)": "₹{:,.2f}",
                    "Active Schemes": "{:,.0f}",
                    "AUM Contribution %": "{:.2f}%"
                }), use_container_width=True, hide_index=True
            )

            code_col1, code_col2 = st.columns(2)
            with code_col1:
                st.markdown("#### AUM Volume Share per Distribution Code (incl. Others)")
                fig_code_pie = px.pie(code_summary_schemes, names='Code_Str', values='Total_AUM', hole=0.4, title="AUM Contribution by License Code", color_discrete_sequence=px.colors.sequential.Plotly3)
                st.plotly_chart(fig_code_pie, use_container_width=True)
            with code_col2:
                st.markdown("#### Asset Class Split per Distribution Code")
                df_code_melt = code_summary_schemes.melt(id_vars=['Code_Str'], value_vars=['Equity_AUM', 'Hybrid_AUM', 'Debt_AUM'], var_name='Asset Class', value_name='AUM')
                fig_code_stack = px.bar(df_code_melt, x='Code_Str', y='AUM', color='Asset Class', title="Asset Mix by Code", barmode='stack', color_discrete_sequence=px.colors.qualitative.Set2)
                st.plotly_chart(fig_code_stack, use_container_width=True)
                
        with tab4:
            st.markdown("### Relationship Executive Performance Review")
            partner_perf = df_master_clients.groupby('Partner/Employee').agg(
                Total_Book_AUM=('Current Total AUM', 'sum'),
                Client_Count=('Investor Name', 'count'),
                Avg_Ticket_Size=('Current Total AUM', 'mean')
            ).reset_index().sort_values(by='Total_Book_AUM', ascending=False)
            
            st.dataframe(
                partner_perf.style.format({
                    "Total_Book_AUM": "₹{:,.2f}",
                    "Client_Count": "{:,.0f}",
                    "Avg_Ticket_Size": "₹{:,.2f}"
                }), use_container_width=True, hide_index=True
            )
            
        with tab5:
            st.markdown("### Asset Placement Tiers Matrix")
            def categorize_client(val):
                if val >= 50000000: return "💎 Ultra HNI (> 5 Cr)"
                elif val >= 10000000: return "🥇 HNI (1 Cr - 5 Cr)"
                elif val >= 2500000: return "🥈 Affluent (25L - 1 Cr)"
                else: return "🥉 Retail (< 25L)"
                
            df_master_clients['Client Tier'] = df_master_clients['Current Total AUM'].apply(categorize_client)
            tier_perf = df_master_clients.groupby('Client Tier').agg(
                Combined_AUM=('Current Total AUM', 'sum'),
                Accounts_Count=('Investor Name', 'count')
            ).reset_index()
            tier_perf['AUM Contribution %'] = (tier_perf['Combined_AUM'] / total_aum_base) * 100
            
            st.dataframe(
                tier_perf.style.format({
                    "Combined_AUM": "₹{:,.0f}",
                    "Accounts_Count": "{:,.0f}",
                    "AUM Contribution %": "{:.2f}%"
                }), use_container_width=True, hide_index=True
            )
            
        with tab6:
            st.markdown("### AMC Portfolio Placement Weights")
            amc_grouped = df_master_schemes.groupby('AMC Name')['Current AUM Total (₹)'].sum().reset_index().sort_values(by='Current AUM Total (₹)', ascending=False).head(10)
            amc_grouped['Share %'] = (amc_grouped['Current AUM Total (₹)'] / total_aum_base) * 100
            st.dataframe(
                amc_grouped.style.format({
                    "Current AUM Total (₹)": "₹{:,.0f}",
                    "Share %": "{:.2f}%"
                }), use_container_width=True, hide_index=True
            )