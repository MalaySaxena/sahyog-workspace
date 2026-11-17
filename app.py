import streamlit as st
import pandas as pd
import datetime
import os
import openpyxl
import plotly.express as px
import json
import re
import numpy as np
import io
import subprocess
import zipfile
import uuid
import threading

FILE_IO_LOCK = threading.Lock()


def ensure_playwright_browsers():
    """Installs Chromium binaries if running in Streamlit Cloud."""
    try:
        subprocess.run(["playwright", "install", "chromium"], check=True)
    except Exception as e:
        print(f"Playwright browser installation notice: {e}")


# Call before launching Playwright browser sessions
ensure_playwright_browsers()

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


def format_date_val(d):
    """Converts Date object or tuple to YYYY-MM-DD string, returning empty string if blank/None."""
    if d is None:
        return ""
    if isinstance(d, (list, tuple)):
        if len(d) == 0 or d[0] is None:
            return ""
        d = d[0]
    if hasattr(d, "strftime"):
        return d.strftime("%Y-%m-%d")
    s = str(d).strip()
    if s.lower() in ["none", "nat", "nan", ""]:
        return ""
    return s


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

    }
    return clean_username in default_users and default_users[clean_username] == password


def render_login_screen():
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("""
            <div style="background-color: #f8f9fa; padding: 30px; border-radius: 12px; border-top: 6px solid #1F4E79; box-shadow: 0 4px 12px rgba(0,0,0,0.1); margin-top: 40px;">
                <h2 style="color: #1F4E79; text-align: center; margin-bottom: 5px;">🏦 Sahyog Financials</h2>
                <p style="text-align: center; color: #6c757d; font-size: 14px;">Enterprise Workspace Console Login (v10.0)</p>
                <hr style="margin-bottom: 20px;">
            </div>
        """, unsafe_allow_html=True)

        if "login_attempts" not in st.session_state:
            st.session_state["login_attempts"] = 0
        if "lockout_until" not in st.session_state:
            st.session_state["lockout_until"] = None

        now = datetime.datetime.now()
        if st.session_state["lockout_until"]:
            if now < st.session_state["lockout_until"]:
                remaining_sec = int((st.session_state["lockout_until"] - now).total_seconds())
                st.error(f"🚫 **Access Temporarily Locked:** Too many failed login attempts. Please wait {remaining_sec}s before retrying.")
                st.stop()
            else:
                st.session_state["lockout_until"] = None
                st.session_state["login_attempts"] = 0

        with st.form("login_form"):
            username_input = st.text_input("👤 Username")
            password_input = st.text_input("🔑 Password", type="password")
            submit_login = st.form_submit_button("Sign In to Workspace", use_container_width=True)
            
            if submit_login:
                if not username_input or not password_input:
                    st.warning("⚠️ Please enter both Username and Password.")
                elif check_credentials(username_input, password_input.strip()):
                    st.session_state["authenticated"] = True
                    st.session_state["username"] = username_input.strip()
                    st.session_state["login_attempts"] = 0
                    st.session_state["lockout_until"] = None
                    st.success("✅ Login successful! Loading workspace...")
                    st.rerun()
                else:
                    st.session_state["login_attempts"] += 1
                    attempts_left = max(0, 5 - st.session_state["login_attempts"])
                    if st.session_state["login_attempts"] >= 5:
                        st.session_state["lockout_until"] = datetime.datetime.now() + datetime.timedelta(minutes=5)
                        st.error("🚫 **Maximum login attempts (5) reached.** Account locked for 5 minutes to prevent unauthorized access.")
                        st.rerun()
                    else:
                        st.error(f"❌ Invalid Username or Password. ({attempts_left} attempt(s) remaining before 5-minute lockout).")


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

# PRIMARY DESIGNATED SYNC ADMINISTRATOR USERS LIST
PRIMARY_SYNC_USERS = ["admin", "malay"]


def is_valid_xlsx(filepath):
    """Validates that an .xlsx file exists, is non-empty, passes ZIP CRC-32 checks, and opens in openpyxl."""
    if not filepath or not os.path.exists(filepath) or os.path.getsize(filepath) == 0:
        return False
    try:
        with zipfile.ZipFile(filepath, "r") as zf:
            if zf.testzip() is not None:
                return False
        wb = openpyxl.load_workbook(filepath, read_only=True)
        _ = wb.sheetnames
        wb.close()
        return True
    except Exception:
        return False


def repair_corrupted_xlsx(corrupt_path, repaired_path):
    """
    Repairs an .xlsx file suffering from Bad CRC-32 on metadata XMLs (e.g. docProps/core.xml)
    by copying intact ZIP entries and regenerating valid metadata XMLs.
    """
    default_core_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        b'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        b'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        b'<dc:creator>Sahyog Financials</dc:creator>'
        b'</cp:coreProperties>'
    )
    default_app_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
        b'<Application>Microsoft Excel</Application>'
        b'</Properties>'
    )
    try:
        with zipfile.ZipFile(corrupt_path, "r") as zin, zipfile.ZipFile(repaired_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                try:
                    data = zin.read(item.filename)
                    zout.writestr(item.filename, data)
                except zipfile.BadZipFile:
                    if item.filename == "docProps/core.xml":
                        zout.writestr(item.filename, default_core_xml)
                    elif item.filename == "docProps/app.xml":
                        zout.writestr(item.filename, default_app_xml)
                    else:
                        return False
        return is_valid_xlsx(repaired_path)
    except Exception:
        return False


def save_workbook_atomic(wb, target_path):
    """Saves an openpyxl Workbook atomically via a temporary .xlsx file to prevent ZIP/CRC corruption."""
    tmp_path = f"{target_path}.{uuid.uuid4().hex}.tmp.xlsx"
    try:
        wb.save(tmp_path)
        wb.close()
        if not is_valid_xlsx(tmp_path):
            raise zipfile.BadZipFile("Generated workbook failed ZIP/CRC integrity check.")
        os.replace(tmp_path, target_path)
    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass


def is_gdrive_configured():
    try:
        return ("gcp_service_account" in st.secrets) and ("gdrive" in st.secrets and "file_id" in st.secrets["gdrive"])
    except Exception:
        return False


def get_gdrive_service():
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    
    scopes = ["https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return build("drive", "v3", credentials=creds)


def sync_gdrive_to_local():
    """Downloads latest Live_Wealth_Manager_Book.xlsx from Google Drive with ZIP/CRC validation & auto-repair."""
    if is_gdrive_configured():
        tmp_dl = f"{EXCEL_FILE}.dl.{uuid.uuid4().hex}.tmp.xlsx"
        tmp_rep = f"{EXCEL_FILE}.rep.{uuid.uuid4().hex}.tmp.xlsx"
        try:
            from googleapiclient.http import MediaIoBaseDownload
            service = get_gdrive_service()
            file_id = st.secrets["gdrive"]["file_id"]
            
            request = service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
                
            fh.seek(0)
            with open(tmp_dl, "wb") as f:
                f.write(fh.read())

            with FILE_IO_LOCK:
                if is_valid_xlsx(tmp_dl):
                    os.replace(tmp_dl, EXCEL_FILE)
                    return True
                elif repair_corrupted_xlsx(tmp_dl, tmp_rep):
                    os.replace(tmp_rep, EXCEL_FILE)
                    return True
                else:
                    st.sidebar.warning("⚠️ Downloaded Google Drive workbook failed CRC check; preserving local copy.")
                    return False
        except Exception as e:
            st.sidebar.error(f"⚠️ GDrive Sync Download Error: {e}")
            return False
        finally:
            for p in (tmp_dl, tmp_rep):
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except Exception:
                        pass
    return False


def sync_local_to_gdrive():
    """Uploads updated local Live_Wealth_Manager_Book.xlsx back to Google Drive only if ZIP/CRC valid."""
    if is_gdrive_configured():
        if not is_valid_xlsx(EXCEL_FILE):
            st.error("⚠️ Skipped Google Drive upload because local workbook failed ZIP integrity verification.")
            return False
        try:
            from googleapiclient.http import MediaFileUpload
            service = get_gdrive_service()
            file_id = st.secrets["gdrive"]["file_id"]
            
            with FILE_IO_LOCK:
                media = MediaFileUpload(
                    EXCEL_FILE,
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    resumable=True
                )
                service.files().update(
                    fileId=file_id,
                    media_body=media
                ).execute()
            return True
        except Exception as e:
            st.error(f"⚠️ GDrive Sync Upload Error: {e}")
            return False
    return False


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


def create_clean_master_workbook(target_path):
    """Creates a clean, valid master workbook containing all required worksheet tabs."""
    wb = openpyxl.Workbook()
    ws_c = wb.active
    ws_c.title = "2. Client AUM Master"
    ws_c.append(CLIENT_MASTER_SCHEMA)

    ws_s = wb.create_sheet("3. Scheme Allocation Master")
    ws_s.append(SCHEME_MASTER_SCHEMA)

    ws_l = wb.create_sheet("4. Transaction Ledger")
    ws_l.append(SCHEMA_COLUMNS)

    save_workbook_atomic(wb, target_path)


def init_excel_ledger():
    """Ensures EXCEL_FILE exists and is a valid .xlsx archive before reading/writing."""
    if not is_valid_xlsx(EXCEL_FILE):
        sync_gdrive_to_local()

    with FILE_IO_LOCK:
        if os.path.exists(EXCEL_FILE) and not is_valid_xlsx(EXCEL_FILE):
            tmp_rep = f"{EXCEL_FILE}.rep.{uuid.uuid4().hex}.tmp.xlsx"
            if repair_corrupted_xlsx(EXCEL_FILE, tmp_rep):
                os.replace(tmp_rep, EXCEL_FILE)
            else:
                if os.path.exists(tmp_rep):
                    try:
                        os.remove(tmp_rep)
                    except Exception:
                        pass
                backup_path = f"{EXCEL_FILE}.corrupt.{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
                try:
                    os.replace(EXCEL_FILE, backup_path)
                except Exception:
                    try:
                        os.remove(EXCEL_FILE)
                    except Exception:
                        pass

        if os.path.exists(EXCEL_FILE):
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                if "4. Transaction Ledger" not in wb.sheetnames:
                    ws = wb.create_sheet("4. Transaction Ledger")
                    ws.append(SCHEMA_COLUMNS)
                    save_workbook_atomic(wb, EXCEL_FILE)
                else:
                    wb.close()
            except Exception:
                create_clean_master_workbook(EXCEL_FILE)
        else:
            create_clean_master_workbook(EXCEL_FILE)


init_excel_ledger()


@st.cache_data(ttl=5)
def load_all_data():
    sync_gdrive_to_local()
    init_excel_ledger()
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
                if is_valid_xlsx(EXCEL_FILE):
                    df_l = pd.read_excel(EXCEL_FILE, sheet_name="4. Transaction Ledger")
    else:
        if is_valid_xlsx(EXCEL_FILE):
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


def save_transaction_records(records_list):
    """Saves single or batch transaction records atomically to GSheets and local/GDrive Excel."""
    if not records_list:
        return False
    if isinstance(records_list, dict):
        records_list = [records_list]

    saved_to_gsheets = False

    if is_gsheets_configured():
        ws = get_gspread_worksheet("4. Transaction Ledger")
        if ws:
            try:
                first_row = ws.row_values(1)
                gs_cols = first_row if (first_row and len(first_row) >= 10) else SCHEMA_COLUMNS
                all_rows = [[rec.get(col, "") for col in gs_cols] for rec in records_list]
                ws.append_rows(all_rows)
                saved_to_gsheets = True
            except Exception as e:
                st.error(f"❌ Failed to write to Google Sheets: {e}")

    try:
        init_excel_ledger()
        with FILE_IO_LOCK:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            if "4. Transaction Ledger" not in wb.sheetnames:
                ws_ex = wb.create_sheet("4. Transaction Ledger")
                ws_ex.append(SCHEMA_COLUMNS)
                ex_cols = SCHEMA_COLUMNS
            else:
                ws_ex = wb["4. Transaction Ledger"]
                raw_h = [cell.value for cell in ws_ex[1]]
                ex_cols = [str(c) if c is not None else "" for c in raw_h] if (raw_h and len(raw_h) >= 10) else SCHEMA_COLUMNS

            for rec in records_list:
                ws_ex.append([rec.get(col, "") for col in ex_cols])
            save_workbook_atomic(wb, EXCEL_FILE)
        sync_local_to_gdrive()
    except Exception as e:
        st.error(f"❌ Local Excel Error: {e}")

    st.cache_data.clear()
    return saved_to_gsheets


def save_transaction_record(record):
    return save_transaction_records([record])


def overwrite_ledger_database(df_updated):
    saved_to_gsheets = False
    save_cols = [c for c in df_updated.columns if c in SCHEMA_COLUMNS or c == "Units"]
    if not save_cols:
        save_cols = SCHEMA_COLUMNS
    df_updated = df_updated[save_cols]
    
    if is_gsheets_configured():
        ws = get_gspread_worksheet("4. Transaction Ledger")
        if ws:
            try:
                ws.clear()
                ws.update([save_cols] + df_updated.fillna("").values.tolist())
                saved_to_gsheets = True
            except Exception as e:
                st.error(f"❌ Failed to update Google Sheets ledger: {e}")
                
    try:
        init_excel_ledger()
        with FILE_IO_LOCK:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            if "4. Transaction Ledger" in wb.sheetnames:
                del wb["4. Transaction Ledger"]
            ws_ex = wb.create_sheet("4. Transaction Ledger")
            ws_ex.append(save_cols)
            for row in df_updated.fillna("").values.tolist():
                ws_ex.append(row)
            save_workbook_atomic(wb, EXCEL_FILE)
        sync_local_to_gdrive()
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

    check_dirs = [".", os.path.join(os.path.dirname(__file__), "temp_downloads")]
    additional_other_files = [
        ("31241_ot", ["client_31241_Others.xlsx", "client_31241_others.xlsx", "31241_Others.xlsx", "31241_others.xlsx"]),
        ("43854_ot", ["client_43854_Others.xlsx", "client_43854_others.xlsx", "43854_Others.xlsx", "43854_others.xlsx"])
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
                                parsed_rows = parse_client_dataframe(df_sheet, target_code)
                                all_rows.extend(parsed_rows)
                                processed_file_sheets.add(key)
                    except Exception as e:
                        print(f"Error reading additional client file {fpath}: {e}")

    df_res = pd.DataFrame(all_rows, columns=CLIENT_MASTER_SCHEMA)

    if not df_res.empty:
        code_order = {31241.0: 1, '31241_ot': 2, 19221.0: 3, 43854.0: 4, '43854_ot': 5, 72021.0: 6}

        def get_client_sort_key(c):
            if c in code_order:
                return code_order[c]
            try:
                f_c = float(c)
                if f_c in code_order:
                    return code_order[f_c]
            except Exception:
                pass
            return 99

        df_res['Sort_Order'] = df_res['Code'].map(get_client_sort_key)
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
            if not c_code or c_code.
