import os
import time
import json
import datetime
import pandas as pd
import re
from playwright.sync_api import sync_playwright

DOWNLOAD_DIR = os.path.join(os.path.dirname(__file__), "temp_downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
SYNC_LOG_FILE = os.path.join(os.path.dirname(__file__), "sync_log.json")


def get_last_sync_status():
    """Returns a dictionary containing the last sync date and timestamp."""
    if os.path.exists(SYNC_LOG_FILE):
        try:
            with open(SYNC_LOG_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {"last_sync_date": None, "last_sync_time": None}


def record_sync_completed():
    """Saves today's date and timestamp upon successful scraping & sheet updates."""
    now = datetime.datetime.now()
    log_data = {
        "last_sync_date": str(now.date()),
        "last_sync_time": now.strftime("%Y-%m-%d %H:%M:%S")
    }
    with open(SYNC_LOG_FILE, "w") as f:
        json.dump(log_data, f, indent=4)


def fetch_account_reports(code, username, password, portal_type="NJ", headless=True):
    """
    Logs into a single partner portal account and downloads both 
    Client Demographics Master and Scheme Allocation/AUM Master reports into temp_downloads.
    """
    print(f"\n🔑 Starting session for {portal_type} Code: {code} (User: {username})")
    
    client_file = os.path.join(DOWNLOAD_DIR, f"client_{code}.xlsx")
    aum_file = os.path.join(DOWNLOAD_DIR, f"aum_{code}.xlsx")
    
    # Check local fallback files if available in root or temp_downloads
    local_c = f"client_{code}.xlsx"
    local_a = f"aum_{code}.xlsx"
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless, slow_mo=500 if not headless else 0)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        
        try:
            if portal_type == "NJ":
                print("  ↳ Navigating to NJ Partner Desk...")
                page.goto("https://pdesk.njwealth.in/pdesk/login", timeout=90000)
                
                page.get_by_role("textbox", name="* Partner ID").click()
                page.get_by_role("textbox", name="* Partner ID").fill(username)
                page.get_by_role("textbox", name="* Password").click()
                page.get_by_role("textbox", name="* Password").fill(password)
                page.get_by_role("button", name="Log In").click()

                page.wait_for_load_state("networkidle")
                print("  ↳ Login authenticated.")

                # Download Report #1: Scheme Allocation / AUM Master
                print("  ↳ Downloading AUM Scheme Report...")
                page.get_by_role("button", name="Reports").click()
                page.get_by_role("button", name="Mutual Funds", exact=True).click()
                page.locator("div").filter(has_text=re.compile(r"^MF AUM ReportView MF AUM with different view and filter options\.$")).nth(2).click()
                page.get_by_text("Group Wise").click()
                page.get_by_text("Scheme Wise").click()
                #page.get_by_role("checkbox", name="Invested Amount").check()
                #page.get_by_role("checkbox", name="Difference In AUM").check()
                page.get_by_role("button", name="APPLY").click()
                with page.expect_download(timeout=45000) as download_info_a:
                    page.get_by_role("button", name="file-excel Excel").click()
                download_info_a.value.save_as(aum_file)
                print(f"  ✅ Saved: {aum_file}")

                # Download Report #2: Client Demographics Master
                print("  ↳ Downloading Client Master Report...")
                with page.expect_popup() as page1_info:
                    page.get_by_role("button", name="CRM").click()
                page1 = page1_info.value
                page1.get_by_text("groups").click()
                page1.get_by_text("Datamodel2closemore_horiz").click()
                page1.get_by_text("file_export").click()
                with page1.expect_download(timeout=45000) as download_info_c:
                    page1.get_by_role("button").first.click()

                page1.get_by_text("logout", exact=True).click()
                page1.close()
                page.get_by_role("button", name="Logout").click()
                page.get_by_role("button", name="Yes").click()
                download_info_c.value.save_as(client_file)
                print(f"  ✅ Saved: {client_file}")

        except Exception as e:
            print(f"❌ Error scraping code {code}: {e}")
            client_file, aum_file = None, None
        finally:
            context.close()
            browser.close()
            
    # Fallback to root files if live download was bypassed or executed locally
    if not os.path.exists(client_file) and os.path.exists(local_c):
        client_file = local_c
    if not os.path.exists(aum_file) and os.path.exists(local_a):
        aum_file = local_a

    return client_file if os.path.exists(str(client_file)) else None, aum_file if os.path.exists(str(aum_file)) else None


def fetch_all_portal_sheets(headless=True):
    """
    [STEP 1]: Downloads portal Excel files for all distribution codes into temp_downloads.
    Does NOT modify Live_Wealth_Manager_Book.xlsx. Returns status dict per code.
    """
    import streamlit as st
    
    results = {}
    default_codes = ["31241", "19221", "43854", "72021"]
    
    if "portal_credentials" in st.secrets:
        creds_items = st.secrets["portal_credentials"].items()
    else:
        creds_items = [(f"nj_{c}", {"code": c, "username": f"user_{c}", "password": "pwd", "portal_type": "NJ"}) for c in default_codes]

    for key, creds in creds_items:
        code = str(creds["code"]).replace('.0', '').strip()
        user = creds.get("username", "")
        pwd = creds.get("password", "")
        p_type = creds.get("portal_type", "NJ")
        
        c_file, a_file = fetch_account_reports(code, user, pwd, portal_type=p_type, headless=headless)
        
        c_ok = c_file is not None and os.path.exists(c_file)
        a_ok = a_file is not None and os.path.exists(a_file)
        
        results[code] = {
            "client_ok": c_ok,
            "aum_ok": a_ok,
            "client_file": c_file,
            "aum_file": a_file
        }
        
    return results


def scan_downloaded_feed_files():
    """
    Scans temp_downloads/ and current root directory for available client_*.xlsx and aum_*.xlsx files.
    Returns two dictionaries: client_files_dict, aum_files_dict.
    """
    check_dirs = [
        DOWNLOAD_DIR,
        os.path.dirname(__file__) if '__file__' in locals() else '.'
    ]
    
    client_dict = {}
    aum_dict = {}
    default_codes = ["31241", "19221", "43854", "72021"]
    
    # Check default codes in temp_downloads first
    for code in default_codes:
        c_fn = f"client_{code}.xlsx"
        a_fn = f"aum_{code}.xlsx"
        
        for cdir in check_dirs:
            if not os.path.exists(cdir): continue
            cp = os.path.join(cdir, c_fn)
            ap = os.path.join(cdir, a_fn)
            
            if code not in client_dict and os.path.exists(cp):
                client_dict[code] = cp
            if code not in aum_dict and os.path.exists(ap):
                aum_dict[code] = ap

    # Scan for all other matching client_*.xlsx and aum_*.xlsx files (e.g. Others, ARN, etc.)
    for cdir in check_dirs:
        if not os.path.exists(cdir): continue
        for fname in sorted(os.listdir(cdir)):
            if not fname.endswith(".xlsx") or fname.startswith("~$") or "Live_Wealth_Manager_Book" in fname:
                continue
            fpath = os.path.join(cdir, fname)
            f_lower = fname.lower()
            
            if f_lower.startswith("client_"):
                key = fname.replace(".xlsx", "").replace("client_", "")
                if key not in client_dict:
                    client_dict[key] = fpath
            elif f_lower.startswith("aum_"):
                key = fname.replace(".xlsx", "").replace("aum_", "")
                if key not in aum_dict:
                    aum_dict[key] = fpath

    return client_dict, aum_dict


def run_full_portal_sync(headless=True):
    """
    Convenience function that fetches all sheets and returns dicts of DataFrames for backward compatibility.
    """
    fetch_all_portal_sheets(headless=headless)
    c_files, a_files = scan_downloaded_feed_files()
    
    c_dfs = {}
    a_dfs = {}
    
    for code, fpath in c_files.items():
        try:
            df_c = pd.read_excel(fpath)
            df_c["Code"] = float(code) if code.isdigit() else code
            c_dfs[code] = df_c
        except Exception as e:
            print(f"Error loading client file {fpath}: {e}")
            
    for code, fpath in a_files.items():
        try:
            df_a = pd.read_excel(fpath)
            df_a["Code"] = float(code) if code.isdigit() else code
            a_dfs[code] = df_a
        except Exception as e:
            print(f"Error loading AUM file {fpath}: {e}")
            
    return c_dfs, a_dfs


if __name__ == "__main__":
    print("=" * 65)
    print("🚀 RUNNING PORTAL SCRAPER & SCANNER TEST (v10.0)")
    print("=" * 65)
    c_dict, a_dict = scan_downloaded_feed_files()
    print(f"Client Feeds Discovered: {len(c_dict)} | AUM Feeds Discovered: {len(a_dict)}")
    print("Client Files:", c_dict)
    print("AUM Files:", a_dict)
    print("=" * 65)