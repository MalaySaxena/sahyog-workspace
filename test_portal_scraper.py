import os
import time
import pandas as pd
import re
from playwright.sync_api import sync_playwright

# Download directory setup
DOWNLOAD_DIR = os.path.join(os.path.dirname(__file__), "temp_downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

def fetch_account_reports(code, username, password, portal_type="NJ", headless=True):
    """
    Logs into a single account and downloads both Client Master and AUM/Scheme reports.
    """
    print(f"\n🔑 Starting session for {portal_type} Code: {code} (User: {username})")
    
    client_file = os.path.join(DOWNLOAD_DIR, f"client_{code}.xlsx")
    aum_file = os.path.join(DOWNLOAD_DIR, f"aum_{code}.xlsx")
    
    with sync_playwright() as p:
        # Launch browser (headless=False allows watching visual execution during test)
        browser = p.chromium.launch(headless=headless, slow_mo=800)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        
        try:
            if portal_type == "NJ":
                print("  ↳ Navigating to NJ Partner Desk...")
                page.goto("https://pdesk.njwealth.in/pdesk/login", timeout=90000)
                
                # --- PASTE YOUR GENERATED PLAYWRIGHT CODEGEN LINES HERE ---
                # Example:
                # page.get_by_placeholder("Partner ID").fill(username)
                # page.get_by_placeholder("Password").fill(password)
                # page.get_by_role("button", name="Sign In").click()
                # ---------------------------------------------------------
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
                page.get_by_role("checkbox", name="Invested Amount").check()
                page.get_by_role("checkbox", name="Difference In AUM").check()
                page.get_by_role("button", name="APPLY").click()
                with page.expect_download() as download_info_a:
                    page.get_by_role("button", name="file-excel Excel").click()
                download_info_a.value.save_as(aum_file)
                print(f"  ✅ Saved: {aum_file}")

                # Download Report #2: Client Demographics Master
                print("  ↳ Downloading Client Master Report...")
                # page.goto("https://pdesk.njwealth.in/reports/client-master")
                with page.expect_popup() as page1_info:
                    page.get_by_role("button", name="CRM").click()
                page1 = page1_info.value
                #page1.goto("https://pdesk.njwealth.in/crm/dashboard")
                page1.get_by_text("groups").click()
                page1.get_by_text("Datamodel2closemore_horiz").click()
                page1.get_by_text("file_export").click()
                with page1.expect_download() as download_info_c:
                    page1.get_by_role("button").first.click()

                page1.get_by_text("logout", exact=True).click()
                page1.close()
                page.get_by_role("button", name="Logout").click()
                page.get_by_role("button", name="Yes").click()
                download_info_c.value.save_as(client_file)
                print(f"  ✅ Saved: {client_file}")

                

        except Exception as e:
            print(f"  ❌ Error scraping code {code}: {e}")
            client_file, aum_file = None, None
        finally:
            context.close()
            browser.close()
            
    return client_file, aum_file

def run_full_portal_sync(headless=True):
    """
    App-facing pipeline that loops through st.secrets credentials.
    """
    import streamlit as st
    
    if "portal_credentials" not in st.secrets:
        st.error("Missing [portal_credentials] in .streamlit/secrets.toml")
        return pd.DataFrame(), pd.DataFrame()

    all_client_dfs = []
    all_aum_dfs = []
    
    for key, creds in st.secrets["portal_credentials"].items():
        code = creds["code"]
        user = creds["username"]
        pwd = creds["password"]
        p_type = creds.get("portal_type", "NJ")
        
        c_file, a_file = fetch_account_reports(code, user, pwd, portal_type=p_type, headless=headless)
        
        if c_file and os.path.exists(c_file):
            df_c = pd.read_excel(c_file)
            df_c["Code"] = f"{p_type}-{code}"
            all_client_dfs.append(df_c)
            
        if a_file and os.path.exists(a_file):
            df_a = pd.read_excel(a_file)
            df_a["Code"] = f"{p_type}-{code}"
            all_aum_dfs.append(df_a)
            
    df_clients_final = pd.concat(all_client_dfs, ignore_index=True) if all_client_dfs else pd.DataFrame()
    df_schemes_final = pd.concat(all_aum_dfs, ignore_index=True) if all_aum_dfs else pd.DataFrame()
    
    return df_clients_final, df_schemes_final


# -----------------------------------------------------------------------------
# ISOLATED TEST RUNNER BLOCK
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    import streamlit as st
    
    print("=" * 65)
    print("🚀 RUNNING ISOLATED PORTAL SCRAPER TEST")
    print("=" * 65)
    
    # Run scraper in visible mode (headless=False) to watch execution
    df_clients, df_schemes = run_full_portal_sync(headless=False)
    
    print("\n" + "=" * 65)
    print("📊 TEST SUMMARY & DATA VALIDATION")
    print("=" * 65)
    print(f"Client Report Records Extracted: {len(df_clients)} rows")
    print(f"Scheme AUM Records Extracted:   {len(df_schemes)} rows")
    
    if not df_clients.empty:
        print("\nClient Report Preview:")
        print(df_clients.head(2))
    if not df_schemes.empty:
        print("\nScheme Report Preview:")
        print(df_schemes.head(2))
    print("=" * 65)