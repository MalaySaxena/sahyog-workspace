import os
import time
import re
from playwright.sync_api import sync_playwright

# Setup temporary download directory
DOWNLOAD_DIR = os.path.join(os.path.dirname(__file__), "temp_downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# -----------------------------------------------------------------------------
# TEST CREDENTIALS & TARGET URL (FILL YOUR TEST DETAILS HERE)
# -----------------------------------------------------------------------------
TEST_CODE = "19221"                             # Branch code being tested
TEST_USERNAME = "sahyogm"         # Replace with test username
TEST_PASSWORD = "sF##2025"         # Replace with test password
TARGET_LOGIN_URL = "https://pdesk.njwealth.in/"  # Replace with target login URL

def run_isolated_scraper_test():
    print("=" * 60)
    print(f"🚀 STARTING ISOLATED SCRAPER TEST FOR CODE: {TEST_CODE}")
    print("=" * 60)
    
    save_file_path_aum = os.path.join(DOWNLOAD_DIR, f"aum_{TEST_CODE}.xlsx")
    save_file_path_client = os.path.join(DOWNLOAD_DIR, f"client_{TEST_CODE}.xlsx")
    
    with sync_playwright() as p:
        # 1. Launch browser in VISIBLE mode with slow_mo delay for visual debugging
        print("\n[1/5] Launching browser window...")
        browser = p.chromium.launch(headless=False, slow_mo=1000)  # 1 second delay between actions
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        
        try:
            # 2. Navigate to Login Page
            print(f"[2/5] Navigating to {TARGET_LOGIN_URL}...")
            #page.goto(TARGET_LOGIN_URL, timeout=30000)
            page.goto("https://pdesk.njwealth.in/pdesk/login", timeout=90000)
            # -----------------------------------------------------------------
            # 3. PASTE YOUR GENERATED PLAYWRIGHT ACTIONS HERE
            # -----------------------------------------------------------------
            print("[3/5] Entering credentials and submitting login...")
            
            # --- REPLACE THE LINES BELOW WITH YOUR GENERATED CODEGEN LINES ---
            # Example generated code structure:
            # page.get_by_placeholder("Partner ID").fill(TEST_USERNAME)
            # page.get_by_placeholder("Password").fill(TEST_PASSWORD)
            # page.get_by_role("button", name="Sign In").click()
            # -----------------------------------------------------------------
            
            page.get_by_role("textbox", name="* Partner ID").click()
            page.get_by_role("textbox", name="* Partner ID").fill("sahyogm")
            page.get_by_role("textbox", name="* Password").click()
            page.get_by_role("textbox", name="* Password").fill("sF##2025")
            page.get_by_role("button", name="Log In").click()
            
            # Wait for network activity to settle after login
            page.wait_for_load_state("networkidle")
            print("  ↳ Login submitted successfully.")

            # 4. Navigate to Report Page & Trigger Download
            print("[4/5] Navigating to report download section...")
            
            # --- PASTE YOUR REPORT NAVIGATION / CLICK DOWNLOAD LINES HERE ---
            # page.goto("https://pdesk.njwealth.in/reports/client-aum")
            page.get_by_role("button", name="Reports").click()
            page.get_by_role("button", name="Mutual Funds", exact=True).click()
            page.locator("div").filter(has_text=re.compile(r"^MF AUM ReportView MF AUM with different view and filter options\.$")).nth(2).click()
            page.get_by_text("Group Wise").click()
            page.get_by_text("Scheme Wise").click()
            page.get_by_role("checkbox", name="Invested Amount").check()
            page.get_by_role("checkbox", name="Difference In AUM").check()
            page.get_by_role("button", name="APPLY").click()
            with page.expect_download() as download_info:
                page.get_by_role("button", name="file-excel Excel").click()
            download = download_info.value
            
            with page.expect_popup() as page1_info:
                page.get_by_role("button", name="CRM").click()
            page1 = page1_info.value
            #page1.goto("https://pdesk.njwealth.in/crm/dashboard")
            page1.get_by_text("groups").click()
            page1.get_by_text("Datamodel2closemore_horiz").click()
            page1.get_by_text("file_export").click()
            with page1.expect_download() as download1_info:
                page1.get_by_role("button").first.click()
            download1 = download1_info.value
            page1.get_by_text("logout", exact=True).click()
            page1.close()
            page.get_by_role("button", name="Logout").click()
            page.get_by_role("button", name="Yes").click()
            #with page.expect_download(timeout=30000) as download_info:
                # Trigger click action on the Export button
                # page.get_by_role("button", name="Export to Excel").click()
            #    pass
                
            #download = download_info.value
            download.save_as(save_file_path_aum)
            download1.save_as(save_file_path_client)
            print(f"  ↳ Download intercepted successfully!")

            # 5. Verification Check
            print("\n[5/5] Verifying downloaded output file...")
            if os.path.exists(save_file_path_aum):
                #file_size_kb = os.path.getsize(save_file_path_aum) / 1024
                print(f"✅ SUCCESS! File saved at: {save_file_path_aum}")
                #print(f"   File Size: {file_size_kb:.2f} KB")
            else:
                print("❌ ERROR: File download was expected but not found on disk.")

            if os.path.exists(save_file_path_client):
                #file_size_kb = os.path.getsize(save_file_path_aum) / 1024
                print(f"✅ SUCCESS! File saved at: {save_file_path_client}")
                #print(f"   File Size: {file_size_kb:.2f} KB")
            else:
                print("❌ ERROR: File download was expected but not found on disk.")

        except Exception as e:
            print(f"\n❌ TEST FAILED WITH EXCEPTION: {e}")
            
        finally:
            print("\n[Closing Browser Context]")
            time.sleep(2)
            context.close()
            browser.close()

if __name__ == "__main__":
    run_isolated_scraper_test()