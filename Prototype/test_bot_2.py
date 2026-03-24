import os
from playwright.sync_api import sync_playwright

PROFILE_DIR = os.path.join(os.getcwd(), "automation_proxy")

# Create a dedicated folder for your downloads
DOWNLOAD_DIR = os.path.join(os.getcwd(), "Teams_Assignments")
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

def run():
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            channel="chrome",
            headless=False,
            accept_downloads=True
        )
        page = context.pages[0]
        
        print("Navigating to Teams...")
        page.goto("https://teams.microsoft.com/v2/")
        
        # Increased initial load time
        page.wait_for_timeout(10000) 
        
        # --- NEW: FORCE MAIN GRID ---
        print("Ensuring we are on the main Teams grid...")
        page.get_by_role("button", name="Teams (Ctrl+Shift+3)").click()
        page.wait_for_timeout(3000)
        
        print("Clicking Class...")
        page.get_by_role("button", name="CE222-COAL-Spring26").click()
        page.wait_for_timeout(3000)

        print("Clicking Channel...")
        page.get_by_role("treeitem", name="General").click()
        page.wait_for_timeout(3000)

        # Click the feed and scroll a bit to ensure assignments are loaded
        page.mouse.click(500, 500)
        page.wait_for_timeout(1000)
        for _ in range(3):
            page.keyboard.press("PageUp")
            # Increased scroll delay
            page.wait_for_timeout(1500) 

        # ---------------------------------------------------------
        # THE SMART LOOP
        # ---------------------------------------------------------
        print("\nScanning for assignments...")
        
        total_assignments = page.locator("[title*='assignment-viewer']").count()
        print(f"Found {total_assignments} assignment(s) in this channel.\n")

        for i in range(total_assignments):
            print(f"--- Processing Assignment {i + 1} of {total_assignments} ---")
            
            # Re-query the locator inside the loop
            current_assignment = page.locator("[title*='assignment-viewer']").nth(i)
            current_assignment.click()
            
            print("  -> Waiting for assignment to load...")
            # Increased iframe load time
            page.wait_for_timeout(7000) 

            try:
                iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame
                options_btn = iframe.locator("[data-test=\"attachment-options-button\"]")
                
                # Increased timeout for finding the options button
                if options_btn.is_visible(timeout=4000):
                    options_btn.click()
                    # Increased menu pop-up delay
                    page.wait_for_timeout(2000) 

                    with page.expect_download() as download_info:
                        iframe.get_by_role("menuitem", name="Download").click()
                    
                    download = download_info.value
                    filename = download.suggested_filename
                    final_path = os.path.join(DOWNLOAD_DIR, filename)

                    if os.path.exists(final_path):
                        print(f"  -> Skipping: '{filename}' is already downloaded.")
                        download.cancel() 
                    else:
                        print(f"  -> Downloading NEW file: '{filename}'")
                        download.save_as(final_path)
                        print(f"  -> Saved to {DOWNLOAD_DIR}")
                else:
                    print("  -> No file attachment found in this assignment.")

            except Exception as e:
                print(f"  -> Error reading this assignment: {e}")

            # --- NEW: RETURN VIA SIDEBAR ---
            print("  -> Returning to channel feed via sidebar click...")
            page.get_by_role("treeitem", name="General").click()
            
            # Increased delay to ensure the feed fully reloads before clicking the next assignment
            page.wait_for_timeout(5000) 

        # ---------------------------------------------------------
        print("\nAll assignments processed!")
        page.wait_for_timeout(3000)
        context.close()

if __name__ == "__main__":
    run()