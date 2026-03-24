import os
from playwright.sync_api import sync_playwright

# --- CONFIGURATION ---
PROFILE_DIR = os.path.join(os.getcwd(), "automation_proxy")
DOWNLOAD_DIR = os.path.join(os.getcwd(), "Teams_Assignments")

# Ensure the download directory exists
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

def run():
    with sync_playwright() as p:
        # Launch Chrome using your saved session
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            channel="chrome",
            headless=False,
            accept_downloads=True
        )
        page = context.pages[0]
        
        print("Navigating to Teams...")
        page.goto("https://teams.microsoft.com/v2/")
        
        # Generous wait time for the initial Teams dashboard to load completely
        page.wait_for_timeout(10000) 
        
        # --- 1. OPEN ASSIGNMENTS TAB ---
        print("Clicking Assignments tab...")
        try:
            page.get_by_role("button", name="Assignments (Ctrl+Shift+4)").click()
        except:
            # Fallback just in case the shortcut text varies
            page.locator("[data-tid='app-bar-edu-assignments']").first.click()
            
        # Give the iframe plenty of time to fetch the assignments list from the server
        page.wait_for_timeout(8000) 

        # --- 2. TARGET THE IFRAME ---
        print("Targeting the Assignments iFrame...")
        iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame

        # THE FIX: Append :visible directly to the CSS selector to ignore hidden <style> tags
        target_card_id = "[id*='CardHeader__headerEDUASSIGN']:visible"
        
        # Grab all the visible assignment cards
        assignment_cards = iframe.locator(target_card_id)
        total_assignments = assignment_cards.count()
        
        print(f"\nFound {total_assignments} clickable upcoming assignment(s)!\n")

        # --- 3. TRAVERSE ASSIGNMENTS ---
        for i in range(total_assignments):
            print(f"--- Processing Assignment {i + 1} of {total_assignments} ---")
            
            # We must re-fetch the iframe and card inside the loop
            # because hitting the "Back" button forces the page to refresh
            iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame
            current_card = iframe.locator(target_card_id).nth(i)
            current_card.click()
            
            print("  -> Waiting for assignment details to load...")
            page.wait_for_timeout(6000) 

            try:
                # --- 4. CHECK FOR REFERENCE MATERIALS ---
                # Find all "3-dots" attachment menus on the page
                attachment_menus = iframe.locator("[data-test=\"attachment-options-button\"]")
                attachment_count = attachment_menus.count()

                if attachment_count > 0:
                    print(f"  -> Found {attachment_count} reference file(s).")
                    
                    for j in range(attachment_count):
                        print(f"  -> Opening 3-dots menu for file {j+1}...")
                        attachment_menus.nth(j).click()
                        
                        # Wait for the dropdown menu to visually appear
                        page.wait_for_timeout(2000)

                        print("  -> Triggering download...")
                        with page.expect_download() as download_info:
                            # Use a try/except to handle the two ways Teams labels the download button
                            try:
                                iframe.get_by_role("menuitem", name="Download").click(timeout=3000)
                            except:
                                iframe.get_by_text("Download").first.click(timeout=3000)
                        
                        # Capture file info
                        download = download_info.value
                        filename = download.suggested_filename
                        final_path = os.path.join(DOWNLOAD_DIR, filename)

                        # --- 5. THE DUPLICATE CHECK ---
                        if os.path.exists(final_path):
                            print(f"    => Skipping: '{filename}' already exists locally.")
                            download.cancel() # Delete the temporary ghost file
                        else:
                            print(f"    => Downloading NEW file: '{filename}'")
                            download.save_as(final_path)
                            print(f"    => SUCCESS! Saved to Teams_Assignments folder.")
                            
                        # Quick breather between multiple files
                        page.wait_for_timeout(2000) 
                else:
                    print("  -> No reference materials attached to this assignment.")

            except Exception as e:
                print(f"  -> Error reading assignment files: {e}")

            # --- 6. USE IFRAME BACK BUTTON ---
            print("  -> Clicking back button to return to the list...")
            try:
                iframe.locator("[data-test=\"back-button\"]").first.click()
            except:
                print("  -> Back button failed, attempting browser back...")
                page.go_back()
            
            # Very important delay so the list fully re-renders before the next loop starts
            page.wait_for_timeout(5000) 

        print("\n=== COMPLETE: All upcoming assignments processed! ===")
        page.wait_for_timeout(3000)
        context.close()

if __name__ == "__main__":
    run()