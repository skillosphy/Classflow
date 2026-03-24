import os
import re
import json
import pyperclip
import pyautogui
import subprocess
import time
from playwright.sync_api import sync_playwright

# --- CONFIGURATION ---
PROFILE_DIR = os.path.join(os.environ['LOCALAPPDATA'], "TeamsBot_Proxy")
USER_DESKTOP = os.path.join(os.environ['USERPROFILE'], 'Desktop')
DOWNLOAD_DIR = os.path.join(USER_DESKTOP, "Assignments")
HISTORY_FILE = os.path.join(os.environ['LOCALAPPDATA'], "TeamsBot_History.json")

if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

# --- COURSE MAPPING ---
COURSE_MAP = {
    "CS224": "FLAT",
    "CS272": "HCI",
    "CE222": "COAL",
    "CS232": "DBMSLab",
    "ES111": "Stats"
}

def load_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_history(history):
    with open(HISTORY_FILE, 'w') as f:
        json.dump(history, f)

def run():
    history = load_history()
    master_deadlines = {}

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            channel="chrome",
            headless=False,
            accept_downloads=True
        )
        page = context.pages[0]
        
        print("🚀 Accessing Teams Web Portal...")
        page.goto("https://teams.microsoft.com/v2/", wait_until="domcontentloaded", timeout=90000)

        try:
            # Navigate to Assignments
            assignments_btn = page.locator("[data-tid='app-bar-edu-assignments']").first
            assignments_btn.wait_for(state="visible", timeout=60000)
            assignments_btn.click()

            iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame
            target_card_id = "[id*='CardHeader__headerEDUASSIGN']:visible"
            iframe.locator(target_card_id).first.wait_for(state="visible", timeout=60000)

            assignment_cards = iframe.locator(target_card_id)
            total_assignments = assignment_cards.count()
            
            print(f"✅ Found {total_assignments} assignments in view.")

            for i in range(total_assignments):
                iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame
                current_card = iframe.locator(target_card_id).nth(i)
                
                # Preliminary Extraction
                full_card_text = current_card.locator("xpath=..").inner_text()
                assignment_title = full_card_text.split('\n')[0].strip()
                
                card_lines = full_card_text.split('\n')
                raw_card_date = card_lines[1].replace("Due ", "").strip() if len(card_lines) > 1 else "N/A"

                course_name = "OTHER" 
                for key, clean_name in COURSE_MAP.items():
                    if key.upper() in full_card_text.upper():
                        course_name = clean_name
                        break
                
                unique_id = f"[{course_name}] {assignment_title}"
                master_deadlines[unique_id] = raw_card_date

                # --- SPEED OPTIMIZATION: SKIP LOGIC ---
                if unique_id in history and history[unique_id] == raw_card_date:
                    print(f"⏩ Skipping {unique_id} (No updates).")
                    continue

                print(f"🔍 Processing {unique_id} (New/Modified)...")
                current_card.click()
                
                try:
                    iframe.locator("[data-test=\"back-button\"]").first.wait_for(state="visible", timeout=15000)
                    
                    # File Retrieval
                    attachment_menus = iframe.locator("[data-test=\"attachment-options-button\"]")
                    for j in range(attachment_menus.count()):
                        attachment_menus.nth(j).click()
                        try:
                            download_btn = iframe.get_by_role("menuitem", name="Download").first
                            with page.expect_download() as download_info:
                                download_btn.click()
                            
                            download = download_info.value
                            safe_course = re.sub(r'[\\/*?:"<>|]', "", course_name)
                            safe_title = re.sub(r'[\\/*?:"<>|]', "", assignment_title)
                            final_path = os.path.join(DOWNLOAD_DIR, f"[{safe_course}] {safe_title} - {download.suggested_filename}")
                            
                            if not os.path.exists(final_path):
                                download.save_as(final_path)
                            else:
                                download.cancel()
                        except:
                            pass
                    
                    # Log Visit in History
                    history[unique_id] = raw_card_date
                    save_history(history)

                except Exception as detail_err:
                    print(f"⚠️ Could not process details for {unique_id}: {detail_err}")

                # Return to list view
                iframe.locator("[data-test=\"back-button\"]").first.click()
                iframe.locator(target_card_id).first.wait_for(state="visible", timeout=20000)

        except Exception as e:
            print(f"❌ Automation encountered an error: {e}")

        context.close()

    # --- PROFESSIONAL TEXT FORMATTING ---
    sorted_deadlines = dict(sorted(master_deadlines.items()))
    formatted_list = "ACADEMIC DASHBOARD | MICROSOFT TEAMS\n"
    formatted_list += "=" * 40 + "\n\n"
    
    if not sorted_deadlines:
        formatted_list += "No pending assignments detected.\n"
    else:
        for name, date in sorted_deadlines.items():
            formatted_list += f"{name}\n"
            formatted_list += f"  > Due: {date}\n"
            formatted_list += "-" * 20 + "\n"
    
    formatted_list += f"\nLast Sync: {time.strftime('%H:%M | %b %d, %Y')}"

    pyperclip.copy(formatted_list)
    
    # --- STICKY NOTES OVERWRITE ---
    print("Updating Sticky Note Dashboard...")
    subprocess.Popen('explorer.exe shell:appsFolder\\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App')
    
    time.sleep(5) # Brief pause for app initialization
    
    # Selecting existing content and overwriting
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.3)
    pyautogui.press('backspace')
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'v')
    
    print("📋 Sync Complete. View updated in Sticky Notes.")

if __name__ == "__main__":
    run()