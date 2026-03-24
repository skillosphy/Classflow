import os
import re
import json
import pyperclip
import pyautogui
import subprocess
import time
from playwright.sync_api import sync_playwright

# --- 1. DYNAMIC CONFIGURATION ---
PROFILE_DIR = os.path.join(os.environ['LOCALAPPDATA'], "TeamsBot_Proxy")
USER_DESKTOP = os.path.join(os.environ['USERPROFILE'], 'Desktop')
DOWNLOAD_DIR = os.path.join(USER_DESKTOP, "Assignments")
# Hidden file to track visited assignments for speed optimization
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
        except: return {}
    return {}

def save_history(history):
    with open(HISTORY_FILE, 'w') as f:
        json.dump(history, f)

def clean_date_string(raw_text):
    return raw_text.replace("Due ", "").strip()

def run():
    history = load_history()
    master_deadlines = {}

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            channel="chrome",
            headless=False, # Set to False if you need to redo Login/2FA
            accept_downloads=True
        )
        page = context.pages[0]
        
        print("🚀 Navigating to Teams...")
        page.goto("https://teams.microsoft.com/v2/", wait_until="domcontentloaded", timeout=90000)
        
        # Check login
        if "login" in page.url or "microsoftonline" in page.url:
            print("🚨 Login Required! Run with headless=False to authenticate.")
            return

        try:
            # Locate Assignment Button
            assignments_btn = page.locator("[data-tid='app-bar-edu-assignments']").first
            assignments_btn.wait_for(state="visible", timeout=60000)
            assignments_btn.click()

            iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame
            target_card_id = "[id*='CardHeader__headerEDUASSIGN']:visible"
            iframe.locator(target_card_id).first.wait_for(state="visible", timeout=60000)

            assignment_cards = iframe.locator(target_card_id)
            total_assignments = assignment_cards.count()
            
            print(f"✅ Found {total_assignments} assignment(s). Check history...")

            for i in range(total_assignments):
                try:
                    iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame
                    current_card = iframe.locator(target_card_id).nth(i)
                    
                    # --- EXTRACTION FOR SKIP LOGIC ---
                    full_card_text = current_card.locator("xpath=..").inner_text()
                    assignment_title = full_card_text.split('\n')[0].strip()
                    card_lines = full_card_text.split('\n')
                    raw_card_date = card_lines[1].replace("Due ", "").strip() if len(card_lines) > 1 else "N/A"

                    course_name = "Other" 
                    for key, clean_name in COURSE_MAP.items():
                        if key.upper() in full_card_text.upper():
                            course_name = clean_name
                            break
                    
                    unique_id = f"[{course_name}] {assignment_title}"
                    master_deadlines[unique_id] = raw_card_date

                    # --- SPEED OPTIMIZATION ---
                    if unique_id in history and history[unique_id] == raw_card_date:
                        print(f"Skipping {unique_id} (Already synced)")
                        continue

                    print(f"🔍 Visiting {unique_id}...")
                    current_card.click()
                    iframe.locator("[data-test=\"back-button\"]").first.wait_for(state="visible", timeout=15000)

                    # Download Files Logic (Core Logic Maintained)
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
                        except: pass
                    
                    # Update History
                    history[unique_id] = raw_card_date
                    save_history(history)

                    # Back to list
                    iframe.locator("[data-test=\"back-button\"]").first.click()
                    iframe.locator(target_card_id).first.wait_for(state="visible", timeout=20000)
                
                except Exception as e:
                    print(f"Error processing assignment {i+1}: {e}")

        except Exception as e:
            print(f"❌ Critical Automation Error: {e}")

        context.close()

    # --- PROFESSIONAL FORMATTING ---
    sorted_deadlines = dict(sorted(master_deadlines.items()))
    formatted_list = "Pending Assignments\n"
    formatted_list += "=" * 35 + "\n\n"
    
    for name, date in sorted_deadlines.items():
        formatted_list += f"{name}\n"
        formatted_list += f"  > Due: {date}\n"
        formatted_list += "-" * 20 + "\n"
    
    formatted_list += f"\nLast Sync: {time.strftime('%H:%M | %b %d')}"
    pyperclip.copy(formatted_list)
    
    # --- STICKY NOTE OVERWRITE ---
    print("Updating Sticky Note...")
    subprocess.Popen('explorer.exe shell:appsFolder\\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App')
    time.sleep(5) # Wait for focus
    
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.3)
    pyautogui.press('backspace')
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'v')
    
    print("Sticky Notes updated")

if __name__ == "__main__":
    run()