import os
import re
import json
import pyperclip
import pyautogui
import subprocess
import time
from playwright.sync_api import sync_playwright

# --- 1. DYNAMIC CONFIGURATION ---
# Stores the browser session so you stay logged in
PROFILE_DIR = os.path.join(os.environ['LOCALAPPDATA'], "TeamsBot_Proxy")

# Dynamically locate the current user's Desktop
USER_DESKTOP = os.path.join(os.environ['USERPROFILE'], 'Desktop')

DOWNLOAD_DIR = os.path.join(USER_DESKTOP, "Assignments")
# Hidden file to track visited assignments for speed skip logic
HISTORY_FILE = os.path.join(os.environ['LOCALAPPDATA'], "TeamsBot_History.json")

if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

# --- YOUR CUSTOM NAMING CONVENTION ---
COURSE_MAP = {
    "CS224": "FLAT",
    "CS272": "HCI",
    "CE222": "COAL",
    "CS232": "DBMS Lab",
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

def run():
    history = load_history()
    master_deadlines = {}

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            channel="chrome",
            headless=False, # Set to True for invisible background runs
            accept_downloads=True
        )
        page = context.pages[0]
        
        print("Navigating to Teams...")
        page.goto("https://teams.microsoft.com/v2/", wait_until="domcontentloaded", timeout=90000)
        
        # --- LOGIN DETECTION ---
        page.wait_for_timeout(5000) 
        if "login" in page.url or "microsoftonline" in page.url:
            print("🚨 Login Required! Bot waiting up to 5 minutes...")
            wait_time = 300000 
        else:
            wait_time = 30000 

        print("Waiting for Teams to load...")
        try:
            assignments_btn = page.get_by_role("button", name="Assignments (Ctrl+Shift+4)")
            assignments_btn.wait_for(state="visible", timeout=wait_time)
            assignments_btn.click()
        except:
            assignments_btn = page.locator("[data-tid='app-bar-edu-assignments']").first
            assignments_btn.wait_for(state="visible", timeout=wait_time)
            assignments_btn.click()

        print("Waiting for assignment list to fetch...")
        iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame
        target_card_id = "[id*='CardHeader__headerEDUASSIGN']:visible"
        iframe.locator(target_card_id).first.wait_for(state="visible", timeout=30000)

        assignment_cards = iframe.locator(target_card_id)
        total_assignments = assignment_cards.count()
        
        print(f"\nFound {total_assignments} assignment(s). Checking history for speed skip...\n")

        for i in range(total_assignments):
            try:
                iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame
                current_card = iframe.locator(target_card_id).nth(i)
                
                # --- ROBUST EXTRACTION FOR SPEED SKIP ---
                full_card_text = current_card.locator("xpath=..").inner_text()
                card_lines = [line.strip() for line in full_card_text.split('\n') if line.strip()]
                
                # Title is always the first non-empty line
                assignment_title = card_lines[0]
                
                # SEARCH for the line that actually contains the due date
                raw_card_date = "N/A"
                for line in card_lines:
                    if "Due " in line:
                        raw_card_date = line.replace("Due ", "").strip()
                        break

                course_name = "Other" 
                for key, clean_name in COURSE_MAP.items():
                    if key.upper() in full_card_text.upper():
                        course_name = clean_name
                        break
                
                unique_id = f"[{course_name}] {assignment_title}"
                master_deadlines[unique_id] = raw_card_date

                # --- SPEED SKIP LOGIC ---
                if unique_id in history and history[unique_id] == raw_card_date:
                    print(f"Skipping {unique_id} (No changes)")
                    continue

                print(f"Processing: {unique_id}")
                current_card.click()
                iframe.locator("[data-test=\"back-button\"]").first.wait_for(state="visible", timeout=15000)

                # --- DOWNLOAD LOGIC ---
                attachment_menus = iframe.locator("[data-test=\"attachment-options-button\"]")
                attachment_count = attachment_menus.count()

                if attachment_count > 0:
                    for j in range(attachment_count):
                        attachment_menus.nth(j).click()
                        try:
                            download_btn = iframe.get_by_role("menuitem", name="Download")
                            download_btn.wait_for(state="visible", timeout=5000)
                        except:
                            download_btn = iframe.get_by_text("Download").first
                            download_btn.wait_for(state="visible", timeout=5000)

                        with page.expect_download() as download_info:
                            download_btn.click()
                        
                        download = download_info.value
                        original_filename = download.suggested_filename
                        safe_course = re.sub(r'[\\/*?:"<>|]', "", course_name)
                        safe_title = re.sub(r'[\\/*?:"<>|]', "", assignment_title)
                        new_filename = f"[{safe_course}] {safe_title} - {original_filename}"
                        final_path = os.path.join(DOWNLOAD_DIR, new_filename)

                        if os.path.exists(final_path):
                            download.cancel() 
                        else:
                            download.save_as(final_path)
                        page.wait_for_timeout(1000)
                
                # Update history after successful visit
                history[unique_id] = raw_card_date
                save_history(history)

                # Re-navigating back to list
                try:
                    iframe.locator("[data-test=\"back-button\"]").first.click()
                except:
                    page.go_back()
                iframe.locator(target_card_id).first.wait_for(state="visible", timeout=15000)

            except Exception as e:
                print(f"Error on card {i}: {e}")

        context.close()

    # --- PROFESSIONAL FORMATTING (NO EMOJIS) ---
    sorted_deadlines = dict(sorted(master_deadlines.items()))
    formatted_list = "ACADEMIC DASHBOARD | TEAMS\n"
    formatted_list += "=" * 35 + "\n\n"
    
    if not sorted_deadlines:
        formatted_list += "No assignments found.\n"
    else:
        for name, date in sorted_deadlines.items():
            formatted_list += f"{name}\n"
            formatted_list += f"  > Due: {date}\n"
            formatted_list += "-" * 20 + "\n"
    
    formatted_list += f"\nLast Sync: {time.strftime('%H:%M | %b %d')}"
    pyperclip.copy(formatted_list)
    
    # --- STICKY NOTES OVERWRITE LOGIC ---
    print("Updating Sticky Notes...")
    subprocess.Popen('explorer.exe shell:appsFolder\\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App')
    time.sleep(5) 
    
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.3)
    pyautogui.press('backspace')
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'v')
    
    print("Note updated.")

if __name__ == "__main__":
    run()