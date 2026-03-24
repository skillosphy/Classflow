import os
import re
import pyperclip
import tkinter as tk
from playwright.sync_api import sync_playwright

# --- CONFIGURATION ---
PROFILE_DIR = os.path.join(os.getcwd(), "automation_proxy")
DOWNLOAD_DIR = r"C:\Users\Lenovo\Desktop\Assignments"
DEADLINE_FILE = r"C:\Users\Lenovo\Desktop\Script-Kiddie\Classflow\deadlines.txt"

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

def clean_date_string(raw_text):
    """Just strips the word 'Due ' from the text."""
    return raw_text.replace("Due ", "").strip()

def run():
    master_deadlines = {}

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            channel="chrome",
            headless=True,
            accept_downloads=True
        )
        page = context.pages[0]
        
        print("Navigating to Teams...")
        page.goto("https://teams.microsoft.com/v2/")
        
        print("Waiting for Teams to load...")
        try:
            assignments_btn = page.get_by_role("button", name="Assignments (Ctrl+Shift+4)")
            assignments_btn.wait_for(state="visible", timeout=30000)
            assignments_btn.click()
        except:
            assignments_btn = page.locator("[data-tid='app-bar-edu-assignments']").first
            assignments_btn.wait_for(state="visible", timeout=30000)
            assignments_btn.click()

        print("Waiting for assignment list to fetch...")
        iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame
        target_card_id = "[id*='CardHeader__headerEDUASSIGN']:visible"
        iframe.locator(target_card_id).first.wait_for(state="visible", timeout=20000)

        assignment_cards = iframe.locator(target_card_id)
        total_assignments = assignment_cards.count()
        
        print(f"\nFound {total_assignments} clickable upcoming assignment(s)!\n")

        for i in range(total_assignments):
            iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame
            current_card = iframe.locator(target_card_id).nth(i)
            
            # --- 1. EXTRACT INFO FOR SORTING ---
            full_card_text = current_card.locator("xpath=..").inner_text()
            
            # The Title is always the very first line before a line break
            assignment_title = full_card_text.split('\n')[0].strip()
            
            # Smart Dictionary Lookup (CASE-INSENSITIVE)
            course_name = "Other" 
            for key, clean_name in COURSE_MAP.items():
                if key.upper() in full_card_text.upper():
                    course_name = clean_name
                    break
            
            unique_display_name = f"[{course_name}] {assignment_title}"
            print(f"--- Processing: {unique_display_name} ---")

            # Click to enter details
            current_card.click()
            
            # Wait for details
            iframe.locator("[data-test=\"back-button\"]").first.wait_for(state="visible", timeout=15000)

            try:
                # --- 2. EXTRACT DEADLINE ---
                date_element = iframe.get_by_text(re.compile(r"^Due ")).first
                if date_element.is_visible(timeout=3000):
                    raw_due_text = date_element.inner_text()
                    final_date = clean_date_string(raw_due_text)
                    print(f"  -> 📅 Deadline: {final_date}")
                    master_deadlines[unique_display_name] = final_date
                else:
                    master_deadlines[unique_display_name] = "No date specified"

                # --- 3. DOWNLOAD & RENAME FILES ---
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
                        
                        # --- THE SMART RENAME ---
                        safe_course = re.sub(r'[\\/*?:"<>|]', "", course_name)
                        safe_title = re.sub(r'[\\/*?:"<>|]', "", assignment_title)
                        
                        new_filename = f"[{safe_course}] {safe_title} - {original_filename}"
                        final_path = os.path.join(DOWNLOAD_DIR, new_filename)

                        if os.path.exists(final_path):
                            print(f"    => Skipping: '{new_filename}' already exists.")
                            download.cancel() 
                        else:
                            print(f"    => Saving as: '{new_filename}'")
                            download.save_as(final_path)
                            
                        page.wait_for_timeout(1000)
                else:
                    print("  -> No files to download.")

            except Exception as e:
                print(f"  -> Error: {e}")

            # Return to list
            try:
                iframe.locator("[data-test=\"back-button\"]").first.click()
            except:
                page.go_back()
            
            iframe.locator(target_card_id).first.wait_for(state="visible", timeout=15000)

        context.close()

    # --- 4. FORMATTED TEXT GENERATION ---
    print("\n=== GENERATING CUSTOM STICKY NOTE ===\n")
    
    sorted_deadlines = dict(sorted(master_deadlines.items()))

    formatted_list = "📅 ASSIGNMENT TRACKER 📅\n"
    formatted_list += "------------------------\n"
    
    for unique_name, date in sorted_deadlines.items():
        formatted_list += f"🔴 {unique_name}\n   Due: {date}\n\n"
        
    with open(DEADLINE_FILE, "w", encoding="utf-8") as f:
        f.write(formatted_list)
        
    pyperclip.copy(formatted_list)
    
    # --- 5. THE PYTHON STICKY NOTE UI ---
    root = tk.Tk()
    root.title("Assignment Deadlines")
    root.configure(bg="#FFF7D1") # Sticky note yellow
    root.attributes("-topmost", True) # Keep it on top of other windows
    
    # Text Label
    label = tk.Label(
        root, 
        text=formatted_list, 
        bg="#FFF7D1", 
        font=("Consolas", 11), 
        justify="left", 
        padx=20, 
        pady=20
    )
    label.pack()
    
    # Dismiss Button
    close_btn = tk.Button(
        root, 
        text="Dismiss", 
        command=root.destroy, 
        bg="#FFD0D0", 
        relief="flat",
        font=("Consolas", 10, "bold"),
        padx=10,
        pady=5
    )
    close_btn.pack(pady=(0, 15))
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    
    # --- NEW MATH FOR TOP RIGHT ---
    # Screen width minus window width, minus 20 pixels for a clean margin
    x = root.winfo_screenwidth() - width - 20
    y = 20 # 20 pixels down from the absolute top of the screen
    
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    # Run the UI loop (this keeps the window open until you click Dismiss)
    root.mainloop()
    
    print("📋 Script Complete!")

if __name__ == "__main__":
    run()