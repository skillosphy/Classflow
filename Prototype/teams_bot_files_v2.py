import os
import re
import time
from playwright.sync_api import sync_playwright

# --- CONFIGURATION ---
PROFILE_DIR = os.path.join(os.getcwd(), "automation_proxy")
MY_CLASSES = ["CE222-COAL-Spring26", "CS232"] # Add your exact class names here

def scroll_feed(page):
    """Clicks into the chat area and pages up to load older posts."""
    print("        -> Scrolling to load older messages...")
    try:
        page.mouse.click(500, 500)
        page.wait_for_timeout(500)
        for _ in range(5):
            page.keyboard.press("PageUp")
            page.wait_for_timeout(1500)
    except:
        pass

def run():
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            channel="chrome",
            headless=False,
            accept_downloads=True
        )
        
        page = context.pages[0]
        page.goto("https://teams.microsoft.com/v2/")
        
        print("Loading Teams Dashboard... (Waiting 15s)")
        page.wait_for_timeout(15000)

        for class_name in MY_CLASSES:
            try:
                print(f"\n=======================================")
                print(f"Entering Team: {class_name}")
                print(f"=======================================")
                
                # --- 1. ENTER THE TEAM ---
                try:
                    # Tries to find the exact class button like codegen did
                    page.get_by_role("button", name=re.compile(class_name)).first.click()
                except:
                    page.get_by_text(class_name).first.click()
                page.wait_for_timeout(5000)

                # --- 2. OPEN MAIN CHANNELS ---
                print("  -> Locating Channels...")
                # Use a safe text fallback instead of the brittle icon code from codegen
                try:
                    main_channels = page.get_by_text("Main Channels").first
                    if main_channels.is_visible(timeout=3000):
                        main_channels.click()
                        page.wait_for_timeout(2000)
                except:
                    pass # It might already be open

                # --- 3. GATHER CHANNELS ---
                # We dynamically read the treeitems so we don't have to hardcode "Chapter 1", "Chapter 2"
                tree_items = page.get_by_role("treeitem").all()
                target_channels = []
                
                for item in tree_items:
                    name = item.inner_text().split('\n')[0].strip()
                    if name and name not in ["Main Channels", "Hidden channels", "See all channels"] and class_name not in name:
                        clean_name = name.replace(" More options", "").strip()
                        if clean_name not in target_channels:
                            target_channels.append(clean_name)

                print(f"  -> Found Channels to Traverse: {target_channels}")

                # --- 4. TRAVERSE CHANNELS ---
                for channel_name in target_channels:
                    print(f"\n    [Channel: {channel_name}]")
                    
                    # Click the channel using the exact format from your Codegen
                    try:
                        page.get_by_role("treeitem", name=f"{channel_name} More options").click()
                    except:
                        page.get_by_text(channel_name).first.click()
                    
                    page.wait_for_timeout(4000)
                    
                    # --- 5. SCROLL ---
                    scroll_feed(page)

                    # --- 6. ASSIGNMENT CHECK (Path A: Deep Links & iFrames) ---
                    print("        -> Scanning for Assignment Links...")
                    assignment_links = page.locator("a[href*='assignment-viewer'], [title*='assignment-viewer']").all()
                    
                    if len(assignment_links) > 0:
                        print(f"          => Found {len(assignment_links)} assignment link(s). Opening...")
                        assignment_links[0].click()
                        page.wait_for_timeout(5000) 
                        
                        try:
                            # Switch to iframe as discovered in Codegen
                            iframe = page.locator("iframe[name='embedded-page-container']").content_frame
                            
                            # Standard assignment 3-dots
                            options_btn = iframe.locator("[data-test='attachment-options-button']").first
                            if options_btn.is_visible(timeout=4000):
                                options_btn.click()
                                page.wait_for_timeout(1000)
                                
                                with page.expect_download() as download_info:
                                    iframe.get_by_role("menuitem", name="Download").click()
                                print(f"          => SUCCESS: Downloaded {download_info.value.suggested_filename}")
                                page.wait_for_timeout(2000)
                            
                            # Exit assignment using the specific iframe back button from Codegen
                            print("          => Returning to feed...")
                            iframe.locator("[data-test='back-button']").first.click()
                            page.wait_for_timeout(2000)
                            
                        except Exception as e:
                            print(f"          => Error inside iframe: {e}")

                    # --- 7. ASSIGNMENT CHECK (Path B: Direct Feed Attachments) ---
                    print("        -> Scanning for Direct File Attachments...")
                    direct_attachments = page.get_by_label("More attachment options").all()
                    
                    if len(direct_attachments) > 0:
                        print(f"          => Found {len(direct_attachments)} direct file(s). Downloading...")
                        direct_attachments[0].click()
                        page.wait_for_timeout(1000)
                        
                        try:
                            with page.expect_download() as download_info:
                                page.get_by_text("Download").first.click()
                            print(f"          => SUCCESS: Downloaded {download_info.value.suggested_filename}")
                            page.wait_for_timeout(2000)
                        except Exception as e:
                            print(f"          => Could not download direct file: {e}")

            except Exception as e:
                print(f"❌ Error in Team '{class_name}': {e}")
            
            # --- 8. RETURN TO DASHBOARD ---
            print("  -> Exiting Team, returning to Dashboard...")
            try:
                page.get_by_role("button", name="Back to All teams").click()
            except:
                page.goto("https://teams.microsoft.com/v2/")
            page.wait_for_timeout(4000)

        print("\n=== CRAWL COMPLETE ===")
        context.close()

if __name__ == "__main__":
    run()