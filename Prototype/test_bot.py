import os
from playwright.sync_api import sync_playwright

PROFILE_DIR = os.path.join(os.getcwd(), "automation_proxy")

def run():
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            channel="chrome",
            headless=False,
            accept_downloads=True  # Added this to ensure Chrome allows the file!
        )
        page = context.pages[0]
        
        print("Navigating to Teams...")
        page.goto("https://teams.microsoft.com/v2/")
        page.wait_for_timeout(7500) 
        
        print("Clicking Class...")
        page.get_by_role("button", name="CE222-COAL-Spring26").click()
        page.wait_for_timeout(2000)

        print("Clicking Channel...")
        page.get_by_role("treeitem", name="General").click()

        page.mouse.click(500, 500)
        page.wait_for_timeout(500)

        print("Opening Assignment...")
        page.get_by_title("https://teams.microsoft.com/l/entity/66aeee93-507d-479a-a3ef-8f494af43945/classroom?context=%7B%22subEntityId%22%3A%22%7B%5C%22version%5C%22%3A%5C%221.0%5C%22,%5C%22config%5C%22%3A%7B%5C%22classes%5C%22%3A%5B%7B%5C%22id%5C%22%3A%5C%226835d677-87fa-4f9d-9055-cd43a0bea0e3%5C%22,%5C%22assignmentIds%5C%22%3A%5B%5C%2260615be9-09b9-442d-8ff4-408d20076fd5%5C%22%5D%7D%5D%7D,%5C%22action%5C%22%3A%5C%22navigate%5C%22,%5C%22view%5C%22%3A%5C%22assignment-viewer%5C%22,%5C%22deeplinkType%5C%22%3A0%7D%22,%22channelId%22%3A%2219%3AQJRqZAhnhVwyqbzal_0PjKj3duhrACY4_CSxuVfpP9Y1@thread.tacv2%22%7D", exact=True).click()
        page.wait_for_timeout(5000)

        print("Opening 3-Dots Menu...")
        page.locator("iframe[name=\"embedded-page-container\"]").content_frame.locator("[data-test=\"attachment-options-button\"]").click()
        page.wait_for_timeout(1000)

        print("Triggering Download...")
        # Start the listener
        with page.expect_download() as download_info:
            page.locator("iframe[name=\"embedded-page-container\"]").content_frame.get_by_role("menuitem", name="Download").click()
        
        # WE ARE NOW OUTSIDE THE 'WITH' BLOCK
        download = download_info.value
        
        # Grab the real name and extension of the file from the server
        real_filename = download.suggested_filename
        print(f"Detected File: {real_filename}")
        
        # Save it to your current working directory
        final_path = os.path.join(os.getcwd(), real_filename)
        download.save_as(final_path)
        
        print(f"SUCCESS! File saved perfectly to: {final_path}")

        page.wait_for_timeout(5000)
        context.close()

if __name__ == "__main__":
    run()