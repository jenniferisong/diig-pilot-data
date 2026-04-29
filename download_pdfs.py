import json
import os
import pandas as pd
from playwright.sync_api import sync_playwright

COOKIES_FILE = "cookies.json"
LINKS_FILE = "links_found.txt"
START_URL = "https://congressional-proquest-com.proxy.lib.duke.edu/profiles/gis/search/basic/basicsearch"

def main():
    # Find the Excel file in the directory
    excel_files = [p for p in os.listdir('.') if p.endswith(('.xls', '.xlsx')) and p != "output_with_pdfs.xlsx"]
    if not excel_files:
        print("No Excel file found!")
        return
    
    df = pd.read_excel(excel_files[0])
    # Target range: 3840 to 6960
    subset = df.iloc[3839:6960]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-setuid-sandbox"])
        context = browser.new_context()

        # Cookie Sanitizer
        if os.path.exists(COOKIES_FILE):
            with open(COOKIES_FILE, 'r') as f:
                cookies = json.load(f)
            for c in cookies:
                ss = str(c.get("sameSite", "Lax")).lower()
                c["sameSite"] = ss.capitalize() if ss in ["strict", "lax", "none"] else "Lax"
                if "id" in c: del c["id"]
            context.add_cookies(cookies)
            print("Cookies injected.")

        page = context.new_page()

        with open(LINKS_FILE, "a") as f:
            for index, row in subset.iterrows():
                uid = str(row.iloc[2]).strip()
                print(f"Row {index + 1} | UID: {uid}")

                try:
                    # Navigate with a long timeout for the proxy
                    page.goto(START_URL, wait_until="domcontentloaded", timeout=60000)
                    
                    # 1. Use the EXACT ID from your provided HTML: searchText_0
                    search_input = page.wait_for_selector("#searchText_0", timeout=20000)
                    search_input.click()
                    search_input.fill(uid)
                    
                    # 2. Click the actual search button from your HTML: #submitbutton
                    page.click("#submitbutton")
                    
                    # 3. Wait for the results to load and the Permalink to appear
                    # ProQuest results can be slow; waiting 30s max
                    permalink_btn = page.wait_for_selector("text=Permalink", timeout=30000)
                    permalink_btn.click()
                    
                    # 4. Extract the link from the popup
                    page.wait_for_timeout(3000)
                    link = page.evaluate("""() => {
                        const el = document.querySelector('input[readonly], #permalinkText, .permalink-url');
                        if (el) return el.value || el.innerText;
                        return document.activeElement ? document.activeElement.value : '';
                    }""")
                    
                    if "http" in str(link):
                        f.write(f"{uid} | {link}\n")
                        f.flush()
                        print(f"  ✅ SUCCESS: {link}")
                    else:
                        print("  ❌ Extraction failed (Empty link).")

                except Exception as e:
                    print(f"  ❌ Error: {type(e).__name__}")
                    # Save a screenshot for every failure to diagnose
                    page.screenshot(path="last_failure.png")

        browser.close()

if __name__ == "__main__":
    main()