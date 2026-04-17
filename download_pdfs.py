import re
import time
from pathlib import Path

import pandas as pd
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


# This is the website you will open manually in your own Edge window before the script starts searching.
START_URL = "https://congressional-proquest-com.proxy.lib.duke.edu/profiles/gis/search/basic/basicsearch"

# This is the local debugging address that Edge exposes when you start it
# with --remote-debugging-port=9222.
EDGE_DEBUG_URL = "http://127.0.0.1:9222"

# Output Excel file name.
OUTPUT_FILE = "output_with_pdfs.xlsx"


def find_excel_file() -> Path:
    """
    Find the first Excel file in the current folder.
    We ignore the output file so we do not accidentally re-read it later.
    """
    current_folder = Path(".")
    excel_files = [
        path
        for path in current_folder.iterdir()
        if path.is_file()
        and path.suffix.lower() in [".xls", ".xlsx"]
        and path.name != OUTPUT_FILE
    ]

    if not excel_files:
        raise FileNotFoundError("No Excel file (.xls or .xlsx) was found in this folder.")

    if len(excel_files) > 1:
        print("More than one Excel file was found.")
        print(f"Using the first one: {excel_files[0].name}")

    return excel_files[0]


def load_excel(excel_path: Path) -> pd.DataFrame:
    """
    Read the Excel file into a pandas DataFrame.

    Important:
    - .xlsx files usually work with openpyxl
    - .xls files usually need xlrd installed
    """
    print(f"Reading Excel file: {excel_path.name}")

    if excel_path.suffix.lower() == ".xls":
        return pd.read_excel(excel_path, engine="xlrd")

    return pd.read_excel(excel_path)


def find_uid_column(df: pd.DataFrame) -> str:
    """
    Automatically find the UID column.

    Simple rules:
    - First try exact matches like UID or uid
    - Then try exact matches like source
    - Then try any column name that contains the word uid
    """
    columns = list(df.columns)
    normalized_map = {
        column: str(column).strip().lower().replace(" ", "").replace("_", "")
        for column in columns
    }

    for column in columns:
        if normalized_map[column] == "uid":
            return column

    for column in columns:
        if normalized_map[column] == "source":
            return column

    for column in columns:
        normalized = normalized_map[column]
        if "uid" in normalized or "source" in normalized:
            return column

    raise ValueError(
        "Could not find a UID/source column automatically.\n"
        f"Columns found: {columns}"
    )


def clean_uid(value) -> str:
    """
    Turn a UID cell value into a clean string.
    Blank values become an empty string.
    """
    if pd.isna(value):
        return ""

    text = str(value).strip()
    if text.lower() == "nan":
        return ""

    return text


def make_safe_filename(text: str) -> str:
    """
    Remove characters that are not safe in Windows file names.
    """
    safe = re.sub(r'[<>:"/\\\\|?*]', "_", text)
    safe = safe.strip().strip(".")
    return safe or "downloaded_file"


def wait_for_manual_login():
    """
    Pause so you can log in manually in your own Edge window and get to the search page.
    """
    print("\nUse your own Edge window for login.")
    print("Please do these steps manually in Edge before pressing Enter:")
    print("1. First close all normal Edge windows")
    print("2. Start Edge with remote debugging enabled")
    print(r'   & "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\Users\jenni\Downloads\diig\edge_debug_profile"')
    print("3. In that Edge window, go to this page:")
    print(f"   {START_URL}")
    print("4. Log in if needed")
    print("5. Deal with any cookie or consent screen")
    print("6. Stay on the search page")
    print("7. Make sure the search box is visible and clickable")
    input("When Edge is ready on the search page, press Enter here to continue...")


def try_fill_search_box(page, uid: str):
    """
    Try several simple search box selectors.
    This keeps the first version flexible without being too complicated.
    """
    possible_search_boxes = [
        "input[type='search']",
        "input[name='q']",
        "input[name='query']",
        "input[name='search']",
        "input[id*='search']",
        "input[placeholder*='Search']",
        "input[placeholder*='search']",
        "input[type='text']",
    ]

    for selector in possible_search_boxes:
        locator = page.locator(selector).first
        if locator.count() > 0:
            locator.fill(uid)
            return True

    return False


def try_click_search_button(page):
    """
    Try common search button patterns.
    If no button is found, the caller can still press Enter.
    """
    possible_buttons = [
        page.get_by_role("button", name=re.compile("search", re.I)),
        page.get_by_role("link", name=re.compile("search", re.I)),
        page.locator("input[type='submit']").first,
        page.locator("button[type='submit']").first,
        page.locator("button").filter(has_text=re.compile("search", re.I)).first,
    ]

    for locator in possible_buttons:
        try:
            if locator.count() > 0:
                locator.click()
                return True
        except Exception:
            pass

    return False


def try_open_first_result(page) -> bool:
    """
    Open the first search result.
    """
    result_links = [
        ("first result title link", page.locator("a.resultTitle").first),
        ("first docview link", page.locator("a[href*='docview']").first),
        ("first record link", page.locator("a[href*='record']").first),
        ("first main result link", page.locator("main a").first),
        ("first visible link with title", page.get_by_role("link").first),
    ]

    for label, locator in result_links:
        try:
            if locator.count() == 0:
                continue

            print(f"  Opening first result using: {label}")
            locator.click()
            page.wait_for_timeout(3000)
            return True
        except Exception:
            pass

    return False


def try_open_details_tab(page) -> bool:
    """
    Open the Details tab on the record page.
    """
    details_tabs = [
        ("button: Details exact", page.get_by_role("button", name=re.compile(r"^details$", re.I)).first),
        ("button: Details", page.get_by_role("button", name=re.compile(r"details", re.I)).first),
        ("css button: Details", page.locator("button:has-text('Details')").first),
        ("tab: Details", page.get_by_role("tab", name=re.compile(r"details", re.I)).first),
        ("link: Details", page.get_by_role("link", name=re.compile(r"details", re.I)).first),
    ]

    for label, locator in details_tabs:
        try:
            if locator.count() == 0:
                continue

            print(f"  Opening details tab using: {label}")
            locator.wait_for(timeout=10000)
            locator.click()
            page.wait_for_timeout(4000)
            return True
        except Exception:
            pass

    return False


def try_get_full_pdf_url(page, uid: str) -> str:
    """
    Open the full PDF option and return its URL.

    Simple strategy:
    - Click a "Complete PDF" or similar option
    - If a new tab opens, use that tab's URL
    - If the current page changes to a PDF/viewer page, use the current URL
    """
    pdf_buttons = [
        ("button: Complete PDF", page.get_by_role("button", name=re.compile(r"complete\s*pdf", re.I)).first),
        ("link: Complete PDF", page.get_by_role("link", name=re.compile(r"complete\s*pdf", re.I)).first),
        ("button: Full text PDF", page.get_by_role("button", name=re.compile(r"full\s*text\s*pdf", re.I)).first),
        ("link: Full text PDF", page.get_by_role("link", name=re.compile(r"full\s*text\s*pdf", re.I)).first),
        ("button: Full PDF", page.get_by_role("button", name=re.compile(r"full.*pdf", re.I)).first),
        ("link: Full PDF", page.get_by_role("link", name=re.compile(r"full.*pdf", re.I)).first),
        ("button: PDF", page.get_by_role("button", name=re.compile(r"\bpdf\b", re.I)).first),
        ("link: PDF", page.get_by_role("link", name=re.compile(r"\bpdf\b", re.I)).first),
    ]

    original_url = page.url

    for label, locator in pdf_buttons:
        try:
            if locator.count() == 0:
                continue

            print(f"  Trying PDF fallback option: {label}")

            try:
                with page.context.expect_page(timeout=10000) as new_page_info:
                    locator.click()
                new_page = new_page_info.value
                new_page.wait_for_load_state("domcontentloaded", timeout=15000)
                new_page.wait_for_timeout(2000)
                pdf_url = new_page.url.strip()
                if pdf_url.startswith("http"):
                    new_page.close()
                    return pdf_url
                new_page.close()
            except Exception:
                # No new tab may have opened. In that case, keep checking below.
                pass

            page.wait_for_timeout(3000)
            if page.url and page.url != original_url and page.url.startswith("http"):
                return page.url
        except Exception:
            pass

    return ""


def process_one_uid(page, uid: str) -> str:
    """
    Search one UID, open the first result, open the details tab,
    and capture the full PDF page URL.
    If anything fails, return an empty string.
    """
    print(f"Searching UID: {uid}")

    if not try_fill_search_box(page, uid):
        print(f"  Could not find a search box for UID {uid}")
        return ""

    search_clicked = try_click_search_button(page)
    if not search_clicked:
        # If no visible search button is found, pressing Enter is the simplest fallback.
        page.keyboard.press("Enter")

    # Give the results page a moment to load before looking for the PDF option.
    page.wait_for_timeout(4000)

    if not try_open_first_result(page):
        print(f"  Could not open the first result for UID {uid}")
        return ""

    if not try_open_details_tab(page):
        print(f"  Could not open the details tab for UID {uid}")
        return ""

    pdf_url = try_get_full_pdf_url(page, uid)
    if pdf_url:
        print(f"  Saved PDF page URL: {pdf_url}")
        return pdf_url

    print(f"  Could not find a full PDF link for UID {uid}")
    return ""


def main():
    """
    Main script flow.
    """
    print("Starting PDF link collection script...")

    excel_path = find_excel_file()
    df = load_excel(excel_path)

    uid_column = find_uid_column(df)
    print(f"UID column found automatically: {uid_column}")

    # Create the output column if it does not already exist.
    if "pdf_link" not in df.columns:
        df["pdf_link"] = ""

    with sync_playwright() as p:
        wait_for_manual_login()

        print(f"Connecting to Edge at {EDGE_DEBUG_URL} ...")
        try:
            browser = p.chromium.connect_over_cdp(EDGE_DEBUG_URL)
        except Exception as exc:
            raise RuntimeError(
                "\nCould not connect to Edge on port 9222.\n"
                "Please do this exactly:\n"
                "1. Close all normal Edge windows\n"
                '2. Run this command:\n'
                '   & "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\\Users\\jenni\\Downloads\\diig\\edge_debug_profile"\n'
                f"3. Open this page in Edge: {START_URL}\n"
                "4. Log in and wait until the search box is visible\n"
                "5. Run python download_pdfs.py again\n"
            ) from exc

        # Reuse the first existing Edge context and tab if possible.
        if browser.contexts:
            context = browser.contexts[0]
        else:
            context = browser.new_context(accept_downloads=True)

        if context.pages:
            page = context.pages[0]
        else:
            page = context.new_page()
            page.goto(START_URL, wait_until="domcontentloaded")

        page.wait_for_timeout(2000)

        total_rows = len(df)

        for index, row in df.iterrows():
            uid = clean_uid(row[uid_column])

            print(f"\nRow {index + 1} of {total_rows}")

            # If UID is blank, skip it.
            if not uid:
                print("  UID is blank. Skipping.")
                df.at[index, "pdf_link"] = ""
                continue

            # If pdf_link already has something, keep it and skip.
            existing_pdf_link = str(row.get("pdf_link", "")).strip()
            if existing_pdf_link and existing_pdf_link.lower() != "nan":
                print("  pdf_link already exists. Skipping.")
                continue

            try:
                pdf_link = process_one_uid(page, uid)
                df.at[index, "pdf_link"] = pdf_link
            except PlaywrightTimeoutError:
                print(f"  Timed out while processing UID {uid}. Leaving pdf_link blank.")
                df.at[index, "pdf_link"] = ""
            except Exception as exc:
                print(f"  Error while processing UID {uid}: {exc}")
                print("  Leaving pdf_link blank and continuing.")
                df.at[index, "pdf_link"] = ""

            # Save after every row so progress is not lost if something stops later.
            df.to_excel(OUTPUT_FILE, index=False)

            # Go back to the search page for the next UID.
            # We may need to go back twice: once from the record page to results,
            # and once from results to the search page state.
            try:
                page.go_back(timeout=10000)
                page.wait_for_timeout(3000)
                page.go_back(timeout=10000)
                page.wait_for_timeout(3000)
            except Exception:
                print("  Could not go back automatically.")
                print("  If the next search fails, manually return to the search page and run again.")

            time.sleep(1)

        browser.close()

    # Save one final time at the end.
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\nDone. Updated spreadsheet saved as: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
