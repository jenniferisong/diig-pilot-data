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
            try:
                locator.click()
                locator.press("Control+A")
                locator.press("Backspace")
            except Exception:
                pass
            locator.fill(uid)
            return True

    return False


def has_search_box(page) -> bool:
    """
    Check whether the current page looks like the search page.
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
        try:
            if page.locator(selector).first.count() > 0:
                return True
        except Exception:
            pass

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


def go_to_search_page(page):
    """
    Make sure we are back on the main search page before starting a UID.
    """
    try:
        page.goto(START_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(3000)
    except Exception:
        pass


def try_get_permalink(page, uid: str) -> str:
    """
    Click the Permalink option on the results page and read the URL.

    Many sites place the link into a focused text box after clicking Permalink.
    This function checks that focused box first, then other obvious input fields.
    """
    permalink_buttons = [
        ("button: Permalink", page.get_by_role("button", name=re.compile(r"permalink", re.I)).first),
        ("link: Permalink", page.get_by_role("link", name=re.compile(r"permalink", re.I)).first),
        ("text: Permalink", page.locator(":text('Permalink')").first),
    ]

    for label, locator in permalink_buttons:
        try:
            if locator.count() == 0:
                continue

            print(f"  Clicking permalink using: {label}")
            locator.click()
            page.wait_for_timeout(1500)

            active_value = page.evaluate(
                """
                () => {
                    const el = document.activeElement;
                    if (!el) return "";
                    if (typeof el.value === "string") return el.value.trim();
                    return "";
                }
                """
            )
            if isinstance(active_value, str) and active_value.startswith("http"):
                return active_value

            input_selectors = [
                "input[type='url']",
                "input[value^='http']",
                "textarea",
                "input",
            ]
            for selector in input_selectors:
                field = page.locator(selector).first
                if field.count() == 0:
                    continue
                try:
                    value = field.input_value().strip()
                    if value.startswith("http"):
                        return value
                except Exception:
                    pass
        except Exception:
            pass

    return ""


def try_read_visible_permalink(page) -> str:
    """
    Read a permalink directly from the current page without clicking anything.

    This is useful on a record page where a Permalink section is already visible.
    """
    try:
        permalink_section_selectors = [
            "label:has-text('Permalink') + input",
            "label:has-text('Permalink') ~ input",
            "text=Permalink:",
        ]

        for selector in permalink_section_selectors:
            try:
                locator = page.locator(selector).first
                if locator.count() == 0:
                    continue
                value = locator.input_value().strip()
                if value.startswith("http"):
                    return value
            except Exception:
                pass
    except Exception:
        pass

    input_selectors = [
        "input[type='url']",
        "input[value^='http']",
        "textarea",
        "input",
    ]

    for selector in input_selectors:
        try:
            field = page.locator(selector).first
            if field.count() == 0:
                continue
            value = field.input_value().strip()
            if value.startswith("http"):
                return value
        except Exception:
            pass

    try:
        page_text = page.locator("body").inner_text(timeout=5000)
        url_match = re.search(r"https?://\S+", page_text)
        if url_match:
            return url_match.group(0).rstrip(".,);")
    except Exception:
        pass

    return ""


def try_filter_to_hearing(page) -> bool:
    """
    Try to filter the results page to Document Type: HEARING.
    """
    hearing_filters = [
        ("checkbox: HEARING", page.get_by_role("checkbox", name=re.compile(r"hearing", re.I)).first),
        ("label: HEARING", page.get_by_text(re.compile(r"^hearing$", re.I)).first),
        ("document type HEARING", page.locator(":has-text('Document Type')").locator(":text('HEARING')").first),
        ("facet HEARING", page.locator("[class*='facet']:has-text('HEARING')").first),
    ]

    for label, locator in hearing_filters:
        try:
            if locator.count() == 0:
                continue

            print(f"  Applying HEARING filter using: {label}")
            locator.click()
            page.wait_for_timeout(4000)
            return True
        except Exception:
            pass

    return False


def process_one_uid(page, uid: str) -> str:
    """
    Search one UID, click Permalink on the results page,
    and capture that link.
    If the results page has no Permalink option, filter to Document Type: HEARING
    and try Permalink again.
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

    # Give the results page a moment to load before looking for the Permalink button.
    page.wait_for_timeout(4000)

    permalink = try_get_permalink(page, uid)
    if permalink:
        print(f"  Saved permalink: {permalink}")
        return permalink

    print("  No permalink on results page. Trying Document Type: HEARING filter...")
    if try_filter_to_hearing(page):
        permalink = try_get_permalink(page, uid)
        if permalink:
            print(f"  Saved permalink after HEARING filter: {permalink}")
            return permalink

    print(f"  Could not find a permalink for UID {uid}")
    return ""


def process_one_uid_with_retry(page, uid: str) -> str:
    """
    Try each UID up to two times.

    Each attempt starts from the search page again so old results do not
    interfere with the next search.
    """
    for attempt in range(1, 3):
        print(f"  Attempt {attempt} for UID {uid}")
        go_to_search_page(page)

        permalink = process_one_uid(page, uid)
        if permalink:
            return permalink

        print(f"  Attempt {attempt} did not succeed for UID {uid}")

    return ""


def main():
    """
    Main script flow.
    """
    print("Starting permalink collection script...")

    excel_path = find_excel_file()
    df = load_excel(excel_path)

    uid_column = find_uid_column(df)
    print(f"UID column found automatically: {uid_column}")

    # Create the output column if it does not already exist.
    if "permalink" not in df.columns:
        df["permalink"] = ""

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
                df.at[index, "permalink"] = ""
                continue

            # If permalink already has something, keep it and skip.
            existing_permalink = str(row.get("permalink", "")).strip()
            if existing_permalink and existing_permalink.lower() != "nan":
                print("  permalink already exists. Skipping.")
                continue

            try:
                permalink = process_one_uid_with_retry(page, uid)
                df.at[index, "permalink"] = permalink
            except PlaywrightTimeoutError:
                print(f"  Timed out while processing UID {uid}. Leaving permalink blank.")
                df.at[index, "permalink"] = ""
            except Exception as exc:
                print(f"  Error while processing UID {uid}: {exc}")
                print("  Leaving permalink blank and continuing.")
                df.at[index, "permalink"] = ""

            # Save after every row so progress is not lost if something stops later.
            df.to_excel(OUTPUT_FILE, index=False)

            # Go back to the search page for the next UID.
            try:
                go_to_search_page(page)
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
