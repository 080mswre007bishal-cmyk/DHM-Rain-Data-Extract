"""
main.py
-------
For every station in stations.json:
  1. Open its DHM page in a headless browser
     (dhm.gov.np/hydrology/rainfallSingle/<id> -- confirmed to carry the
     same "Date / Period / VIEW / Accumulated Rainfall" widget as the
     rainfall-watch-map page, via a screenshot you shared).
  2. Set the Date field to YESTERDAY (Nepal time), because today's
     figure is still updating throughout the day.
  3. Set Period to "Daily" and click VIEW. This reveals a table with
     three columns: Date, Daily (mm), Total (mm) -- Total is just the
     running sum of Daily across the visible rows.
  4. Read every row of that table.
  5. For each station's output/<station>.xlsx, skip any date already
     recorded and append only the new ones. Running this twice, or
     missing a scheduled run, does not create duplicates and
     self-heals small gaps, since each run re-reads the whole visible
     window (usually the last 8 days).

Google Drive upload is DISABLED for now, per your request -- focus is
on getting the extraction right first. The functions are still here,
commented out of the main loop, so it's a one-line change to turn
back on later. In the meantime, output/*.xlsx is picked up by the
"Upload run artifacts" step already in rainfall.yml, so you can
download it from the Actions run summary page after every run.

IMPORTANT -- what's confirmed vs. not (this has already been through
one debug round with real output, see the comments in
set_period_to_daily() and extract_rainfall_history() below for what
that revealed):
CONFIRMED: the table's three columns (Date, Daily, Total) and their
date format ("Fri, Sep 25, 2026, 12:00 AM"); that Period is a row of
clickable tabs (Point / Hourly / Daily / Point for 7 days), NOT a
dropdown; and that the Date field defaults to today, so today's row
is filtered out after extraction rather than trying to set the field.
STILL NOT FULLY CONFIRMED: whether clicking "Daily" reliably lands on
that tab every time (vs. some other tab/label elsewhere on the page
also matching). Run DEBUG_STATIONS=3 and check output/debug/*.txt --
specifically the "all_rows_extracted" line and whether the real page
text's table header reads "Date Daily Total" (correct) rather than
"Date Point" (still wrong) -- before trusting a live run.

Usage:
    pip install playwright openpyxl
    playwright install chromium
    python main.py
"""

import datetime
import json
import os
import re
import time

from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright

STATIONS_FILE = "stations.json"
OUTPUT_DIR = "output"
DEBUG_DIR = os.path.join(OUTPUT_DIR, "debug")
STATION_URL = "https://dhm.gov.np/hydrology/rainfallSingle/{}"

# Set via env var for a quick test, e.g. DEBUG_STATIONS=3 python main.py
DEBUG_STATIONS = int(os.environ.get("DEBUG_STATIONS", "0"))

# Nepal is UTC+5:45.
NEPAL_OFFSET = datetime.timedelta(hours=5, minutes=45)

# Date formats confirmed / guessed for parsing table cells and the top
# Date field. Confirmed format is first.
DATE_FORMATS = [
    "%a, %b %d, %Y, %I:%M %p",   # Fri, Sep 25, 2026, 12:00 AM  (confirmed)
    "%Y-%m-%d",                   # 2026-09-25
    "%d %b, %Y", "%d %B, %Y", "%d %b %Y", "%d %B %Y",
    "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y",
]
DATE_PATTERNS = [
    r"[A-Za-z]{3},\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4}(?:,\s+\d{1,2}:\d{2}\s*[AP]M)?",
    r"\d{4}-\d{2}-\d{2}",
    r"\d{1,2}\s+[A-Za-z]{3,9},?\s+\d{4}",
    r"\d{1,2}/\d{1,2}/\d{4}",
    r"\d{1,2}-\d{1,2}-\d{4}",
]


def safe_filename(name, station_index, url_id):
    base = station_index or url_id
    clean_name = re.sub(r"[^\w\-]+", "_", name).strip("_")
    return f"{base}_{clean_name}.xlsx"


def nepal_today_iso():
    now_npt = datetime.datetime.utcnow() + NEPAL_OFFSET
    return now_npt.date().strftime("%Y-%m-%d")


def parse_date_token(token):
    token = token.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.datetime.strptime(token, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def parse_number(text):
    matches = re.findall(r"\d+\.\d+|\d+", text)
    if not matches:
        return None
    val = float(matches[0])
    return val if 0 <= val < 1000 else None


def set_period_to_daily(page):
    """Click the 'Daily' tab among Point / Hourly / Daily / Point for 7
    days. CONFIRMED (from a real debug run): Period is a row of plain
    clickable tabs, not a <select> dropdown -- an earlier version of
    this code tried select_option() first, which silently matched some
    unrelated hidden <select> on the page and reported false success,
    while the page's actual Period stayed on its default, "Point" (the
    raw 10-minute AWS feed). Exact match ("Daily") is tried first to
    avoid accidentally matching "Point for 7 days"."""
    selectors = [
        "text='Daily'",
        "button:has-text('Daily')",
        "a:has-text('Daily')",
        "li:has-text('Daily')",
        "label:has-text('Daily')",
    ]
    for selector in selectors:
        try:
            page.click(selector, timeout=2000)
            return True
        except Exception:
            continue
    return False


def table_rows_from_page(page):
    """Read the Accumulated Rainfall table: Date | Daily | Total.
    Returns a list of (date_str, daily_value, total_value)."""
    rows_data = []
    try:
        rows = page.query_selector_all("table tr")
    except Exception:
        rows = []

    for row in rows:
        try:
            cells = row.query_selector_all("td, th")
            texts = [c.inner_text().strip() for c in cells]
        except Exception:
            continue

        if len(texts) < 2:
            continue

        date_str = None
        for pattern in DATE_PATTERNS:
            m = re.search(pattern, texts[0])
            if m:
                date_str = parse_date_token(m.group(0))
                if date_str:
                    break
        if not date_str:
            continue

        daily_val = parse_number(texts[1]) if len(texts) > 1 else None
        total_val = parse_number(texts[2]) if len(texts) > 2 else None

        if daily_val is not None:
            rows_data.append((date_str, daily_val, total_val))

    return rows_data


def extract_rainfall_history(page, debug_tag=None):
    """
    Clicks the "Daily" period tab and VIEW, then reads the resulting
    table (Date | Daily | Total).

    Doesn't touch the Date field: it's confirmed to default to today,
    and rather than fight an unconfirmed date-picker element, today's
    row (which may still be accumulating) is simply filtered out of
    the result below instead.

    Returns a list of (date_str, daily_value, total_value) tuples,
    excluding today's row.
    """
    period_set_ok = set_period_to_daily(page)

    for selector in ["text=VIEW", "button:has-text('VIEW')", "input[value='VIEW']"]:
        try:
            page.click(selector, timeout=3000)
            break
        except Exception:
            continue

    try:
        page.wait_for_selector("text=Accumulated Rainfall", timeout=8000)
    except Exception:
        pass
    try:
        page.wait_for_load_state("networkidle", timeout=8000)
    except Exception:
        page.wait_for_timeout(3000)

    all_rows = table_rows_from_page(page)
    today = nepal_today_iso()
    kept_rows = [r for r in all_rows if r[0] != today]
    skipped_today = [r for r in all_rows if r[0] == today]

    if debug_tag:
        os.makedirs(DEBUG_DIR, exist_ok=True)
        page.screenshot(path=os.path.join(DEBUG_DIR, f"{debug_tag}.png"))
        with open(os.path.join(DEBUG_DIR, f"{debug_tag}.txt"), "w", encoding="utf-8") as f:
            f.write(f"period_set_ok (clicked 'Daily' tab): {period_set_ok}\n")
            f.write(f"today (Nepal): {today}\n")
            f.write(f"all_rows_extracted: {all_rows}\n")
            f.write(f"skipped_today_row (still accumulating): {skipped_today}\n")
            f.write(f"kept_rows (written to Excel): {kept_rows}\n\n")
            f.write("=== FULL PAGE TEXT AFTER VIEW ===\n")
            f.write(page.inner_text("body"))

    return kept_rows


def get_existing_dates(path):
    if not os.path.exists(path):
        return set()
    wb = load_workbook(path)
    ws = wb.active
    dates = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            dates.add(str(row[0]))
    return dates


def append_new_rows(path, new_entries, run_timestamp):
    """new_entries: list of (date_str, daily_value, total_value) tuples
    NOT already in the file. Appends one row per entry and saves."""
    if os.path.exists(path):
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["Date", "Daily Rainfall (mm)", "Running Total on page (mm)", "First recorded (UTC)"])

    for date_str, daily_val, total_val in sorted(new_entries):
        ws.append([date_str, daily_val, total_val, run_timestamp])

    wb.save(path)


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(STATIONS_FILE, encoding="utf-8") as f:
        stations = json.load(f)

    if DEBUG_STATIONS:
        stations = stations[:DEBUG_STATIONS]
        print(f"DEBUG MODE: running only {len(stations)} station(s)")

    run_timestamp = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")

    with sync_playwright() as p:
        # Always headless: this runs on GitHub's servers, which have no
        # display. Debug mode still works -- it saves a screenshot and
        # the page text/state to output/debug/ so you can inspect the
        # result afterward instead of watching a live browser window.
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        for station in stations:
            url_id = station["url_id"]
            name = station["name"]
            print(f"-> {name} (url_id={url_id})")

            filename = safe_filename(name, station.get("station_index"), url_id)
            filepath = os.path.join(OUTPUT_DIR, filename)

            try:
                page.goto(STATION_URL.format(url_id), timeout=20000)
                debug_tag = None
                if DEBUG_STATIONS:
                    clean_name_for_tag = re.sub(r"[^\w-]+", "_", name)
                    debug_tag = f"{url_id}_{clean_name_for_tag}"
                history = extract_rainfall_history(page, debug_tag=debug_tag)
            except Exception as e:
                print(f"   failed: {e}")
                history = []

            print(f"   read {len(history)} row(s): {history}")

            existing_dates = get_existing_dates(filepath)
            new_entries = [(d, daily, total) for d, daily, total in history if d not in existing_dates]

            if new_entries:
                append_new_rows(filepath, new_entries, run_timestamp)
                print(f"   added {len(new_entries)} new row(s): {sorted(new_entries)}")
            else:
                print("   no new rows to add (all already recorded, or nothing parsed)")

            # Google Drive upload is disabled for now -- output/*.xlsx is
            # available as a downloadable workflow artifact instead. To
            # re-enable later, restore get_drive_service()/upload_to_drive()
            # from the previous version and call them here.

            time.sleep(0.5)  # be polite to the server

        browser.close()

    print("Done.")


if __name__ == "__main__":
    main()
