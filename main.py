"""
main.py
-------
For every station in stations.json:
  1. Open its DHM page in a headless browser.
  2. Select the "Daily" period and click VIEW. On the pages I could
     confirm (dhm.gov.np/hydrology/rainfallSingle/<id>), this reveals
     an "Accumulated Rainfall" table -- the same widget you found on
     the rainfall-watch-map page. This is loaded by JavaScript, so a
     real browser is used instead of a plain HTTP request.
  3. Read every (date, value) pair that appears -- normally the last
     8 days.
  4. For each station's output/<station>.xlsx, skip any date already
     recorded and append only the new ones. Running this twice, or
     missing a scheduled run, does not create duplicates and
     self-heals small gaps because each run re-reads the last 8 days.
  5. If any new rows were added, upload/update that file in a Google
     Drive folder.

IMPORTANT -- please read before relying on this in production:
I confirmed the STATIC HTML of the DHM station pages by fetching them
directly (the "Period: Point / Hourly / Daily / Point for 7 days" and
"VIEW" / "Accumulated Rainfall" labels are really there). What I could
NOT confirm is exactly how the 8-day table is laid out in the page
text right after you click VIEW -- my environment can't reach
dhm.gov.np or run a real browser against it, so I never got to see
that table's actual text. extract_rainfall_history() below is a
best-effort, generic parser: it snapshots the page text before/after
clicking, looks for date-shaped tokens in the new text, and pairs each
one with the nearest plausible number (0-999). It is very likely to
need a tweak.

Run with DEBUG_STATIONS=3 first (see bottom of file). It saves a
screenshot AND the full before/after page text to output/debug/ for
each station -- open the .txt file, find the real table text, and
check whether the (date, value) pairs printed in the log actually
match. If not, adjust extract_rainfall_history() below using that
saved text as your reference. That's the one function you should
expect to tune; the Excel/dedup/Drive-upload logic under it doesn't
depend on getting the parsing exactly right on the first try.

Usage:
    pip install playwright openpyxl google-api-python-client google-auth
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

# How many days back the DHM "Daily" + VIEW table is expected to show.
# You said 8 -- adjust here if the debug output shows a different count.
HISTORY_DAYS = 8

# Date formats DHM might use in the table (tried in this order). The
# site's own header shows dates like "10 Jun, 2026", so that's first.
DATE_PATTERNS = [
    r"\d{1,2}\s+[A-Za-z]{3,9},?\s+\d{4}",   # 10 Jun, 2026 / 10 June 2026
    r"\d{4}-\d{2}-\d{2}",                    # 2026-06-10
    r"\d{1,2}/\d{1,2}/\d{4}",                # 10/06/2026
    r"\d{1,2}-\d{1,2}-\d{4}",                # 10-06-2026
]
DATE_FORMATS = [
    "%d %b, %Y", "%d %B, %Y", "%d %b %Y", "%d %B %Y",
    "%Y-%m-%d",
    "%d/%m/%Y", "%m/%d/%Y",
    "%d-%m-%Y",
]


def safe_filename(name, station_index, url_id):
    base = station_index or url_id
    clean_name = re.sub(r"[^\w\-]+", "_", name).strip("_")
    return f"{base}_{clean_name}.xlsx"


def parse_date_token(token):
    token = token.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.datetime.strptime(token, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def find_date_and_rest(line):
    """Find the first date-shaped token in `line`. Returns
    (normalized_date_str, line_with_date_removed) or (None, None)."""
    for pattern in DATE_PATTERNS:
        m = re.search(pattern, line)
        if m:
            parsed = parse_date_token(m.group(0))
            if parsed:
                rest = line[:m.start()] + line[m.end():]
                return parsed, rest
    return None, None


def extract_rainfall_history(page, debug_tag=None):
    """
    Best-effort extraction of the multi-day rainfall table that appears
    after selecting "Daily" and clicking VIEW.
    Returns a list of (date_str, value) tuples, date_str normalized to
    YYYY-MM-DD, newest/oldest order as found on the page.
    """
    before = page.inner_text("body")

    for selector in ["text=Daily", "label:has-text('Daily')", "[value='daily']"]:
        try:
            page.click(selector, timeout=2000)
            break
        except Exception:
            continue

    for selector in ["text=VIEW", "button:has-text('VIEW')", "input[value='VIEW']"]:
        try:
            page.click(selector, timeout=3000)
            break
        except Exception:
            continue

    try:
        page.wait_for_load_state("networkidle", timeout=8000)
    except Exception:
        page.wait_for_timeout(3000)

    after = page.inner_text("body")

    if debug_tag:
        os.makedirs(DEBUG_DIR, exist_ok=True)
        page.screenshot(path=os.path.join(DEBUG_DIR, f"{debug_tag}.png"))
        with open(os.path.join(DEBUG_DIR, f"{debug_tag}.txt"), "w", encoding="utf-8") as f:
            f.write("=== BEFORE ===\n" + before + "\n\n=== AFTER ===\n" + after)

    before_lines = set(l.strip() for l in before.split("\n") if l.strip())
    after_lines = [l.strip() for l in after.split("\n") if l.strip()]
    new_lines = [l for l in after_lines if l not in before_lines]

    entries = []
    seen_dates = set()
    for i, line in enumerate(new_lines):
        date_str, rest = find_date_and_rest(line)
        if not date_str or date_str in seen_dates:
            continue

        # Look for the rainfall number on the same line (after removing
        # the date itself) first, then the next couple of lines -- the
        # exact layout (same row vs. next row) isn't confirmed, see the
        # module docstring.
        value = None
        for text in [rest] + new_lines[i + 1:i + 3]:
            for match in re.findall(r"\d+\.\d+|\d+", text):
                val = float(match)
                if 0 <= val < 1000:
                    value = val
                    break
            if value is not None:
                break

        if value is not None:
            entries.append((date_str, value))
            seen_dates.add(date_str)

    return entries[:HISTORY_DAYS]


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
    """new_entries: list of (date_str, value) tuples NOT already in the
    file. Appends one row per entry and saves."""
    if os.path.exists(path):
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["Date", "Rainfall (mm)", "First recorded (UTC)"])

    for date_str, value in sorted(new_entries):
        ws.append([date_str, value, run_timestamp])

    wb.save(path)


def get_drive_service():
    """Build a Google Drive API client from a service account.
    Expects GOOGLE_SERVICE_ACCOUNT_JSON env var to contain the *contents*
    of the service account JSON key (not a file path) -- this is how the
    GitHub Actions workflow supplies it from a secret.
    """
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    creds_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    creds_dict = json.loads(creds_json)
    creds = service_account.Credentials.from_service_account_info(
        creds_dict, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)


def upload_to_drive(service, folder_id, filepath, filename):
    from googleapiclient.http import MediaFileUpload

    query = f"name='{filename}' and '{folder_id}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id)").execute()
    files = results.get("files", [])

    media = MediaFileUpload(
        filepath,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if files:
        service.files().update(fileId=files[0]["id"], media_body=media).execute()
    else:
        metadata = {"name": filename, "parents": [folder_id]}
        service.files().create(body=metadata, media_body=media, fields="id").execute()


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(STATIONS_FILE, encoding="utf-8") as f:
        stations = json.load(f)

    if DEBUG_STATIONS:
        stations = stations[:DEBUG_STATIONS]
        print(f"DEBUG MODE: running only {len(stations)} station(s)")

    drive_service = None
    folder_id = os.environ.get("DRIVE_FOLDER_ID")
    if not DEBUG_STATIONS and "GOOGLE_SERVICE_ACCOUNT_JSON" in os.environ and folder_id:
        drive_service = get_drive_service()

    run_timestamp = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")

    with sync_playwright() as p:
        # Always headless: this runs on GitHub's servers, which have no
        # display. Debug mode still works -- it saves a screenshot and
        # the before/after page text to output/debug/ so you can inspect
        # the result afterward instead of watching a live browser window.
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
                debug_tag = f"{url_id}_{re.sub(r'[^\w-]+', '_', name)}" if DEBUG_STATIONS else None
                history = extract_rainfall_history(page, debug_tag=debug_tag)
            except Exception as e:
                print(f"   failed: {e}")
                history = []

            print(f"   read {len(history)} day(s): {history}")

            existing_dates = get_existing_dates(filepath)
            new_entries = [(d, v) for d, v in history if d not in existing_dates]

            if new_entries:
                append_new_rows(filepath, new_entries, run_timestamp)
                print(f"   added {len(new_entries)} new day(s): {sorted(new_entries)}")
            else:
                print("   no new days to add (all already recorded, or nothing parsed)")

            if drive_service and new_entries:
                try:
                    upload_to_drive(drive_service, folder_id, filepath, filename)
                except Exception as e:
                    print(f"   drive upload failed: {e}")

            time.sleep(0.5)  # be polite to the server

        browser.close()

    print("Done.")


if __name__ == "__main__":
    main()
