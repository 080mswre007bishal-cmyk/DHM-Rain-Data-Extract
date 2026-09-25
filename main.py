"""
main.py
-------
For every station in stations.json:
  1. Open its DHM page in a headless browser.
  2. Select the "Daily" period and click VIEW to trigger the 24-hour
     accumulated rainfall reading (this is loaded by JavaScript, so a
     real browser is used instead of a plain HTTP request).
  3. Read the value that appears.
  4. Append it as a new row to output/<station>.xlsx (one file per
     station, so it becomes a running history over time).
  5. Upload/update that file in a Google Drive folder.

IMPORTANT — please read before relying on this in production:
I built this against the *static* HTML of the DHM pages (confirmed by
fetching them), but I could not load the live JavaScript app myself to
watch exactly what the page looks like right after you click VIEW —
my environment can't reach dhm.gov.np, and the underlying data call
isn't visible to a text-based fetch. So step 3 (extract_daily_rainfall)
uses a generic, resilient approach: it snapshots the page text before
and after clicking VIEW and pulls the new number(s) out of the diff,
rather than depending on one exact CSS selector that might be wrong.

Run with DEBUG_STATIONS=3 first (see bottom of file) to check a few
stations, look at the saved screenshots/text in output/debug/, and
confirm the number it picked is actually the 24h rainfall figure. If
it's picking the wrong number (e.g. a date or station index), adjust
`extract_daily_rainfall()` — that's the one function you may need to
tune, everything else (Excel writing, Drive upload, scheduling) is
independent of that and doesn't need touching.

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

# Set via env var for a quick local test, e.g. DEBUG_STATIONS=3 python main.py
DEBUG_STATIONS = int(os.environ.get("DEBUG_STATIONS", "0"))


def safe_filename(name, station_index, url_id):
    base = station_index or url_id
    clean_name = re.sub(r"[^\w\-]+", "_", name).strip("_")
    return f"{base}_{clean_name}.xlsx"


def extract_daily_rainfall(page, debug_tag=None):
    """
    Best-effort extraction of the 24h accumulated rainfall value.
    Returns a float (mm) or None if nothing usable was found.
    """
    before = page.inner_text("body")

    # Select "Daily" period — try a few likely ways the site exposes this,
    # since the exact element (radio/tab/button) wasn't confirmed live.
    for selector in ["text=Daily", "label:has-text('Daily')", "[value='daily']"]:
        try:
            page.click(selector, timeout=2000)
            break
        except Exception:
            continue

    # Click VIEW
    for selector in ["text=VIEW", "button:has-text('VIEW')", "input[value='VIEW']"]:
        try:
            page.click(selector, timeout=3000)
            break
        except Exception:
            continue

    # Give the AJAX call time to complete
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

    # Look for a plausible rainfall number (decimal, 0-999 mm range) in
    # the newly-appeared text.
    candidates = []
    for line in new_lines:
        for match in re.findall(r"\d+\.\d+|\d+", line):
            val = float(match)
            if 0 <= val < 1000:
                candidates.append(val)

    return candidates[0] if candidates else None


def write_station_excel(path, timestamp, date_str, value):
    if os.path.exists(path):
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["Timestamp (UTC)", "Date", "24h Rainfall (mm)"])

    ws.append([timestamp, date_str, value])
    wb.save(path)


def get_drive_service():
    """Build a Google Drive API client from a service account.
    Expects GOOGLE_SERVICE_ACCOUNT_JSON env var to contain the *contents*
    of the service account JSON key (not a file path) — this is how the
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

    now = datetime.datetime.utcnow()
    timestamp = now.strftime("%Y-%m-%d %H:%M:%S")
    date_str = now.strftime("%Y-%m-%d")

    with sync_playwright() as p:
        # Always headless: this runs on GitHub's servers, which have no
        # display. Debug mode still works — it saves a screenshot and the
        # before/after page text to output/debug/ so you can inspect the
        # result afterward instead of watching a live browser window.
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        for station in stations:
            url_id = station["url_id"]
            name = station["name"]
            print(f"-> {name} (url_id={url_id})")

            try:
                page.goto(STATION_URL.format(url_id), timeout=20000)
                debug_tag = f"{url_id}_{re.sub(r'[^\w-]+', '_', name)}" if DEBUG_STATIONS else None
                value = extract_daily_rainfall(page, debug_tag=debug_tag)
            except Exception as e:
                print(f"   failed: {e}")
                value = None

            print(f"   24h rainfall: {value}")

            filename = safe_filename(name, station.get("station_index"), url_id)
            filepath = os.path.join(OUTPUT_DIR, filename)
            write_station_excel(filepath, timestamp, date_str, value)

            if drive_service:
                try:
                    upload_to_drive(drive_service, folder_id, filepath, filename)
                except Exception as e:
                    print(f"   drive upload failed: {e}")

            time.sleep(0.5)  # be polite to the server

        browser.close()

    print("Done.")


if __name__ == "__main__":
    main()
