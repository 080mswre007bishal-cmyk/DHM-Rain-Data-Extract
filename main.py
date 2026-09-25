
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

IMPORTANT:
Run with DEBUG_STATIONS=3 first. Check the screenshots and text files
in output/debug/ to confirm that the extracted number is actually the
24-hour rainfall figure displayed on the DHM page.

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

# Set via environment variable for a quick test:
# DEBUG_STATIONS=3 python main.py
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

    # Capture page text before selecting Daily and clicking VIEW.
    before = page.inner_text("body")

    # Select the Daily period.
    # Try a few likely ways the site exposes this element.
    for selector in [
        "text=Daily",
        "label:has-text('Daily')",
        "[value='daily']",
    ]:
        try:
            page.click(selector, timeout=2000)
            break
        except Exception:
            continue

    # Click VIEW.
    for selector in [
        "text=VIEW",
        "button:has-text('VIEW')",
        "input[value='VIEW']",
    ]:
        try:
            page.click(selector, timeout=3000)
            break
        except Exception:
            continue

    # Give the AJAX request time to complete.
    try:
        page.wait_for_load_state("networkidle", timeout=8000)
    except Exception:
        page.wait_for_timeout(3000)

    # Capture page text after clicking VIEW.
    after = page.inner_text("body")

    # Save screenshot and page text in debug mode.
    if debug_tag:
        os.makedirs(DEBUG_DIR, exist_ok=True)

        screenshot_path = os.path.join(
            DEBUG_DIR, f"{debug_tag}.png"
        )

        text_path = os.path.join(
            DEBUG_DIR, f"{debug_tag}.txt"
        )

        page.screenshot(path=screenshot_path)

        with open(text_path, "w", encoding="utf-8") as f:
            f.write(
                "=== BEFORE ===\n"
                + before
                + "\n\n=== AFTER ===\n"
                + after
            )

    # Identify lines that appeared after clicking VIEW.
    before_lines = set(
        line.strip()
        for line in before.split("\n")
        if line.strip()
    )

    after_lines = [
        line.strip()
        for line in after.split("\n")
        if line.strip()
    ]

    new_lines = [
        line
        for line in after_lines
        if line not in before_lines
    ]

    # Find plausible rainfall numbers in the newly appeared text.
    # NOTE: This is a best-effort method and must be verified
    # against the screenshots and saved page text.
    candidates = []

    for line in new_lines:
        matches = re.findall(r"\d+\.\d+|\d+", line)

        for match in matches:
            val = float(match)

            if 0 <= val < 1000:
                candidates.append(val)

    return candidates[0] if candidates else None


def write_station_excel(path, timestamp, date_str, value):
    """Append a rainfall record to the station's Excel file."""

    if os.path.exists(path):
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append([
            "Timestamp (UTC)",
            "Date",
            "24h Rainfall (mm)",
        ])

    ws.append([timestamp, date_str, value])
    wb.save(path)


def get_drive_service():
    """
    Build a Google Drive API client from a service account.

    Expects GOOGLE_SERVICE_ACCOUNT_JSON to contain the contents
    of the service account JSON key.
    """

    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    creds_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    creds_dict = json.loads(creds_json)

    creds = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/drive"],
    )

    return build("drive", "v3", credentials=creds)


def upload_to_drive(service, folder_id, filepath, filename):
    """Upload a new file or update an existing file in Google Drive."""

    from googleapiclient.http import MediaFileUpload

    query = (
        f"name='{filename}' and "
        f"'{folder_id}' in parents and trashed=false"
    )

    results = (
        service.files()
        .list(q=query, fields="files(id)")
        .execute()
    )

    files = results.get("files", [])

    media = MediaFileUpload(
        filepath,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    if files:
        service.files().update(
            fileId=files[0]["id"],
            media_body=media,
        ).execute()
    else:
        metadata = {
            "name": filename,
            "parents": [folder_id],
        }

        service.files().create(
            body=metadata,
            media_body=media,
            fields="id",
        ).execute()


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(STATIONS_FILE, encoding="utf-8") as f:
        stations = json.load(f)

    # In debug mode, process only the requested number of stations.
    if DEBUG_STATIONS:
        stations = stations[:DEBUG_STATIONS]

        # Clear old debug files so the artifact contains only
        # results from the current run.
        os.makedirs(DEBUG_DIR, exist_ok=True)

        for filename in os.listdir(DEBUG_DIR):
            filepath = os.path.join(DEBUG_DIR, filename)

            if os.path.isfile(filepath):
                os.remove(filepath)

        print(
            f"DEBUG MODE: running only "
            f"{len(stations)} station(s)"
        )

    # Google Drive is disabled in debug mode.
    drive_service = None
    folder_id = os.environ.get("DRIVE_FOLDER_ID")

    if (
        not DEBUG_STATIONS
        and "GOOGLE_SERVICE_ACCOUNT_JSON" in os.environ
        and folder_id
    ):
        drive_service = get_drive_service()

    now = datetime.datetime.utcnow()

    timestamp = now.strftime("%Y-%m-%d %H:%M:%S")
    date_str = now.strftime("%Y-%m-%d")

    with sync_playwright() as p:

        # GitHub Actions runs without a display, so use headless mode.
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        for station in stations:
            url_id = station["url_id"]
            name = station["name"]

            print(f"-> {name} (url_id={url_id})")

            try:
                page.goto(
                    STATION_URL.format(url_id),
                    timeout=20000,
                )

                # FIX: Build the cleaned station name separately.
                # This avoids backslashes inside an f-string
                # expression, which cause a SyntaxError in Python 3.11.
                clean_name = re.sub(
                    r"[^\w-]+",
                    "_",
                    name,
                )

                debug_tag = (
                    f"{url_id}_{clean_name}"
                    if DEBUG_STATIONS
                    else None
                )

                value = extract_daily_rainfall(
                    page,
                    debug_tag=debug_tag,
                )

            except Exception as e:
                print(f"   failed: {e}")
                value = None

            print(f"   24h rainfall: {value}")

            filename = safe_filename(
                name,
                station.get("station_index"),
                url_id,
            )

            filepath = os.path.join(
                OUTPUT_DIR,
                filename,
            )

            write_station_excel(
                filepath,
                timestamp,
                date_str,
                value,
            )

            # Upload only during normal production runs.
            if drive_service:
                try:
                    upload_to_drive(
                        drive_service,
                        folder_id,
                        filepath,
                        filename,
                    )
                except Exception as e:
                    print(f"   drive upload failed: {e}")

            time.sleep(0.5)  # Be polite to the server.

        browser.close()

    print("Done.")


if __name__ == "__main__":
    main()