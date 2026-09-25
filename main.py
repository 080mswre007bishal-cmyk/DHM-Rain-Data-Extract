"""
main.py
-------
For every station in stations.json:

  1. Open the DHM rainfall station page.
  2. Select "Daily".
  3. Click "VIEW".
  4. Read the "Accumulated Rainfall" table.
  5. Identify the maximum accumulated rainfall value for the
     24-hour period.
  6. Save one record per date to output/<station_index>.xlsx.
  7. Upload/update that file in Google Drive.

Example:

    Station index: 1009
    Date:          2026/09/25
    24h rainfall:  49.20 mm

Output:

    output/1009.xlsx

Excel:

    Date          24h Rainfall (mm)
    2026/09/25    49.20

Debug:

    DEBUG_STATIONS=3

This creates screenshots and text files in:

    output/debug/

Usage locally:

    pip install playwright openpyxl google-api-python-client google-auth
    playwright install chromium

    python main.py

Debug:

    DEBUG_STATIONS=3 python main.py
"""

import datetime
import json
import os
import re
import time

from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright


# ============================================================
# SETTINGS
# ============================================================

STATIONS_FILE = "stations.json"

OUTPUT_DIR = "output"

DEBUG_DIR = os.path.join(
    OUTPUT_DIR,
    "debug",
)

STATION_URL = (
    "https://dhm.gov.np/hydrology/rainfallSingle/{}"
)

# Set DEBUG_STATIONS=3 to process only the first 3 stations.
DEBUG_STATIONS = int(
    os.environ.get(
        "DEBUG_STATIONS",
        "0",
    )
)


# ============================================================
# FILE NAME
# ============================================================

def safe_filename(name, station_index, url_id):
    """
    Use station index as the Excel filename.

    Example:
        station_index = 1009
        result = 1009.xlsx

    If station_index is unavailable, url_id is used.
    """

    base = station_index or url_id

    return f"{base}.xlsx"


# ============================================================
# EXTRACT DAILY RAINFALL
# ============================================================

def extract_daily_rainfall(page, debug_tag=None):
    """
    Extract the 24-hour accumulated rainfall from the DHM page.

    The DHM page contains an "Accumulated Rainfall" section.
    The accumulated rainfall values increase backwards through
    the displayed 24-hour period.

    Example:

        08:35 PM     0.20
        08:25 PM     0.40
        ...
        05:55 AM    48.80
        05:45 AM    49.20

    Therefore the maximum accumulated value is selected as
    the 24-hour rainfall.

    Returns:
        float
            24-hour rainfall in mm.

        None
            If no reliable value is found.
    """

    print("   Selecting Daily period...")

    # --------------------------------------------------------
    # Select DAILY
    # --------------------------------------------------------

    daily_selected = False

    daily_selectors = [
        "text=Daily",
        "label:has-text('Daily')",
        "[value='daily']",
        "input[value='Daily']",
        "input[value='daily']",
    ]

    for selector in daily_selectors:

        try:

            page.click(
                selector,
                timeout=3000,
            )

            daily_selected = True

            print(
                f"   Daily selected using: {selector}"
            )

            break

        except Exception:
            continue

    if not daily_selected:

        print(
            "   Warning: Could not explicitly select Daily."
        )

    # --------------------------------------------------------
    # Click VIEW
    # --------------------------------------------------------

    print("   Clicking VIEW...")

    view_clicked = False

    view_selectors = [
        "text=VIEW",
        "button:has-text('VIEW')",
        "input[value='VIEW']",
        "input[value='View']",
        "button:has-text('View')",
    ]

    for selector in view_selectors:

        try:

            page.click(
                selector,
                timeout=5000,
            )

            view_clicked = True

            print(
                f"   VIEW clicked using: {selector}"
            )

            break

        except Exception:
            continue

    if not view_clicked:

        print(
            "   Warning: Could not explicitly click VIEW."
        )

    # --------------------------------------------------------
    # Wait for JavaScript/AJAX data
    # --------------------------------------------------------

    print("   Waiting for rainfall data...")

    try:

        page.wait_for_load_state(
            "networkidle",
            timeout=8000,
        )

    except Exception:

        pass

    # Give the DHM JavaScript application additional time.
    page.wait_for_timeout(5000)

    # --------------------------------------------------------
    # Capture page text
    # --------------------------------------------------------

    try:

        body_text = page.inner_text(
            "body"
        )

    except Exception as e:

        print(
            f"   Could not read page text: {e}"
        )

        return None

    # --------------------------------------------------------
    # Save debug information
    # --------------------------------------------------------

    if debug_tag:

        os.makedirs(
            DEBUG_DIR,
            exist_ok=True,
        )

        screenshot_path = os.path.join(
            DEBUG_DIR,
            f"{debug_tag}.png",
        )

        text_path = os.path.join(
            DEBUG_DIR,
            f"{debug_tag}.txt",
        )

        try:

            page.screenshot(
                path=screenshot_path,
                full_page=True,
            )

            print(
                f"   Debug screenshot saved: "
                f"{screenshot_path}"
            )

        except Exception as e:

            print(
                f"   Screenshot failed: {e}"
            )

        try:

            with open(
                text_path,
                "w",
                encoding="utf-8",
            ) as f:

                f.write(body_text)

            print(
                f"   Debug text saved: {text_path}"
            )

        except Exception as e:

            print(
                f"   Debug text save failed: {e}"
            )

    # --------------------------------------------------------
    # Split page into lines
    # --------------------------------------------------------

    lines = [
        line.strip()
        for line in body_text.splitlines()
        if line.strip()
    ]

    # --------------------------------------------------------
    # Locate "Accumulated Rainfall"
    # --------------------------------------------------------

    accumulated_index = None

    for i, line in enumerate(lines):

        if "Accumulated Rainfall" in line:

            accumulated_index = i

            break

    if accumulated_index is None:

        print(
            "   ERROR: 'Accumulated Rainfall' "
            "section was not found."
        )

        return None

    print(
        f"   Found 'Accumulated Rainfall' "
        f"at line {accumulated_index}."
    )

    # Everything after the accumulated rainfall heading.
    accumulated_lines = lines[
        accumulated_index:
    ]

    # --------------------------------------------------------
    # Extract decimal rainfall values
    # --------------------------------------------------------

    rainfall_values = []

    for line in accumulated_lines:

        # Find decimal numbers such as:
        #
        # 0.20
        # 49.20
        # 12.50
        #
        # We deliberately do not use integers because dates,
        # hours and other page numbers could otherwise be picked.
        matches = re.findall(
            r"(?<![\d.])\d+\.\d{1,2}(?![\d.])",
            line,
        )

        for match in matches:

            try:

                value = float(match)

            except ValueError:

                continue

            # Reasonable rainfall range.
            if 0 <= value <= 1000:

                rainfall_values.append(
                    value
                )

    # --------------------------------------------------------
    # Check if rainfall values were found
    # --------------------------------------------------------

    if not rainfall_values:

        print(
            "   ERROR: No rainfall values were found "
            "in the Accumulated Rainfall section."
        )

        return None

    # --------------------------------------------------------
    # Select final 24-hour accumulated rainfall
    # --------------------------------------------------------

    total_rainfall = max(
        rainfall_values
    )

    print(
        f"   Number of rainfall values found: "
        f"{len(rainfall_values)}"
    )

    print(
        f"   24h rainfall selected: "
        f"{total_rainfall:.2f} mm"
    )

    return total_rainfall


# ============================================================
# WRITE EXCEL
# ============================================================

def write_station_excel(
    path,
    date_str,
    value,
):
    """
    Write one rainfall record per date.

    If the date already exists, update the existing value.

    Example:

        Date          24h Rainfall (mm)
        2026/09/25    49.20
    """

    # --------------------------------------------------------
    # Open existing workbook
    # --------------------------------------------------------

    if os.path.exists(path):

        wb = load_workbook(
            path
        )

        ws = wb.active

    # --------------------------------------------------------
    # Create new workbook
    # --------------------------------------------------------

    else:

        wb = Workbook()

        ws = wb.active

        ws.title = "Rainfall"

        ws.append(
            [
                "Date",
                "24h Rainfall (mm)",
            ]
        )

    # --------------------------------------------------------
    # Find whether this date already exists
    # --------------------------------------------------------

    existing_row = None

    for row in range(
        2,
        ws.max_row + 1,
    ):

        existing_date = ws.cell(
            row=row,
            column=1,
        ).value

        if existing_date is None:
            continue

        existing_date = str(
            existing_date
        ).strip()

        if existing_date == date_str:

            existing_row = row

            break

    # --------------------------------------------------------
    # Update existing date
    # --------------------------------------------------------

    if existing_row is not None:

        ws.cell(
            row=existing_row,
            column=2,
            value=value,
        )

        print(
            f"   Updated Excel record: "
            f"{date_str} = {value:.2f} mm"
        )

    # --------------------------------------------------------
    # Add new date
    # --------------------------------------------------------

    else:

        ws.append(
            [
                date_str,
                value,
            ]
        )

        print(
            f"   Added Excel record: "
            f"{date_str} = {value:.2f} mm"
        )

    # --------------------------------------------------------
    # Format rainfall cell
    # --------------------------------------------------------

    for row in range(
        2,
        ws.max_row + 1,
    ):

        ws.cell(
            row=row,
            column=2,
        ).number_format = "0.00"

    # --------------------------------------------------------
    # Save workbook
    # --------------------------------------------------------

    wb.save(path)


# ============================================================
# GOOGLE DRIVE
# ============================================================

def get_drive_service():
    """
    Build a Google Drive API client from a service account.

    The GitHub secret must contain the complete service-account
    JSON.

    Environment variable:

        GOOGLE_SERVICE_ACCOUNT_JSON
    """

    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    creds_json = os.environ[
        "GOOGLE_SERVICE_ACCOUNT_JSON"
    ]

    creds_dict = json.loads(
        creds_json
    )

    creds = (
        service_account
        .Credentials
        .from_service_account_info(
            creds_dict,
            scopes=[
                "https://www.googleapis.com/auth/drive"
            ],
        )
    )

    return build(
        "drive",
        "v3",
        credentials=creds,
    )


def upload_to_drive(
    service,
    folder_id,
    filepath,
    filename,
):
    """
    Upload a new file or update an existing file
    in Google Drive.
    """

    from googleapiclient.http import (
        MediaFileUpload
    )

    # --------------------------------------------------------
    # Find existing file
    # --------------------------------------------------------

    query = (
        f"name='{filename}' and "
        f"'{folder_id}' in parents and "
        f"trashed=false"
    )

    results = (
        service.files()
        .list(
            q=query,
            fields="files(id,name)",
        )
        .execute()
    )

    files = results.get(
        "files",
        []
    )

    # --------------------------------------------------------
    # Prepare Excel file
    # --------------------------------------------------------

    media = MediaFileUpload(
        filepath,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    # --------------------------------------------------------
    # Update existing file
    # --------------------------------------------------------

    if files:

        file_id = files[0]["id"]

        print(
            f"   Updating Google Drive file: "
            f"{filename}"
        )

        (
            service.files()
            .update(
                fileId=file_id,
                media_body=media,
            )
            .execute()
        )

    # --------------------------------------------------------
    # Create new file
    # --------------------------------------------------------

    else:

        print(
            f"   Creating Google Drive file: "
            f"{filename}"
        )

        metadata = {
            "name": filename,
            "parents": [
                folder_id
            ],
        }

        (
            service.files()
            .create(
                body=metadata,
                media_body=media,
                fields="id",
            )
            .execute()
        )


# ============================================================
# MAIN
# ============================================================

def main():

    # --------------------------------------------------------
    # Create output directory
    # --------------------------------------------------------

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True,
    )

    # --------------------------------------------------------
    # Load stations
    # --------------------------------------------------------

    with open(
        STATIONS_FILE,
        encoding="utf-8",
    ) as f:

        stations = json.load(f)

    # --------------------------------------------------------
    # DEBUG MODE
    # --------------------------------------------------------

    if DEBUG_STATIONS:

        stations = stations[
            :DEBUG_STATIONS
        ]

        os.makedirs(
            DEBUG_DIR,
            exist_ok=True,
        )

        # Remove previous debug files.
        for filename in os.listdir(
            DEBUG_DIR
        ):

            filepath = os.path.join(
                DEBUG_DIR,
                filename,
            )

            if os.path.isfile(
                filepath
            ):

                os.remove(
                    filepath
                )

        print(
            f"DEBUG MODE: running only "
            f"{len(stations)} station(s)"
        )

    # --------------------------------------------------------
    # Google Drive
    # --------------------------------------------------------

    drive_service = None

    folder_id = os.environ.get(
        "DRIVE_FOLDER_ID"
    )

    if (
        not DEBUG_STATIONS
        and "GOOGLE_SERVICE_ACCOUNT_JSON"
        in os.environ
        and folder_id
    ):

        print(
            "Connecting to Google Drive..."
        )

        drive_service = (
            get_drive_service()
        )

    # --------------------------------------------------------
    # Nepal time
    # --------------------------------------------------------

    nepal_timezone = datetime.timezone(
        datetime.timedelta(
            hours=5,
            minutes=45,
        )
    )

    now = datetime.datetime.now(
        nepal_timezone
    )

    # Example:
    #
    # 2026/09/25
    #

    date_str = now.strftime(
        "%Y/%m/%d"
    )

    print(
        f"Nepal date: {date_str}"
    )

    # --------------------------------------------------------
    # Start browser
    # --------------------------------------------------------

    with sync_playwright() as p:

        browser = p.chromium.launch(
            headless=True
        )

        page = browser.new_page()

        # ----------------------------------------------------
        # Process each station
        # ----------------------------------------------------

        for station in stations:

            url_id = station[
                "url_id"
            ]

            name = station[
                "name"
            ]

            station_index = station.get(
                "station_index"
            )

            print()
            print(
                "========================================"
            )

            print(
                f"Station: {name}"
            )

            print(
                f"URL ID: {url_id}"
            )

            print(
                f"Station index: "
                f"{station_index}"
            )

            print(
                "========================================"
            )

            value = None

            try:

                # ------------------------------------------------
                # Open DHM page
                # ------------------------------------------------

                page.goto(
                    STATION_URL.format(
                        url_id
                    ),
                    timeout=20000,
                    wait_until="domcontentloaded",
                )

                # Allow page JavaScript to initialise.
                page.wait_for_timeout(
                    3000
                )

                # ------------------------------------------------
                # Debug filename
                # ------------------------------------------------

                clean_name = re.sub(
                    r"[^\w-]+",
                    "_",
                    name,
                ).strip("_")

                debug_tag = (
                    f"{url_id}_{clean_name}"
                    if DEBUG_STATIONS
                    else None
                )

                # ------------------------------------------------
                # Extract rainfall
                # ------------------------------------------------

                value = extract_daily_rainfall(
                    page,
                    debug_tag=debug_tag,
                )

            except Exception as e:

                print(
                    f"   FAILED: {e}"
                )

                value = None

            # ----------------------------------------------------
            # Print result
            # ----------------------------------------------------

            if value is not None:

                print(
                    f"   FINAL 24h RAINFALL: "
                    f"{value:.2f} mm"
                )

            else:

                print(
                    "   FINAL 24h RAINFALL: "
                    "NOT FOUND"
                )

            # ----------------------------------------------------
            # Excel filename
            # ----------------------------------------------------

            filename = safe_filename(
                name,
                station_index,
                url_id,
            )

            filepath = os.path.join(
                OUTPUT_DIR,
                filename,
            )

            print(
                f"   Excel file: {filepath}"
            )

            # ----------------------------------------------------
            # Only write if rainfall was successfully extracted
            # ----------------------------------------------------

            if value is not None:

                write_station_excel(
                    filepath,
                    date_str,
                    value,
                )

            else:

                print(
                    "   Excel NOT updated because "
                    "rainfall value was not found."
                )

            # ----------------------------------------------------
            # Google Drive
            # ----------------------------------------------------

            if (
                drive_service
                and value is not None
            ):

                try:

                    upload_to_drive(
                        drive_service,
                        folder_id,
                        filepath,
                        filename,
                    )

                except Exception as e:

                    print(
                        f"   Drive upload failed: "
                        f"{e}"
                    )

            # ----------------------------------------------------
            # Small delay between stations
            # ----------------------------------------------------

            time.sleep(
                0.5
            )

        # --------------------------------------------------------
        # Close browser
        # --------------------------------------------------------

        browser.close()

    print()
    print(
        "========================================"
    )
    print(
        "Done."
    )
    print(
        "========================================"
    )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    main()