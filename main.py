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
DEBUG_DIR = os.path.join(OUTPUT_DIR, "debug")

STATION_URL = "https://dhm.gov.np/hydrology/rainfallSingle/{}"

# Set DEBUG_STATIONS to a number such as 3 to test only the
# first 3 stations.
#
# If DEBUG_STATIONS is not set, all stations will be processed.
DEBUG_STATIONS = int(
    os.environ.get("DEBUG_STATIONS", "0")
)


# ============================================================
# DIRECTORIES
# ============================================================

os.makedirs(OUTPUT_DIR, exist_ok=True)

if DEBUG_STATIONS:
    os.makedirs(DEBUG_DIR, exist_ok=True)


# ============================================================
# LOAD STATIONS
# ============================================================

with open(
    STATIONS_FILE,
    "r",
    encoding="utf-8",
) as f:

    stations = json.load(f)


# ============================================================
# SAFE EXCEL FILENAME
# ============================================================

def safe_filename(
    name,
    station_index,
    url_id,
):

    base = station_index or url_id

    return f"{base}.xlsx"


# ============================================================
# EXTRACT DAILY RAINFALL
# ============================================================

def extract_daily_rainfall(
    page,
    debug_tag=None,
):

    # --------------------------------------------------------
    # Select Daily
    # --------------------------------------------------------

    daily_selectors = [
        "text=Daily",
        "label:has-text('Daily')",
        "[value='daily']",
        "input[value='Daily']",
        "input[value='daily']",
    ]

    daily_selected = False

    for selector in daily_selectors:

        try:

            locator = page.locator(selector)

            if locator.count() > 0:

                locator.first.click(
                    timeout=3000
                )

                daily_selected = True

                print(
                    f"   Selected Daily using: {selector}"
                )

                break

        except Exception:
            continue

    if not daily_selected:

        print(
            "   WARNING: Could not explicitly select Daily."
        )


    # --------------------------------------------------------
    # Click VIEW
    # --------------------------------------------------------

    view_selectors = [
        "text=VIEW",
        "button:has-text('VIEW')",
        "input[value='VIEW']",
        "input[value='View']",
        "button:has-text('View')",
    ]

    view_clicked = False

    for selector in view_selectors:

        try:

            locator = page.locator(selector)

            if locator.count() > 0:

                locator.first.click(
                    timeout=5000
                )

                view_clicked = True

                print(
                    f"   Clicked VIEW using: {selector}"
                )

                break

        except Exception:
            continue

    if not view_clicked:

        print(
            "   ERROR: Could not find/click VIEW."
        )

        return None


    # --------------------------------------------------------
    # Wait for data
    # --------------------------------------------------------

    try:

        page.wait_for_load_state(
            "networkidle",
            timeout=8000,
        )

    except Exception:
        pass

    time.sleep(5)


    # --------------------------------------------------------
    # Get page text
    # --------------------------------------------------------

    body_text = page.locator(
        "body"
    ).inner_text()


    # --------------------------------------------------------
    # Save debug screenshot and text
    # --------------------------------------------------------

    if debug_tag:

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

        except Exception as e:

            print(
                f"   WARNING: Could not save screenshot: {e}"
            )

        try:

            with open(
                text_path,
                "w",
                encoding="utf-8",
            ) as f:

                f.write(body_text)

        except Exception as e:

            print(
                f"   WARNING: Could not save debug text: {e}"
            )


    # --------------------------------------------------------
    # Find Accumulated Rainfall section
    # --------------------------------------------------------

    lines = body_text.splitlines()

    accumulated_index = None

    for i, line in enumerate(lines):

        if "Accumulated Rainfall" in line:

            accumulated_index = i

            break


    if accumulated_index is None:

        print(
            "   ERROR: 'Accumulated Rainfall' section not found."
        )

        return None


    print(
        "   Found Accumulated Rainfall section."
    )


    # --------------------------------------------------------
    # Extract rainfall values
    # --------------------------------------------------------
    #
    # The DHM table contains rainfall observations over time.
    #
    # We extract decimal values from the accumulated rainfall
    # line and use the maximum value as the 24-hour accumulated
    # rainfall figure.
    # --------------------------------------------------------

    line = lines[accumulated_index]

    matches = re.findall(
        r"(?<![\d.])\d+\.\d{1,2}(?![\d.])",
        line,
    )


    # Sometimes the values may appear in the lines immediately
    # following the heading, so if the heading line itself does
    # not contain numbers, inspect the next few lines.

    if not matches:

        nearby_lines = lines[
            accumulated_index:
            accumulated_index + 15
        ]

        nearby_text = " ".join(
            nearby_lines
        )

        matches = re.findall(
            r"(?<![\d.])\d+\.\d{1,2}(?![\d.])",
            nearby_text,
        )


    rainfall_values = []

    for value_text in matches:

        try:

            value = float(value_text)

            if 0 <= value <= 1000:

                rainfall_values.append(
                    value
                )

        except ValueError:

            continue


    if not rainfall_values:

        print(
            "   ERROR: No valid rainfall values found."
        )

        return None


    total = max(
        rainfall_values
    )


    print(
        f"   Rainfall values found: {rainfall_values}"
    )

    print(
        f"   24h rainfall selected: {total:.2f} mm"
    )


    return total


# ============================================================
# WRITE / UPDATE EXCEL
# ============================================================

def write_station_excel(
    path,
    date_str,
    value,
):

    # --------------------------------------------------------
    # Open existing workbook or create new workbook
    # --------------------------------------------------------

    if os.path.exists(path):

        wb = load_workbook(path)

        ws = wb.active

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
    # Check whether today's date already exists
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

    # --------------------------------------------------------
    # Otherwise append new date
    # --------------------------------------------------------

    else:

        ws.append(
            [
                date_str,
                value,
            ]
        )


    # --------------------------------------------------------
    # Number formatting
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
# GOOGLE DRIVE CONNECTION
# ============================================================

def get_drive_service():

    from google.oauth2 import service_account
    from googleapiclient.discovery import build


    # --------------------------------------------------------
    # Read service-account JSON from GitHub Secret
    # --------------------------------------------------------

    creds_json = os.environ[
        "GOOGLE_SERVICE_ACCOUNT_JSON"
    ]


    # --------------------------------------------------------
    # Convert JSON string into dictionary
    # --------------------------------------------------------

    creds_dict = json.loads(
        creds_json
    )


    # --------------------------------------------------------
    # Authenticate
    # --------------------------------------------------------

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


    # --------------------------------------------------------
    # Build Google Drive API service
    # --------------------------------------------------------

    return build(
        "drive",
        "v3",
        credentials=creds,
    )


# ============================================================
# UPLOAD / UPDATE FILE IN GOOGLE DRIVE
# ============================================================

def upload_to_drive(
    service,
    folder_id,
    filepath,
    filename,
):

    from googleapiclient.http import MediaFileUpload


    # --------------------------------------------------------
    # Search for an existing file with the same filename
    # inside the target folder.
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
    # Prepare Excel file for upload
    # --------------------------------------------------------

    media = MediaFileUpload(
        filepath,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


    # --------------------------------------------------------
    # UPDATE existing file
    # --------------------------------------------------------

    if files:

        file_id = files[0]["id"]

        print(
            f"   Updating Google Drive file: {filename}"
        )


        (
            service.files()
            .update(
                fileId=file_id,
                media_body=media,
            )
            .execute()
        )


        print(
            f"   SUCCESS: {filename} updated in Google Drive."
        )


    # --------------------------------------------------------
    # CREATE new file
    # --------------------------------------------------------

    else:

        print(
            f"   Creating Google Drive file: {filename}"
        )


        metadata = {
            "name": filename,
            "parents": [folder_id],
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


        print(
            f"   SUCCESS: {filename} created in Google Drive."
        )


# ============================================================
# MAIN
# ============================================================

def main():

    print(
        "=================================================="
    )

    print(
        "DHM Rainfall Extraction"
    )

    print(
        "=================================================="
    )


    # --------------------------------------------------------
    # Determine number of stations
    # --------------------------------------------------------

    if DEBUG_STATIONS:

        stations_to_process = stations[
            :DEBUG_STATIONS
        ]

        print(
            f"DEBUG MODE: Processing first "
            f"{DEBUG_STATIONS} stations."
        )

    else:

        stations_to_process = stations

        print(
            f"PRODUCTION MODE: Processing all "
            f"{len(stations_to_process)} stations."
        )


    # --------------------------------------------------------
    # Nepal date/time
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


    date_str = now.strftime(
        "%Y/%m/%d"
    )


    print(
        f" Nepal date: {date_str}"
    )


    # --------------------------------------------------------
    # Google Drive setup
    # --------------------------------------------------------

    drive_service = None


    folder_id = os.environ.get(
        "DRIVE_FOLDER_ID"
    )


    google_credentials = os.environ.get(
        "GOOGLE_SERVICE_ACCOUNT_JSON"
    )


    if google_credentials and folder_id:

        print(
            "Connecting to Google Drive..."
        )


        try:

            drive_service = (
                get_drive_service()
            )


            print(
                "Google Drive connection successful."
            )


        except Exception as e:

            print(
                f"Google Drive connection failed: {e}"
            )

            drive_service = None


    else:

        print(
            "Google Drive upload is NOT available."
        )

        if not google_credentials:

            print(
                "   Missing secret: GOOGLE_SERVICE_ACCOUNT_JSON"
            )

        if not folder_id:

            print(
                "   Missing secret: DRIVE_FOLDER_ID"
            )


    # --------------------------------------------------------
    # Start Playwright
    # --------------------------------------------------------

    with sync_playwright() as p:

        browser = p.chromium.launch(
            headless=True
        )


        page = browser.new_page()


        # ----------------------------------------------------
        # Process stations
        # ----------------------------------------------------

        for station_number, station in enumerate(
            stations_to_process,
            start=1,
        ):

            # ------------------------------------------------
            # Read station information
            # ------------------------------------------------

            url_id = str(
                station.get(
                    "url_id",
                    ""
                )
            ).strip()


            station_index = str(
                station.get(
                    "station_index",
                    ""
                )
            ).strip()


            station_name = str(
                station.get(
                    "station_name",
                    station.get(
                        "name",
                        ""
                    ),
                )
            ).strip()


            if not url_id:

                print(
                    f"\n[{station_number}] "
                    f"ERROR: Station has no url_id."
                )

                continue


            # ------------------------------------------------
            # Display station
            # ------------------------------------------------

            print(
                "\n--------------------------------------------------"
            )

            print(
                f"[{station_number}/{len(stations_to_process)}] "
                f"{station_index} - {station_name}"
            )

            print(
                f"URL ID: {url_id}"
            )


            # ------------------------------------------------
            # Construct DHM URL
            # ------------------------------------------------

            url = STATION_URL.format(
                url_id
            )


            print(
                f"Opening: {url}"
            )


            # ------------------------------------------------
            # Debug filename
            # ------------------------------------------------

            debug_tag = None

            if DEBUG_STATIONS:

                debug_tag = (
                    f"{station_index}_{url_id}"
                )


            # ------------------------------------------------
            # Open station page
            # ------------------------------------------------

            try:

                page.goto(
                    url,
                    wait_until="domcontentloaded",
                    timeout=30000,
                )

            except Exception as e:

                print(
                    f"   ERROR opening page: {e}"
                )

                continue


            # ------------------------------------------------
            # Extract rainfall
            # ------------------------------------------------

            try:

                value = extract_daily_rainfall(
                    page,
                    debug_tag=debug_tag,
                )

            except Exception as e:

                print(
                    f"   ERROR extracting rainfall: {e}"
                )

                value = None


            # ------------------------------------------------
            # If rainfall was successfully extracted
            # ------------------------------------------------

            if value is not None:

                # --------------------------------------------
                # Excel filename
                # --------------------------------------------

                filename = safe_filename(
                    station_name,
                    station_index,
                    url_id,
                )


                # --------------------------------------------
                # Local Excel path
                # --------------------------------------------

                filepath = os.path.join(
                    OUTPUT_DIR,
                    filename,
                )


                # --------------------------------------------
                # Write/update Excel
                # --------------------------------------------

                try:

                    write_station_excel(
                        filepath,
                        date_str,
                        value,
                    )


                    print(
                        f"   Excel saved: {filepath}"
                    )


                except Exception as e:

                    print(
                        f"   ERROR saving Excel: {e}"
                    )

                    continue


                # --------------------------------------------
                # Upload/update Google Drive
                # --------------------------------------------

                if (
                    drive_service
                    and folder_id
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
                            "   Drive upload failed: "
                            f"{e}"
                        )


                else:

                    print(
                        "   Google Drive upload skipped."
                    )


            else:

                print(
                    "   No rainfall value extracted."
                )


        # ----------------------------------------------------
        # Close browser
        # ----------------------------------------------------

        browser.close()


    # --------------------------------------------------------
    # Finish
    # --------------------------------------------------------

    print(
        "\n=================================================="
    )

    print(
        "Rainfall extraction completed."
    )

    print(
        "=================================================="
    )


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    main()