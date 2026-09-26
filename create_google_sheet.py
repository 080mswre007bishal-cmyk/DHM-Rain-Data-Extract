import os

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

SPREADSHEET_NAME = "DHM Rainfall Data"
FOLDER_NAME = "DHM Rainfall"


if not os.path.exists("token.json"):
    raise FileNotFoundError("token.json was not found.")

creds = Credentials.from_authorized_user_file(
    "token.json",
    SCOPES
)

if creds.expired and creds.refresh_token:
    print("Access token expired. Refreshing...")
    creds.refresh(Request())

print("Google credentials loaded.")


drive = build(
    "drive",
    "v3",
    credentials=creds
)

sheets = build(
    "sheets",
    "v4",
    credentials=creds
)


print("\nSearching for Google Drive folder...")

folder_results = drive.files().list(
    q=(
        f"name = '{FOLDER_NAME}' "
        "and mimeType = 'application/vnd.google-apps.folder' "
        "and trashed = false"
    ),
    spaces="drive",
    fields="files(id,name)"
).execute()

folders = folder_results.get("files", [])

if not folders:
    raise RuntimeError(
        f"Could not find Google Drive folder: {FOLDER_NAME}"
    )

folder_id = folders[0]["id"]

print(f"Found folder: {FOLDER_NAME}")
print(f"Folder ID: {folder_id}")


print("\nChecking whether spreadsheet already exists...")

sheet_results = drive.files().list(
    q=(
        f"name = '{SPREADSHEET_NAME}' "
        "and mimeType = 'application/vnd.google-apps.spreadsheet' "
        "and trashed = false"
    ),
    spaces="drive",
    fields="files(id,name,parents)"
).execute()

existing_sheets = sheet_results.get("files", [])


if existing_sheets:

    spreadsheet_id = existing_sheets[0]["id"]

    print("Spreadsheet already exists.")
    print(f"Spreadsheet ID: {spreadsheet_id}")

else:

    print("\nCreating new Google Sheet...")

    spreadsheet = sheets.spreadsheets().create(
        body={
            "properties": {
                "title": SPREADSHEET_NAME
            },
            "sheets": [
                {
                    "properties": {
                        "title": "Rainfall"
                    }
                }
            ]
        }
    ).execute()

    spreadsheet_id = spreadsheet["spreadsheetId"]

    print("Spreadsheet created.")
    print(f"Spreadsheet ID: {spreadsheet_id}")


    print("\nMoving spreadsheet into DHM Rainfall folder...")

    file_info = drive.files().get(
        fileId=spreadsheet_id,
        fields="parents"
    ).execute()

    old_parents = file_info.get("parents", [])

    update_kwargs = {
        "fileId": spreadsheet_id,
        "addParents": folder_id,
        "fields": "id,parents"
    }

    if old_parents:
        update_kwargs["removeParents"] = ",".join(old_parents)

    drive.files().update(**update_kwargs).execute()

    print("Spreadsheet moved into DHM Rainfall folder.")


spreadsheet_info = sheets.spreadsheets().get(
    spreadsheetId=spreadsheet_id
).execute()

sheet_names = [
    s["properties"]["title"]
    for s in spreadsheet_info.get("sheets", [])
]

if "Rainfall" not in sheet_names:

    print("\nCreating Rainfall tab...")

    sheets.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": "Rainfall"
                        }
                    }
                }
            ]
        }
    ).execute()

    print("Rainfall tab created.")

else:
    print("\nRainfall tab already exists.")


header = [
    "Station Index",
    "Station Name",
    "Date",
    "Daily Rainfall (mm)",
    "Running Total (mm)"
]

existing_values = sheets.spreadsheets().values().get(
    spreadsheetId=spreadsheet_id,
    range="Rainfall!A1:E1"
).execute()

values = existing_values.get("values", [])

if not values:

    print("\nAdding header row...")

    sheets.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Rainfall!A1:E1",
        valueInputOption="RAW",
        body={
            "values": [header]
        }
    ).execute()

    print("Header row added.")

else:

    print("\nHeader row already exists.")


print("\n======================================")
print("GOOGLE SHEET SETUP COMPLETE")
print("======================================")
print(f"Spreadsheet: {SPREADSHEET_NAME}")
print(f"Spreadsheet ID: {spreadsheet_id}")
print("Sheet: Rainfall")
print("Folder: DHM Rainfall")
print("======================================")