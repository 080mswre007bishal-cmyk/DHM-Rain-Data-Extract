import os
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

SPREADSHEET_NAME = "DHM Rainfall Data"
SHEET_NAME = "Rainfall"

# --------------------------------------------------
# Load Google credentials
# --------------------------------------------------

creds = Credentials.from_authorized_user_file(
    "token.json",
    SCOPES
)

if creds.expired and creds.refresh_token:
    print("Access token expired. Refreshing...")
    creds.refresh(Request())

print("Google credentials loaded.")

# --------------------------------------------------
# Connect to Google Drive
# --------------------------------------------------

drive = build("drive", "v3", credentials=creds)

print("Searching for Google Sheet...")

results = drive.files().list(
    q=(
        "name = 'DHM Rainfall Data' "
        "and mimeType = 'application/vnd.google-apps.spreadsheet' "
        "and trashed = false"
    ),
    pageSize=10,
    fields="files(id,name)"
).execute()

files = results.get("files", [])

if not files:
    raise RuntimeError("Could not find Google Sheet: DHM Rainfall Data")

spreadsheet_id = files[0]["id"]

print(f"Google Sheet found: {SPREADSHEET_NAME}")
print(f"Spreadsheet ID: {spreadsheet_id}")

# --------------------------------------------------
# Connect to Google Sheets
# --------------------------------------------------

sheets = build("sheets", "v4", credentials=creds)

# --------------------------------------------------
# Test record
# --------------------------------------------------

test_row = [
    "TEST",
    "Test Station",
    "2026-09-26",
    12.34,
    56.78,
]

print()
print("Writing test rainfall record...")
print(test_row)

result = sheets.spreadsheets().values().append(
    spreadsheetId=spreadsheet_id,
    range=f"{SHEET_NAME}!A:E",
    valueInputOption="USER_ENTERED",
    insertDataOption="INSERT_ROWS",
    body={
        "values": [test_row]
    }
).execute()

print()
print("======================================")
print("GOOGLE SHEET WRITE TEST PASSED")
print("======================================")
print()
print("Updated range:")
print(result.get("updates", {}).get("updatedRange"))