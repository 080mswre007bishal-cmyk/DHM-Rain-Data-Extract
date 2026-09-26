import os
import json
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]


# ---------------------------------------------------------
# 1. Read Google OAuth token
# ---------------------------------------------------------

if not os.path.exists("token.json"):
    raise FileNotFoundError("token.json was not found.")

creds = Credentials.from_authorized_user_file(
    "token.json",
    SCOPES
)

# Refresh token if necessary
if creds.expired and creds.refresh_token:
    print("Access token expired. Refreshing...")
    creds.refresh(Request())

print("Google OAuth credentials loaded successfully.")


# ---------------------------------------------------------
# 2. Test Google Drive
# ---------------------------------------------------------

print("\nTesting Google Drive API...")

drive = build(
    "drive",
    "v3",
    credentials=creds
)

results = drive.files().list(
    pageSize=10,
    fields="files(id,name,mimeType)"
).execute()

files = results.get("files", [])

print("Google Drive API connected successfully.")
print(f"Files visible: {len(files)}")

for file in files:
    print(
        f"  - {file['name']} "
        f"({file['mimeType']})"
    )


# ---------------------------------------------------------
# 3. Test Google Sheets
# ---------------------------------------------------------

print("\nTesting Google Sheets API...")

sheets = build(
    "sheets",
    "v4",
    credentials=creds
)

print("Google Sheets API connected successfully.")


print("\n======================================")
print("GOOGLE CONNECTION TEST PASSED")
print("======================================")