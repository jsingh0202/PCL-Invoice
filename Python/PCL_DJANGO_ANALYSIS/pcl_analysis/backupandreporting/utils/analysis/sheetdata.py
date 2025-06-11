import os.path
from pathlib import Path
from decimal import Decimal

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# The ID and range of a sample spreadsheet.
OVERHAUL = "1pmzLJFCFhaiNhGL4wHo4dpRlGq1FAt59g065cfL1B1k"
MAY_SHEET = "2025.05!A2:M"
BASE_DIR = Path(__file__).resolve().parent


def get_data() -> dict:
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    token_path = BASE_DIR / "token.json"
    creds_path = BASE_DIR / "credentials.json"
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(token_path, "w") as token:
            token.write(creds.to_json())

    overhaul_data = {}

    try:
        service = build("sheets", "v4", credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=OVERHAUL, range=MAY_SHEET).execute()
        values = result.get("values", [])

        if not values:
            print("No data found.")
            raise ValueError("No data found in the specified range.")

        # print("Name, Major:")
        for row in values:
            # print(f"{row[0]}, {row[8]}")
            overhaul_data[row[0]] = Decimal(row[8].replace(",", ""))

    except HttpError as err:
        print(err)

    return overhaul_data


if __name__ == "__main__":
    get_data()
