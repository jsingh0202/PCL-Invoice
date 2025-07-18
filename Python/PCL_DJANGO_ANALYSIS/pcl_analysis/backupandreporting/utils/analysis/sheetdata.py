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
COST_SUMMARY = "1R_lUhoo05uPWMsLXFUwkgUxzJa6yXi2I8pNp_xPSI_Y"
PCL_CODE_DICT = "PCL_CODE_DICT"
BASE_DIR = Path(__file__).resolve().parent


def validate():
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

    return creds


def get_data(creds):
    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = (
        sheet.values().get(spreadsheetId=COST_SUMMARY, range="PCL_CODE_DICT").execute()
    )
    values = result.get("values", [])
    if not values:
            print("No data found.")
            raise ValueError("No data found in the specified range.")
    return values


def get_code_and_budget() -> dict:
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = validate()
    cost_summary_data = {}

    try:
        values = get_data(creds)

        # get the code and budget values
        headers = values[0]
        code_idx = headers.index("PCL_Code")
        budget_idx = headers.index("Budget")

        code_budget = [
            (row[code_idx], row[budget_idx])
            for row in values[1:]
            if len(row) > max(code_idx, budget_idx)
        ]

        # print("Code and Budget Values:")
        # print(code_budget)
        # print("Name, Major:")

        for code, budget in code_budget:
            # print(f"{code}, {budget}")
            cleaned_budget = budget.replace("$", "").replace(",", "").strip()
            cost_summary_data[code] = Decimal(cleaned_budget)

    except HttpError as err:
        print(err)

    return cost_summary_data


def get_code_and_delete() -> dict:
    creds = validate()
    code_delete_data = {}

    try:
        values = get_data(creds)

        # get the code values
        headers = values[0]
        code_idx = headers.index("PCL_Code")
        deleted_idx = headers.index("Deleted")
        
        code_delete = [
            (row[code_idx], row[deleted_idx])
            for row in values[1:]
            if len(row) > max(code_idx, deleted_idx)
        ]

        for code, delete in code_delete:
            code_delete_data[code] = delete

    except HttpError as err:
        print(err)

    return code_delete_data


if __name__ == "__main__":
    get_code_and_budget()
