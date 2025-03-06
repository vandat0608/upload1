from google.oauth2 import service_account
from googleapiclient.discovery import build
import re

def connect_to_google_sheets(spreadsheet_id):
    try:
        credentials = service_account.Credentials.from_service_account_file(
            CREDENTIALS_FILE,
            scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        service = build('sheets', 'v4', credentials=credentials)
        return service.spreadsheets()
    except Exception as e:
        logging.error(f"Error connecting to Google Sheets: {e}")
        return None

def extract_spreadsheet_id(google_sheet_url):
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', google_sheet_url)
    if match:
        return match.group(1)
    return None

def upload_to_google_sheets(file_path, spreadsheet_id, sheet_name, service):
    try:
        # Load the data from the Excel file
        workbook = openpyxl.load_workbook(file_path)
        summary_sheet = workbook["Thống kê nghỉ học"]
        data = summary_sheet.values

        # Prepare the data for Google Sheets
        body = {
            'values': list(data)
        }

        # Clear the existing data in the specified sheet
        service.values().clear(spreadsheetId=spreadsheet_id, range=sheet_name).execute()

        # Update the Google Sheet with new data
        service.values().update(
            spreadsheetId=spreadsheet_id,
            range=sheet_name,
            valueInputOption='RAW',
            body=body
        ).execute()

        logging.info(f"Data uploaded successfully to {sheet_name} in spreadsheet {spreadsheet_id}")
    except Exception as e:
        logging.error(f"Error uploading data to Google Sheets: {e}")