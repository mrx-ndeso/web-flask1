import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

def get_gspread_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    return gspread.authorize(creds)

def upload_to_google_sheets(df, spreadsheet_id, worksheet_name):
    client = get_gspread_client()
    sheet = client.open_by_key(spreadsheet_id).worksheet(worksheet_name)
    sheet.update([df.columns.values.tolist()] + df.values.tolist())

def download_from_google_sheets(spreadsheet_id, worksheet_name):
    client = get_gspread_client()
    sheet = client.open_by_key(spreadsheet_id).worksheet(worksheet_name)
    data = sheet.get_all_records()
    return pd.DataFrame(data)
