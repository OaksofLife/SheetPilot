import gspread
from oauth2client.service_account import ServiceAccountCredentials

def get_sheet_data(spreadsheet_url, start_row, end_row, columns):
    # Define scope
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

    # Authenticate
    creds = ServiceAccountCredentials.from_json_keyfile_name('service_account.json', scope)
    client = gspread.authorize(creds)

    # Open the sheet by URL
    sheet = client.open_by_url(spreadsheet_url).sheet1

    # Read the desired rows
    data = []
    for row in sheet.get_all_values()[start_row-1:end_row]:
        selected = [row[c-1] for c in columns]  # Adjust to 0-indexed for Python lists
        data.append(selected)

    return data
