import os.path
import time
from googlesearch import search
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
# The ID and range of a sample spreadsheet.
INPUT_SPREADSHEET_ID = "13fI1G9J3Onrf_ZcLRNHBDzZcyHMW54uMJ5zMs3hwN4I"
OUTPUT_SHEET_ID = "1QGju0RRnI67EkktnosKDiVQGFfpUch6a_wFTaeb3GxQ"
INPUT_RANGE_NAME = "List of NBFCs!A1:I9328"
OUTPUT_RANGE_NAME = "output_sheet!A2:F19330"


# Searches for the website link of a given query.

def search_website(query):
    for url in search(query, tld="co.in", num=1):
        return url
    return "No URL found"


# Gets the Google Sheets service.
def get_sheets_service():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return build("sheets", "v4", credentials=creds)


# Reads data from the input sheet.
def read_data(service):
    try:
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=INPUT_SPREADSHEET_ID, range=INPUT_RANGE_NAME).execute()
        values = result.get("values", [])
        return values
    except HttpError as err:
        print(f"An error occurred: {err}")
        return []


# Processes the data by searching for the website link of each row.
def process_data(data: list) -> list:
    output_values = []
    sr_no = 1
    for row in data[1:]:  # Skip header row
        if len(row) > 1:
            name = row[1]
            email = row[8]
            regional_office = row[2]
            address = row[7]
            website_link = search_website(name)
            print(f" Website Name: {name}, Link: {website_link}")
            time.sleep(1)  # Reduced delay between searches
            if website_link != "No URL found":
                output_values.append([sr_no, name, regional_office, address, email, website_link])
                sr_no += 1
    return output_values


# Writes the data to the output sheet.
def write_data(service, data):
    try:
        sheet = service.spreadsheets()
        body = {'values': data}
        sheet.values().update(
            spreadsheetId=OUTPUT_SHEET_ID,
            range=OUTPUT_RANGE_NAME,
            valueInputOption="RAW",
            body=body
        ).execute()
        print('Data successfully written to output sheet.')
    except HttpError as err:
        print(f"An error occurred: {err}")


# The main function that reads data, processes it, and writes it to the output sheet.
def main():
    service = get_sheets_service()
    data = read_data(service)
    if not data:
        print("No data found.")
        return None
    processed_data = process_data(data)
    write_data(service, processed_data)


if __name__ == "__main__":
    main()
