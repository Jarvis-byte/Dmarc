import csv
from openpyxl import Workbook
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import checkdmarc as checkdmarc
from datetime import datetime
from DataFromExcel import DataFromExcel
import requests
from io import BytesIO
def main(filename=None, data=None):
    # Replace with the ID of your Google Sheet
    SPREADSHEET_ID = '188lKNuzMXplzkO6UpZLsFOHx-NgDXA3lfZpvkSV-neE'

    # Replace with the path to your service account JSON key file
    SERVICE_ACCOUNT_FILE = 'gis-bcdr-ops-fe35-84aa16f1f12c.json'

    # Authenticate and build the service object
    creds = None
    try:
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
        )
    except Exception as e:
        print('Error loading credentials: {}'.format(str(e)))

    service = build('drive', 'v3', credentials=creds)

    # Retrieve the Google Sheet file object
    # Export the Google Sheet file as a CSV file
    try:
        response = service.files().export_media(fileId=SPREADSHEET_ID, mimeType='text/csv').execute()

        # Write the CSV data to a local file
        with open('output.csv', 'wb') as f:
            f.write(response)

        print('Google Sheet downloaded as output.csv')
    except HttpError as error:
        print('An error occurred: {}'.format(error))

    # Read data from CSV file
    with open('output.csv', 'r') as f:
        reader = csv.reader(f)
        data = list(reader)

    # Create a new workbook and sheet
    workbook = Workbook()
    sheet = workbook.active

    # Write data to sheet
    for row in data:
        sheet.append(row)

    # Save the workbook as an XLSX file
    workbook.save('output.xlsx')

    fis = "output.xlsx"
    now = datetime.now()

    current_time = now.strftime("%H:%M:%S")
    print("Start Time =", current_time)
    # Obj for read data
    data_from_excel = DataFromExcel()
    data_from_excel.readData(fis)


if __name__ == '__main__':
    main(filename=None, data=bytes)
