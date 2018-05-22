# coding: utf-8
from pprint import pprint

def main():
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials

    scope = ['https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive']

    credentials = ServiceAccountCredentials.from_json_keyfile_name('API Project-a38c7273d3d8.json', scope)

    gc = gspread.authorize(credentials)

    SPREADSHEET_ID = '1Or3lnAR1UAIZmzUY7_iQRgNlGiVV5EEBy5NTPrzurG4'
    for sheet in gc.open_by_key(SPREADSHEET_ID).worksheets():
        cells = sheet.get_all_records()
        pprint(cells)

if __name__=='__main__':
    main()
