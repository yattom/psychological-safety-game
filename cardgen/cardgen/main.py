# coding: utf-8
from pprint import pprint
import datetime

SPREADSHEET_ID = '1Or3lnAR1UAIZmzUY7_iQRgNlGiVV5EEBy5NTPrzurG4'


def auth():
    from oauth2client.service_account import ServiceAccountCredentials

    scope = ['https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/presentations',
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/drive.appdata',
            'https://www.googleapis.com/auth/drive.file',
            ]

    credentials = ServiceAccountCredentials.from_json_keyfile_name('API Project-a38c7273d3d8.json', scope)
    return credentials


def read_card_data(credentials):
    import gspread

    gc = gspread.authorize(credentials)

    all_data = {}
    for sheet in gc.open_by_key(SPREADSHEET_ID).worksheets():
        cells = sheet.get_all_records()
        all_data[sheet.title] = cells
    return all_data


def create_service_for_drive(credentials):
    import httplib2
    from googleapiclient.discovery import build
    http = httplib2.Http()
    http = credentials.authorize(http)
    service = build("drive", "v3", http=http)
    return service


def create_service_for_slides(credentials):
    import httplib2
    from googleapiclient.discovery import build
    http = httplib2.Http()
    http = credentials.authorize(http)
    service = build("slides", "v1", http=http)
    return service


def create_new_presentation(service, metadata, title, now):
    fileId = metadata[title]
    #new_filename = '/' + metadata['出力フォルダ'] + '/' + now + '_' + title
    new_filename = now + '_' + title
    body = { 'name': new_filename }
    new_file = service.files().copy(fileId=fileId, body=body).execute()
    return new_file


def create_presentation(credentials, data, now):
    slides_service = create_service_for_slides(credentials)
    drive_service = create_service_for_drive(credentials)
    metadata = {}
    for entry in data['metadata']:
        key = entry['キー']
        value = entry['値']
        metadata[key] = value
    del data['metadata']
    for title in data:
        if title == 'metadata':
            continue
        if title not in metadata:
            continue
        new_file = create_new_presentation(drive_service, metadata, title, now)


def main():
    now = datetime.datetime.now().isoformat(timespec='seconds')
    credentials = auth()
    data = read_card_data(credentials)
    create_presentation(credentials, data, now)


if __name__=='__main__':
    main()
