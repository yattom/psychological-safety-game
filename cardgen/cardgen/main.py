# coding: utf-8
from pprint import pprint
import datetime
import sys


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


def read_card_data(spreadsheet_id, credentials):
    import gspread

    gc = gspread.authorize(credentials)

    all_data = {}
    for sheet in gc.open_by_key(spreadsheet_id).worksheets():
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


def create_new_presentation(service, metadata, title, timestamp):
    fileId = metadata[title]
    #new_filename = '/' + metadata['出力フォルダ'] + '/' + timestamp + '_' + title
    new_filename = timestamp + '_' + title
    body = { 'name': new_filename }
    new_file = service.files().copy(fileId=fileId, body=body).execute()
    return new_file


def detect_elements_to_modify(keywords, presentation):
    # TODO: グループ化した要素に未対応
    elements_to_modify = {}
    for key in keywords:
        for element in presentation['slides'][0]['pageElements']:
            try:
                if len(element['shape']['text']['textElements']) > 0:
                    for text_element in element['shape']['text']['textElements']:
                        try:
                            if text_element['textRun']['content'].strip() == key:
                                elements_to_modify[key] = element['objectId']
                        except KeyError:
                            continue
            except KeyError:
                continue
    return elements_to_modify


def add_slides(service, gfile, slide_data):
    requests = []
    presentation = service.presentations().get(presentationId=gfile['id']).execute()
    template_slide_id = presentation['slides'][0]['objectId']
    elements_to_modify = detect_elements_to_modify(slide_data[0].keys(), presentation)
    for i, slide in enumerate(slide_data):
        new_element_ids = { key:elements_to_modify[key]+'_'+str(i) for key in elements_to_modify }
        objectIds = { elements_to_modify[key]:new_element_ids[key] for key in elements_to_modify }
        requests.append({ 'duplicateObject': { 'objectId': template_slide_id, 'objectIds': objectIds }})
        for key in elements_to_modify:
            requests.append({ 'deleteText':
                                { 'objectId': new_element_ids[key],
                                  'textRange': { 'type': 'ALL' } 
                                }
                            })
            requests.append({ 'insertText':
                                { 'objectId': new_element_ids[key],
                                  'text': slide[key],
                                  'insertionIndex': 0,
                                }
                            })
    requests.append({ 'deleteObject': { 'objectId': template_slide_id } })
    service.presentations().batchUpdate(presentationId=gfile['id'], body={ 'requests': requests }).execute()


def delete_presensation(service, gfile):
    service.files().delete(fileId=gfile['id']).execute()


def create_presentation(credentials, data, timestamp):
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
        new_file = create_new_presentation(drive_service, metadata, title, timestamp)
        add_slides(slides_service, new_file, data[title])
        export_pdf(drive_service, new_file)
        delete_presensation(drive_service, new_file)


def export_pdf(service, file_to_export):
    raw = service.files().export(fileId=file_to_export['id'], mimeType='application/pdf').execute()
    with open(file_to_export['name'] + ".pdf", 'wb') as f:
        f.write(raw)


def main():
    if len(sys.argv) < 2:
        print("usage: python cardgen/main.py 1Or3lnAR1UAIZmzUY7_iQRgNlGiVV5EEBy5NTPrzurG4")
        exit()
    spreadsheet_id = sys.argv[1]
    timestamp = datetime.datetime.now().isoformat(timespec='seconds')
    timestamp = timestamp.replace(':', '')
    credentials = auth()
    data = read_card_data(spreadsheet_id, credentials)
    create_presentation(credentials, data, timestamp)


if __name__=='__main__':
    main()
