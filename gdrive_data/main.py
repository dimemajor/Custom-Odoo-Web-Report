import os
from pprint import pprint

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import gdrive_data.auth as auth
from constants import SCOPES

def search_file(service, categories):
    drive_pics = {}
    try:
        page_token = None
        pg_token = None
        for category in categories:
            while True:
                response = service.files().list(q=f"name='{category}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
                                                fields='nextPageToken,'
                                                    'files(id, name, webViewLink)', pageSize=1000,
                                                pageToken=page_token).execute()
                for folder in response.get('files', []):
                    while True:
                        drive_pics[folder.get('name')] = {}
                        response = service.files().list(q=f"'{folder.get('id')}' in parents",
                                                        fields='nextPageToken,'
                                                            'files(id, name, webViewLink)', pageSize=1000,
                                                        pageToken=pg_token).execute()
                        for file in response.get('files', []):
                            drive_pics[folder.get('name')][os.path.splitext(file.get('name'))[0]] = file.get("webViewLink")
                        pg_token = response.get('nextPageToken', None)
                        if pg_token is None:
                            break
                page_token = response.get('nextPageToken', None)
                if page_token is None:
                    break


    except HttpError as error:
        print(F'An error occurred: {error}')
        drive_pics = None

    return drive_pics
