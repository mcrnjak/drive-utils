from apiclient import discovery
from httplib2 import Http
from oauth2client import file, client, tools
from openpyxl import Workbook


files_cache = {}


def cache_put(files_list):
    for f in files_list:
        key = f['id']
        files_cache[key] = f


def get_credentials():
    scope = 'https://www.googleapis.com/auth/drive.metadata.readonly'
    store = file.Storage('storage.json')
    credentials = store.get()

    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets('client_id.json', scope)
        credentials = tools.run_flow(flow, store)

    return credentials


def get_drive_service():
    credentials = get_credentials()
    service = discovery.build('drive', 'v2', http=credentials.authorize(Http()))
    return service


def get_file_path(f, service):
    path = get_file_path_recursively(f, service, [])
    return "/".join(path)


def get_file_path_recursively(f, service, path):
    title = f['title']
    path.insert(0, title)

    if f['parents'][0].get('isRoot', None):
        return path

    parent_id = f['parents'][0]['id']
    parent = files_cache.get(parent_id, None)

    if not parent:
        parent = service.files().get(fileId=parent_id).execute()
        cache_put([parent])

    return get_file_path_recursively(parent, service, path)


def write_to_spreadsheet(include_full_file_path):
    service = get_drive_service()

    wb = Workbook()
    ws = wb.active
    ws.title = "Shared Files"

    header_row = ["File Id", "Full Path", "File Name", "Permission Id", "Permission Owner", "Permission Role"]
    ws.append(header_row)

    page_token = None
    while True:
        request = service.files().list(
            q="trashed=false and 'me' in owners",
            pageToken=page_token,
            fields="nextPageToken, items(id, title, permissions, parents)")

        response = request.execute()
        files = response.get('items', [])
        cache_put(files)

        for f in files:
            permissions = f['permissions']
            if len(permissions) > 1:

                file_path = None
                if include_full_file_path:
                    file_path = get_file_path(f, service)

                for perm in permissions:
                    row_data = [f['id'], file_path, f['title'], perm['id'], perm.get('name', perm['id']), perm['role']]
                    ws.append(row_data)

        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break

    wb.save("shared_drive_files.xlsx")


if __name__ == '__main__':
    write_to_spreadsheet(True)
