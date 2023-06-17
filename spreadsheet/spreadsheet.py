# spreadsheet với tên là titleSpreadsheet và tên của sheet là titleSheet
# xóa, chỉnh sửa dữ liệu, ...
# done
import base.auth as auth
from googleapiclient.errors import HttpError


def create(titleSpreadsheet, titleSheet):
    try:
        service = auth.sheet_service
        spreadsheet = {
            'properties': {
                'title': titleSpreadsheet
            },
            'sheets': {
                'properties': {
                    'title': titleSheet,
                    'hidden': False
                }
            }
        }
        request = service.spreadsheets().create(body=spreadsheet)
        response = request.execute()
        id = response.get('spreadsheetId')
        permission(id)
        msg = 'đã tạo thành công sheet với id: ' + id

    except HttpError as err:
        msg = 'không thể tạo form lúc này!\n' + err
        print(msg)

    return id


def addSheet(id, titleSheet):
    ibody = {
        "requests": [
            {
                "addSheet": {
                    "properties": {
                        "title": titleSheet
                    }
                }
            }
        ]
    }
    request = auth.sheet_service.spreadsheets().batchUpdate(spreadsheetId=id, body=ibody)
    response = request.execute()
    return response


def getSheet(id, sheetName):
    range = "'" + sheetName + "'" + "!A1:T100"
    request = auth.sheet_service.spreadsheets().values().batchGet(
        spreadsheetId=id,
        ranges=range
    )

    result = request.execute()
    if 'values' in result['valueRanges'][0]:
        return result['valueRanges'][0]['values']
    return result


def getColumnSheet(id, sheetName, column):
    range = "'" + sheetName + "'" + "!" + column + "1:" + column + "100"
    request = auth.sheet_service.spreadsheets().values().batchGet(
        spreadsheetId=id,
        ranges=range
    )

    result = request.execute()
    if 'values' in result['valueRanges'][0]:
        return result['valueRanges'][0]['values']
    return result

def permission(id):
    resource = {
        "role": "writer",
        "type": "anyone"
    }
    # for email
    # permission1 = {
    #     'type': 'user',
    #     'role': 'writer',
    #     'emailAddress': 'sangproject123@gmail.com'
    # }
    # auth.drive_service.permissions().create(  fileId=key,
    #                                     body=permission1).execute()
    auth.drive_service.permissions().create(fileId=id, body=resource).execute()


def updateValue(id, nameSheet, data):
    range = "'" + nameSheet + "'" + '!A1:T100'
    value_input_option = 'RAW'
    value_range = {
        'values': data
    }
    # xóa dữ liệu trước đó
    clear(id, nameSheet)

    request = auth.sheet_service.spreadsheets().values().update(spreadsheetId=id, range=range,
                                                                valueInputOption=value_input_option,
                                                                body=value_range)
    response = request.execute()
    return response


def clear(id, nameSheet):
    range = "'" + nameSheet + "'" + '!A1:T100'

    request = auth.sheet_service.spreadsheets().values().clear(spreadsheetId=id, range=range)
    response = request.execute()
    return response


def getLink(id):
    try:
        request = auth.sheet_service.spreadsheets().get(spreadsheetId=id)
        response = request.execute()
    except HttpError as err:
        msg = 'link sheet bị lỗi!\n'
        print(err)
        return ''
    return response['spreadsheetUrl']

