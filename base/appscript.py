import base.auth as auth
service = auth.script_service
scriptId = 'AKfycbwEKil2-jLUc1Gd1_5a_zt9W3mEMd_r4FwcyyPeYFyIyCAmfoWuJ4P7gKHAoiuVvg6McQ'


def test():
    body = {
        'function': 'test'
    }
    re = service.scripts().run(scriptId=scriptId, body=body)
    return re.execute()


def createSheet():
    body = {
        'function': 'createSheet'
    }
    request = service.scripts().run(scriptId=scriptId, body=body)
    request = request.execute()
    return request['response']['result']


def updateSheet(studenList, topicList, datafile):
    localdata = datafile
    id = localdata['spreadsheetId']
    if not id:
        id = createSheet()
        localdata['spreadsheetId'] = id
    if not checkSheetId(id):
        id = createSheet()
        localdata['spreadsheetId'] = id

    body1 = {
        'function': 'updateSheet1',
        'parameters': [
            id,
            studenList
        ]
    }
    body2 = {
        'function': 'updateSheet2',
        'parameters': [
            id,
            topicList
        ]
    }
    request = service.scripts().run(scriptId=scriptId, body=body1)
    request = request.execute()

    request = service.scripts().run(scriptId=scriptId, body=body2)
    request = request.execute()
    return localdata


def addRegisterSheet(datafile):
    localdata = datafile
    id = localdata['spreadsheetId']
    if not id:
        return 'Trang tính không được tạo! Hãy cập nhật lại trang tính'
    if not checkSheetId(id):
        return 'Trang tính không tồn tại.\nHãy tạo trang tính mới'
    body = {
        'function': 'addRegisterSheet',
        'parameters': [
            id
        ]
    }
    service.scripts().run(scriptId=scriptId, body=body).execute()
    return ''


def getLinkSheet(datafile):
    localdata = datafile
    id = localdata['spreadsheetId']
    if not id:
        return [False, 'Trang tính không được tạo! Hãy cập nhật lại trang tính']
    if not checkSheetId(id):
        return [False, 'Trang tính không tồn tại.\nHãy tạo trang tính mới']
    return [True, 'https://docs.google.com/spreadsheets/d/' + id]


def updateForm(datafile):
    localdata = datafile
    sheetId = localdata['spreadsheetId']
    formId = localdata['formId']
    if not sheetId:
        return [localdata, ' Dữ liệu trống! Hãy cập nhật lại danh sách đăng ký!']
    body = {
        'function': 'updateForm',
        'parameters': [
            sheetId,
            formId
        ]
    }
    request = service.scripts().run(scriptId=scriptId, body=body)
    response = request.execute()
    print(response)
    result = response['response']['result']
    if not result[0]:
        return [localdata, 'Sinh viên chưa đăng ký!']
    formId = result[1]
    localdata['formId'] = formId
    return [localdata, '']


def getFormLink(datafile):
    localdata = datafile
    id = localdata['formId']
    if not id:
        return [False, 'Không có dữ liệu form.\nHãy tạo lại form mới!']
    if not checkFormId(id):
        return [False, 'Form không còn tồn tại.\nHãy tạo lại form mới!']
    return [True, 'https://docs.google.com/forms/d/' + id + '/viewform']


def getResponseForm(datafile):
    localdata = datafile
    sheetId = localdata['spreadsheetId']
    if not sheetId:
        return [1]
    formId = localdata['formId']
    if not formId:
        return [1]
    body = {
        'function': 'getResponseFromForm',
        'parameters': [
            formId,
            sheetId
        ]
    }
    request = service.scripts().run(scriptId=scriptId, body=body)
    response = request.execute()
    print(response)
    result = response['response']['result']
    return result


def getAverageMark(datafile):
    localdata = datafile
    sheetId = localdata['spreadsheetId']
    formId = localdata['formId']
    body = {
        'function': 'getAverageMark',
        'parameters': [
            formId,
            sheetId
        ]
    }
    request = service.scripts().run(scriptId=scriptId, body=body)
    response = request.execute()
    # print(response)
    result = response['response']['result']
    return result


def checkFormId(id):
    body = {
        'function': 'checkFormId',
        'parameters': [
            id
        ]
    }
    request = service.scripts().run(scriptId=scriptId, body=body)
    response = request.execute()
    return response


def checkSheetId(id):
    body = {
        'function': 'checkSheetId',
        'parameters': [
            id
        ]
    }
    request = service.scripts().run(scriptId=scriptId, body=body)
    response = request.execute()
    return response



