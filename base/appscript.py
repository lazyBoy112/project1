import base.auth as auth
import json
import os
service = auth.script_service
scriptId = 'AKfycbw5hOug2GV9UctYxqfclZPLgEnMhbGSkhaMkFvOjCNeijiFN1I5oM1AwCYj07FqdAeG0Q'


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


def updateSheet(studenList, topicList):
    localdata = getLocalData()
    id = localdata['spreadsheetId']
    if not id:
        id = createSheet()
        setSpreadSheetId(id)
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
    return


def addRegisterSheet():
    localdata = getLocalData()
    id = localdata['spreadsheetId']
    if not id:
        return 'Trang tính không được tạo! Hãy cập nhật lại trang tính'
    body = {
        'function': 'addRegisterSheet',
        'parameters': [
            id
        ]
    }
    service.scripts().run(scriptId=scriptId, body=body).execute()
    return ''


def getLinkSheet():
    localdata = getLocalData()
    id = localdata['spreadsheetId']
    if not id:
        return [False, 'Trang tính không được tạo! Hãy cập nhật lại trang tính']
    return [True, 'https://docs.google.com/spreadsheets/d/' + id]


def updateForm():
    localdata = getLocalData();
    sheetId = localdata['spreadsheetId']
    formId = localdata['formId']
    if not sheetId:
        return [False, ' Dữ liệu trên localdata bị mật! Hãy cập nhật lại danh sách đăng ký!']
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
        return [False, 'Sinh viên chưa đăng ký!']
    formId = result[1]
    setFormId(formId)
    return [True, '']


def getFormLink():
    id = getLocalData()
    id = id['formId']
    if not id:
        return ''
    return 'https://docs.google.com/forms/d/' + id + '/viewform'

def getResultForm():
    localdata = getLocalData()
    sheetId = localdata['spreadsheetId']
    if not sheetId:
        return []
    formId = localdata['formId']
    if not formId:
        return []
    body = {
        'function': 'getResult',
        'parameters': [
            formId,
            sheetId
        ]
    }
    request = service.scripts().run(scriptId=scriptId, body=body)
    response = request.execute()
    # print(response)
    result = response['response']['result']
    data = []
    # data.append(['Nhóm', 'Tên sinh viên', 'Điểm trung bình'])
    for i in range(len(result)):
        data.append([result[i][0], result[i][1], result[i][-1]])
    print(data)
    return data


def getLocalData():
    if not os.path.exists('base/data.json'):
        createLocalData()
    with open('base/data.json', 'r') as file:
        data = json.load(file)
    return data


def setSpreadSheetId(id):
    with open('base/data.json', 'r') as file:
        data = json.load(file)
    data['spreadsheetId'] = id
    file.close()
    data = json.dumps(data)
    with open('base/data.json', 'w') as file:
        file.write(data)
    file.close


def setFormId(id):
    with open('base/data.json', 'r') as file:
        data = json.load(file)
    data['formId'] = id
    file.close()
    data = json.dumps(data)
    with open('base/data.json', 'w') as file:
        file.write(data)
    file.close

def createLocalData():
    data = {"spreadsheetId": "", "formId": ""}
    data = json.dumps(data)
    with open('data.json', 'w') as file:
        file.write(data)
    file.close()

