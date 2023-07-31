from django.shortcuts import render, HttpResponse
from django.http import HttpResponseRedirect
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from io import BytesIO
import json
import pandas
import os.path


SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/script.projects'
]
scriptId = 'AKfycbwEKil2-jLUc1Gd1_5a_zt9W3mEMd_r4FwcyyPeYFyIyCAmfoWuJ4P7gKHAoiuVvg6McQ'


def service():
    creds = None
    if os.path.exists('static\\json\\token.json'):
        creds = Credentials.from_authorized_user_file('static\\json\\token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('static\\json\\key.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('static\\json\\token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        script_service = build('script', 'v1', credentials=creds)
        return script_service
    except HttpError as err:
        print(err)


def createSheet():
    body = {
        'function': 'createSheet'
    }
    request = service().scripts().run(scriptId=scriptId, body=body)
    request = request.execute()
    return request['response']['result']


def getUpdateSheet(studenList, topicList, datafile):
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
    request = service().scripts().run(scriptId=scriptId, body=body1)
    request = request.execute()

    request = service().scripts().run(scriptId=scriptId, body=body2)
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
    service().scripts().run(scriptId=scriptId, body=body).execute()
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
    request = service().scripts().run(scriptId=scriptId, body=body)
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
    request = service().scripts().run(scriptId=scriptId, body=body)
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
    request = service().scripts().run(scriptId=scriptId, body=body)
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
    request = service().scripts().run(scriptId=scriptId, body=body)
    response = request.execute()
    return response


def checkSheetId(id):
    body = {
        'function': 'checkSheetId',
        'parameters': [
            id
        ]
    }
    request = service().scripts().run(scriptId=scriptId, body=body)
    response = request.execute()
    return response
# view.py

# def createNewProject(request):
#     return render(request, 'pageSheet.html')
# def openProject(request):
#     return render(request, 'pageSheet.html')
# def updateSheet(request):
#     return render(request, 'pageSheet.html')
# def addSheet(request):
#     return render(request, 'pageSheet.html')
# def gotoSheet(request):
#     return render(request, 'pageSheet.html')
# def createForm(request):
#     return render(request, 'pageSheet.html')
# def copyLinkForm(request):
#     return render(request, 'pageSheet.html')
# def downloadResult(request):
#     return render(request, 'pageSheet.html')


def createNewProject(request):
    datafile = {"spreadsheetId": "", "formId": ""}
    response = render(request, 'pageSheet.html')
    response.set_cookie('spreadsheetId', datafile['spreadsheetId'])
    response.set_cookie('formId', datafile['formId'])
    return response


def openProject(request):
    datafile = {}
    response = render(request, 'pageSheet.html')
    if request.method == 'POST':
        projectFile = request.FILES.get('projectFile')
        if projectFile is None:
            return HttpResponse('Bạn chưa chọn file')
        datafile = json.load(projectFile)
        # writeLocal(request, datafile)
        response.set_cookie('spreadsheetId', datafile['spreadsheetId'])
        response.set_cookie('formId', datafile['formId'])
    # print(datafile)
    return response


def updateSheet(request):
    spreadsheetId = request.COOKIES.get('spreadsheetId', 1)
    formId = request.COOKIES.get('formId', 1)
    if (spreadsheetId == 1):
        spreadsheetId = ''
    if (formId == 1):
        formId = ''
    datafile = {'spreadsheetId': spreadsheetId, 'formId': formId}
    if request.method == 'POST':
        f_student = request.FILES.get('studentList')
        f_topic = request.FILES.get('topicList')
        # kiem tra file chua duoc tai len
        if (f_student is None) or (f_topic is None):
            return HttpResponse('Chưa chọn file')
        # chuyen doi du lieu tu excel qua [[]]
        tmp = pandas.read_excel(f_student)
        tmp = tmp.fillna('')
        d_student = tmp.values.tolist()
        h_student = tmp.columns.tolist()
        d_student.insert(0, h_student)
        for i in range(len(d_student)):
            if not i:
                d_student[i].insert(6, 'Nhóm')
            else:
                d_student[i].insert(6, '')

        tmp = pandas.read_excel(f_topic)
        tmp = tmp.fillna('')
        d_topic = tmp.values.tolist()
        h_topic = tmp.columns.tolist()
        d_topic.insert(0, h_topic)

        datafile = getUpdateSheet(d_student, d_topic, datafile)

        data = json.dumps(datafile)
        response = HttpResponse(data, content_type='application/json')
        response['Content-Disposition'] = 'attachment; filename=my_file.json'
        response.set_cookie('spreadsheetId', datafile['spreadsheetId'])
        response.set_cookie('formId', datafile['formId'])
        return response


def addSheet(request):
    # datafile = readLocal(request)
    spreadsheetId = request.COOKIES.get('spreadsheetId', 1)
    formId = request.COOKIES.get('formId', 1)
    if (spreadsheetId == 1):
        spreadsheetId = ''
    if (formId == 1):
        formId = ''
    datafile = {'spreadsheetId': spreadsheetId, 'formId': formId}
    print(datafile)
    msg = addRegisterSheet(datafile)
    print(msg)
    if msg:
        return HttpResponse(msg)
    return render(request, 'pageSheet.html')


def gotoSheet(request):
    spreadsheetId = request.COOKIES.get('spreadsheetId', 1)
    formId = request.COOKIES.get('formId', 1)
    if (spreadsheetId == 1):
        spreadsheetId = ''
    if (formId == 1):
        formId = ''
    datafile = {'spreadsheetId': spreadsheetId, 'formId': formId}
    print(datafile)
    link = getLinkSheet(datafile)
    if not link[0]:
        return HttpResponse(link[1])
    return HttpResponseRedirect(link[1])


def createForm(request):
    spreadsheetId = request.COOKIES.get('spreadsheetId', 1)
    formId = request.COOKIES.get('formId', 1)
    if (spreadsheetId == 1):
        spreadsheetId = ''
    if (formId == 1):
        formId = ''
    datafile = {'spreadsheetId': spreadsheetId, 'formId': formId}
    print(datafile)
    msg = updateForm(datafile)
    if msg[1]:
        return HttpResponse(msg[1])
    datafile = msg[0]
    # writeLocal(request, datafile)
    data = json.dumps(datafile)
    response = HttpResponse(data, content_type='application/json')
    response['Content-Disposition'] = 'attachment; filename=my_file.json'
    response.set_cookie('spreadsheetId', datafile['spreadsheetId'])
    response.set_cookie('formId', datafile['formId'])
    return response


def gotoForm(request):
    spreadsheetId = request.COOKIES.get('spreadsheetId', 1)
    formId = request.COOKIES.get('formId', 1)
    if (spreadsheetId == 1):
        spreadsheetId = ''
    if (formId == 1):
        formId = ''
    datafile = {'spreadsheetId': spreadsheetId, 'formId': formId}
    print(datafile)
    link = getFormLink(datafile)
    if not link[0]:
        return HttpResponse('Hãy tạo Form!')
    return HttpResponseRedirect(link[1])


def downloadResult(request):
    spreadsheetId = request.COOKIES.get('spreadsheetId', 1)
    formId = request.COOKIES.get('formId', 1)
    if (spreadsheetId == 1):
        spreadsheetId = ''
    if (formId == 1):
        formId = ''
    datafile = {'spreadsheetId': spreadsheetId, 'formId': formId}
    print(datafile)
    result1 = getResponseForm(datafile)
    if not result1:
        return HttpResponse('Chưa có kết quả đánh giá')
    if result1[0] == 1:
        return HttpResponse(' Có lỗi khi lấy kết quả!\n hãy tạo lại đánh giá mới')
    max_lenght = 0
    for i in range(len(result1)):
        if (max_lenght < len(result1[i])):
            max_lenght = len(result1[i])
    num = (max_lenght - 5) / 2
    num = int(num)
    column1 = ['Được chấp nhận', 'Email đã đăng ký', 'Email đánh giá', 'Nhóm', 'Sinh viên đánh giá']
    for i in range(num):
        column1.append('Thành viên thứ ' + str(i + 1) )
        column1.append('Điểm')

    # print(column1)
    df1 = pandas.DataFrame(result1, columns=column1)

    result2 = getAverageMark(datafile)
    if not result2:
        df2 = [[]]
    else:
        column2 = ['Nhóm', 'Tên', 'Điểm trung bình']
        df2 = pandas.DataFrame(result2, columns=column2)

    # df = pandas.DataFrame(result, columns=['Nhóm', 'Tên', 'Điểm'])
    with BytesIO() as b:
        # Use the StringIO object as the filehandle.
        writer = pandas.ExcelWriter(b, engine='xlsxwriter')
        df1.to_excel(writer, sheet_name='Phản hồi', index=False)
        df2.to_excel(writer, sheet_name='Điểm trung bình', index=False)
        writer._save()
        filename = 'test'
        content_type = 'application/vnd.ms-excel'
        response = HttpResponse(b.getvalue(), content_type=content_type)
        response['Content-Disposition'] = 'attachment; filename="' + filename + '.xlsx"'
        return response


# các page chính và page testa
def mainPage(request):
    return render(request, 'mainPage.html')


def ssheet(request):
    return render(request, 'pageSheet.html')


def form(request):
    return render(request, 'pageForm.html')
