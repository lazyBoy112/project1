from django.shortcuts import render, HttpResponse
import pandas as pandas
import base.appscript as ascript
import webbrowser
import pyperclip
from io import BytesIO
import json


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
    print(datafile)
    return response


def updateSheet(request):
    # datafile = readLocal(request)
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
        d_student = tmp.to_csv()
        d_student = d_student.split('\r')
        for i in range(len(d_student)):
            if i:
                d_student[i] = d_student[i][3:]
            else:
                d_student[i] = d_student[i][1:]
        for i in range(len(d_student)):
            d_student[i] = d_student[i].split(',')
        for i in range(len(d_student)):
            if not i:
                d_student[i].insert(6, 'Nhóm')
            else:
                d_student[i].insert(6, '')

        tmp = pandas.read_excel(f_topic)
        d_topic = tmp.to_csv()
        d_topic = d_topic.split('\r')
        for i in range(len(d_topic)):
            if i:
                d_topic[i] = d_topic[i][3:]
            else:
                d_topic[i] = d_topic[i][1:]
        for i in range(len(d_topic)):
            d_topic[i] = d_topic[i].split(',')

        datafile = ascript.updateSheet(d_student, d_topic, datafile)
        # writeLocal(request, datafile)

        data = json.dumps(datafile)
        response = HttpResponse(data, content_type='application/json')
        response['Content-Disposition'] = 'attachment; filename=my_file.json'
        response.set_cookie('spreadsheetId', datafile['spreadsheetId'])
        response.set_cookie('formId', datafile['formId'])
        # return render(request, 'pageSheet.html')
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
    msg = ascript.addRegisterSheet(datafile)
    print(msg)
    if msg:
        return HttpResponse(msg)
    return render(request, 'pageSheet.html')


def gotoSheet(request):
    # datafile = readLocal(request)
    spreadsheetId = request.COOKIES.get('spreadsheetId', 1)
    formId = request.COOKIES.get('formId', 1)
    if (spreadsheetId == 1):
        spreadsheetId = ''
    if (formId == 1):
        formId = ''
    datafile = {'spreadsheetId': spreadsheetId, 'formId': formId}
    print(datafile)
    link = ascript.getLinkSheet(datafile)
    if not link[0]:
        return HttpResponse(link[1])
    webbrowser.open(link[1])
    return render(request, 'pageSheet.html')


def createForm(request):
    # datafile = readLocal(request)
    spreadsheetId = request.COOKIES.get('spreadsheetId', 1)
    formId = request.COOKIES.get('formId', 1)
    if (spreadsheetId == 1):
        spreadsheetId = ''
    if (formId == 1):
        formId = ''
    datafile = {'spreadsheetId': spreadsheetId, 'formId': formId}
    print(datafile)
    msg = ascript.updateForm(datafile)
    if msg[1]:
        return HttpResponse(msg[1])
    datafile = msg[0]
    # writeLocal(request, datafile)
    data = json.dumps(datafile)
    response = HttpResponse(data, content_type='application/json')
    response['Content-Disposition'] = 'attachment; filename=my_file.json'
    response.set_cookie('spreadsheetId', datafile['spreadsheetId'])
    response.set_cookie('formId', datafile['formId'])
    # return render(request, 'pageForm.html')
    return response


def copyLinkForm(request):
    # datafile = readLocal(request)
    spreadsheetId = request.COOKIES.get('spreadsheetId', 1)
    formId = request.COOKIES.get('formId', 1)
    if (spreadsheetId == 1):
        spreadsheetId = ''
    if (formId == 1):
        formId = ''
    datafile = {'spreadsheetId': spreadsheetId, 'formId': formId}
    print(datafile)
    link = ascript.getFormLink(datafile)
    if not link[0]:
        return HttpResponse('Hãy tạo Form!')
    pyperclip.copy(link[1])
    return render(request, 'pageForm.html')


def downloadResult(request):
    # datafile = readLocal(request)
    spreadsheetId = request.COOKIES.get('spreadsheetId', 1)
    formId = request.COOKIES.get('formId', 1)
    if (spreadsheetId == 1):
        spreadsheetId = ''
    if (formId == 1):
        formId = ''
    datafile = {'spreadsheetId': spreadsheetId, 'formId': formId}
    print(datafile)
    result1 = ascript.getResponseForm(datafile)
    if result1[0] == 1:
        return HttpResponse(' Có lỗi khi lấy kết quả!\n hãy tạo lại đánh giá mới')
    max_lenght = 0
    if result1:
        for i in range(len(result1)):
            if (max_lenght < len(result1[i])):
                max_lenght = len(result1[i])
        num = (max_lenght - 5) / 2
        num = int(num)
        column1 = ['Được chấp nhận', 'Email đã đăng ký', 'Email đánh giá', 'Nhóm', 'Sinh viên đánh giá']
        for i in range(num):
            column1.append('Thành viên thứ ' + str(i + 1) )
            column1.append('Điểm')

        print(column1)
        df1 = pandas.DataFrame(result1, columns=column1)
    else:
        return HttpResponse(' Chưa có phản hồi từ sinh viên!')

    result2 = ascript.getAverageMark(datafile)
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
