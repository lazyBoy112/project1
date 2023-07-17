from django.shortcuts import render, HttpResponse
import pandas as pandas
import base.appscript as ascript
import webbrowser
import pyperclip
from io import BytesIO


def index(request):
    return render(request, "index.html")


def updateSheet(request):
    if request.method == 'POST':
        f_student = request.FILES.get('studentList')
        f_topic = request.FILES.get('topicList')
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

        print(d_topic)
        ascript.updateSheet(d_student, d_topic)

        return render(request, 'pageSheet.html')


def addSheet(request):
    msg = ascript.addRegisterSheet()
    print(msg)
    if msg:
        return HttpResponse(msg)
    return render(request, 'pageSheet.html')


def gotoSheet(request):
    link = ascript.getLinkSheet()
    if not link[0]:
        return HttpResponse(link[1])
    webbrowser.open(link[1])
    return render(request, 'pageSheet.html')


def createForm(request):
    msg = ascript.updateForm()
    if not msg[0]:
        return HttpResponse(msg[1])
    return render(request, 'pageForm.html')


def copyLinkForm(request):
    link = ascript.getFormLink()
    if not link:
        return HttpResponse(' Hãy tạo Form!')
    pyperclip.copy(link)
    return render(request, 'pageForm.html')


def downloadResult(request):
    result = ascript.getResultForm()
    if not result:
        return HttpResponse(' Có lỗi khi lấy kết quả!')
    df = pandas.DataFrame(result, columns=['Nhóm', 'Tên', 'Điểm'])
    # response = HttpResponse(content_type='application/xlsx')
    # response['Content-Disposition'] = f'attachment; filename="FILENAME.xlsx"'
    # with pandas.ExcelWriter(response) as writer:
    #     df.to_excel(writer, sheet_name='SHEET NAME')
    # return response
    # return render(request, 'pageForm.html')
    with BytesIO() as b:
        # Use the StringIO object as the filehandle.
        writer = pandas.ExcelWriter(b, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', index=False)
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


def testa(request):
    print(ascript.test())
    return render(request, 'pageSheet.html')

