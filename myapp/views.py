from django.shortcuts import render, HttpResponse


def createNewProject(request):
    return render(request, 'pageSheet.html')


def openProject(request):
    return render(request, 'pageSheet.html')


def createNewFile(request):
    return render(render, "")


def updateSheet(request):
    return render(request, 'pageSheet.html')


def addSheet(request):
    return render(request, 'pageSheet.html')


def gotoSheet(request):
    return render(request, 'pageSheet.html')


def createForm(request):
    return render(request, 'pageForm.html')


def copyLinkForm(request):
    return render(request, 'pageForm.html')


def downloadResult(request):
    return render(request, 'pageForm.html')


# các page chính và page testa
def mainPage(request):
    return render(request, 'mainPage.html')


def ssheet(request):
    return render(request, 'pageSheet.html')


def form(request):
    return render(request, 'pageForm.html')


def testa(request):
    if request.method == 'POST':
        f_student = request.FILES.get('studentList')
        print(f_student.get_path())

    return render(request, 'mainPage.html')

