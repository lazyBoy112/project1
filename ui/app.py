import tkinter as tk
from tkinter import ttk, filedialog
import spreadsheet.spreadsheet as ssheet
import form.form as form
import pandas as pandas
import json as json
import webbrowser as webbrowser
from tkinter import messagebox
import pyperclip
import os


def start():
    mainWindow = tk.Tk()
    mainWindow.geometry('500x800')

    # kiểm tra tồn tại file data.json
    if not os.path.exists('data.json'):
        data = {"spreadsheet": "", "form": "", "linkSheet": ""}
        data = json.dumps(data)
        with open('data.json', 'w') as file:
            file.write(data)

    notebook = ttk.Notebook(mainWindow)
    notebook.pack()

    tab1 = Tab1(notebook)
    tab2 = Tab2(notebook)

    tab1.pack(fill='both', expand=True)
    tab2.pack(fill='both', expand=True)

    notebook.add(tab1, text='Danh sách đăng ký')
    notebook.add(tab2, text='Form đánh giá')

    mainWindow.mainloop()


def Tab1(notebook):
    tab1 = ttk.Frame(notebook)
    # student list Btn
    uploadBtn = ttk.Button(tab1, text='Tải lên danh sách sinh viên')   # upload excel file
    # student list label
    uploadStatusLabel = ttk.Label(tab1)
    # update Btn
    updateBtn = ttk.Button(tab1, text='Cập nhật')
    # copy link
    copyLinkBtn = ttk.Button(tab1, text='copy link sheet')
    # get link
    getLinkBtn = ttk.Button(tab1, text='đi đến sheet')
    # topic btn
    uploadTopicBtn = ttk.Button(tab1, text='Tải lên danh sách đề tài')
    # topic status
    uploadTopicLabel = ttk.Label(tab1)
    # cập nhật danh sách đề tài 
    updateTopicBtn = ttk.Button(tab1, text='Cập nhật danh sách đề tài')
    # tạo danh sách đăng ký bài tập lớn
    createTopicBtn = ttk.Button(tab1, text='Tạo danh sách đăng ký chủ đề')
    # add event

    def browseFile():
        filename = filedialog.askopenfilename(title='Select file',
                                              filetypes=(('Excel files', '*.xlxs, *.xls'),
                                                         ('all files', '*.*'))
                                              )
        uploadStatusLabel.configure(text=filename)

    def update():
        path = uploadStatusLabel['text']
        if not path:
            messagebox.showerror('Lỗi', 'Không có file danh sách tải lên!')
            return
        if not getLink():
            createTemplateSheet()
        with open('data.json', 'r') as datafile:
            data = json.load(datafile)
        # lấy danh sách từ file excel và thêm cột nhóm vào
        spreadsheetId = data['spreadsheet']
        excelData = pandas.read_excel(path)
        bodyData = excelData.values.tolist()
        column = excelData.columns.values
        sheet_data = bodyData
        column = list(column)
        sheet_data.insert(0, column)
        for i in range(len(sheet_data)):
            if not i:
                sheet_data[i].insert(6, 'Nhóm')
            else:
                sheet_data[i].insert(6, '')
            for j in range(len(sheet_data[i])):
                if str(sheet_data[i][j]) == "nan":
                    sheet_data[i][j] = ''
        ssheet.clear(spreadsheetId, 'Danh sách nhóm')
        ssheet.updateValue(spreadsheetId, 'Danh sách nhóm', sheet_data)
        # sheet_data

    def createTemplateSheet():
        with open('data.json', 'r') as datafile:
            data = json.load(datafile)
        datafile.close()
        new_id = ssheet.create('Danh sách đăng ký nhóm', 'Danh sách nhóm')
        ssheet.addSheet(new_id, 'Danh sách đề tài')
        ssheet.addSheet(new_id, 'Đăng ký đề tài')
        data['spreadsheet'] = new_id
        jsondata = json.dumps(data)
        with open('data.json', 'w') as datafile:
            datafile.write(jsondata)
        return new_id


    def getLink():
        with open('data.json', 'r') as datafile:
            data = json.load(datafile)
        spreadsheetId = data['spreadsheet']
        if not spreadsheetId:
            return ''
        link = ssheet.getLink(spreadsheetId)
        return link

    def copyLink():
        link = getLink()
        if not link:
            # khoi tao msg box lỗi
            messagebox.showerror('lỗi googlesheet', 'không có sự phản hồi từ sheet!')
            messagebox.showwarning('Khởi tạo sheet mới', 'Đang khởi tạo sheet mới trong localdata')
            createTemplateSheet()
            messagebox.showinfo('Nhắc lệnh', 'Vui lòng cập nhật lại sheet')
        pyperclip.copy(link)

    def gotoLink():
        link = getLink()
        if not link:
            # khoi tao msg box lỗi
            messagebox.showerror('lỗi googlesheet', 'không có sự phản hồi từ sheet!')
            messagebox.showwarning('Khởi tạo sheet mới', 'Đang khởi tạo sheet mới trong localdata')
            createTemplateSheet()
            messagebox.showinfo('Nhắc lệnh', 'Vui lòng cập nhật lại sheet')
            return
        webbrowser.open(link)

    def updateTopic():
        path = uploadTopicLabel['text']
        if not path:
            messagebox.showerror('Lỗi', 'Không có file danh sách tải lên!')
            return
        if not getLink():
            createTemplateSheet()
        with open('data.json', 'r') as datafile:
            data = json.load(datafile)
        # lấy danh sách từ file excel và thêm cột nhóm vào
        spreadsheetId = data['spreadsheet']
        excelData = pandas.read_excel(path)
        bodyData = excelData.values.tolist()
        column = excelData.columns.values
        sheet_data = bodyData
        column = list(column)
        sheet_data.insert(0, column)
        for i in range(len(sheet_data)):
            for j in range(len(sheet_data[i])):
                if str(sheet_data[i][j]) == "nan":
                    sheet_data[i][j] = ''
        ssheet.clear(spreadsheetId, 'Danh sách đề tài')
        ssheet.updateValue(spreadsheetId, 'Danh sách đề tài', sheet_data)

    def TopicResisger():
        if not getLink():
            createTemplateSheet()
        with open('data.json', 'r') as datafile:
            data = json.load(datafile)
        # lấy danh sách từ file excel và thêm cột nhóm vào
        spreadsheetId = data['spreadsheet']
        sheet_data = list()
        raw_value = ssheet.getColumnSheet(spreadsheetId, 'Danh sách nhóm', 'G')
        raw_value.pop(0)
        tmp_1 = list()
        for i in range(len(raw_value)):
            tmp_1.insert(0, raw_value[i][0])

        tmp_1.reverse()
        tmp_2 = dict.fromkeys(tmp_1)
        tmp_1 = list(tmp_2)
        sheet_data.insert(0, ['Tên nhóm', 'Đề tài đăng ký'])
        for i in range(len(tmp_1)):
            sheet_data.insert(0, [tmp_1[i], ''])

        sheet_data.reverse()

        ssheet.clear(spreadsheetId, 'Đăng ký đề tài')
        ssheet.updateValue(spreadsheetId, 'Đăng ký đề tài', sheet_data)

    uploadBtn.configure(command=browseFile)
    updateBtn.configure(command=update)
    uploadTopicBtn.configure(command=browseFile)
    updateTopicBtn.configure(command=updateTopic)
    createTopicBtn.configure(command=TopicResisger)
    copyLinkBtn.configure(command=copyLink)
    getLinkBtn.configure(command=gotoLink)

    uploadBtn.pack()
    uploadStatusLabel.pack()
    updateBtn.pack()
    uploadTopicBtn.pack()
    uploadTopicLabel.pack()
    updateTopicBtn.pack()
    copyLinkBtn.pack()
    getLinkBtn.pack()
    createTopicBtn.pack()
    return tab1


def Tab2(notebook):
    tab2 = ttk.Frame(notebook)
    # nhập email
    inputEmail = ttk.Entry(tab2)
    # tạo form 
    createFormBtn = ttk.Button(tab2, text='Tạo form')
    # Xem kết quả 
    gotoResultBtn = ttk.Button(tab2, text='Xem kết quả')
    # copyLink
    copylinkBtn = ttk.Button(tab2, text='copylink')
    # Tải xuống kết quả
    downloadBtn = ttk.Button(tab2, text='Tải xuống kết quả')
    # xóa form cũ
    deleteFormBtn = ttk.Button(tab2, text='Xóa form cũ')

    def createForm():
        email = inputEmail.get()
        if not email:
            messagebox.showinfo('Email', 'chưa nhập email')
            return
        if not checkEmail(email):
            messagebox.showinfo('Email', 'email phải có định dạng @gmail.com')
            return
        if not getLink():
            createTemplateForm()
        with open('data.json', 'r') as datafile:
            data = json.load(datafile)
        formId = data['form']
        spreadsheetId = data['spreadsheet']

        raw_value = ssheet.getColumnSheet(spreadsheetId, 'Đăng ký đề tài', 'A')
        raw_value.pop(0)
        group = list()
        for i in range(len(raw_value)):
            group.insert(0, raw_value[i][0])
        group.reverse()

        raw_value = ssheet.getColumnSheet(spreadsheetId, 'Danh sách đề tài', 'B')
        raw_value.pop(0)
        topic = list()
        for i in range(len(raw_value)):
            topic.insert(0, raw_value[i][0])
        topic.reverse()
        form.update(formId, group, topic)
        form.permission(formId, email)

    def checkEmail(email):
        t = '@gmail.com'
        if len(email) < 10:
            return False
        for i in range(10):
            if t[-i - 1] != email[-i - 1]:
                return False
        return True

    def createTemplateForm():
        with open('data.json', 'r') as datafile:
            data = json.load(datafile)
        datafile.close()
        new_id = form.create()  # rewrite
        data['form'] = new_id
        jsondata = json.dumps(data)
        with open('data.json', 'w') as datafile:
            datafile.write(jsondata)
        return new_id

    def getLink():
        with open('data.json', 'r') as datafile:
            data = json.load(datafile)
        formId = data['form']
        if not formId:
            return ''
        link = form.getFormLink(formId)
        return link

    def gotoResult():
        with open('data.json', 'r') as datafile:
            data = json.load(datafile)
        formId = data['form']
        if not formId:
            formId = createTemplateForm()
        link = form.getResultLink(formId)
        if not link:
            # khoi tao msg box lỗi
            messagebox.showerror('lỗi google form', 'không có sự phản hồi từ form!')
            messagebox.showwarning('Khởi tạo form mới', 'Đang khởi tạo form mới trong localdata')
            createTemplateForm()
            messagebox.showinfo('Nhắc lệnh', 'Vui lòng cập nhật lại form')
            return
        webbrowser.open(link)


    def copyLink():
        link = getLink()
        if not link:
            # khoi tao msg box lỗi
            messagebox.showerror('lỗi googlesheet', 'không có sự phản hồi từ sheet!')
            messagebox.showwarning('Khởi tạo sheet mới', 'Đang khởi tạo sheet mới trong localdata')
            createTemplateForm()
            messagebox.showinfo('Nhắc lệnh', 'Vui lòng cập nhật lại sheet')
        pyperclip.copy(link)

    def download():
        pass

    def deleteForm():
        if not getLink():
            return
        with open('data.json', 'r') as datafile:
            data = json.load(datafile)
        formId = data['form']
        form.deleteFile(formId)
        data['form'] = ''
        jsondata = json.dumps(data)
        with open('data.json', 'w') as datafile:
            datafile.write(jsondata)



    createFormBtn.configure(command=createForm)
    gotoResultBtn.configure(command=gotoResult)
    copylinkBtn.configure(command=copyLink)
    downloadBtn.configure(command=download)
    deleteFormBtn.configure(command=deleteForm)

    inputEmail.pack()
    createFormBtn.pack()
    gotoResultBtn.pack()
    copylinkBtn.pack()
    downloadBtn.pack()
    deleteFormBtn.pack()

    return tab2
