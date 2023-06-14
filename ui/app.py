import tkinter as tk
from tkinter import ttk, filedialog
import spreadsheet.spreadsheet as ssheet
import pandas as pandas
import json as json
import webbrowser as webbrowser
from tkinter import messagebox
import pyperclip
import os


def start():
    mainWindow = tk.Tk()
    mainWindow.geometry('500x800')
    if not os.path.exists('data.json'):
        data = {"spreadsheet": "", "form": "", "linkSheet": ""}
        data = json.dumps(data)
        with open('data.json', 'w') as file:
            file.write(data)

    notebook = ttk.Notebook(mainWindow)
    notebook.pack()

    # tab1 = ttk.Frame(notebook) # width, height
    tab1 = Tab1(notebook)
    tab2 = Tab2(notebook)  # width, height
    tab3 = Tab3(notebook)  # width, height

    tab1.pack(fill='both', expand=True)
    tab2.pack(fill='both', expand=True)
    tab3.pack(fill='both', expand=True)

    notebook.add(tab1, text='Tab1')
    notebook.add(tab2, text='Tab2')
    notebook.add(tab3, text='Tab3')

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
        ssheet.addSheet(new_id, 'Đăng ký đề tài')
        ssheet.addSheet(new_id, 'Danh sách đề tài')
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

    uploadBtn.configure(command=browseFile)
    updateBtn.configure(command=update)
    copyLinkBtn.configure(command=copyLink)
    getLinkBtn.configure(command=gotoLink)

    uploadBtn.pack()
    uploadStatusLabel.pack()
    updateBtn.pack()
    copyLinkBtn.pack()
    getLinkBtn.pack()
    return tab1


def Tab2(notebook):
    tab2 = ttk.Frame(notebook)
    # update Group
    updateGroupBtn = ttk.Button(tab2, text='Cập nhật')
    # label update
    updateGroupLabel = ttk.Label(tab2)
    # topic Btn
    uploadTopicBtn = ttk.Button(tab2, text='Tải lên danh sách chủ đề')
    # label topic
    uploadTopicLabel = ttk.Label(tab2)
    # update Btn
    updateBtn = ttk.Button(tab2, text='Cập nhật')
    # get link
    copyLinkBtn = ttk.Button(tab2, text='Copy link sheet')
    # goto link
    gotoLinkBtn = ttk.Button(tab2, text='Đi đến sheet')

    updateGroupBtn.pack()
    updateGroupLabel.pack()
    uploadTopicBtn.pack()
    uploadTopicLabel.pack()
    updateBtn.pack()
    copyLinkBtn.pack()
    gotoLinkBtn.pack()

    return tab2


def Tab3(notebook):
    tab3 = ttk.Frame(notebook)
    # update form  =>> msg
    # viewform
    # viewresult
    # get link

    return tab3
