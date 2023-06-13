import tkinter as tk
from tkinter import ttk, filedialog
import spreadsheet.spreadsheet as ssheet
import pandas as pandas
import json as json
import webbrowser as webbrowser
from tkinter import messagebox


def start():
    mainWindow = tk.Tk()
    mainWindow.geometry('500x800')

    notebook = ttk.Notebook(mainWindow)
    notebook.pack()

    # tab1 = ttk.Frame(notebook) # width, height
    tab1 = Tab1(notebook)
    tab2 = Tab2(notebook)  # width, height
    tab3 = Tab3(notebook)  # width, height

    tab1.pack(fill=tk.BOTH, expand=True)
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
    # get link
    getLinkBtn = ttk.Button(tab1, text='lấy link sheet')

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
            with open('data.json', 'r') as datafile:
                data = json.load(datafile)
            datafile.close()
            new_id = ssheet.create('Danh sách đăng ký nhóm', 'Danh sách nhóm')
            data['spreadsheet'] = new_id
            jsondata = json.dumps(data)
            with open('data.json', 'w') as datafile:
                datafile.write(jsondata)
            pass
        with open('data.json', 'r') as datafile:
            data = json.load(datafile)
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
        ssheet.clear(spreadsheetId, 'Danh sách nhóm')
        ssheet.updateValue(spreadsheetId, 'Danh sách nhóm', sheet_data)
        # sheet_data

    def getLink():
        with open('data.json', 'r') as datafile:
            data = json.load(datafile)
        spreadsheetId = data['spreadsheet']
        if not spreadsheetId:
            return ''
        link = ssheet.getLink(spreadsheetId)
        return link

    def gotoLink():
        link = getLink()
        if not link:
            pass
            # khoi tao msg box lỗi
        webbrowser.open(link)

    uploadBtn.configure(command=browseFile)
    updateBtn.configure(command=update)
    getLinkBtn.configure(command=gotoLink)

    uploadBtn.pack()
    uploadStatusLabel.pack()
    updateBtn.pack()
    getLinkBtn.pack()
    return tab1


def Tab2(notebook):
    tab2 = ttk.Frame(notebook)
    # update Group
    # label update
    # topic Btn
    # label topic
    # update Btn
    # get link

    return tab2


def Tab3(notebook):
    tab3 = ttk.Frame(notebook)
    # update form  =>> msg
    # viewform
    # viewresult
    # get link

    return tab3
