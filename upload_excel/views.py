from django.shortcuts import render
import pandas as pd
# Create your views here.
from django.shortcuts import render, redirect
from django.contrib import messages

import gspread
from oauth2client.service_account import ServiceAccountCredentials

def upload_excel(request): # Upload 1 file excel lên
    if request.method == 'POST':
        file = request.FILES['excel_file']
        df = pd.read_excel(file) 
        # Xử lý dữ liệu trong DataFrame       
        # Ghi dữ liệu lên Google Sheets
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(credentials)
        spreadsheet_id = '1f6bMQfcU9W4IKLhpcX_fe3e6askDl0EZkaAt4HufiCg'
        spreadsheet = client.open_by_key(spreadsheet_id)
        sheet = spreadsheet.worksheet('sheet1')
        worksheet = spreadsheet.worksheet('sheet1') 
        # Ghi dữ liệu từ DataFrame vào Google Sheets
        worksheet.update('A1', df.values.tolist())   
        #return render(request, 'home1.html')
        messages.success(request, 'Đã tải lên tệp Excel')
        return redirect('home1')
    return render(request, 'upload_excel.html')


def upload_excel_another(request): # Upload file excel thứ hai lên
    if request.method == 'POST':
        file = request.FILES['excel_file']
        df = pd.read_excel(file) 
        # Xử lý dữ liệu trong DataFrame       
        # Ghi dữ liệu lên Google Sheets
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(credentials)
        spreadsheet_id = '1f6bMQfcU9W4IKLhpcX_fe3e6askDl0EZkaAt4HufiCg'
        spreadsheet = client.open_by_key(spreadsheet_id)
        sheet = spreadsheet.worksheet('sheet2')
        worksheet = spreadsheet.worksheet('sheet2') 
        # Ghi dữ liệu từ DataFrame vào Google Sheets
        worksheet.update('A1', df.values.tolist())   
    #return render(request, 'home1.html')
        messages.success(request, 'Đã tải lên tệp Excel')
        return redirect('home1')
    return render(request, 'upload_excel_another.html')


def create_sheet(request):
    if request.method == 'POST':
        # Thực hiện xử lý khi nhấp vào nút
        # Kết nối và xác thực tới Google Sheets API
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(credentials)      
        # Mở tài liệu Google Sheets
        spreadsheet_id = '1f6bMQfcU9W4IKLhpcX_fe3e6askDl0EZkaAt4HufiCg'
        spreadsheet = client.open_by_key(spreadsheet_id)    
        # Tạo một trang tính mới theo mẫu
        new_sheet_name = 'Sheet3'
        template_sheet_name = 'sheet1'
        template_sheet = spreadsheet.worksheet(template_sheet_name)
        #spreadsheet.duplicate_sheet(template_sheet.id, new_sheet_name,insert_sheet_index=2)  
        spreadsheet.duplicate_sheet( template_sheet.id, new_sheet_name, insert_sheet_index= 1)

        # Thực hiện các thao tác khác với trang tính mới
        # Chuyển hướng về trang gốc sau khi tạo xong
        return redirect('home1')
    return render(request, 'create_sheet.html')


def home1(request):
    return render(request, 'home1.html')    