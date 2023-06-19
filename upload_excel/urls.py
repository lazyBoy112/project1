from django.urls import path
from .views import upload_excel
from . import views

urlpatterns = [
    path('', views.home1, name='home1'),
    path('upload_excel/', views.upload_excel, name='upload_excel'),
    path('upload_excel_another/', views.upload_excel_another, name='upload_excel_another'),
    path('create_sheet/', views.create_sheet, name='create_sheet'),
    
]
