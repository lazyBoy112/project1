from django.urls import path
from . import views 

urlpatterns = [
    path('', views.tai_xuong_ket_qua, name='tai_xuong_ket_qua'),
]