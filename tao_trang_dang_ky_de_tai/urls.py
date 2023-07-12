from django.urls import path
from . import views

urlpatterns = [
    path('', views.tao_trang_dang_ky_de_tai, name='tao_trang_dang_ky_de_tai'),
]