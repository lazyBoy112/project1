from django.urls import path
from . import views

urlpatterns = [
    path('', views.di_den_trang, name='di_den_trang'),
]