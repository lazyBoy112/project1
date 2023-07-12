from django.urls import path
from . import views

urlpatterns = [
    path('', views.xoa, name='xoa'),
]