from django.urls import path
from . import views 

urlpatterns = [
    path('', views.tao_form, name='tao_form'),
]