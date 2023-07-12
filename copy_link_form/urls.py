from django.urls import path
from . import views 

urlpatterns = [
    path('', views.copy_link_form, name='copy_link_form'),
]