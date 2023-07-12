from django.urls import path
from . import views 

urlpatterns = [
    path('', views.cap_nhat, name='cap_nhat'),
    
]