from django.urls import path
from . import views

app_name = 'myapp'

urlpatterns = [
    path('', views.mainPage),
    path('upload', views.updateSheet, name='upload'),
    path('addSheet', views.addSheet, name='addSheet'),
    path('gotoSheet', views.gotoSheet, name='gotoSheet'),
    path('createForm', views.createForm, name='createForm'),
    path('copyLinkForm', views.copyLinkForm, name='copyLinkForm'),
    path('downloadResult', views.downloadResult, name='downloadResult'),

    path('mainPage', views.mainPage, name='mainPage'),
    path('pageSheet', views.ssheet, name='pageSheet'),
    path('pageForm', views.form, name='pageForm'),

    path('testa', views.testa, name='testa'),
]
