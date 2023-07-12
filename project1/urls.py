"""project1 URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/4.1/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path, include

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('home.urls')),
    path('upload/', include('upload.urls',)),
    path('download/', include('download.urls',)),
    path('upload/upload_dssv/', include('upload_dssv.urls')),
    path('upload/upload_detai/', include('upload_detai.urls')),
    path('upload/cap_nhat/', include('cap_nhat.urls')),
    path('upload/tao_trang_dang_ky_de_tai/', include('tao_trang_dang_ky_de_tai.urls')),
    path('upload/di_den_trang/', include('di_den_trang.urls')),
    path('upload/xoa/', include('xoa.urls')),
    path('download/tao_form/', include('tao_form.urls',)),
    path('download/copy_link_form/', include('copy_link_form.urls',)),
    path('download/tai_xuong_ket_qua/', include('tai_xuong_ket_qua.urls',)),

]
