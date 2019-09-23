"""wechat_punch URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/2.2/topics/http/urls/
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
from django.urls import path
from wechat_punch import views as wechat_punch_views
from convert_excel import views as hello_world_views 

urlpatterns = [
    path('',wechat_punch_views.index),
    path('admin/', admin.site.urls),
    path('convert/', hello_world_views.index),
    path('upload_file/',hello_world_views.upload_file)
]
