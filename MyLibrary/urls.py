from django.contrib import admin
from django.urls import path, include
from MyApp import views
from django.views.generic.base import TemplateView

urlpatterns = [
    path('admin/', admin.site.urls),
    path('api/', include("MyApp.urls")),  # 添加一个新的路由
]
