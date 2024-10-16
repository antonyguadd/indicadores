"""
URL configuration for indicadores_api project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/5.1/topics/http/urls/
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
from indicadores.views import upload_excel, upload_and_generate_indicators, download_file, ordenes_mes, download_ordenes_mes_file, UserCreate
from rest_framework_simplejwt.views import TokenObtainPairView, TokenRefreshView

urlpatterns = [
    path('admin/', admin.site.urls),
    path('api/token/', TokenObtainPairView.as_view(), name='token_obtain_pair'),
    path('api/token/refresh/', TokenRefreshView.as_view(), name='token_refresh'),
    path('api/upload/', upload_excel, name='upload_excel'),
    path('api/upload_and_generate/', upload_and_generate_indicators, name='upload_and_generate'),
    path('api/download/', download_file, name='download_file'),
    path('api/ordenes_mes/', ordenes_mes, name='ordenes_mes'),
    path('api/download_ordenes_mes/', download_ordenes_mes_file, name='download_ordenes_mes_file'),
    path('api/register/', UserCreate.as_view(), name='register'),
]