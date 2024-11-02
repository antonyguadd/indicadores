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
from indicadores.views.vistas_dia import upload_excel, upload_and_generate_indicators, download_file, UserCreate
from indicadores.views.vistas_mes import upload_excel_mes, upload_and_generate_monthly_indicators, download_file_mes
from rest_framework_simplejwt.views import TokenObtainPairView, TokenRefreshView
from indicadores.views.vistas_reportes import upload_excel_report, upload_and_generate_report_indicators, download_file_report

urlpatterns = [
    path('admin/', admin.site.urls),
    path('api/token/', TokenObtainPairView.as_view(), name='token_obtain_pair'),
    path('api/token/refresh/', TokenRefreshView.as_view(), name='token_refresh'),
    path('api/upload/', upload_excel, name='upload_excel'),
    path('api/upload_and_generate/', upload_and_generate_indicators, name='upload_and_generate'),
    path('api/download/', download_file, name='download_file'),
    path('api/register/', UserCreate.as_view(), name='register'),
    # Ruta para subir el archivo Excel
    path('api/upload_mes/', upload_excel_mes, name='upload_excel_mes'),

    # Ruta para procesar el archivo y generar indicadores mensuales
    path('api/generate_monthly_indicators/', upload_and_generate_monthly_indicators, name='generate_monthly_indicators'),

    # Ruta para descargar el archivo Excel generado con indicadores mensuales
    path('api/download_mes/', download_file_mes, name='download_file_mes'),
    
    # Rutas para reportes
    path('api/upload_report/', upload_excel_report, name='upload_excel_report'),
    path('api/generate_report_indicators/', upload_and_generate_report_indicators, name='generate_report_indicators'),
    path('api/download_report/', download_file_report, name='download_file_report'),
    

]