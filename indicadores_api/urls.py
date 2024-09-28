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
from indicadores.views import upload_excel, upload_form

urlpatterns = [
    path('admin/', admin.site.urls),
    path('upload/', upload_form, name='upload_form'),  # Formulario de subida
    path('api/upload/', upload_excel, name='upload_excel'),  # Ruta para procesar el archivo
    path('api/', include('indicadores.urls')),  # Incluye rutas de la app indicadores
]
