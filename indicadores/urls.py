from django.urls import path
from .views import upload_excel, upload_and_generate_indicators, download_file

urlpatterns = [
    path('upload/', upload_excel, name='upload_excel'),  # Subir y procesar archivo
    path('upload_and_generate/', upload_and_generate_indicators, name='upload_and_generate'),  # Procesar archivo y generar indicadores
    path('download/', download_file, name='download_file'),  # URL para la descarga del archivo generado
]