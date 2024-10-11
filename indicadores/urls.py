from django.urls import path
from .views import upload_excel, upload_and_generate_indicators, download_file, ordenes_mes, download_ordenes_mes_file, UserCreate

urlpatterns = [
    path('upload/', upload_excel, name='upload_excel'),  # Subir y procesar archivo
    path('upload_and_generate/', upload_and_generate_indicators, name='upload_and_generate'),  # Procesar archivo y generar indicadores
    path('download/', download_file, name='download_file'),  # URL para la descarga del archivo generado
    path('ordenes_mes/', ordenes_mes, name='ordenes_mes'),  # Generar otros indicadores
    path('download_ordenes_mes/', download_ordenes_mes_file, name='download_ordenes_mes_file'),
    
    # URL para la descarga del archivo de otros indicadores
]