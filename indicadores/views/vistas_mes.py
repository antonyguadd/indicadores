from django.shortcuts import render
from rest_framework.response import Response
from rest_framework.decorators import api_view
import logging
from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
import pandas as pd
import os
from ..utils.utils_mes import procesar_excel_mes, indicadores_por_mes_y_servicio, generar_excel_mes

logger = logging.getLogger(__name__)

# Directorio temporal para guardar archivos
TEMP_DIR = 'temp_files'
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

# Vista para subir un archivo y almacenarlo temporalmente
@csrf_exempt
@api_view(['POST'])
def upload_excel_mes(request):
    if 'file' not in request.FILES:
        return JsonResponse({'error': 'No file provided'}, status=400)
    
    file = request.FILES['file']
    file_path = os.path.join(TEMP_DIR, file.name)
    
    try:
        # Guardar el archivo en el directorio temporal
        with open(file_path, 'wb') as f:
            for chunk in file.chunks():
                f.write(chunk)
        
        # Contar las líneas del archivo Excel
        df = pd.read_excel(file_path)
        line_count = len(df)
        
        return JsonResponse({'message': 'File uploaded successfully', 'file_path': file_path, 'file_name': file.name, 'line_count': line_count})
    except Exception as e:
        logger.error(f"Error uploading file: {str(e)}")
        return JsonResponse({'error': f"Error uploading file: {str(e)}"}, status=500)

# Vista para procesar el archivo almacenado y generar indicadores por mes
@csrf_exempt
@api_view(['POST'])
def upload_and_generate_monthly_indicators(request):
    file_path = request.POST.get('file_path')
    
    if not file_path or not os.path.exists(file_path):
        return JsonResponse({'error': 'File not found'}, status=400)
    
    try:
        # Procesar el archivo Excel y generar indicadores
        df = procesar_excel_mes(file_path)
        logger.info(f"Datos procesados: {df.head()}")  # Para verificar el contenido del DataFrame
        
        # Generar los indicadores por mes y servicio
        indicadores = indicadores_por_mes_y_servicio(df)
        
        # Generar el archivo Excel con los indicadores
        output_file = generar_excel_mes(indicadores)

        # Guardar el archivo generado temporalmente en el servidor
        generated_file_path = os.path.join(TEMP_DIR, 'indicadores_por_mes.xlsx')
        with open(generated_file_path, 'wb') as f:
            f.write(output_file)

        # URL de descarga del archivo
        download_url = request.build_absolute_uri('/api/download_mes/')
        return JsonResponse({'message': 'File processed successfully', 'download_url': download_url})

    except ValueError as e:
        logger.error(f"Error processing file: {str(e)}")
        return JsonResponse({"error": f"Error processing file: {str(e)}"}, status=400)
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        return JsonResponse({"error": f"Unexpected error: {str(e)}"}, status=500)

# Vista para descargar el archivo Excel generado con indicadores mensuales
@csrf_exempt
def download_file_mes(request):
    file_path = os.path.join(TEMP_DIR, 'indicadores_por_mes.xlsx')

    try:
        with open(file_path, 'rb') as f:
            response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = 'attachment; filename="indicadores_por_mes.xlsx"'
            return response
    except FileNotFoundError:
        return JsonResponse({'error': 'File not found'}, status=404)
