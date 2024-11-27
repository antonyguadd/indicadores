import pandas as pd
from io import BytesIO
import xlsxwriter
from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
from rest_framework.decorators import api_view
import logging
import os
from indicadores.utils.utils_suspendidas import procesar_excel_suspendidas, indicadores_criterio, indicadores_por_zona, generar_excel_suspendidas

logger = logging.getLogger(__name__)

# Directorio temporal para guardar archivos
TEMP_DIR = 'temp_files'
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

@csrf_exempt
@api_view(['POST'])
def upload_excel_suspendidas(request):
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

@csrf_exempt
@api_view(['POST'])
def upload_and_generate_report_suspendidas(request):
    file_path = request.POST.get('file_path')
    
    if not file_path or not os.path.exists(file_path):
        return JsonResponse({'error': 'File not found'}, status=400)
    
    try:
        # Procesar el archivo Excel y generar indicadores
        df = procesar_excel_suspendidas(file_path)
        logger.info(f"Datos procesados: {df.head()}")  # Para verificar el contenido del DataFrame
        
        # Generar los indicadores de criterios
        indicadores = indicadores_criterio(df)
        
        # Generar los indicadores por zona
        indicadores_zona = indicadores_por_zona(df)
        
        # Generar el archivo Excel con los indicadores
        output_file = generar_excel_suspendidas(indicadores, indicadores_zona)

        # Guardar el archivo generado temporalmente en el servidor
        generated_file_path = os.path.join(TEMP_DIR, 'indicadores_suspendidas.xlsx')
        with open(generated_file_path, 'wb') as f:
            f.write(output_file)

        # URL de descarga del archivo
        download_url = request.build_absolute_uri('/api/download_suspendidas_report/')
        return JsonResponse({'message': 'File processed successfully', 'download_url': download_url})

    except ValueError as e:
        logger.error(f"Error processing file: {str(e)}")
        return JsonResponse({"error": f"Error processing file: {str(e)}"}, status=400)
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        return JsonResponse({"error": f"Unexpected error: {str(e)}"}, status=500)

@csrf_exempt
def download_suspendidas_report(request):
    file_path = os.path.join(TEMP_DIR, 'indicadores_suspendidas.xlsx')

    try:
        with open(file_path, 'rb') as f:
            response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = 'attachment; filename="indicadores_suspendidas.xlsx"'
            return response
    except FileNotFoundError:
        return JsonResponse({'error': 'File not found'}, status=404)